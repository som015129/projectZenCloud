"""
context_store.py — LangGraph State Definition + In-Memory SSE Queue Store

CascadeState is the typed state object that flows through the LangGraph
workflow nodes. Every field is explicitly typed for correctness.

The SSE queue store is separate from LangGraph state because asyncio.Queue
objects cannot be serialized by the LangGraph checkpointer.

Usage:
  from context_store import CascadeState, create_sse_queue, emit_event
"""

import asyncio
from typing import Any, Dict, List, Optional, TypedDict


# ═══════════════════════════════════════════════════════════════════════════════
# SUPPORTING TYPES
# ═══════════════════════════════════════════════════════════════════════════════

class GeneratedDoc(TypedDict):
    """Record of one generated document within a cascade session."""
    node_id:        str
    file_id:        str
    file_name:      str
    output_format:  str
    content_summary:str
    version:        int
    model_used:     str
    is_finalised:   bool
    finalised_at:   Optional[str]
    finalised_by:   Optional[str]


class DeltaItem(TypedDict):
    """One atomic change identified by the Delta Analysis Agent."""
    type:        str   # "added" | "modified" | "removed"
    section:     str   # Document section or field name
    description: str   # Human-readable description of the change


class AffectedNode(TypedDict):
    """A downstream document impacted by a delta change."""
    node_id:         str
    label:           str
    phase:           str
    tier:            str
    current_version: int
    next_version:    int
    is_finalised:    bool
    impact_type:     str   # Mandatory | Conditional | Review
    impact_reason:   str   # Why this node is affected


class ProjectContext(TypedDict):
    """Structured project facts extracted by Agent 2 from the input document."""
    project_name:          str
    client_name:           str
    delivery_partner:      str
    project_type:          str   # e.g. "Salesforce Sales Cloud Implementation"
    scope_summary:         str
    in_scope_modules:      List[str]
    out_of_scope:          List[str]
    key_stakeholders:      List[Dict[str, str]]  # [{name, role, org}]
    timeline:              Dict[str, str]         # {start, end, go_live, key_milestones}
    integration_points:    List[str]
    data_migration_scope:  str
    key_requirements:      List[str]
    geographic_scope:      str
    budget_reference:      str
    delivery_methodology:  str
    additional_context:    str


# ═══════════════════════════════════════════════════════════════════════════════
# LANGGRAPH MAIN STATE
# ═══════════════════════════════════════════════════════════════════════════════

class CascadeState(TypedDict):
    """
    The typed state object passed between all LangGraph nodes.
    Every field is read/written by one or more nodes in the workflow.
    """

    # ── Session identity ───────────────────────────────────────────────────
    session_id:     str
    mode:           str      # "new" | "delta"
    username:       str

    # ── Input document (mode="new") ────────────────────────────────────────
    input_text:      str     # Extracted text from the source document
    input_file_name: str     # Original filename (e.g. "SOW_ClientX.docx")

    # ── Delta mode inputs (mode="delta") ──────────────────────────────────
    delta_node_id:      Optional[str]   # Which document was changed
    delta_from_version: Optional[int]   # Version to compare against
    delta_new_text:     Optional[str]   # New document text or change description
    delta_input_type:   Optional[str]   # "full_doc" | "delta_section" | "prompt"

    # ── Agent 2 output: extracted project context ──────────────────────────
    project_context: Dict[str, Any]   # ProjectContext dict

    # ── Graph traversal (mode="new") ─────────────────────────────────────
    selected_nodes:   List[str]         # Node IDs user selected to generate
    wave_order:       List[List[str]]   # BFS waves: [[wave0_nodes], [wave1_nodes], ...]
    current_wave_idx: int               # Which wave is currently being processed

    # ── Generated documents: node_id -> GeneratedDoc ───────────────────────
    generated_docs: Dict[str, GeneratedDoc]

    # ── Agent 4 output: delta analysis (mode="delta") ─────────────────────
    delta_items:      Optional[List[DeltaItem]]
    affected_nodes:   Optional[List[AffectedNode]]

    # ── Human-in-the-loop: user scope selection (mode="delta") ────────────
    user_delta_nodes:       Optional[List[str]]         # node_ids user confirmed for update
    user_finalised_choices: Optional[Dict[str, bool]]   # node_id -> True = include finalised

    # ── Delta wave tracking ───────────────────────────────────────────────
    delta_waves:          Optional[List[List[str]]]  # BFS sub-waves for delta propagation
    current_delta_wave:   int

    # ── Progress counters ─────────────────────────────────────────────────
    total_docs:     int
    completed_docs: int

    # ── Status ────────────────────────────────────────────────────────────
    error:       Optional[str]
    is_complete: bool


# ═══════════════════════════════════════════════════════════════════════════════
# DEFAULT STATE FACTORY
# ═══════════════════════════════════════════════════════════════════════════════

def make_initial_state(
    session_id: str,
    mode: str,
    username: str,
    input_text: str,
    input_file_name: str,
    selected_nodes: List[str],
    delta_node_id: Optional[str] = None,
    delta_from_version: Optional[int] = None,
    delta_new_text: Optional[str] = None,
    delta_input_type: Optional[str] = None,
) -> CascadeState:
    return CascadeState(
        session_id=session_id,
        mode=mode,
        username=username,
        input_text=input_text,
        input_file_name=input_file_name,
        delta_node_id=delta_node_id,
        delta_from_version=delta_from_version,
        delta_new_text=delta_new_text,
        delta_input_type=delta_input_type,
        project_context={},
        selected_nodes=selected_nodes,
        wave_order=[],
        current_wave_idx=0,
        generated_docs={},
        delta_items=None,
        affected_nodes=None,
        user_delta_nodes=None,
        user_finalised_choices=None,
        delta_waves=None,
        current_delta_wave=0,
        total_docs=len(selected_nodes),
        completed_docs=0,
        error=None,
        is_complete=False,
    )


# ═══════════════════════════════════════════════════════════════════════════════
# SSE QUEUE STORE  (separate from LangGraph state — not serializable)
# ═══════════════════════════════════════════════════════════════════════════════

_sse_queues: Dict[str, asyncio.Queue] = {}


def create_sse_queue(session_id: str) -> asyncio.Queue:
    """Create and register an SSE queue for a cascade session."""
    q: asyncio.Queue = asyncio.Queue(maxsize=1000)
    _sse_queues[session_id] = q
    return q


def get_sse_queue(session_id: str) -> Optional[asyncio.Queue]:
    return _sse_queues.get(session_id)


async def emit_event(session_id: str, event_type: str, data: Dict[str, Any]) -> None:
    """
    Push an SSE event dict into the session queue.
    The SSE route drains this queue and streams to the browser.
    """
    q = _sse_queues.get(session_id)
    if q:
        try:
            await asyncio.wait_for(
                q.put({"event": event_type, "data": data}),
                timeout=2.0,
            )
        except (asyncio.TimeoutError, asyncio.QueueFull):
            pass  # Non-blocking — never block the agent on a slow client


def signal_sse_done(session_id: str) -> None:
    """
    Push a None sentinel to signal stream end.
    The SSE route exits its generator loop on receiving None.
    """
    q = _sse_queues.get(session_id)
    if q:
        try:
            q.put_nowait(None)
        except asyncio.QueueFull:
            pass


def cleanup_sse_queue(session_id: str) -> None:
    """Remove the queue after the SSE stream closes."""
    _sse_queues.pop(session_id, None)
