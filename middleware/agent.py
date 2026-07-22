"""
agent.py — Genuine Claude tool-use agent loop.
Handles the "Ask Library" feature: natural language Q&A grounded
in the document library stored in SQLite.

Tools Claude can call autonomously:
  1. search_library(query, format, date_range)
  2. read_document(file_id)
  3. list_recent(limit, username)
  4. get_download_link(file_id)

The agent loop:
  1. Send user question + tool definitions to Claude
  2. Claude returns tool_use block
  3. Execute the tool (SQLite / file read)
  4. Send tool_result back to Claude
  5. Claude calls another tool OR returns final answer
  6. Return final answer + any download links to caller
"""

import json
from typing import Any, Dict, List, Optional

import anthropic

from db import search_documents, get_document, list_documents, get_content_summary

MODEL    = "claude-sonnet-4-20250514"
MAX_ITER = 6   # max tool-call rounds before forcing a final answer

# ── Tool definitions sent to Claude ───────────────────────────────────────

TOOLS: List[Dict] = [
    {
        "name": "search_library",
        "description": (
            "Search the document library by natural language query. "
            "Returns a list of matching documents with id, file_name, template, "
            "output_format, size_kb, username, created_at, and a snippet of content_summary. "
            "Use this when the user asks to find documents by topic, project name, content, or keywords."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "query": {
                    "type": "string",
                    "description": "Keywords or phrase to search for across document names, templates, and content.",
                },
                "format": {
                    "type": "string",
                    "description": "Optional filter by output format: xlsx, docx, pdf, pptx, xml",
                    "enum": ["xlsx", "docx", "pdf", "pptx", "xml"],
                },
                "date_from": {
                    "type": "string",
                    "description": "Optional ISO date string (YYYY-MM-DD) to filter documents created on or after this date.",
                },
                "date_to": {
                    "type": "string",
                    "description": "Optional ISO date string (YYYY-MM-DD) to filter documents created on or before this date.",
                },
            },
            "required": ["query"],
        },
    },
    {
        "name": "read_document",
        "description": (
            "Read the stored content of a specific document by its file_id. "
            "Returns the full content_summary which contains the document's actual data — "
            "use this to answer questions about a specific document's contents, "
            "summarise it, or extract specific fields."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "file_id": {
                    "type": "string",
                    "description": "The document ID (uuid) returned by search_library or list_recent.",
                },
            },
            "required": ["file_id"],
        },
    },
    {
        "name": "list_recent",
        "description": (
            "List the most recently generated documents in the library. "
            "Returns id, file_name, template, output_format, size_kb, username, created_at. "
            "Use this when the user asks about recent documents, latest files, or when no specific query is given."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "limit": {
                    "type": "integer",
                    "description": "Number of documents to return (default 10, max 50).",
                },
                "username": {
                    "type": "string",
                    "description": "Optional: filter by a specific user.",
                },
                "format": {
                    "type": "string",
                    "description": "Optional: filter by output format.",
                    "enum": ["xlsx", "docx", "pdf", "pptx", "xml"],
                },
            },
            "required": [],
        },
    },
    {
        "name": "get_download_link",
        "description": (
            "Get the download URL for a specific document by its file_id. "
            "Always call this when the user wants to download a document "
            "or when you want to include a download link in your answer."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "file_id": {
                    "type": "string",
                    "description": "The document ID (uuid).",
                },
                "file_name": {
                    "type": "string",
                    "description": "The file name (used to label the download link).",
                },
            },
            "required": ["file_id"],
        },
    },
]

# ── Tool execution ─────────────────────────────────────────────────────────

async def _execute_tool(tool_name: str, tool_input: Dict, base_url: str) -> Any:
    """Execute a tool call and return the result as a string."""

    if tool_name == "search_library":
        results = await search_documents(
            query       = tool_input.get("query", ""),
            output_format = tool_input.get("format"),
            date_from   = tool_input.get("date_from"),
            date_to     = tool_input.get("date_to"),
            limit       = 15,
        )
        if not results:
            return "No documents found matching that query."
        # Return compact representation — enough for Claude to reason with
        out = []
        for d in results:
            out.append({
                "file_id":       d["id"],
                "file_name":     d["file_name"],
                "template":      d["template"],
                "format":        d["output_format"],
                "size_kb":       d["size_kb"],
                "username":      d["username"],
                "created_at":    d["created_at"],
                "content_snippet": (d.get("content_summary") or "")[:400],
            })
        return json.dumps(out, indent=2)

    if tool_name == "read_document":
        file_id = tool_input.get("file_id", "")
        doc     = await get_document(file_id)
        if not doc:
            return f"Document with id '{file_id}' not found."
        summary = doc.get("content_summary") or ""
        return (
            f"File: {doc['file_name']}\n"
            f"Template: {doc['template']}\n"
            f"Format: {doc['output_format']}\n"
            f"Created: {doc['created_at']}\n"
            f"User: {doc['username']}\n\n"
            f"Content:\n{summary[:8000]}"
        )

    if tool_name == "list_recent":
        limit    = min(int(tool_input.get("limit", 10)), 50)
        username = tool_input.get("username")
        fmt      = tool_input.get("format")
        docs     = await list_documents(username=username, output_format=fmt, limit=limit)
        if not docs:
            return "No documents found in the library."
        out = []
        for d in docs:
            out.append({
                "file_id":    d["id"],
                "file_name":  d["file_name"],
                "template":   d["template"],
                "format":     d["output_format"],
                "size_kb":    d["size_kb"],
                "username":   d["username"],
                "created_at": d["created_at"],
            })
        return json.dumps(out, indent=2)

    if tool_name == "get_download_link":
        file_id   = tool_input.get("file_id", "")
        file_name = tool_input.get("file_name", file_id)
        url       = f"{base_url}/download/{file_id}?name={file_name}"
        return json.dumps({"download_url": url, "file_name": file_name})

    return f"Unknown tool: {tool_name}"


# ── System prompt ──────────────────────────────────────────────────────────

SYSTEM_PROMPT = """You are ProjectZen's document library assistant.
You help users find, summarise, and download documents stored in the library.

Rules:
- ONLY answer based on what the tools return. Never invent document names, IDs, or content.
- Always include download links when returning documents to the user. Use get_download_link for each document.
- Format download links clearly as: [filename](download_url)
- When listing multiple documents, present them as a numbered list with download links.
- If no documents are found, say so clearly and suggest the user generate one.
- Keep answers concise and structured.
- If the user asks about document content, use read_document to get the actual content first.
- Dates are in UTC ISO format. When users say "today", "this week", "yesterday" — interpret relative to current context."""


# ── Main agent loop ────────────────────────────────────────────────────────

async def ask_library(
    question: str,
    username: Optional[str] = None,
    base_url: str = "http://localhost:8000",
) -> Dict[str, Any]:
    """
    Run the agent loop for a user question.
    Returns: { answer: str, documents: [...], tool_calls_made: int }
    """
    client = anthropic.AsyncAnthropic()

    # Build initial message
    user_content = question
    if username:
        user_content += f"\n\n[Current user: {username}]"

    messages: List[Dict] = [
        {"role": "user", "content": user_content}
    ]

    documents_surfaced: List[Dict] = []   # collect docs Claude found
    tool_calls_made = 0
    final_answer    = ""

    print(f"\n🤖 Agent loop started | question: {question[:80]}...")

    for iteration in range(MAX_ITER):
        print(f"   → Iteration {iteration + 1}/{MAX_ITER}")

        response = await client.messages.create(
            model      = MODEL,
            max_tokens = 2048,
            system     = SYSTEM_PROMPT,
            tools      = TOOLS,
            messages   = messages,
        )

        # Append assistant response to messages
        messages.append({"role": "assistant", "content": response.content})

        # Check stop reason
        if response.stop_reason == "end_turn":
            # Claude is done — extract text answer
            for block in response.content:
                if hasattr(block, "text"):
                    final_answer = block.text
            print(f"   → Agent finished after {tool_calls_made} tool calls")
            break

        if response.stop_reason != "tool_use":
            # Unexpected stop
            for block in response.content:
                if hasattr(block, "text"):
                    final_answer = block.text
            break

        # Process tool calls
        tool_results = []
        for block in response.content:
            if block.type != "tool_use":
                continue

            tool_calls_made += 1
            tool_name  = block.name
            tool_input = block.input
            tool_id    = block.id

            print(f"   → Tool call: {tool_name}({json.dumps(tool_input)[:120]})")

            result = await _execute_tool(tool_name, tool_input, base_url)

            # Track documents returned by get_download_link
            if tool_name == "get_download_link":
                try:
                    parsed = json.loads(result)
                    documents_surfaced.append({
                        "file_id":      tool_input.get("file_id"),
                        "file_name":    parsed.get("file_name"),
                        "download_url": parsed.get("download_url"),
                    })
                except Exception:
                    pass

            tool_results.append({
                "type":        "tool_result",
                "tool_use_id": tool_id,
                "content":     str(result),
            })

        # Append tool results as user message
        messages.append({"role": "user", "content": tool_results})

    else:
        # Hit MAX_ITER — force a final answer from whatever we have
        final_answer = (
            "I searched the library but could not complete the full analysis. "
            "Please try a more specific question."
        )

    return {
        "answer":          final_answer,
        "documents":       documents_surfaced,
        "tool_calls_made": tool_calls_made,
    }
