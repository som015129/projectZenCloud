"""
knowledge_graph.py — SF Document Dependency Knowledge Graph
Python representation of the 32-document Salesforce project delivery hierarchy.

Provides:
  - Complete node/edge definitions for all 32 SF project documents
  - BFS wave ordering (Kahn's algorithm) for cascade generation
  - Parent / child / downstream traversal
  - Default output format per document type (admin-configurable via DB)
  - Document tier classification: T1=critical → T4=low blast radius
  - Model selection by tier (Opus for T1/T2, Sonnet for T3/T4)
  - RTM change-wave definitions
  - Change impact metadata per document
"""

from typing import Dict, List, Optional


# ═══════════════════════════════════════════════════════════════════════════════
# 1. NODE DEFINITIONS  (32 documents across 5 phases)
# ═══════════════════════════════════════════════════════════════════════════════

NODES: List[Dict[str, str]] = [
    # ── Initiate ────────────────────────────────────────────────────────────
    {
        "id": "sow", "label": "SOW / SP51", "phase": "Initiate",
        "type": "Source / Contract", "owner": "Account / Sales Lead",
        "trigger": "Contract signature", "notes": "", "tier": "T1",
    },
    {
        "id": "project-plan", "label": "Project Plan", "phase": "Initiate",
        "type": "Schedule / Master", "owner": "Project / Delivery Manager",
        "trigger": "SOW signed; team mobilised", "notes": "Cross-phase document", "tier": "T1",
    },
    {
        "id": "comm-plan", "label": "Communication Plan", "phase": "Initiate",
        "type": "Plan", "owner": "Project / Delivery Manager",
        "trigger": "Project Plan baselined", "notes": "", "tier": "T3",
    },
    {
        "id": "gov-matrix", "label": "Governance Matrix", "phase": "Initiate",
        "type": "Plan / RACI", "owner": "Project / Delivery Manager",
        "trigger": "Kickoff", "notes": "Informs every phase", "tier": "T3",
    },
    {
        "id": "status-templates", "label": "Status Report Templates", "phase": "Initiate",
        "type": "Template", "owner": "PMO",
        "trigger": "Governance setup", "notes": "", "tier": "T4",
    },
    {
        "id": "risk-log", "label": "Risk / Issue Log", "phase": "Initiate",
        "type": "Living register", "owner": "Project / Delivery Manager",
        "trigger": "Kickoff (initial population)", "notes": "", "tier": "T4",
    },
    {
        "id": "onboarding-deck", "label": "Onboarding Deck", "phase": "Initiate",
        "type": "Template", "owner": "Project / Delivery Manager",
        "trigger": "Team mobilisation", "notes": "Used per new joiner", "tier": "T4",
    },
    {
        "id": "milestone-format", "label": "Milestone Sign-off Format", "phase": "Initiate",
        "type": "Template", "owner": "PMO",
        "trigger": "Governance setup", "notes": "Phase gates throughout project", "tier": "T4",
    },
    # ── Confirm ─────────────────────────────────────────────────────────────
    {
        "id": "workshop-notes", "label": "Workshop Notes", "phase": "Confirm",
        "type": "Raw input", "owner": "Functional Consultants",
        "trigger": "Workshop session", "notes": "", "tier": "T2",
    },
    {
        "id": "rtm", "label": "RTM", "phase": "Confirm",
        "type": "Traceability", "owner": "Functional Lead",
        "trigger": "Workshop closure", "notes": "Requirement Traceability Matrix", "tier": "T1",
    },
    {
        "id": "to-be", "label": "To-Be Process", "phase": "Confirm",
        "type": "Process design", "owner": "Functional Lead",
        "trigger": "Design workshop", "notes": "", "tier": "T2",
    },
    {
        "id": "fit-gap", "label": "Fit-Gap Analysis", "phase": "Confirm",
        "type": "Analysis", "owner": "Solution Architect",
        "trigger": "Requirements baselined", "notes": "Mutual relationship with RTM", "tier": "T2",
    },
    {
        "id": "solution-design", "label": "Solution Design", "phase": "Confirm",
        "type": "Design / Bridge", "owner": "Solution Architect",
        "trigger": "Fit-Gap signed off", "notes": "Hub document", "tier": "T1",
    },
    {
        "id": "interface-list", "label": "Interface List", "phase": "Confirm",
        "type": "Inventory", "owner": "Integration Lead",
        "trigger": "Solution Design draft", "notes": "Mutual with Integration Strategy", "tier": "T3",
    },
    {
        "id": "int-strategy", "label": "Integration Strategy", "phase": "Confirm",
        "type": "Strategy", "owner": "Integration Architect",
        "trigger": "Solution Design draft", "notes": "", "tier": "T3",
    },
    {
        "id": "dm-strategy", "label": "Data Migration Strategy", "phase": "Confirm",
        "type": "Strategy", "owner": "Data Migration Lead",
        "trigger": "Solution Design draft", "notes": "", "tier": "T3",
    },
    {
        "id": "func-spec", "label": "Functional Spec", "phase": "Confirm",
        "type": "Spec", "owner": "Functional Consultants",
        "trigger": "Solution Design baselined", "notes": "", "tier": "T2",
    },
    # ── Design-Build ────────────────────────────────────────────────────────
    {
        "id": "config-workbook", "label": "Configuration Workbook", "phase": "Design-Build",
        "type": "Build artefact", "owner": "Functional Consultants",
        "trigger": "Functional Spec signed off", "notes": "", "tier": "T3",
    },
    {
        "id": "config-specs", "label": "Configuration Specs", "phase": "Design-Build",
        "type": "Build artefact", "owner": "Functional Consultants",
        "trigger": "Workbook update", "notes": "", "tier": "T3",
    },
    {
        "id": "int-td", "label": "Integration TD", "phase": "Design-Build",
        "type": "Build artefact", "owner": "Integration Developer",
        "trigger": "Interface signed off", "notes": "Technical Design", "tier": "T3",
    },
    {
        "id": "unit-test-cases", "label": "Unit Test Cases", "phase": "Design-Build",
        "type": "Test artefact", "owner": "Functional Consultants",
        "trigger": "Build complete", "notes": "", "tier": "T4",
    },
    # ── Integrate ────────────────────────────────────────────────────────────
    {
        "id": "test-strategy", "label": "Test Strategy", "phase": "Integrate",
        "type": "Strategy", "owner": "Test Lead",
        "trigger": "Build phase complete", "notes": "", "tier": "T2",
    },
    {
        "id": "test-cases", "label": "Test Cases", "phase": "Integrate",
        "type": "Test artefact", "owner": "Test Lead",
        "trigger": "Test Strategy approved", "notes": "", "tier": "T3",
    },
    {
        "id": "int-test-cases", "label": "Integration Test Cases", "phase": "Integrate",
        "type": "Test artefact", "owner": "Integration Lead",
        "trigger": "Interfaces deployed to SIT", "notes": "", "tier": "T4",
    },
    {
        "id": "uat-test-cases", "label": "UAT Test Cases", "phase": "Integrate",
        "type": "Test artefact", "owner": "Business Lead",
        "trigger": "SIT exit", "notes": "", "tier": "T3",
    },
    {
        "id": "defect-log", "label": "Defect Log", "phase": "Integrate",
        "type": "Living register", "owner": "Test Lead",
        "trigger": "Test execution starts", "notes": "", "tier": "T4",
    },
    {
        "id": "uat-signoff", "label": "UAT Sign-off", "phase": "Integrate",
        "type": "Gate document", "owner": "Business Sponsor",
        "trigger": "UAT exit criteria met", "notes": "", "tier": "T3",
    },
    # ── Deploy ───────────────────────────────────────────────────────────────
    {
        "id": "cutover-plan", "label": "Cutover Plan", "phase": "Deploy",
        "type": "Runbook", "owner": "Project / Cutover Manager",
        "trigger": "UAT Sign-off", "notes": "", "tier": "T2",
    },
    {
        "id": "golive-checklist", "label": "Go-Live Checklist", "phase": "Deploy",
        "type": "Gate document", "owner": "Project / Delivery Manager",
        "trigger": "Day before go-live", "notes": "", "tier": "T3",
    },
    {
        "id": "kt-plan", "label": "KT Plan + Train-the-Trainer", "phase": "Deploy",
        "type": "Plan", "owner": "Functional Lead",
        "trigger": "Build complete (drafted earlier)", "notes": "", "tier": "T3",
    },
    {
        "id": "hypercare-tracker", "label": "Hypercare Tracker", "phase": "Deploy",
        "type": "Living register", "owner": "Hypercare Lead",
        "trigger": "Go-Live cutover complete", "notes": "", "tier": "T4",
    },
]

# ═══════════════════════════════════════════════════════════════════════════════
# 2. EDGE DEFINITIONS  (source → target = upstream feeds downstream)
# ═══════════════════════════════════════════════════════════════════════════════

EDGES: List[Dict[str, str]] = [
    # SOW outgoing
    {"source": "sow",             "target": "project-plan"},
    {"source": "sow",             "target": "comm-plan"},
    {"source": "sow",             "target": "gov-matrix"},
    {"source": "sow",             "target": "risk-log"},
    {"source": "sow",             "target": "onboarding-deck"},
    {"source": "sow",             "target": "workshop-notes"},
    {"source": "sow",             "target": "rtm"},
    {"source": "sow",             "target": "config-workbook"},   # SOW → Configuration Workbook
    # Project Plan outgoing
    {"source": "project-plan",    "target": "comm-plan"},
    {"source": "project-plan",    "target": "gov-matrix"},
    {"source": "project-plan",    "target": "status-templates"},
    {"source": "project-plan",    "target": "onboarding-deck"},
    {"source": "project-plan",    "target": "milestone-format"},
    {"source": "project-plan",    "target": "workshop-notes"},
    {"source": "project-plan",    "target": "test-strategy"},
    {"source": "project-plan",    "target": "cutover-plan"},
    # Communication Plan outgoing
    {"source": "comm-plan",       "target": "status-templates"},
    # Governance Matrix outgoing
    {"source": "gov-matrix",      "target": "status-templates"},
    {"source": "gov-matrix",      "target": "onboarding-deck"},
    {"source": "gov-matrix",      "target": "milestone-format"},
    # Risk Log outgoing
    {"source": "risk-log",        "target": "cutover-plan"},
    # Workshop Notes outgoing
    {"source": "workshop-notes",  "target": "rtm"},
    {"source": "workshop-notes",  "target": "to-be"},
    {"source": "workshop-notes",  "target": "fit-gap"},
    {"source": "workshop-notes",  "target": "solution-design"},
    {"source": "workshop-notes",  "target": "func-spec"},
    # RTM <-> Fit-Gap mutual: keep both edges, break cycle in wave computation
    {"source": "fit-gap",         "target": "rtm"},
    {"source": "rtm",             "target": "fit-gap"},
    # Fit-Gap outgoing
    {"source": "fit-gap",         "target": "to-be"},
    {"source": "fit-gap",         "target": "solution-design"},
    # RTM outgoing
    {"source": "rtm",             "target": "solution-design"},
    {"source": "rtm",             "target": "func-spec"},
    {"source": "rtm",             "target": "config-specs"},
    {"source": "rtm",             "target": "unit-test-cases"},
    {"source": "rtm",             "target": "test-strategy"},
    {"source": "rtm",             "target": "test-cases"},
    {"source": "rtm",             "target": "uat-test-cases"},
    # To-Be outgoing
    {"source": "to-be",           "target": "solution-design"},
    {"source": "to-be",           "target": "uat-test-cases"},
    {"source": "to-be",           "target": "kt-plan"},
    # Solution Design outgoing
    {"source": "solution-design", "target": "func-spec"},
    {"source": "solution-design", "target": "interface-list"},
    {"source": "solution-design", "target": "int-strategy"},
    {"source": "solution-design", "target": "dm-strategy"},
    {"source": "solution-design", "target": "config-workbook"},
    {"source": "solution-design", "target": "int-td"},
    {"source": "solution-design", "target": "test-strategy"},
    {"source": "solution-design", "target": "test-cases"},
    {"source": "solution-design", "target": "cutover-plan"},
    {"source": "solution-design", "target": "kt-plan"},
    # Interface List <-> Integration Strategy mutual
    {"source": "interface-list",  "target": "int-strategy"},
    {"source": "int-strategy",    "target": "interface-list"},
    # Interface List outgoing
    {"source": "interface-list",  "target": "int-td"},
    {"source": "interface-list",  "target": "int-test-cases"},
    # Integration Strategy outgoing
    {"source": "int-strategy",    "target": "int-td"},
    {"source": "int-strategy",    "target": "cutover-plan"},
    # Functional Spec outgoing
    {"source": "func-spec",       "target": "config-workbook"},
    {"source": "func-spec",       "target": "unit-test-cases"},
    {"source": "func-spec",       "target": "test-cases"},
    # Configuration Workbook outgoing
    {"source": "config-workbook", "target": "config-specs"},
    {"source": "config-workbook", "target": "unit-test-cases"},
    {"source": "config-workbook", "target": "kt-plan"},
    # Configuration Specs outgoing
    {"source": "config-specs",    "target": "unit-test-cases"},
    # Integration TD outgoing
    {"source": "int-td",          "target": "int-test-cases"},
    # Test Strategy outgoing
    {"source": "test-strategy",   "target": "test-cases"},
    {"source": "test-strategy",   "target": "int-test-cases"},
    {"source": "test-strategy",   "target": "uat-test-cases"},
    # Test Cases outgoing
    {"source": "test-cases",      "target": "int-test-cases"},
    {"source": "test-cases",      "target": "uat-test-cases"},
    {"source": "test-cases",      "target": "defect-log"},
    # Integration Test Cases outgoing
    {"source": "int-test-cases",  "target": "defect-log"},
    # UAT Test Cases outgoing
    {"source": "uat-test-cases",  "target": "defect-log"},
    {"source": "uat-test-cases",  "target": "uat-signoff"},
    # Defect Log outgoing
    {"source": "defect-log",      "target": "uat-signoff"},
    {"source": "defect-log",      "target": "hypercare-tracker"},
    # UAT Sign-off outgoing
    {"source": "uat-signoff",     "target": "cutover-plan"},
    {"source": "uat-signoff",     "target": "golive-checklist"},
    # Cutover Plan outgoing
    {"source": "cutover-plan",    "target": "golive-checklist"},
    # Go-Live Checklist outgoing
    {"source": "golive-checklist","target": "hypercare-tracker"},
]

# ═══════════════════════════════════════════════════════════════════════════════
# 3. CHANGE IMPACT METADATA  (what to update when each document changes)
# ═══════════════════════════════════════════════════════════════════════════════

CHANGE_IMPACT: Dict[str, Dict[str, str]] = {
    "sow":              {"impactType": "Mandatory",   "governance": "CR (CCB-level) + Steerco approval + revised SOW signed",        "documentsToUpdate": "Project Plan -> Communication Plan -> Governance Matrix -> Risk Log -> RTM -> cascade through Solution Design and downstream", "notes": "Highest blast radius. SOW change is rare and contractual. Treat as full project re-baseline."},
    "project-plan":     {"impactType": "Mandatory",   "governance": "CR + Steerco approval if dates slip",                          "documentsToUpdate": "Communication Plan -> Status Report Templates -> Workshop Notes -> Test Strategy -> Cutover Plan -> KT Plan -> Milestone Sign-off Format", "notes": "Cross-phase document. Effort or date changes ripple to every dependent timeline."},
    "comm-plan":        {"impactType": "Conditional", "governance": "PMO update; Steerco notice if material",                        "documentsToUpdate": "Status Report Templates -> ad-hoc stakeholder notice if audience changes", "notes": "Low blast radius. Self-contained unless cadence changes."},
    "gov-matrix":       {"impactType": "Conditional", "governance": "Steerco notice; signed acknowledgment from new approvers",      "documentsToUpdate": "Status Report Templates -> Milestone Sign-off Format -> all gate documents re-route to new approvers", "notes": "Low blast radius unless approvers change. Then medium impact."},
    "status-templates": {"impactType": "Conditional", "governance": "PMO update",                                                   "documentsToUpdate": "Future weekly + steerco reports adopt new format", "notes": "No downstream artefact impact. Just template version control."},
    "risk-log":         {"impactType": "Review",      "governance": "Steerco visibility per cadence",                               "documentsToUpdate": "Steerco Report -> Cutover Plan (residual risks list) -> Hypercare Tracker", "notes": "Living document - always changing. Materially-changed entries go to steerco."},
    "onboarding-deck":  {"impactType": "Review",      "governance": "None required",                                                "documentsToUpdate": "No downstream artefact", "notes": "Self-contained. Update for accuracy as project evolves."},
    "milestone-format": {"impactType": "Review",      "governance": "PMO update",                                                   "documentsToUpdate": "Future sign-offs adopt new format", "notes": "No downstream artefact impact."},
    "workshop-notes":   {"impactType": "Mandatory",   "governance": "Internal review by Functional Lead",                           "documentsToUpdate": "RTM -> To-Be Process -> Fit-Gap -> Functional Spec", "notes": "Pre-RTM artefact. Mid-project workshops typically tied to a CR or clarification."},
    "rtm":              {"impactType": "Mandatory",   "governance": "CR raised; Steerco approval mandatory before rework",           "documentsToUpdate": "Fit-Gap -> Solution Design -> Functional Spec -> Config Workbook -> Config Specs -> Integration TD -> Unit Test Cases -> Test Cases -> Integration Test Cases -> UAT Test Cases", "notes": "Largest predictable blast radius. Most common change scenario. See RTM cascade waves."},
    "to-be":            {"impactType": "Mandatory",   "governance": "CR + Steerco for material changes",                            "documentsToUpdate": "Solution Design -> UAT Test Cases -> KT Plan", "notes": "Process-shape changes are medium-to-high impact, especially close to UAT."},
    "fit-gap":          {"impactType": "Mandatory",   "governance": "CR + Steerco if effort delta material",                        "documentsToUpdate": "Solution Design -> Risk Log -> Project Plan (effort delta if standard->custom flips)", "notes": "Reclassification triggers downstream design rework."},
    "solution-design":  {"impactType": "Mandatory",   "governance": "CR mandatory; Steerco if scope/effort impact",                 "documentsToUpdate": "Functional Spec -> Interface List -> Configuration Workbook -> Integration TD -> Test Cases -> Cutover Plan -> KT Plan", "notes": "Hub document. Changes here have the widest design-to-build cascade. Update with care."},
    "interface-list":   {"impactType": "Mandatory",   "governance": "CR if scope changes; internal review otherwise",               "documentsToUpdate": "Integration TD -> Integration Test Cases -> Cutover Plan (if cutover-relevant interface)", "notes": "Adding/removing interfaces is high impact. Field-level changes are medium."},
    "int-strategy":     {"impactType": "Conditional", "governance": "CR + Architecture review",                                     "documentsToUpdate": "Integration TD -> Cutover Plan (deployment approach)", "notes": "Strategy changes are rare mid-project. High impact when they happen."},
    "dm-strategy":      {"impactType": "Conditional", "governance": "CR + DM review",                                               "documentsToUpdate": "Cutover Plan (load steps, mock cycles) -> Integration TD (if interface-driven)", "notes": "Mock-load failures often trigger this."},
    "func-spec":        {"impactType": "Mandatory",   "governance": "Internal review; CR if scope-bearing",                         "documentsToUpdate": "Configuration Workbook -> Configuration Specs -> Unit Test Cases -> KT material", "notes": "Common change scenario during build iterations."},
    "config-workbook":  {"impactType": "Mandatory",   "governance": "Internal review; CR if scope-bearing",                         "documentsToUpdate": "Configuration Specs -> Unit Test Cases -> KT material", "notes": "Iterative throughout build."},
    "config-specs":     {"impactType": "Mandatory",   "governance": "Internal review",                                              "documentsToUpdate": "Unit Test Cases -> Test Cases (if config visible at SIT)", "notes": "Frequent change during build. Low-medium blast radius."},
    "int-td":           {"impactType": "Mandatory",   "governance": "Internal review; CR if scope-bearing",                         "documentsToUpdate": "Integration code (rework) -> Integration Test Cases", "notes": "Affects integration build and SIT only."},
    "unit-test-cases":  {"impactType": "Conditional", "governance": "Internal review",                                              "documentsToUpdate": "Test Cases (regression coverage - only if defect indicates coverage gap)", "notes": "Self-contained at developer level."},
    "test-strategy":    {"impactType": "Mandatory",   "governance": "Test Lead + Steerco approval if cycles or dates change",        "documentsToUpdate": "Test Cases -> Integration Test Cases -> UAT Test Cases (cycle structure, environments, criteria)", "notes": "Strategy-level changes affect every test artefact downstream."},
    "test-cases":       {"impactType": "Conditional", "governance": "Test Lead review",                                             "documentsToUpdate": "Defect Log -> Integration Test Cases -> UAT Test Cases (if scenario reused)", "notes": "Mid-execution updates common when defects expose coverage gaps."},
    "int-test-cases":   {"impactType": "Conditional", "governance": "Integration Lead review",                                      "documentsToUpdate": "Defect Log", "notes": "Self-contained for the affected interface."},
    "uat-test-cases":   {"impactType": "Mandatory",   "governance": "Business Lead approval",                                       "documentsToUpdate": "Defect Log -> UAT Sign-off (new test execution required)", "notes": "Adding a UAT case post-cycle-start may extend UAT timeline."},
    "defect-log":       {"impactType": "Review",      "governance": "Steerco visibility on severity 1/2",                           "documentsToUpdate": "UAT Sign-off (closure status) -> Hypercare Tracker (open at go-live)", "notes": "Living register. Status changes are continuous."},
    "uat-signoff":      {"impactType": "Mandatory",   "governance": "Steerco confirmation",                                         "documentsToUpdate": "Cutover Plan (gate to start) -> Go-Live Checklist (confirms business approval)", "notes": "Gate document. If withdrawn, cutover halts immediately."},
    "cutover-plan":     {"impactType": "Mandatory",   "governance": "Cutover Manager + Steerco for material change",                "documentsToUpdate": "Go-Live Checklist (activity readiness)", "notes": "Re-baselined after each mock cutover. Final version locks 1-2 days pre-go-live."},
    "golive-checklist": {"impactType": "Mandatory",   "governance": "Steerco Go/No-Go meeting",                                     "documentsToUpdate": "Go/No-Go decision -> Hypercare Tracker (initial state of issues)", "notes": "Final readiness gate. Each item green/amber/red."},
    "kt-plan":          {"impactType": "Conditional", "governance": "Functional Lead + Business Lead",                              "documentsToUpdate": "Trained super-user readiness -> AMS handover pack", "notes": "Updated when training content or audience changes."},
    "hypercare-tracker":{"impactType": "Review",      "governance": "Hypercare Lead -> AMS lead",                                   "documentsToUpdate": "AMS handover pack (open items)", "notes": "Lowest downstream blast radius. End of chain."},
}

# ═══════════════════════════════════════════════════════════════════════════════
# 4. RTM CHANGE WAVES  (pre-defined cascade waves for RTM changes)
# ═══════════════════════════════════════════════════════════════════════════════

RTM_WAVES: List[Dict] = [
    {
        "wave": "Wave 1 - Governance trigger", "color": "#fee2e2",
        "docs": ["risk-log"],
        "steps": ["Raise Change Request (CR) ticket", "Update Risk / Issue Log", "Steerco approval of the CR"],
    },
    {
        "wave": "Wave 2 - Design rebaseline", "color": "#fef3c7",
        "docs": ["rtm", "fit-gap", "solution-design", "func-spec", "interface-list", "int-td"],
        "steps": ["Update RTM", "Update Fit-Gap", "Update Solution Design", "Create/amend Functional Spec", "Update Interface List + Integration TD"],
    },
    {
        "wave": "Wave 3 - Build rework", "color": "#fed7aa",
        "docs": ["config-workbook", "config-specs"],
        "steps": ["Update Configuration Workbook", "Update Configuration Specs", "Re-execute build in environment"],
    },
    {
        "wave": "Wave 4 - Test rework", "color": "#bbf7d0",
        "docs": ["unit-test-cases", "test-cases", "int-test-cases", "uat-test-cases"],
        "steps": ["Update Unit Test Cases", "Update SIT Test Cases", "Update Integration Test Cases", "Update UAT Test Cases"],
    },
    {
        "wave": "Wave 5 - Plan and comms", "color": "#bfdbfe",
        "docs": ["project-plan", "comm-plan", "status-templates"],
        "steps": ["Re-baseline Project Plan", "Trigger Communication Plan - stakeholder alert", "Log CR in Status Report"],
    },
    {
        "wave": "Wave 6 - Deploy artefacts", "color": "#ddd6fe",
        "docs": ["cutover-plan", "kt-plan"],
        "steps": ["Update Cutover Plan (if requirement in current go-live scope)", "Update KT Plan + training content"],
    },
    {
        "wave": "Closure", "color": "#e5e7eb",
        "docs": ["milestone-format"],
        "steps": ["Re-baseline using Milestone Sign-off Format - formally close the change cycle"],
    },
]

# ═══════════════════════════════════════════════════════════════════════════════
# 5. DEFAULT OUTPUT FORMATS  (admin-configurable; these are system defaults)
# ═══════════════════════════════════════════════════════════════════════════════

DEFAULT_OUTPUT_FORMATS: Dict[str, str] = {
    "sow":              "docx",   # Contract narrative
    "project-plan":     "xlsx",   # Schedule / Gantt
    "comm-plan":        "docx",   # Plan narrative
    "gov-matrix":       "xlsx",   # RACI matrix
    "status-templates": "pptx",   # Presentation template
    "risk-log":         "xlsx",   # Risk register
    "onboarding-deck":  "pptx",   # Presentation
    "milestone-format": "docx",   # Sign-off form
    "workshop-notes":   "docx",   # Notes narrative
    "rtm":              "xlsx",   # Traceability matrix
    "to-be":            "docx",   # Process narrative
    "fit-gap":          "xlsx",   # Comparison matrix
    "solution-design":  "docx",   # Design document
    "interface-list":   "xlsx",   # Interface inventory
    "int-strategy":     "docx",   # Strategy narrative
    "dm-strategy":      "docx",   # Strategy narrative
    "func-spec":        "docx",   # Spec document
    "config-workbook":  "xlsx",   # Configuration table
    "config-specs":     "xlsx",   # Spec table
    "int-td":           "docx",   # Technical design
    "unit-test-cases":  "xlsx",   # Test case table
    "test-strategy":    "docx",   # Strategy narrative
    "test-cases":       "xlsx",   # Test case table
    "int-test-cases":   "xlsx",   # Test case table
    "uat-test-cases":   "xlsx",   # UAT test table
    "defect-log":       "xlsx",   # Defect register
    "uat-signoff":      "docx",   # Sign-off document
    "cutover-plan":     "docx",   # Runbook
    "golive-checklist": "xlsx",   # Checklist
    "kt-plan":          "docx",   # Plan narrative
    "hypercare-tracker":"xlsx",   # Tracker register
}

# ═══════════════════════════════════════════════════════════════════════════════
# 6. MODEL SELECTION BY TIER
# ═══════════════════════════════════════════════════════════════════════════════

TIER_MODELS: Dict[str, str] = {
    "T1": "claude-opus-4-8",    # Critical: SOW, Project Plan, RTM, Solution Design
    "T2": "claude-opus-4-8",    # High: Fit-Gap, Func Spec, Test Strategy, Cutover Plan
    "T3": "claude-sonnet-4-6",  # Medium: Config WB, Int Strategy, Test Cases, etc.
    "T4": "claude-sonnet-4-6",  # Low: Status Templates, Risk Log, Defect Log, etc.
}

# Phase display order
PHASES: List[str] = ["Initiate", "Confirm", "Design-Build", "Integrate", "Deploy"]

# ═══════════════════════════════════════════════════════════════════════════════
# 7. INTERNAL INDEXES
# ═══════════════════════════════════════════════════════════════════════════════

_node_index: Dict[str, Dict] = {n["id"]: n for n in NODES}

# Cycle-breaking edges (reverse of mutual relationships — excluded from BFS)
_CYCLE_BREAKS = {
    ("fit-gap", "rtm"),              # rtm -> fit-gap is primary direction
    ("int-strategy", "interface-list"),  # interface-list -> int-strategy is primary
}

# ═══════════════════════════════════════════════════════════════════════════════
# 8. PUBLIC API
# ═══════════════════════════════════════════════════════════════════════════════

def get_node(node_id: str) -> Optional[Dict]:
    """Return node metadata dict, or None."""
    return _node_index.get(node_id)


def get_all_nodes() -> List[Dict]:
    return list(NODES)


def get_all_node_ids() -> List[str]:
    return [n["id"] for n in NODES]


def get_nodes_by_phase(phase: str) -> List[Dict]:
    return [n for n in NODES if n["phase"] == phase]


def get_nodes_grouped_by_phase() -> Dict[str, List[Dict]]:
    """Return {phase: [nodes]} ordered by PHASES constant."""
    result: Dict[str, List[Dict]] = {}
    for phase in PHASES:
        nodes = get_nodes_by_phase(phase)
        if nodes:
            result[phase] = nodes
    return result


def get_children(node_id: str) -> List[str]:
    """Direct downstream dependencies (nodes that need THIS node as input)."""
    return [e["target"] for e in EDGES if e["source"] == node_id]


def get_parents(node_id: str) -> List[str]:
    """Direct upstream dependencies (nodes that THIS node depends on)."""
    return [e["source"] for e in EDGES if e["target"] == node_id]


def get_downstream_nodes(node_id: str) -> List[str]:
    """
    All downstream nodes reachable from node_id via BFS.
    Returns in BFS order (shallowest first). Does not include node_id itself.
    """
    visited: set = set()
    queue = list(get_children(node_id))
    result: List[str] = []
    for nid in queue:
        if nid not in visited and nid != node_id:
            visited.add(nid)
            result.append(nid)
            queue.extend(get_children(nid))
    return result


def get_change_impact(node_id: str) -> Dict[str, str]:
    return CHANGE_IMPACT.get(node_id, {})


def get_default_format(node_id: str) -> str:
    return DEFAULT_OUTPUT_FORMATS.get(node_id, "docx")


def get_model_for_node(node_id: str) -> str:
    node = _node_index.get(node_id)
    tier = node["tier"] if node else "T3"
    return TIER_MODELS.get(tier, "claude-sonnet-4-6")


def compute_bfs_waves(selected_nodes: List[str]) -> List[List[str]]:
    """
    Compute BFS generation waves for a subset of nodes using Kahn's algorithm.

    Nodes whose parents (within the selected set) are all done become available
    in the next wave. Cycle-breaking edges are excluded to handle mutual pairs.

    Returns: list of waves, each wave is a list of node_ids safe to generate
             in parallel (all their in-selected parents are already done).
    """
    selected = set(selected_nodes)

    # Build in-degree and adjacency for selected nodes only
    in_degree: Dict[str, int] = {n: 0 for n in selected}
    adjacency: Dict[str, List[str]] = {n: [] for n in selected}

    for edge in EDGES:
        src, tgt = edge["source"], edge["target"]
        if src not in selected or tgt not in selected:
            continue
        if (src, tgt) in _CYCLE_BREAKS:
            continue
        adjacency[src].append(tgt)
        in_degree[tgt] += 1

    # Kahn's algorithm
    waves: List[List[str]] = []
    current: List[str] = [n for n in selected if in_degree[n] == 0]

    while current:
        # Sort deterministically within each wave: by phase then by id
        current.sort(key=lambda n: (
            PHASES.index(_node_index[n]["phase"]) if n in _node_index else 99,
            n,
        ))
        waves.append(current)
        nxt: List[str] = []
        for node in current:
            for child in adjacency.get(node, []):
                in_degree[child] -= 1
                if in_degree[child] == 0:
                    nxt.append(child)
        current = nxt

    return waves


def get_graph_data_for_frontend() -> Dict:
    """
    Return the full graph data structure compatible with the frontend
    Cytoscape.js visualization (mirrors data.js structure).
    """
    return {
        "nodes": NODES,
        "edges": EDGES,
        "changeImpact": CHANGE_IMPACT,
        "rtmWaves": RTM_WAVES,
        "defaultFormats": DEFAULT_OUTPUT_FORMATS,
        "phases": PHASES,
    }
