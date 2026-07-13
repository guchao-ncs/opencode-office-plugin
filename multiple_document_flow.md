# Multi-Document & Multi-Instance Workflow

This diagram illustrates how the task pane, OpenCode agent, and `word_mcp_launcher` wrapper coordinate to ensure mutations target the correct file when you have multiple documents or instances of Word desktop running simultaneously.

## 1. High-Level Architecture & Binding Flow

```mermaid
sequenceDiagram
    autonumber
    actor User as User in Word Document B
    participant TP as Taskpane UI (Doc B)
    participant OC as OpenCode Server
    participant L as word_mcp_launcher.py
    participant COM as Windows ROT (COM)
    participant WordB as winword.exe (Doc B Instance)
    
    User->>TP: Type message: "Add Heading 1 'Section 2'"
    
    Note over TP: 1. Office.js retrieves Document Identity
    TP->>TP: Get URL: "C:\path\to\DocB.docx"<br/>Compute text hash of Doc B
    
    Note over TP: 2. Session Integrity Gating
    alt URL or Hash changed since last message
        TP->>TP: Reset sessionId to null (Session rollover)
    end
    
    TP->>OC: POST /session (Creates/returns session ID)
    
    Note over TP: 3. Inject Identity Block & Contract
    TP->>OC: POST /session/{id}/prompt_async<br/>(Prompt + DocB URL + hash contract)
    
    Note over OC: 4. Agent intends to mutate document
    OC->>L: MCP Request: word_live_add_heading(text="Section 2", filename="DocB.docx")
    
    Note over L: 5. Stack Inspection
    L->>L: Inspect call stack: detect filename="DocB.docx"
    
    Note over L: 6. ROT Discovery & Disambiguation
    L->>COM: Query Running Object Table
    COM-->>L: Return all active Document & Application instances
    L->>L: Loop instances & match "DocB.docx"<br/>Find Application B owning Doc B
    
    Note over L: 7. Target-Bound Mutation
    L->>WordB: Execute COM call specifically on Application B / Document B
    WordB-->>User: Physical mutation visible instantly
    WordB-->>L: Success
    L-->>OC: MCP Response
    OC-->>TP: SSE event updates UI
    TP-->>User: Assistant bubble: "Added heading to DocB.docx"
```

## 2. Key Safety Guardrails

### A. Saved Documents
For a document saved to a path (e.g. `DocB.docx`):
1. **Identify**: The task pane gets `Office.context.document.url` (`C:\path\to\DocB.docx`).
2. **List**: The agent calls `word_live_list_open` to list documents known to COM.
3. **Match**: The agent matches `C:\path\to\DocB.docx` or `DocB.docx` to exactly one open COM document.
4. **Target**: The agent passes `filename="DocB.docx"` to all subsequent edits.
5. **Fail Closed**: If no match or multiple matches exist, the agent stops and prompts you to resolve the conflict.

### B. Unsaved Documents
For a newly created unsaved document (e.g. `Document1`):
1. **Nonce Generation**: The task pane generates a session nonce (e.g. `nonce-3g7k9s2a`).
2. **List Check**: The agent calls `word_live_list_open`.
3. **Guard**: 
   * If **exactly one** unsaved document is open globally, the agent prompts: *"I see only one unsaved document. Please confirm if you want me to edit this."*
   * If **two or more** unsaved documents are open, the agent **fails closed** and asks you to save the document first to prevent guessing target windows.
