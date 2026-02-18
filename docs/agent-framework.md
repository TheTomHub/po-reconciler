# Agent Framework — Order Processing Suite

## Vision
Harvey.ai for Order Processing. An agentic suite inside Microsoft 365 that handles the full order processing workflow: Capture → Validate → Enter → Reconcile → Resolve → Predict. Operator-first, bottom-up distribution via Copilot + Excel + Outlook.

## Architecture

### Dual-Mode Operation
The suite runs in two modes from the same codebase:

1. **Add-in mode** (current): Task pane UI with buttons. Works without Copilot license. Manual workflow.
2. **Agent mode** (new): Copilot declarative agent with natural language interface. Same Office.js functions underneath. Guided workflow with reasoning transparency.

Both modes share the same engine functions. The agent wraps them with Copilot-native interaction.

### Key Microsoft Requirements
- **Unified manifest** (JSON) — need to convert from current XML manifest
- **declarativeAgent.json** — agent personality, instructions, conversation starters
- **Office-API-local-plugin.json** — function definitions with parameter schemas and reasoning states
- **Office.actions.associate()** — maps function names to agent actions
- Preview: Windows + Web only (Mac coming). Excel/PowerPoint/Word (Outlook coming).

## Agent Design (Harvey Model)

Each action follows the Harvey pattern:

| Principle | Implementation |
|-----------|---------------|
| **Plan** | Agent breaks task into steps, shows plan before executing |
| **Adapt** | Learns from operator corrections (future: intelligence layer) |
| **Interact** | Asks for context when unsure (e.g., "Which column is the SKU?") |
| **Transparent** | Reasoning + responding states show what the agent decided and why |

## Phase 1: Agent-Enable Existing Add-in

Expose current reconciliation functions as Copilot agent actions:

### Actions

| Action ID | Description | Parameters | Returns |
|-----------|-------------|------------|---------|
| `ReconcilePO` | Run PO vs ERP reconciliation | `{ tolerance, currency }` | Summary + exception count |
| `GenerateCreditNote` | Create credit note sheet | `{}` | Confirmation + sheet name |
| `GenerateReInvoice` | Create re-invoice sheet | `{}` | Confirmation + sheet name |
| `DraftExceptionEmail` | Generate email draft | `{}` | Email subject + body |

### Conversation Starters
- "Reconcile the PO against ERP data"
- "Generate a credit note for the exceptions"
- "Draft an email about the pricing discrepancies"
- "Show me the reconciliation summary"

## Phase 2: Capture Module

New agent actions for PO data extraction:

### Actions

| Action ID | Description | Parameters | Returns |
|-----------|-------------|------------|---------|
| `ExtractPOData` | Extract structured data from uploaded PO | `{ format }` | Structured rows in staging sheet |
| `ValidatePOData` | Check extracted data against rules | `{ rules }` | Validation results |
| `StagePOForEntry` | Format staging sheet for ERP push | `{ erpFormat }` | Ready-to-push sheet |

### Workflow
1. User: "I received a PO from Tesco, extract the order data"
2. Agent: reads the active sheet or uploaded file
3. Agent: extracts SKU, price, qty, name into structured format
4. Agent: writes staging sheet with standardized columns
5. Agent: "I extracted 45 line items. 3 items have unusual quantities — please review rows 12, 28, 41."
6. User corrects if needed
7. Agent learns from corrections (future)

## Phase 3+: Validate, Enter, Resolve, Predict

Each phase adds new actions to the same agent. The declarative agent grows its action set while keeping the same unified interface.

## File Structure (Target)

```
src/
  commands/
    commands.js          — Office.actions.associate() bindings
    commands.html         — loader for commands.js
  reconcile/
    reconcile.js          — core reconciliation engine (unchanged)
    results.js            — Excel sheet writer (unchanged)
    creditnote.js         — credit note generator (unchanged)
    creditnote-results.js — credit note sheet writer (unchanged)
  capture/
    extractor.js          — PO data extraction engine (new)
    staging.js            — staging sheet writer (new)
  validate/
    rules.js              — validation rules engine (future)
  email/
    email.js              — email draft generator (unchanged)
  utils/
    format.js             — currency formatting (unchanged)
  taskpane/
    taskpane.html         — add-in UI (unchanged, backward compat)
    taskpane.js           — add-in logic (unchanged)
appPackage/
  manifest.json           — unified manifest (converted from XML)
  declarativeAgent.json   — agent config
  Office-API-local-plugin.json — plugin config with all actions
  assets/
    color.png
    outline.png
```

## Conversion Path

1. Convert XML manifest to unified JSON manifest
2. Add `copilotAgents` section to manifest
3. Create declarativeAgent.json with instructions + conversation starters
4. Create Office-API-local-plugin.json with function definitions
5. Update commands.js with Office.actions.associate() for each action
6. Test agent in Excel with Copilot
7. Keep task pane working for non-Copilot users
