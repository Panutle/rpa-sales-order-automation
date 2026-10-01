# Sales Order Automation with TagUI & Tkinter

[![Python 3.8+](https://img.shields.io/badge/Python-3.8+-blue.svg?style=flat-square&logo=python)](https://python.org)
[![TagUI](https://img.shields.io/badge/RPA-TagUI%20Automation-FF6F00?style=flat-square)](https://github.com/tebelorg/TagUI-Python)
[![Azure Logic Apps / Power Automate](https://img.shields.io/badge/Cloud%20Integration-Power%20Automate-0078D4?style=flat-square&logo=microsoftazure)](https://powerautomate.microsoft.com/)
[![Tkinter](https://img.shields.io/badge/GUI-Tkinter%20Desktop-2D8C3C?style=flat-square)](https://docs.python.org/3/library/tkinter.html)

A Robotic Process Automation (RPA) desktop application designed to streamline the batch confirmation lifecycle of Sales Orders (SO) within internal ERP and OASYS web portals. The engine ingests tabular records across legacy and modern spreadsheet formats, automates deterministic browser navigation with DOM-state synchronization, records granular status audit trails, and synchronizes reconciliation summaries directly to cloud endpoints via Power Automate / Azure Logic Apps.

---

## ⚙️ Core Technical Highlights

### 1. Multi-Format Spreadsheet Extraction & Validation
* **Dual Format Ingestion**: Natively reads modern OpenXML workbooks (`.xlsx` via `openpyxl`) and legacy BIFF8 spreadsheets (`.xls` via `xlrd`).
* **Schema Validation & Deduplication**: Enforces strict worksheet naming (`DATA`) and target identifier headers (`เลขที่เอกสาร`). De-duplicates order IDs in-memory to prevent duplicate transaction overhead.

### 2. Event-Driven Web Automation & DOM Synchronization
* **DOM-State Polling**: Utilizes deterministic element condition checks (`r.present`) to accommodate asynchronous AJAX re-rendering and dynamic modal confirmations without brittle static timers.
* **Environment-Aware Target Routing**: Supports multi-tier runtime targets (`DEV`, `UAT`, `PRODUCTION`) with dynamically mapped base URLs and element selector configurations.

### 3. Non-Intrusive Access Control & Config Integrity
* **Configuration Gate**: Compares the `code` value in `config.json` with an allowlist in the source before execution. This local string comparison is not an authentication boundary; portal access must be controlled separately.

### 4. Automated Cloud Telemetry & Reconciliation
* **Result Classification**: Automatically tallies operational outcomes into categorized metrics:
  - `Completed`: Confirmed and submitted.
  - `Pending Manual Review`: Blocked by validation state or non-standard order status.
  - `Not Found`: Sales order ID absent from target system.
* **Webhook Ingestion**: Encodes resulting execution logs into Base64 strings and dispatches a JSON payload directly to Azure Logic Apps / Power Automate for enterprise archiving and team notifications.

---

## 📊 Data Schema Specification

### 1. Input Specification
Input files must be provided in either `.xlsx` or `.xls` format containing the following structural layout:

| Target Property | Required Value / Type | Description |
| :--- | :--- | :--- |
| **Sheet Name** | `DATA` | Target worksheet container |
| **Target Column** | `เลขที่เอกสาร` (Document Number) | Sales Order (SO) identifiers to process |

### 2. Output Reconciliation Schema (`summary_result_YYYY-MM-DD-HH-MM.xlsx`)

| Column Header | Value Example | Description |
| :--- | :--- | :--- |
| **เลขที่เอกสาร** | `SO-2026-0001` | Evaluated Sales Order ID |
| **Status** | `Completed` / `Pending Check` | Status text captured from target system |
| **Comment** | `✅ (กด Confirms เรียบร้อย)` | Diagnostic execution status summary |

---

## 📂 Repository Structure

```text
rpa-sales-order-automation/
├── README.md
├── requirements.txt
├── .gitignore
├── config.json.example       # Sanitized environment access token template
└── src/
    └── auto_confirms_so_oasys.py  # Main desktop GUI and RPA orchestration script
```

---

## Setup and execution

This application targets internal ERP/OASYS portals. To reproduce its browser flow, you need access to an appropriate test portal, matching selectors, Google Chrome, and a desktop Python installation with Tkinter.

### Install dependencies

Run these commands from a terminal with Python available:

```bash
git clone https://github.com/Panutle/rpa-sales-order-automation.git
cd rpa-sales-order-automation
python -m venv .venv
```

Activate the environment using the command for your shell:

| Shell | Command |
| --- | --- |
| Windows PowerShell | `.\.venv\Scripts\Activate.ps1` |
| macOS / Linux | `source .venv/bin/activate` |

```bash
python -m pip install -r requirements.txt
```

### Configure before running

1. Copy `config.json.example` to `config.json` in the repository root and set `code` to a value accepted by your adapted `Variable.LIST_CONFIG_CODE`.
2. Review `Variable` in `src/auto_confirms_so_oasys.py`: replace portal URLs, environment defaults, the configuration allowlist, and `POWERAUTOMATE_URL` for your deployment. The example config only supplies `code`; it does not configure those services.
3. Prepare `.xlsx` or `.xls` input with worksheet `DATA` and column `เลขที่เอกสาร`.
4. Verify TagUI/Chrome setup and portal selectors in a test environment. Running the automation confirms orders and submits reconciliation data to the configured webhook.

```bash
python src/auto_confirms_so_oasys.py
```

Run from the repository root so the application can locate `config.json`. Use a small test workbook first and inspect the generated reconciliation workbook before processing a larger batch.

### Optional Windows packaging

```powershell
pyinstaller src/auto_confirms_so_oasys.py --onefile --windowed --hidden-import=tagui --distpath ./dist
```

The executable still requires its configuration, browser/TagUI runtime, and access to the target portal. Packaging alone does not provision those dependencies.

---

## 🛡️ Exception Handling & Operational Safeguards

- **GUI State Lock**: Automatically locks interactive trigger buttons while the automation loop runs to avoid race conditions and parallel session conflicts.
- **Graceful Process Teardown**: Hooks `r.close()` into application termination routines, ensuring spawned Chrome driver subprocesses are reliably terminated upon exit.
- **Pre-flight Error Prompts**: Validates spreadsheet schema prior to launching browser sessions, informing operators immediately via native UI dialogs if sheets or required columns are missing.
