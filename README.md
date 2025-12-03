# 📧 Data Collection OFT Generator (Multi-Function)

Streamlit app to generate **Outlook .oft templates** for Data Collection outreach across multiple functions:

- **ER&D** (Engineering & R&D) – uses firm-approved wording (exact copy preserved)
- **Supply Chain** – placeholder “lorem ipsum” text with the same structure
- **Procurement** – placeholder “lorem ipsum” text with the same structure
- **Manufacturing** – placeholder “lorem ipsum” text with the same structure

For each function, the app:

1. Takes an Excel file with case / contact details  
2. Renders **three email templates** per row:
   - **Sebastian – Initial**
   - **POC – Follow-up**
   - **Aseem – Escalation**
3. Generates real **Outlook .oft files** via Outlook COM
4. Zips the templates into three folders (one per template) for download

> ⚠️ **Environment requirement:** This app only works on **Windows** with **Outlook (pywin32)** installed and accessible.

---

## 🔧 Features

- **Multi-function support**
  - Landing page lets you choose a function: `ER&D`, `Supply Chain`, `Procurement`, `Manufacturing`
  - ER&D uses **approved Bain wording** (do not edit in code unless re-approved)
  - Non-ER&D functions use **placeholder “lorem ipsum”** bodies but with consistent structure, subject pattern, and routing logic

- **Outlook OFT generation via COM**
  - Uses `win32com.client` to create `.oft` files
  - Ensures:
    - `BodyFormat = 2` (HTML)
    - `SaveAs(..., 2)` → Outlook Template

- **Recipient handling & hardening**
  - Excel columns read as **strings** (`dtype=str`) to avoid type surprises
  - Normalization of recipient lists:
    - Accepts both comma and semicolon separated addresses
    - Merges CC chunks, de-dups (case-insensitive), preserves order
  - Supports `Display Name <email@...>` or plain emails
  - Validates addresses with a conservative email regex
  - Skips rows with **invalid recipients**, with per-row status messages

- **Jinja2 templating with auto-escape**
  - Uses a hardened Jinja environment
  - Renders HTML bodies safely with autoescaping behavior
  - Simple context: `client_name`, `case_code`, `case_manager_name`, `poc_display_name`, and `today`

- **Single Outlook session reuse**
  - Uses a context-managed `OutlookSession` that:
    - Initializes COM (`CoInitialize`) once
    - Reuses a single `Outlook.Application` instance for all rows
    - Uninitializes COM at the end

- **Per-row status & preview**
  - Preview panel shows:
    - Subject, To, CC, BCC
    - Rendered HTML body for each template (from the **first row**)
  - Status log after generation:
    - `Row X: OK – <client> - <code>`
    - `Row X: SKIPPED – reason`
    - `Row X: FAILED – error`

---

## 📦 Requirements

### OS & Apps

- **Windows** (required for Outlook COM)
- **Microsoft Outlook** installed and configured

### Python Packages

- Python 3.8+ (recommended)
- `streamlit`
- `pandas`
- `jinja2`
- `pywin32` (for Outlook COM integration)

You can install dependencies with:

```bash
pip install streamlit pandas jinja2 pywin32
