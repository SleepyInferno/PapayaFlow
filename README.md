# PapayaFlow

A local web-based tool that transforms raw PaperCut Hive print activity reports into department-level cost and usage summaries. Upload a PaperCut report and a Microsoft Intune/Entra ID user export, and PapayaFlow maps each user to their department, aggregates print statistics, and displays an interactive dashboard — no manual spreadsheet work required.

## Features

- **PDF and CSV parsing** — Accepts PaperCut Hive "User Activity Summary" PDFs (via `pdftotext`) or CSV exports (transaction logs, activity summaries, Hive activity exports)
- **Department mapping** — Matches users to departments using a Microsoft Intune/Entra ID user export (case-insensitive UPN matching, "Unassigned" fallback)
- **Interactive dashboard** — Sortable department table with expandable per-user detail rows, org-wide stat cards, date range header
- **Cost-driver flagging** — Departments and users exceeding a configurable color-printing threshold are visually flagged with a "HIGH COLOR" badge
- **Dark mode** — Toggle between light and dark themes with persistence across sessions
- **Drag-and-drop upload** — Drop files directly onto the upload area
- **CSV export** — One-click download of a 16-column flat CSV with department summary rows and per-user detail rows
- **Zero dependencies** — PowerShell 5.1 backend, vanilla JavaScript frontend, no npm/pip/gems required

## Screenshots

### Upload View
Upload a PaperCut report (PDF or CSV) and your Intune user export to generate the dashboard.

### Dashboard View
Org-wide stat cards, sortable department table with expandable user rows, color-threshold flagging, and CSV export.

## Requirements

- **Windows** with PowerShell 5.1+ (ships with Windows 10/11)
- **pdftotext.exe** (from [Poppler for Windows](https://github.com/oschwartz10612/poppler-windows/releases)) — only required if uploading PDF reports; CSV uploads work without it

## Setup

1. Clone the repository:
   ```
   git clone https://github.com/SleepyInferno/PapayaFlow.git
   cd PapayaFlow
   ```

2. **(PDF support only)** Download [Poppler for Windows](https://github.com/oschwartz10612/poppler-windows/releases), extract it, and copy `pdftotext.exe` into the `bin/` directory:
   ```
   mkdir bin
   copy path\to\poppler\Library\bin\pdftotext.exe bin\
   ```

3. Start the server:
   ```powershell
   .\Start-PapayaFlow.ps1
   ```

4. Your browser opens automatically to `http://localhost:8080`.

## Usage

1. **Upload files** — Select or drag-and-drop two files:
   - A **PaperCut Hive report** (PDF or CSV)
   - A **Microsoft Intune/Entra ID user export** (CSV)
2. Click **Process Report**
3. Browse the dashboard — sort columns, expand departments to see individual users, adjust the color threshold
4. Click **Export CSV** to download the results

## Preparing Your Input Files

### PaperCut Hive Report

PapayaFlow accepts the PaperCut Hive report in three formats:

**Option A: PDF (User Activity Summary)**

1. Log in to PaperCut Hive admin console
2. Navigate to **Reports** > **User activity summary**
3. Select the desired date range
4. Export/download as **PDF**
5. The PDF should contain columns: Rank, User (email), Pages, Print, Copy, B&W, Color, Jobs, 1-Sided, 2-Sided, Scans, Fax, Cost

> Requires `pdftotext.exe` in the `bin/` folder.

**Option B: CSV (Hive Activity Export)**

1. Log in to PaperCut Hive admin console
2. Navigate to **Reports** > **User activity summary**
3. Export as **CSV**
4. Expected columns include:
   - `User` — user's email address
   - `Print+Copy total pages`
   - `Print job BW pages`, `Print job color pages`, `Print job total pages`
   - `Copy job BW pages`, `Copy job color pages`, `Copy job total pages`
   - `Print jobs 1-sided`, `Print jobs 2-sided`, `Print jobs total`
   - `Copy jobs 1-sided`, `Copy jobs 2-sided`, `Copy jobs total`
   - `Scan pages`, `Fax pages`
   - `Cost`

**Option C: CSV (Transaction Log)**

1. Export a PaperCut transaction log CSV
2. Required columns:
   - `User Email` — user's email address
   - `Transaction type` — must include `Print job` or `Copy job` entries
   - `Amount` — transaction cost
   - `Date` — transaction date (used to derive the date range)

> PapayaFlow auto-detects which CSV format you uploaded and parses accordingly.

### Microsoft Intune / Entra ID User Export

This CSV maps each user's email (UPN) to their department. Export it from the Microsoft Intune admin center or Entra ID (Azure AD) portal.

**From Microsoft Intune Admin Center:**

1. Go to [intune.microsoft.com](https://intune.microsoft.com)
2. Navigate to **Users** > **All users**
3. Click **Export users** (top toolbar) or use **Columns** to ensure the right fields are visible
4. Download the CSV

**From Entra ID (Azure AD) Portal:**

1. Go to [entra.microsoft.com](https://entra.microsoft.com)
2. Navigate to **Users** > **All users**
3. Click **Download users** (or **Bulk operations** > **Download users**)
4. Download the CSV

**From Microsoft Graph / PowerShell:**

```powershell
# Using Microsoft Graph PowerShell
Connect-MgGraph -Scopes "User.Read.All"
Get-MgUser -All -Property UserPrincipalName, Department |
    Select-Object UserPrincipalName, Department |
    Export-Csv -Path "users.csv" -NoTypeInformation
```

**Required columns:**

| Column | Description | Required |
|--------|-------------|----------|
| `UserPrincipalName` (or `UPN`) | The user's email address — must match the emails in the PaperCut report | Yes |
| `Department` | The department name (e.g., "Marketing", "Finance", "IT") | Yes |

- Column name matching is **case-insensitive**
- Any column containing `"principal"` is accepted as the UPN field
- Any column containing `"department"` is accepted as the department field
- Users with a blank or missing department are grouped under **"Unassigned"**
- Extra columns in the CSV are ignored — only UPN and Department are used

**Example CSV:**

```csv
UserPrincipalName,DisplayName,Department,JobTitle
user1@example.org,Alex Taylor,Marketing,Manager
user2@example.org,Jordan Lee,Finance,Analyst
user3@example.org,Sam Rivera,,Intern
```

In this example, Bob Wilson would appear under "Unassigned" because his Department is empty.

## CSV Export Format

The exported CSV contains 16 columns with two row types:

| Column | Department Row | User Row |
|--------|---------------|----------|
| `Type` | "Department" | "User" |
| `Department` | Department name | Department name (repeated) |
| `User` | *(empty)* | User's email (UPN) |
| `ActiveUsers` | Count of active users | *(empty)* |
| `TotalPages` | Department total | User's pages |
| `TotalPrint` | Department total | User's print pages |
| `TotalCopy` | Department total | User's copy pages |
| `TotalBW` | Department total | User's B&W pages |
| `TotalColor` | Department total | User's color pages |
| `PctColor` | Department % | User's color % |
| `TotalOneSided` | Department total | User's 1-sided |
| `TotalTwoSided` | Department total | User's 2-sided |
| `TotalScans` | Department total | User's scans |
| `TotalFax` | Department total | User's fax |
| `TotalJobs` | Department total | User's jobs |
| `TotalCost` | Department total | User's cost |

## Project Structure

```
PapayaFlow/
  Start-PapayaFlow.ps1    # Entry point — starts the local web server
  lib/
    Server.ps1             # HTTP listener, routing, multipart form parsing
    PdfParser.ps1          # PaperCut PDF and CSV report parsing
    EntraMapper.ps1        # Intune/Entra ID CSV parsing (UPN -> Department)
    Aggregator.ps1         # Department grouping and summary computation
  web/
    index.html             # Dashboard HTML structure
    styles.css             # Design system with light/dark theme tokens
    app.js                 # Upload flow, dashboard rendering, CSV export
  bin/                     # Place pdftotext.exe here (not included)
```

## Tech Stack

- **Backend:** PowerShell 5.1 with `System.Net.HttpListener` — no external modules
- **Frontend:** Vanilla JavaScript (ES5-compatible), plain HTML/CSS with CSS custom properties for theming
- **PDF parsing:** `pdftotext -layout` from Poppler (user-supplied binary)
- **Architecture:** Single-page app served by a local PowerShell HTTP server on port 8080; all processing happens on your machine — no data leaves localhost

## License

MIT
