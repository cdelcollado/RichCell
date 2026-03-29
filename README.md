# RichCell — Excel 365 Add-in

![RichCell](assets/icon-80.png)

A task pane add-in for Microsoft Excel 365 that lets you write and store formatted HTML directly inside cells, using a WYSIWYG editor powered by [Quill.js](https://quilljs.com/).

Designed for workflows where rich HTML content needs to be authored and maintained inside a spreadsheet.

---

## Features

### Editor
- **WYSIWYG editor** — headings (H1/H2), bold, italic, bullet lists, numbered lists, hyperlinks
- **Auto-load** — selecting a cell that contains HTML automatically loads it into the editor
- **Write to cell** — sends clean, sanitised HTML to the active cell with one click (`Ctrl+Enter`)
- **Undo** — restores the original cell content before any edits were made

### Previews
- **Rendered preview** — collapsible panel showing how the HTML will look when rendered
- **Raw HTML preview** — collapsible panel showing the clean HTML source

### Safety & feedback
- **Unsaved indicator** — a pulsing dot appears next to the cell address when there are unsaved changes
- **Status bar** — real-time feedback for every operation (load, save, error)
- **XSS protection** — strips `on*` event attributes and `javascript:` URIs before writing to the cell

### Infrastructure
- **Runs locally** — no cloud dependency, no data leaves your machine
- **Auto-start** — optional Windows scheduled task starts the server at login
- **Team deployment** — one-command installer for additional machines

---

## Requirements

| Requirement | Version |
|---|---|
| Microsoft Excel | 365 (desktop, Windows) |
| Node.js | LTS (18 or later) |
| Windows | 10 or 11 |

---

## Installation

### 1. Clone the repository

```bash
git clone https://github.com/cdelcollado/RichCell.git
cd RichCell
```

### 2. Install Node.js dependencies

```bash
npm install
```

### 3. Install the SSL certificate

The local HTTPS server requires a trusted certificate. Install the included CA certificate to your system trust store:

Open PowerShell **as Administrator** and run:

```powershell
$store = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root","LocalMachine")
$store.Open("ReadWrite")
$store.Add((New-Object System.Security.Cryptography.X509Certificates.X509Certificate2("certs\ca.crt")))
$store.Close()
```

Also install it for the current user (no admin required):

```powershell
$store = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root","CurrentUser")
$store.Open("ReadWrite")
$store.Add((New-Object System.Security.Cryptography.X509Certificates.X509Certificate2("certs\ca.crt")))
$store.Close()
```

### 4. Start the server

```bash
node server.js
```

The server runs at `https://localhost:3000`. Keep this terminal open while using the add-in.

> To start the server automatically on Windows login, see [Auto-start](#auto-start-optional).

### 5. Load the add-in in Excel

1. Open Excel 365
2. Go to **Insert** → **Add-ins** → **Upload My Add-in** (or configure a shared folder catalog — see below)
3. Select `manifest.xml` from the project folder
4. The **RichCell** task pane will open on the right side

#### Shared folder catalog (recommended for teams)

This method avoids re-uploading the manifest after each Excel restart:

1. Share the project folder on your local network (or locally)
2. In Excel: **File** → **Options** → **Trust Center** → **Trust Center Settings** → **Trusted Add-in Catalogs**
3. Enter the UNC path (e.g. `\\YOUR-PC\RichCell`) and click **Add to List**
4. Check **Show in Menu**, click OK
5. Restart Excel
6. **Insert** → **Add-ins** → **SHARED FOLDER** tab → select **RichCell**

---

## Usage

1. **Select a cell** — the editor automatically loads any existing HTML or plain text from that cell
2. **Write or edit** content using the toolbar:
   - `H1` / `H2` — headings
   - **B** / *I* — bold and italic
   - List buttons — bullet or numbered lists
   - Chain icon — insert a hyperlink
   - Eraser — remove all formatting
3. **Send to Excel** (or `Ctrl+Enter`) — writes clean HTML to the cell and auto-fits the row height
4. **Undo** — restores the editor to the cell's original content
5. **Clear** — resets the editor without modifying the cell
6. Expand **Rendered preview** to see how the HTML will look in a browser
7. Expand **HTML source** to inspect or copy the raw HTML

A pulsing orange dot `●` next to the cell address means there are unsaved changes. The status bar confirms every operation.

---

## Auto-start (optional)

To have the server start automatically when you log in to Windows:

```powershell
$action   = New-ScheduledTaskAction -Execute "node.exe" -Argument "server.js" -WorkingDirectory "C:\RichCell"
$trigger  = New-ScheduledTaskTrigger -AtLogon
$settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Hours 0) -RestartCount 3
Register-ScheduledTask -TaskName "RichCell_Server" -Action $action -Trigger $trigger -Settings $settings -RunLevel Highest -Force
```

To remove the scheduled task:

```powershell
Unregister-ScheduledTask -TaskName "RichCell_Server" -Confirm:$false
```

---

## Team deployment

Use the automated installer to deploy to additional machines:

```powershell
powershell -ExecutionPolicy Bypass -File "\\YOUR-PC\RichCell\deploy\install.ps1"
```

See [`deploy/install.ps1`](deploy/install.ps1) for details.

---

## Project structure

```
RichCell/
├── manifest.xml          # Office Add-in manifest
├── server.js             # Node.js HTTPS server (port 3000)
├── taskpane.html         # Task pane UI
├── taskpane.js           # Excel JS API logic
├── taskpane.css          # Styles
├── generate-icons.js     # Generates assets/icon-{16,32,80}.png
├── assets/               # Add-in icons
├── certs/                # SSL certificate (CA + server, localhost)
└── deploy/               # Team deployment scripts
    ├── install.ps1       # Automated installer for new machines
    └── create-zip.ps1    # Creates a distributable ZIP package
```

---

## Tech stack

- [Office.js](https://learn.microsoft.com/en-us/office/dev/add-ins/) — Excel JavaScript API
- [Quill.js 1.3.7](https://quilljs.com/) — WYSIWYG editor
- Node.js — local HTTPS static file server

---

## License

MIT
