# EnvForge

A Windows GUI tool for managing User and System environment variables side by side — with drill-down support for multi-value variables like `Path`.

![Platform](https://img.shields.io/badge/platform-Windows-blue) ![Python](https://img.shields.io/badge/python-3.x-blue) ![License](https://img.shields.io/badge/license-MIT-green)

---

## Features

- **Side-by-side view** — User and System variables displayed in two panels simultaneously
- **Drill-down for multi-value variables** — Double-click any variable containing `;` (e.g. `Path`, `PATHEXT`) to expand its entries line by line in both panels at once
- **Entry-level copy/move** — Transfer individual entries between User and System scopes while in drill-down view
- **Full CRUD** — Add, edit, delete, and reorder entries within drill-down; add, edit, delete whole variables in the main view
- **Copy/Move whole variables** — Transfer an entire variable from one scope to the other
- **Filter** — Real-time search/filter within each panel
- **Sortable columns** — Click column headers to sort by name or value
- **Auto-elevates to Administrator** — UAC prompt on launch to enable System variable editing
- **Live broadcast** — Changes are immediately broadcast to Windows so new processes pick them up

---

## Requirements

- Windows 10/11
- Python 3.x
- No third-party packages (uses only `tkinter` and `winreg` from the standard library)

---

## Usage

### Recommended: run via the batch file

```
EnvForge.bat
```

This handles admin elevation automatically. If not already running as Administrator, it relaunches via PowerShell `RunAs`.

### Or run directly

```
python EnvForge.py
```

If not admin, the script triggers a UAC prompt and relaunches itself elevated.

---

## Drill-Down

Variables whose value contains `;` are highlighted in **blue** in the main list. Double-clicking one opens drill-down mode in both panels:

| Action | How |
|---|---|
| Edit an entry | Double-click the row, or select + **Edit** |
| Add an entry | **+ Add** button |
| Delete an entry | Select + **Delete** |
| Reorder entries | **▲ / ▼** buttons |
| Copy entry to other scope | Select + **Copy →** |
| Move entry to other scope | Select + **Move →** |
| Return to variable list | **← Back** in the panel header |

If the variable doesn't exist in the other scope yet, you'll be prompted to create it as empty before drill-down activates.

---

## Keyboard Shortcuts

| Key | Action |
|---|---|
| `F5` | Refresh all variables |
| `Delete` | Delete selected variable |
| `Double-click` | Edit variable (or drill down if multi-value) |

---

## File Structure

```
EnvForge.py      # Main application
EnvForge.bat     # Launcher with admin elevation
```
