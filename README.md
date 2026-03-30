# Microsoft 365 Environment Manager

A lightweight Windows desktop app for managing Microsoft 365 tenants. Authenticate with your admin account (full MFA supported) and manage users, licenses, and run raw Graph API queries — all from one simple interface.

## Quick Start (Windows)

1. Install [Python 3.10+](https://www.python.org/downloads/) — check **"Add Python to PATH"** during setup
2. Download / clone this repo
3. Double-click **`run.bat`**

That's it. `run.bat` creates a virtual environment, installs all dependencies, and launches the app automatically. First run takes about 30 seconds; after that it starts instantly.

## Features

- **Users tab** — list, search, create, edit, and delete users; manage per-user licenses
- **Licenses tab** — view all SKUs in your tenant with total, consumed, and remaining seat counts
- **Graph Explorer tab** — run any Microsoft Graph API call (GET / POST / PATCH / DELETE) with a JSON response viewer
- **Interactive MFA login** — clicks "Sign In", your browser opens to the Microsoft login page, complete MFA as normal — the app connects automatically

## Requirements

- Python 3.10 or later
- A Microsoft 365 account with appropriate admin permissions (Global Admin or User Admin recommended)
- Internet connection

## Auth

Uses Microsoft's interactive browser login flow (the same mechanism as `Connect-ExchangeOnline` in PowerShell). No passwords are stored — authentication is handled entirely by Microsoft's login page in your browser. Supports all MFA methods (Authenticator app, SMS, hardware keys, etc.).

## Manual Setup

```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python main.py
```

## Build a Standalone .exe

```powershell
pip install pyinstaller
pyinstaller --noconsole --onefile --name M365Manager main.py
```

The executable will be in the `dist\` folder. No Python installation required to run it.
