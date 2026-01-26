# Gold, Silver & GDXJ Desktop App

A modern Windows desktop dashboard that shows live prices for gold, silver, and GDXJ with an interactive chart.

## Features
- Snapshot cards for gold, silver, and GDXJ.
- Interactive chart with zoom/pan and quick range selection.
- Auto-refresh every minute.

## Run locally
```bash
python -m venv .venv
. .venv/bin/activate  # On Windows: .venv\Scripts\activate
pip install -r requirements.txt
python main.py
```

On Windows, double-clicking `main.py` will launch the app and automatically hide the console window.

## Build a Windows executable
```powershell
python -m venv .venv
. .venv\Scripts\activate
pip install -r requirements.txt
pip install pyinstaller
pyinstaller --noconsole --onefile --name MetalsDashboard main.py
```

The executable will be in the `dist/` folder.
