$ErrorActionPreference = "Stop"

python -m PyInstaller --noconfirm --onefile --console --name ReportCommesse .\src\report_commesse.py
