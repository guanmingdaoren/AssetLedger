@echo off
setlocal
set "PYTHONPATH=%~dp0src"
python -m asset_ledger.app
endlocal

