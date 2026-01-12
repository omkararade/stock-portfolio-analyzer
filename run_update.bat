@echo off
pushd %~dp0

echo ===== Updating Excel Dashboard =====

python\python.exe Backend\update_excel.py

echo ===== Update Complete =====
pause
popd