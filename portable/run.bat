@echo off
echo ========================================
echo    TRA CUU DIEM REN LUYEN
echo ========================================
echo.
echo Dang khoi dong server...
echo Mo trinh duyet: http://localhost:5000
echo.
start http://localhost:5000
python_embed\python.exe app.py
pause
