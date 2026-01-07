@echo off
echo ========================================
echo    CAI DAT MOI TRUONG PORTABLE
echo ========================================
echo.

REM Tai Python Embedded (khong can cai dat)
echo Dang tai Python portable...
curl -L -o python.zip https://www.python.org/ftp/python/3.11.7/python-3.11.7-embed-amd64.zip
tar -xf python.zip -C python_embed
del python.zip

REM Tai pip
echo Dang cai pip...
curl -L -o get-pip.py https://bootstrap.pypa.io/get-pip.py
python_embed\python.exe get-pip.py
del get-pip.py

REM Sua file pth de cho phep pip
echo import site >> python_embed\python311._pth

REM Cai cac thu vien
echo Dang cai thu vien...
python_embed\python.exe -m pip install flask requests openpyxl unidecode --target=python_embed\Lib\site-packages

echo.
echo ========================================
echo    HOAN TAT! Copy folder nay vao USB
echo ========================================
pause
