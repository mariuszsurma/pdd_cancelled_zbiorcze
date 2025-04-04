pyinstaller -F --paths=../LIB --onefile --hidden-import win32timezone pdd_anulowane_zbiorcze.py
copy .\dist\pdd_anulowane_zbiorcze.exe .
copy .\dist\pdd_anulowane_zbiorcze.exe d:\export
pause
