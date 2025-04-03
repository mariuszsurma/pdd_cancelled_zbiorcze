pyinstaller -F --paths=../LIB --onefile --hidden-import win32timezone wysylka_raportow_brakow.py
copy .\dist\wysylka_raportow_brakow.exe .
copy .\dist\wysylka_raportow_brakow.exe d:\export
pause
