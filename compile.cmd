cd C:\Python\winGetPro
pyinstaller --icon=wingetpro.ico --add-data "wingetpro.ico;." wingetpro.py -n WinGetPro --onefile --noconfirm --noconsole --windowed
del wingetpro.spec
rmdir /S /Q build
copy dist\*.exe .
rmdir /S /Q dist

