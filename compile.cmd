cd C:\Python\wingetpro
pyinstaller --icon=it4home.ico --add-data "wingetpro.ico;." wingetpro.py -n Bookmark --onefile --noconfirm --noconsole --windowed
del bookmark.spec
rmdir /S /Q build
copy dist\*.exe .
rmdir /S /Q dist

