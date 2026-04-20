@echo off
python -m PyInstaller --onefile --noconsole --icon=lunch.ico lunch_menu_extractor.py
if not exist release mkdir release
copy /Y dist\lunch_menu_extractor.exe release\급식표_추출기.exe
pause