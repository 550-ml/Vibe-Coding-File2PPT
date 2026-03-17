@echo off
cd /d %~dp0..
python -m pip install -r requirements.txt
pyinstaller --noconfirm --clean --onefile --windowed build\File2PPT.spec
