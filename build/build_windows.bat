@echo off
cd /d %~dp0..
python -m pip install -r requirements.txt
pyinstaller --noconfirm --clean build\File2PPT.spec
