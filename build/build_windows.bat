@echo off
cd /d %~dp0..
python -m pip install -r requirements.txt
pyinstaller --noconfirm --clean build\File2PPT.spec
if exist dist\File2PPT.zip del /f /q dist\File2PPT.zip
powershell -Command "Compress-Archive -Path dist\File2PPT\* -DestinationPath dist\File2PPT.zip -Force"
