@echo off
echo Cleaning previous build...
if exist dist rmdir /s /q dist
if exist build rmdir /s /q build
if exist TibiaFrames.spec del TibiaFrames.spec

echo Building TibiaFrames.exe...
pyinstaller --onefile --windowed --icon=tibia_icon.ico --name=TibiaFrames tibiaframes_v1_2_4.pyw

if exist dist\TibiaFrames.exe (
    echo.
    echo Build successful: dist\TibiaFrames.exe
) else (
    echo.
    echo Build FAILED.
)
pause
