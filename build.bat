@echo off
REM ============================================================
REM  一次性打包腳本：在 Windows 上執行，產生單一 .exe 檔案
REM  執行完後，exe 會出現在 dist\FileConverter.exe
REM  之後把 dist\FileConverter.exe 複製到任何 Windows 電腦即可執行，
REM  不需要再安裝 Python 或任何 pip 套件。
REM ============================================================

echo [1/3] 安裝所需套件...
pip install -r requirements.txt
pip install pyinstaller

echo.
echo [2/3] 開始打包成 exe...
pyinstaller --onefile --windowed --name FileConverter ^
    --hidden-import=win32timezone ^
    --hidden-import=win32com.client ^
    file_converter_gui.py

echo.
echo [3/3] 完成！
echo exe 檔案位置： dist\FileConverter.exe
echo 你可以把這個 exe 複製到其他 Windows 電腦直接使用。
echo.
pause
