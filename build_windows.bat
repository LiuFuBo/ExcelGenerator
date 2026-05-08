@echo off
chcp 65001 >nul
echo === Excel订单生成器 Windows构建 (PySide6) ===

echo 1. 安装Python依赖...
pip install PySide6 openpyxl pypinyin

echo 2. 安装PyInstaller...
pip install pyinstaller

echo 3. 清理旧构建文件...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist *.spec del /q *.spec

echo 4. 构建Windows .exe（单文件模式）...
pyinstaller --onefile --windowed ^
    --name "ExcelGenerator" ^
    --icon=header.ico ^
    --hidden-import=pypinyin ^
    --hidden-import=openpyxl ^
    --hidden-import=openpyxl.cell._writer ^
    --add-data="jj_header_rounded.png;." ^
    --collect-data=pypinyin ^
    --noconfirm ^
    app.py

if not exist "dist\ExcelGenerator.exe" (
    echo 构建失败! .exe未生成
    pause
    exit /b 1
)

echo === 构建完成 ===
echo .exe文件: dist\ExcelGenerator.exe
echo 可以将ExcelGenerator.exe复制到任何Windows电脑上直接运行
pause