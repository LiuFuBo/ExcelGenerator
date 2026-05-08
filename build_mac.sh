#!/bin/bash
# Excel订单生成器 - Mac .dmg 构建脚本 (PySide6版本)
# 注意: 必须使用 Homebrew Python 3.11+ 构建（自带Python 3.8的tkinter在Monterey上有bug）

echo "=== Excel订单生成器 Mac构建 (PySide6) ==="

# 安装依赖
echo "1. 安装Python依赖..."
python3.11 -m pip install PySide6 openpyxl pypinyin pyinstaller

# 清理旧构建
echo "2. 清理旧构建文件..."
rm -rf build dist ExcelGenerator.spec ExcelGenerator.dmg

# 构建Mac .app
echo "3. 构建Mac应用..."
python3.11 -m PyInstaller --windowed \
    --name "ExcelGenerator" \
    --icon=header.icns \
    --hidden-import=pypinyin \
    --hidden-import=openpyxl \
    --hidden-import=openpyxl.cell._writer \
    --add-data="jj_header_rounded.png:." \
    --collect-data=pypinyin \
    --noconfirm \
    app.py

if [ ! -d "dist/ExcelGenerator.app" ]; then
    echo "构建失败! .app未生成"
    exit 1
fi

echo "4. .app构建成功，创建.dmg..."
hdiutil create -volname "ExcelGenerator" \
    -srcfolder dist/ExcelGenerator.app \
    -ov -format UDZO \
    ExcelGenerator.dmg

if [ -f "ExcelGenerator.dmg" ]; then
    echo "=== 构建完成 ==="
    echo ".dmg文件: ExcelGenerator.dmg (约50MB)"
    echo ".app文件: dist/ExcelGenerator.app"
    echo "双击.dmg安装，或将.app拖到应用程序文件夹"
else
    echo ".dmg创建失败"
    exit 1
fi