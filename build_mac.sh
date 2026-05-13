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

echo "4. 创建安装说明文件..."
cat > dist/安装说明.txt << 'INSTRUCTIONS'
========================================
  ExcelGenerator 安装说明 (Mac版)
========================================

首次打开时，macOS 可能提示"无法验证开发者"，
这是因为应用未经过 Apple 公证。请按以下步骤操作：

方法一（推荐）：
  1. 将 ExcelGenerator.app 拖到"应用程序"文件夹
  2. 右键点击 ExcelGenerator.app → 选择"打开"
  3. 弹出确认窗口后，点击"打开"即可

方法二：
  1. 将 ExcelGenerator.app 拖到"应用程序"文件夹
  2. 打开"终端"（在启动台搜索"终端"）
  3. 输入以下命令并回车：
     xattr -cr /Applications/ExcelGenerator.app
  4. 之后即可正常双击打开

========================================
INSTRUCTIONS

echo "5. .app构建成功，创建.dmg..."
# 创建临时目录用于dmg内容，包含.app和安装说明
dmg_temp="dist/dmg_temp"
rm -rf "$dmg_temp"
mkdir -p "$dmg_temp"
cp -R dist/ExcelGenerator.app "$dmg_temp/"
cp dist/安装说明.txt "$dmg_temp/"

hdiutil create -volname "ExcelGenerator" \
    -srcfolder "$dmg_temp" \
    -ov -format UDZO \
    ExcelGenerator.dmg

rm -rf "$dmg_temp"

if [ -f "ExcelGenerator.dmg" ]; then
    echo "=== 构建完成 ==="
    echo ".dmg文件: ExcelGenerator.dmg (约50MB)"
    echo ".app文件: dist/ExcelGenerator.app"
    echo "DMG内包含: ExcelGenerator.app + 安装说明.txt"
    echo "双击.dmg安装，将.app拖到应用程序文件夹"
else
    echo ".dmg创建失败"
    exit 1
fi