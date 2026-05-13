---
name: excel-generator
description: 跨平台Excel订单生成器GUI应用（PySide6框架），支持Mac(.dmg)和Windows(.exe)打包
type: project
---

# Excel订单生成器

跨平台GUI应用（PySide6/Qt框架），自动生成包含销售日期和销售单号的Excel文件，支持新建和追加已有文件。

## 功能列表

### 1. 新建Excel文件
- 在"Excel名称"输入框填写文件名（默认"订单表"），点击"开始生成"后在软件所在目录创建 `.xlsx` 文件
- 文件名自动追加 `.xlsx` 后缀
- 新文件会创建"销售日期"、"销售单号"和"Pos单号"三个列标题，并在下方写入对应数据

### 2. 选择已有Excel文件追加数据
- 点击"选择文件"按钮，打开文件选择对话框（仅允许选择 `.xlsx` 文件，非Excel文件不可选中）
- 选中文件后，路径显示在输入框中，"Excel名称"输入框自动灰显禁用（不可编辑，明确不生效）
- 点击"清除"按钮可取消文件选择，恢复"Excel名称"输入框为可用状态
- 选择已有文件时：
  - 若Excel中已有"销售日期"列 → 不重复添加标题，数据追加到该列末尾
  - 若Excel中已有"销售单号"列 → 不重复添加标题，数据追加到该列末尾
  - 若Excel中已有"Pos单号"列 → 不重复添加标题，数据追加到该列末尾
  - 若列不存在 → 新建列并添加标题
  - 追加数据后自动按销售日期降序重新排序所有数据行（最近的日期排在最前）

### 3. 数据生成规则
- **销售日期**：格式 `YYYY-MM-DD HH:MM:SS`
  - 日期范围：当月1号到今日
  - 时间范围：9:00-21:00（随机）
  - 用户填写的总条数即生成条数，追加到已有数据尾部
  - 生成后按时间降序重新排序（从近到远）
- **销售单号**：格式 `S+YYMMDD+14位1开头随机数`，如 `S26050812345678901234`
  - YYMMDD取自同行销售日期，排序后重新生成覆盖
  - 生成分2种情况：
    - **情况1（无已有单号列）**：实际生成数 = min(用户输入数, 销售日期总长度)
    - **情况2（有已有单号列）**：a = 已有单号数 + 用户输入数
      - a >= 日期总长度 → 从上到下重新生成日期总长度个单号
      - a < 日期总长度 → 从上到下重新生成a个单号
- **Pos单号**：格式 `OS+YYYYMMDD-5位随机数`，如 `OS20260508-12345`
  - YYYYMMDD取自同行销售日期，排序后重新生成覆盖
  - 生成分2种情况：
    - **情况1（无已有Pos单号列）**：实际生成数 = min(用户输入数, 销售日期总长度)
    - **情况2（有已有Pos单号列）**：a = 已有Pos单号数 + 用户输入数
      - a >= 日期总长度 → 从上到下重新生成日期总长度个Pos单号
      - a < 日期总长度 → 从上到下重新生成a个Pos单号
  - 生成顺序：先生成销售日期 → 再生成销售单号 → 最后生成Pos单号

### 4. 数据居中对齐
- 销售日期、销售单号和Pos单号列的内容（标题下方数据）均设置为水平+垂直居中对齐

### 5. 输入校验
- Excel名称和选择文件互斥验证：两者至少需有一个，否则提示"请填写Excel名称或选择已有Excel文件"
- 选择文件后，Excel名称输入框的空值校验自动跳过
- 月份必须在1-12之间，必须为整数
- 销售日期、销售单号和Pos单号总条数至少需要一个大于0
- 条数必须为正整数
- 仅生成单号时，文件中必须已有销售日期数据

### 6. 操作日志
- 操作信息和错误日志在底部日志区域实时显示，错误信息红色标注，普通信息黑色标注
- 文件保存失败时提示"可能被其他程序占用，请关闭Excel后重试"

### 7. 封面图片
- 应用窗口顶部展示林俊杰封面图片（圆角处理），等比缩放显示

## 使用方法

### 界面操作流程

**方式一：新建Excel**
1. 在"Excel名称"输入框填写文件名（如"订单表"）
2. 在"月份"输入框填写月份（默认当前月份）
3. 填写"销售日期总条数"和/或"销售单号总条数"和/或"Pos单号总条数"
4. 点击"开始生成"

**方式二：追加已有Excel**
1. 点击"选择文件"按钮，选择一个 `.xlsx` 文件
2. 在"月份"输入框填写月份
3. 填写需要追加的条数
4. 点击"开始生成"，数据追加后自动按销售日期排序

### 文件保存位置
- 新建模式：Excel文件保存到用户桌面目录
- 选择文件模式：直接保存到选中的文件路径

## 技术架构

GUI框架：**PySide6**（Qt for Python）

**Why:** macOS Monterey 12.7.x 系统自带 Tcl/Tk 有版本检查bug导致 tkinter 闪退，PySide6/Qt 不受此影响

**How to apply:** 所有GUI代码使用 PySide6 的 QtWidgets 模块

## 文件结构

| 文件 | 说明 |
|------|------|
| `app.py` | 主应用代码（PySide6 GUI + 数据生成逻辑） |
| `jj_header.jpeg` | 封面原图（林俊杰照片） |
| `jj_header_rounded.png` | 封面圆角版（四角40px圆角，PNG透明） |
| `header.icns` | Mac应用图标文件 |
| `header.ico` | Windows应用图标文件 |
| `build_mac.sh` | Mac .dmg 构建脚本 |
| `build_windows.bat` | Windows .exe 构建脚本（需在Windows上运行） |

## 构建方法

### Mac .dmg
```bash
bash build_mac.sh
```
产物：`ExcelGenerator.dmg`、`dist/ExcelGenerator.app`

**重要:** Mac构建必须使用 Homebrew Python 3.11+，不能用系统 Python 3.8

### Windows .exe
将项目文件复制到Windows电脑，运行：
```bat
build_windows.bat
```
产物：`dist\ExcelGenerator\ExcelGenerator.exe`（目录型打包，需整个文件夹分发）

### 构建配置（Mac和Windows一致）
- `--windowed`（目录型打包，无控制台窗口）
- `--icon=header.icns` / `--icon=header.ico`（应用图标）
- `--add-data="jj_header_rounded.png"`（封面图片打包进应用）
- `--hidden-import=pypinyin,openpyxl,openpyxl.cell._writer`
- `--collect-data=pypinyin`
- `--noconfirm`

## 运行方式

### 开发模式
```bash
python3.11 app.py
```

### 打包模式
- Mac：双击 `ExcelGenerator.dmg` → 拖拽 `ExcelGenerator.app` 到应用程序文件夹 → 双击运行
- Windows：将 `ExcelGenerator` 文件夹复制到目标电脑 → 双击 `ExcelGenerator.exe`

## 依赖

- `PySide6`：Qt GUI框架
- `openpyxl`：Excel读写
- `pypinyin`：中文转拼音
- `PyInstaller`：打包为独立应用