#!/usr/bin/env python3
"""
Excel订单生成器 - 跨平台GUI应用 (PySide6版本)
功能：
1. 输入Excel名称
2. 输入销售日期总条数
3. 输入销售单号总条数
4. 自动生成销售日期（当月1号到今日，9:00-21:00）
5. 自动生成销售单号（S+YYMMDD+14位1开头随机数）
6. 错误日志显示
7. 输入校验
"""

import sys
import os
import random
import calendar
from datetime import datetime, timedelta

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTextEdit, QMessageBox, QFileDialog
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QPixmap

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter


class ExcelGeneratorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel订单生成器")
        self.setFixedSize(520, 830)
        self.selected_file_path = ""
        self._build_ui()

    def _resource_path(self, filename):
        if getattr(sys, 'frozen', False):
            if sys.platform == 'darwin' and 'Contents/MacOS' in sys.executable:
                base = os.path.dirname(os.path.dirname(os.path.dirname(sys.executable)))
                return os.path.join(base, 'Resources', filename)
            else:
                return os.path.join(os.path.dirname(sys.executable), filename)
        else:
            return os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(20, 15, 20, 15)
        layout.setSpacing(4)

        # 封面图片
        cover_label = QLabel()
        img_path = self._resource_path('jj_header_rounded.png')
        pixmap = QPixmap(img_path)
        if not pixmap.isNull():
            scaled = pixmap.scaled(480, 200, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            cover_label.setPixmap(scaled)
            cover_label.setAlignment(Qt.AlignCenter)
            layout.addWidget(cover_label)
            layout.addSpacing(8)

        # Excel名称
        layout.addWidget(QLabel("Excel名称:"))
        self.excel_name_input = QLineEdit("订单表")
        self.excel_name_input.textChanged.connect(self._on_excel_name_changed)
        layout.addWidget(self.excel_name_input)
        layout.addSpacing(8)

        # 选择已有Excel文件
        layout.addWidget(QLabel("选择已有Excel文件:"))
        file_select_layout = QHBoxLayout()
        self.file_path_display = QLineEdit()
        self.file_path_display.setReadOnly(True)
        self.file_path_display.setStyleSheet("background-color: #f0f0f0; color: #555;")
        self.file_path_display.setPlaceholderText("未选择文件")
        file_select_layout.addWidget(self.file_path_display)
        self.file_select_btn = QPushButton("选择文件")
        self.file_select_btn.setFixedWidth(80)
        self.file_select_btn.clicked.connect(self._on_select_file)
        file_select_layout.addWidget(self.file_select_btn)
        self.file_clear_btn = QPushButton("清除")
        self.file_clear_btn.setFixedWidth(60)
        self.file_clear_btn.setStyleSheet("QPushButton { background-color: #e0e0e0; color: #666; border-radius: 3px; } QPushButton:hover { background-color: #d0d0d0; }")
        self.file_clear_btn.clicked.connect(self._on_clear_file)
        self.file_clear_btn.setVisible(False)
        file_select_layout.addWidget(self.file_clear_btn)
        layout.addLayout(file_select_layout)
        layout.addSpacing(8)

        # 月份
        layout.addWidget(QLabel("月份(1-12):"))
        month_layout = QHBoxLayout()
        self.month_input = QLineEdit()
        self.month_input.setFixedWidth(60)
        self.month_input.setMaxLength(2)
        now_month = str(datetime.now().month)
        self.month_input.setText(now_month)
        month_layout.addWidget(self.month_input)
        month_layout.addWidget(QLabel(f"（当前月份: {now_month})"))
        month_layout.addStretch()
        layout.addLayout(month_layout)
        layout.addSpacing(8)

        # 销售日期总条数
        layout.addWidget(QLabel("销售日期总条数:"))
        self.date_count_input = QLineEdit()
        layout.addWidget(self.date_count_input)
        layout.addSpacing(8)

        # 销售单号总条数
        layout.addWidget(QLabel("销售单号总条数:"))
        self.order_count_input = QLineEdit()
        layout.addWidget(self.order_count_input)
        layout.addSpacing(12)

        # 开始生成按钮
        self.start_btn = QPushButton("开始生成")
        self.start_btn.setFixedHeight(40)
        self.start_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50; color: white;
                font-size: 14px; font-weight: bold; border-radius: 5px;
            }
            QPushButton:hover { background-color: #45a049; }
            QPushButton:pressed { background-color: #3d8b40; }
        """)
        self.start_btn.clicked.connect(self.start_generation)
        layout.addWidget(self.start_btn)
        layout.addSpacing(12)

        # 报错日志
        layout.addWidget(QLabel("报错日志:"))
        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setStyleSheet("background-color: #f5f5f5; font-size: 12px;")
        layout.addWidget(self.log_area)

    def _on_select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择Excel文件", "", "Excel文件 (*.xlsx)"
        )
        if file_path:
            self.selected_file_path = file_path
            self.file_path_display.setText(file_path)
            self.excel_name_input.setEnabled(False)
            self.excel_name_input.clear()
            self.excel_name_input.setStyleSheet("background-color: #e8e8e8; color: #999;")
            self.file_clear_btn.setVisible(True)

    def _on_clear_file(self):
        self.selected_file_path = ""
        self.file_path_display.clear()
        self.excel_name_input.setEnabled(True)
        self.excel_name_input.setText("订单表")
        self.excel_name_input.setStyleSheet("")
        self.file_clear_btn.setVisible(False)

    def _on_excel_name_changed(self, text):
        pass

    def log_message(self, message, is_error=False):
        prefix = "[错误] " if is_error else "[信息] "
        timestamp = datetime.now().strftime("%H:%M:%S")
        color = "red" if is_error else "gray"
        self.log_area.append(
            f'<span style="color:{color}">[{timestamp}]</span> '
            f'<span style="color:{color if is_error else "black"}">{prefix}{message}</span>'
        )

    def generate_random_datetime(self, month):
        now = datetime.now()
        year = now.year
        max_day = calendar.monthrange(year, month)[1]
        # 如果是当前月份，日期不超过今天
        if month == now.month:
            max_day = now.day
        day = random.randint(1, max_day)
        base = datetime(year, month, day, 9, 0, 0)
        random_seconds = random.randint(0, 43200)
        result = base + timedelta(seconds=random_seconds)
        return result.strftime("%Y-%m-%d %H:%M:%S")

    def generate_order_number(self, month):
        now = datetime.now()
        year = now.year
        max_day = calendar.monthrange(year, month)[1]
        if month == now.month:
            max_day = now.day
        day = random.randint(1, max_day)
        random_date = datetime(year, month, day)
        yymmdd = random_date.strftime("%y%m%d")
        random_part = "1" + "".join([str(random.randint(0, 9)) for _ in range(13)])
        return f"S{yymmdd}{random_part}"

    def generate_order_number_from_date(self, date_str):
        dt = datetime.strptime(date_str, "%Y-%m-%d %H:%M:%S")
        yymmdd = dt.strftime("%y%m%d")
        random_part = "1" + "".join([str(random.randint(0, 9)) for _ in range(13)])
        return f"S{yymmdd}{random_part}"

    def find_column_by_header(self, ws, header_name):
        for col in range(1, ws.max_column + 2):
            if ws.cell(row=1, column=col).value == header_name:
                return col
        return None

    def find_first_empty_column(self, ws):
        for col in range(1, ws.max_column + 2):
            val = ws.cell(row=1, column=col).value
            if val is None or val == "":
                return col
        return ws.max_column + 1

    def get_last_data_row(self, ws, col):
        for row in range(ws.max_row, 0, -1):
            val = ws.cell(row=row, column=col).value
            if val is not None and val != "":
                return row
        return 0

    def _sort_rows_by_date(self, ws, date_col, order_col=None):
        last_row = ws.max_row
        if last_row <= 1:
            return
        max_col = ws.max_column

        # 收集所有数据行的值和对齐方式
        rows_data = []
        for row in range(2, last_row + 1):
            row_data = []
            for col in range(1, max_col + 1):
                cell = ws.cell(row=row, column=col)
                row_data.append((cell.value, cell.alignment))
            rows_data.append(row_data)

        # 按销售日期降序排序，无日期的行排在末尾
        date_idx = date_col - 1
        rows_data.sort(
            key=lambda r: (1, r[date_idx][0] or '') if r[date_idx][0] else (0, ''),
            reverse=True
        )

        # 写回排序后的数据
        for i, row_data in enumerate(rows_data):
            for col_idx, (value, orig_alignment) in enumerate(row_data):
                col = col_idx + 1
                cell = ws.cell(row=i + 2, column=col)
                cell.value = value
                if col == date_col or col == order_col:
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                else:
                    cell.alignment = orig_alignment

    def start_generation(self):
        # === 校验Excel名称 ===
        excel_name = self.excel_name_input.text().strip()
        if not excel_name and not self.selected_file_path:
            QMessageBox.warning(self, "提示", "请填写Excel名称或选择已有Excel文件")
            self.log_message("Excel名称为空且未选择文件", is_error=True)
            return

        # === 校验月份 ===
        month_str = self.month_input.text().strip()
        if not month_str:
            QMessageBox.warning(self, "提示", "请填写月份(1-12)")
            self.log_message("月份为空", is_error=True)
            return
        try:
            month = int(month_str)
            if month < 1 or month > 12:
                QMessageBox.warning(self, "提示", "月份必须在1-12之间")
                self.log_message("月份不在1-12范围", is_error=True)
                return
        except ValueError:
            QMessageBox.warning(self, "提示", "月份必须为整数")
            self.log_message("月份格式错误", is_error=True)
            return

        # === 校验销售日期总条数 ===
        date_count_str = self.date_count_input.text().strip()
        if date_count_str:
            try:
                date_count = int(date_count_str)
                if date_count <= 0:
                    QMessageBox.warning(self, "提示", "销售日期总条数必须大于0")
                    self.log_message("销售日期总条数必须大于0", is_error=True)
                    return
            except ValueError:
                QMessageBox.warning(self, "提示", "销售日期总条数必须为整数")
                self.log_message("销售日期总条数格式错误", is_error=True)
                return
        else:
            date_count = 0

        # === 校验销售单号总条数 ===
        order_count_str = self.order_count_input.text().strip()
        if order_count_str:
            try:
                order_count = int(order_count_str)
                if order_count <= 0:
                    QMessageBox.warning(self, "提示", "销售单号总条数必须大于0")
                    self.log_message("销售单号总条数必须大于0", is_error=True)
                    return
            except ValueError:
                QMessageBox.warning(self, "提示", "销售单号总条数必须为整数")
                self.log_message("销售单号总条数格式错误", is_error=True)
                return
        else:
            order_count = 0

        # 至少一个总条数大于0
        if date_count <= 0 and order_count <= 0:
            QMessageBox.warning(self, "提示", "销售日期和销售单号总条数至少需要一个大于0")
            self.log_message("两个总条数均为0", is_error=True)
            return

        # 确定文件路径
        if self.selected_file_path:
            file_path = self.selected_file_path
        else:
            # 确保文件名以.xlsx结尾
            if not excel_name.endswith(".xlsx"):
                excel_name += ".xlsx"
            # 定保存路径：放到软件exe/.app所在目录
            if getattr(sys, 'frozen', False):
                exe_dir = os.path.dirname(sys.executable)
                if sys.platform == 'darwin' and 'Contents/MacOS' in exe_dir:
                    base_dir = os.path.dirname(os.path.dirname(os.path.dirname(exe_dir)))
                else:
                    base_dir = exe_dir
            else:
                base_dir = os.path.dirname(os.path.abspath(__file__))
            file_path = os.path.join(base_dir, excel_name)

        try:
            # === 加载或创建Excel ===
            if os.path.exists(file_path):
                wb = load_workbook(file_path)
                ws = wb.active
                self.log_message(f"加载已有文件: {file_path}")
            else:
                wb = Workbook()
                ws = wb.active
                ws.title = "订单数据"
                self.log_message(f"创建新文件: {file_path}")

            # === 处理销售日期列 ===
            date_col = None
            if date_count > 0:
                date_col = self.find_column_by_header(ws, "销售日期")
                if date_col is not None:
                    last_row = self.get_last_data_row(ws, date_col)
                    date_start_row = last_row + 1
                    self.log_message(f"销售日期列已存在(列{date_col})，追加{date_count}条数据")
                else:
                    date_col = self.find_first_empty_column(ws)
                    ws.cell(row=1, column=date_col, value="销售日期")
                    date_start_row = 2
                    self.log_message(f"在列{date_col}新建销售日期标题")

                # 生成日期并按时间降序排序（从近到远）
                dates = [self.generate_random_datetime(month) for _ in range(date_count)]
                dates.sort(key=lambda x: datetime.strptime(x, "%Y-%m-%d %H:%M:%S"), reverse=True)

                for i, dt in enumerate(dates):
                    cell = ws.cell(row=date_start_row + i, column=date_col, value=dt)
                    cell.alignment = Alignment(horizontal='center', vertical='center')

                self.log_message(
                    f"写入{date_count}条销售日期(行{date_start_row}-{date_start_row + date_count - 1})")

            # 调整列宽
            if date_col is not None:
                ws.column_dimensions[get_column_letter(date_col)].width = 22
            if order_col is not None:
                ws.column_dimensions[get_column_letter(order_col)].width = 25

            # 按销售日期重新排序所有数据（在生成销售单号之前排序，确保单号与同行日期对应）
            if date_col is not None:
                self._sort_rows_by_date(ws, date_col, order_col)
                self.log_message("已按销售日期降序重新排序所有数据")

            # === 处理销售单号列（在排序后生成，确保单号年月日与同行日期对应）===
            order_col = None
            if order_count > 0:
                order_col = self.find_column_by_header(ws, "销售单号")
                if order_col is not None:
                    last_row = self.get_last_data_row(ws, order_col)
                    order_start_row = last_row + 1
                    self.log_message(f"销售单号列已存在(列{order_col})，追加{order_count}条数据")
                else:
                    if date_col is not None:
                        order_col = date_col + 1
                    else:
                        order_col = self.find_first_empty_column(ws)
                    next_val = ws.cell(row=1, column=order_col).value
                    if next_val is not None and next_val != "":
                        ws.insert_cols(order_col)
                        self.log_message(f"在列{order_col}插入新列(原有数据右移)")
                    ws.cell(row=1, column=order_col, value="销售单号")
                    order_start_row = 2
                    self.log_message(f"在列{order_col}新建销售单号标题")

                for i in range(order_count):
                    row = order_start_row + i
                    if date_col is not None:
                        date_value = ws.cell(row=row, column=date_col).value
                        if date_value:
                            order_num = self.generate_order_number_from_date(date_value)
                        else:
                            order_num = self.generate_order_number(month)
                    else:
                        order_num = self.generate_order_number(month)
                    cell = ws.cell(row=row, column=order_col, value=order_num)
                    cell.alignment = Alignment(horizontal='center', vertical='center')

                self.log_message(
                    f"写入{order_count}条销售单号(行{order_start_row}-{order_start_row + order_count - 1})")
                ws.column_dimensions[get_column_letter(order_col)].width = 25

            wb.save(file_path)
            self.log_message(f"文件保存成功: {file_path}")
            QMessageBox.information(self, "成功", f"Excel文件已生成!\n路径: {file_path}")

        except PermissionError:
            msg = "文件保存失败: 可能被其他程序占用，请关闭Excel后重试"
            self.log_message(f"{msg}\n路径: {file_path}", is_error=True)
            QMessageBox.critical(self, "错误", msg)
        except Exception as e:
            msg = f"生成失败: {str(e)}"
            self.log_message(f"{msg}\n{traceback.format_exc()}", is_error=True)
            QMessageBox.critical(self, "错误", msg)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelGeneratorApp()
    window.show()
    sys.exit(app.exec())