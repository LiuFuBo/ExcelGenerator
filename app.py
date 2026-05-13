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
import traceback
from datetime import datetime, timedelta

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTextEdit, QMessageBox, QFileDialog
)
from PySide6.QtCore import Qt

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter


class ExcelGeneratorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel订单生成器")
        self.setFixedSize(520, 700)
        self.selected_file_path = ""
        self._build_ui()

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(20, 15, 20, 15)
        layout.setSpacing(4)

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
        layout.addSpacing(8)

        # Pos单号总条数
        layout.addWidget(QLabel("Pos单号总条数:"))
        self.pos_count_input = QLineEdit()
        layout.addWidget(self.pos_count_input)
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

        # 操作日志
        log_header = QHBoxLayout()
        log_header.addWidget(QLabel("操作日志:"))
        log_header.addStretch()
        self.copy_log_btn = QPushButton("复制日志")
        self.copy_log_btn.setFixedWidth(80)
        self.copy_log_btn.setStyleSheet("QPushButton { background-color: #e0e0e0; color: #666; border-radius: 3px; } QPushButton:hover { background-color: #d0d0d0; }")
        self.copy_log_btn.clicked.connect(self._on_copy_log)
        log_header.addWidget(self.copy_log_btn)
        layout.addLayout(log_header)
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

    def _on_copy_log(self):
        text = self.log_area.toPlainText()
        if text:
            clipboard = QApplication.clipboard()
            clipboard.setText(text)
            QMessageBox.information(self, "提示", "复制日志成功")

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

    def generate_pos_number_from_date(self, date_str):
        dt = datetime.strptime(date_str, "%Y-%m-%d %H:%M:%S")
        yyyymmdd = dt.strftime("%Y%m%d")
        random_part = "".join([str(random.randint(0, 9)) for _ in range(5)])
        return f"OS{yyyymmdd}-{random_part}"

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

    def count_data_cells(self, ws, col):
        """计算某列实际非空数据条数（排除第1行标题）"""
        count = 0
        for row in range(2, ws.max_row + 1):
            val = ws.cell(row=row, column=col).value
            if val is not None and val != "":
                count += 1
        return count

    def _sort_date_column(self, ws, date_col):
        """仅排序销售日期列的数据，其他列保持不动"""
        last_row = ws.max_row
        if last_row <= 1:
            return

        dates = []
        for row in range(2, last_row + 1):
            dates.append(ws.cell(row=row, column=date_col).value)

        dates.sort(
            key=lambda d: (1, d or '') if d else (0, ''),
            reverse=True
        )

        for i, date_val in enumerate(dates):
            row = i + 2
            if date_val is not None:
                cell = ws.cell(row=row, column=date_col, value=date_val)
                cell.alignment = Alignment(horizontal='center', vertical='center')
            else:
                ws.cell(row=row, column=date_col).value = None

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

        # === 校验Pos单号总条数 ===
        pos_count_str = self.pos_count_input.text().strip()
        if pos_count_str:
            try:
                pos_count = int(pos_count_str)
                if pos_count <= 0:
                    QMessageBox.warning(self, "提示", "Pos单号总条数必须大于0")
                    self.log_message("Pos单号总条数必须大于0", is_error=True)
                    return
            except ValueError:
                QMessageBox.warning(self, "提示", "Pos单号总条数必须为整数")
                self.log_message("Pos单号总条数格式错误", is_error=True)
                return
        else:
            pos_count = 0

        # 至少一个总条数大于0
        if date_count <= 0 and order_count <= 0 and pos_count <= 0:
            QMessageBox.warning(self, "提示", "销售日期、销售单号和Pos单号总条数至少需要一个大于0")
            self.log_message("三个总条数均为0", is_error=True)
            return

        # 确定文件路径
        if self.selected_file_path:
            file_path = self.selected_file_path
        else:
            # 确保文件名以.xlsx结尾
            if not excel_name.endswith(".xlsx"):
                excel_name += ".xlsx"
            # 新文件保存到桌面
            base_dir = os.path.expanduser("~/Desktop")
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

                dates = [self.generate_random_datetime(month) for _ in range(date_count)]
                for i, dt in enumerate(dates):
                    cell = ws.cell(row=date_start_row + i, column=date_col, value=dt)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                self.log_message(f"写入{date_count}条销售日期")

            # 如果只生成单号（没有新日期），必须找到已有日期列
            if date_count <= 0 and order_count > 0:
                date_col = self.find_column_by_header(ws, "销售日期")
                if date_col is None:
                    QMessageBox.warning(self, "提示", "无法生成销售单号：文件中没有销售日期数据")
                    self.log_message("没有销售日期列，无法生成单号", is_error=True)
                    return
                self.log_message(f"使用已有销售日期列(列{date_col})")

            # === 排序前先统计已有单号条数（排序会移动位置，需提前计数）===
            existing_order_col = self.find_column_by_header(ws, "销售单号")
            existing_order_length = 0
            if existing_order_col is not None:
                existing_order_length = self.count_data_cells(ws, existing_order_col)

            # === 仅排序销售日期列（不排单号列）===
            if date_col is not None:
                self._sort_date_column(ws, date_col)
                self.log_message("已按销售日期降序排序（仅日期列）")

            # === 计算总日期长度 ===
            total_date_length = 0
            if date_col is not None:
                total_date_length = self.count_data_cells(ws, date_col)
                self.log_message(f"销售日期总长度: {total_date_length}条")

            # === 处理销售单号列 ===
            order_col = None
            actual_order_count = 0

            if order_count > 0:
                if existing_order_col is not None:
                    # 情况2：已有销售单号列，a = 旧单号长度 + 用户输入数
                    order_col = existing_order_col
                    # 清空旧单号（排序后旧单号与日期不再对应，全部清空再重新生成）
                    for row in range(2, ws.max_row + 1):
                        ws.cell(row=row, column=order_col).value = None

                    a = existing_order_length + order_count
                    if a >= total_date_length:
                        actual_order_count = total_date_length
                    else:
                        actual_order_count = a

                    self.log_message(
                        f"已有单号{existing_order_length}条+新增{order_count}={a}，"
                        f"实际重新生成{actual_order_count}条单号"
                    )
                else:
                    # 情况1：没有销售单号列，新建列
                    order_col = date_col + 1 if date_col else self.find_first_empty_column(ws)
                    next_val = ws.cell(row=1, column=order_col).value
                    if next_val is not None and next_val != "":
                        ws.insert_cols(order_col)
                        self.log_message(f"在列{order_col}插入新列(原有数据右移)")
                    ws.cell(row=1, column=order_col, value="销售单号")

                    actual_order_count = min(order_count, total_date_length)
                    self.log_message(
                        f"新建单号列(列{order_col})，生成{actual_order_count}条"
                        f"(请求{order_count}条，日期长度{total_date_length}条)"
                    )

                # 从上到下重新生成销售订单编号（根据同行销售日期覆盖）
                for i in range(actual_order_count):
                    row = 2 + i
                    date_value = ws.cell(row=row, column=date_col).value if date_col else None
                    if date_value:
                        order_num = self.generate_order_number_from_date(date_value)
                    else:
                        order_num = self.generate_order_number(month)
                    cell = ws.cell(row=row, column=order_col, value=order_num)
                    cell.alignment = Alignment(horizontal='center', vertical='center')

                self.log_message(f"已从上到下重新生成{actual_order_count}条销售单号，单号日期与同行日期一致")

            # === 处理Pos单号列 ===
            pos_col = None
            actual_pos_count = 0

            if pos_count > 0:
                existing_pos_col = self.find_column_by_header(ws, "Pos单号")
                existing_pos_length = 0
                if existing_pos_col is not None:
                    existing_pos_length = self.count_data_cells(ws, existing_pos_col)

                if existing_pos_col is not None:
                    # 情况2：已有Pos单号列，a = 旧长度 + 用户输入数
                    pos_col = existing_pos_col
                    # 清空旧Pos单号，全部重新生成
                    for row in range(2, ws.max_row + 1):
                        ws.cell(row=row, column=pos_col).value = None

                    a = existing_pos_length + pos_count
                    if a >= total_date_length:
                        actual_pos_count = total_date_length
                    else:
                        actual_pos_count = a

                    self.log_message(
                        f"已有Pos单号{existing_pos_length}条+新增{pos_count}={a}，"
                        f"实际重新生成{actual_pos_count}条Pos单号"
                    )
                else:
                    # 情况1：没有Pos单号列，新建列
                    pos_col = (order_col or date_col) + 1 if (order_col or date_col) else self.find_first_empty_column(ws)
                    next_val = ws.cell(row=1, column=pos_col).value
                    if next_val is not None and next_val != "":
                        ws.insert_cols(pos_col)
                        self.log_message(f"在列{pos_col}插入新列(原有数据右移)")
                    ws.cell(row=1, column=pos_col, value="Pos单号")

                    actual_pos_count = min(pos_count, total_date_length)
                    self.log_message(
                        f"新建Pos单号列(列{pos_col})，生成{actual_pos_count}条"
                        f"(请求{pos_count}条，日期长度{total_date_length}条)"
                    )

                # 从上到下根据同行销售日期生成Pos单号
                for i in range(actual_pos_count):
                    row = 2 + i
                    date_value = ws.cell(row=row, column=date_col).value if date_col else None
                    if date_value:
                        pos_num = self.generate_pos_number_from_date(date_value)
                    else:
                        pos_num = self.generate_pos_number_from_date(
                            self.generate_random_datetime(month)
                        )
                    cell = ws.cell(row=row, column=pos_col, value=pos_num)
                    cell.alignment = Alignment(horizontal='center', vertical='center')

                self.log_message(f"已从上到下重新生成{actual_pos_count}条Pos单号，Pos单号日期与同行日期一致")

            # 调整列宽
            if date_col is not None:
                ws.column_dimensions[get_column_letter(date_col)].width = 22
            if order_col is not None:
                ws.column_dimensions[get_column_letter(order_col)].width = 25
            if pos_col is not None:
                ws.column_dimensions[get_column_letter(pos_col)].width = 22

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