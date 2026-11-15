from PySide6.QtWidgets import (
    QApplication, QMainWindow, QLabel, QWidget,
    QVBoxLayout, QHBoxLayout, QGridLayout,
    QPushButton, QCheckBox, QInputDialog, QLineEdit, QComboBox,
    QTableWidget, QTableWidgetItem, QCompleter, QMessageBox,
    QHeaderView, QStyle, QTabWidget, QFileDialog,
    QGraphicsView, QGraphicsScene, QGraphicsPixmapItem,
    QDateEdit, QCalendarWidget, QAbstractItemView, QTextEdit,
    QSpacerItem, QSizePolicy, QFrame, QScrollArea, QGraphicsDropShadowEffect,
    QSplitter, QGroupBox,
)
import sys
import sqlite3
import os
import pandas as pd
from PySide6.QtGui import QIcon, QFont, QPixmap, QPainter, QColor, QIntValidator, QImage, QShortcut, QKeySequence
from PySide6.QtCore import Qt, QEvent, QStandardPaths, QDate, QPoint, QPropertyAnimation, QTimer
from datetime import date, datetime
from pathlib import Path
from docx import Document
from tkinter import filedialog, messagebox
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from data_base import db_path_default, backup_db
from data_base import (add_staff, search_staff, modify_staff, delete_staff, export_staff_to_excel,
                       import_staff_from_excel)
from data_base import (add_assets, search_assets, modify_assets, remove_assets, export_assets_to_excel,
                       import_assets_from_excel)
from data_base import (assign_assets, return_assets, search_assets_of_staff, search_serial_number_in_warehouse,
                       search_serial_number_of_staff, export_assets_assignments_to_excel,
                       import_assets_assignments_from_excel,export_assets_of_staffs_to_excel,
                       export_two_tables_to_excel, import_two_tables_from_excel_to_widgets,add_table_from_qtable,
                       add_sign_table, fill_date_placeholders, fill_sign_placeholder_in_tables)
from data_base import (search_asset_assignments_id_and_handover_documents_id, add_handover_documents,
                       add_handover_document_to_assignment)



class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Chương trình quản lý tài sản")
        self.setWindowIcon(QIcon("avatar_app.png"))
        self.setGeometry(0, 0, 400, 420)
        self.setStyleSheet("""
            QMainWindow {
                background: white;
            }

            QPushButton {
                background: qlineargradient(
                    spread:pad, x1:0, y1:0, x2:1, y2:1,
                    stop:0 #016b67, stop:1 #0b9f87, stop:2 #00c3cc
                );
                color: white;
                border: none;
                border-radius: 6px;
                padding: 6px;
                font-weight: bold;
                font-size: 12px;
            }
            QPushButton:hover {
                background-color: #016b67; /* nền khi hover */
                color: white;              /* chữ khi hover */
            }
            QLabel {
                color: black;
                font-weight: bold;
            }

            /* ComboBox, LineEdit, TextEdit viền màu #016b67 */
            QComboBox, QLineEdit, QTextEdit, QTableWidget {
                background-color: white;        /* nền trắng bên trong */
                border: 1px solid #016b67;      /* viền xanh */
                border-radius: 4px;             /* bo tròn */
                color: black;                   /* chữ đen */
            }
        """)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        grid = QGridLayout(central_widget)
        grid.setContentsMargins(30, 20, 30, 16)
        grid.setHorizontalSpacing(20)
        grid.setVerticalSpacing(12)

        # Quản lý cán bộ
        self.button1 = QPushButton("Quản lý Cán bộ", self)
        self.button1.setFixedSize(160, 46)
        self.button1.clicked.connect(self.button1_clicked)
        self.staff_window = None
        grid.addWidget(self.button1, 0, 0)
        # Quản lý tài sản
        self.button2 = QPushButton("Quản lý Tài sản", self)
        self.button2.setFixedSize(160, 46)
        self.button2.clicked.connect(self.button2_clicked)
        self.assets_window = None
        grid.addWidget(self.button2, 0, 1)
        # Bàn giao tài sản
        self.button3 = QPushButton("Bàn giao tài sản", self)
        self.button3.setFixedSize(160, 46)
        self.button3.clicked.connect(self.button3_clicked)
        self.asset_assignments_window = None
        grid.addWidget(self.button3, 1, 0)
        # Biên bản bàn giao
        self.button4 = QPushButton("Biên bản bàn giao", self)
        self.button4.setFixedSize(160, 46)
        self.button4.clicked.connect(self.button4_clicked)
        self.handover_documents_window = None
        grid.addWidget(self.button4, 1, 1)
        # Hướng dẫn sử dụng và ghi chú
        self.button5 = QPushButton("Hướng dẫn sử dụng", self)
        self.button5.setFixedSize(160, 46)
        self.button5.clicked.connect(self.button5_clicked)
        self.user_manual_window = None
        grid.addWidget(self.button5, 2, 0)
        # Back up
        self.button7 = QPushButton("Back up DB", self)
        self.button7.setFixedSize(160, 46)
        self.button7.clicked.connect(self.back_up)
        grid.addWidget(self.button7, 2, 1)
        # Kiểm kê tài sản (QR)
        self.button_inventory = QPushButton("📷 Kiểm kê QR", self)
        self.button_inventory.setFixedSize(160, 46)
        self.button_inventory.clicked.connect(self.button_inventory_clicked)
        self.inventory_qr_window = None
        grid.addWidget(self.button_inventory, 3, 0)
        # === Thêm logo ===
        self.logo_label = QLabel(self)
        self.logo_label.setFixedSize(90, 90)
        pixmap = QPixmap("avatar_app.png")  # đường dẫn logo
        self.logo_label.setPixmap(pixmap.scaled(
            90, 90,
            Qt.AspectRatioMode.KeepAspectRatio,
            Qt.TransformationMode.SmoothTransformation
        ))
        self.logo_label.setStyleSheet("background: transparent;")
        grid.addWidget(self.logo_label, 4, 0, alignment=Qt.AlignmentFlag.AlignCenter)
        # mark
        self.mark = QLabel("Design and develop by \nThangdm1", self)
        font = QFont("Roboto", 8, QFont.Weight.Bold)
        font.setItalic(True)  # chữ nghiêng
        self.mark.setFont(font)
        self.mark.setStyleSheet("color: #016b67;")
        self.mark.setWordWrap(True)
        self.mark.setFixedWidth(160)
        self.mark.setAlignment(Qt.AlignmentFlag.AlignCenter)
        grid.addWidget(self.mark, 4, 1, alignment=Qt.AlignmentFlag.AlignCenter)


        # # --- Khu vực chat ---
        # self.chat_display = QTextEdit(self)
        # self.chat_display.setGeometry(20, 20, 500, 350)
        # self.chat_display.setReadOnly(True)
        #
        # self.user_label = QLabel("Tên:", self)
        # self.user_label.setGeometry(20, 380, 40, 25)
        #
        # self.user_input  "Nguyễn Hương Ngọc",
        #         #                           "Nguyễn Thị Như Trang",
        #         #                           "Người khác"])
        #         # self.user_input.setCurrentIndex(-1)
        #         # = QComboBox(self)
        # self.user_input.setGeometry(60, 380, 180, 25)
        # self.user_input.addItems(["Đỗ Mạnh Thắng",
        #
        # self.message_input = QTextEdit(self)
        # self.message_input.setGeometry(20, 410, 300, 60)
        # self.message_input.setPlaceholderText("Nhập tin nhắn...")
        #
        # self.send_button = QPushButton("Gửi", self)
        # self.send_button.setGeometry(330, 410, 60, 25)
        # self.send_button.clicked.connect(self.send_message)
        #
        # self.load_messages()

    def button1_clicked(self):
        if self.staff_window is None:
            self.staff_window = staff_window()
        self.staff_window.show()

    def button2_clicked(self):
        if self.assets_window is None:
            self.assets_window = assets_window()
        self.assets_window.show()

    def button3_clicked(self):
        if self.asset_assignments_window is None:
            self.asset_assignments_window = asset_assignments_window()
        self.asset_assignments_window.show()

    def button4_clicked(self):
        if self.handover_documents_window is None:
            self.handover_documents_window = handover_documents_window()
        self.handover_documents_window.show()

    def button5_clicked(self):
        if self.user_manual_window is None:
            self.user_manual_window = user_manual_window()
        self.user_manual_window.show()

    def button_inventory_clicked(self):
        if self.inventory_qr_window is None:
            self.inventory_qr_window = inventory_qr_window()
        else:
            self.inventory_qr_window.reload_from_db(preserve_scanned=True)
        self.inventory_qr_window.show()
        self.inventory_qr_window.raise_()
        self.inventory_qr_window.activateWindow()

    def load_messages(self):
        """Load tin nhắn từ DB vào chat_display"""
        conn = sqlite3.connect(db_path_default)
        cursor = conn.cursor()
        cursor.execute("SELECT user_chat, chat, date_and_time FROM chat_box ORDER BY id ASC")
        rows = cursor.fetchall()
        conn.close()

        self.chat_display.clear()
        for user, chat, dt in rows:
            chat_html = chat.replace("\n", "<br>")  # xử lý xuống dòng
            html = f"<b>[{dt}] {user}:</b><br>{chat_html}<br><br>"
            self.chat_display.append(html)

    def send_message(self):
        """Gửi tin nhắn mới"""
        try:
            user = self.user_input.currentText().strip()
            message = self.message_input.toPlainText().strip()

            if not user or not message:
                return

            dt = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            conn = sqlite3.connect(db_path_default)
            cursor = conn.cursor()
            cursor.execute(
                "INSERT INTO chat_box (chat, date_and_time, user_chat) VALUES (?, ?, ?)",
                (message, dt, user)
            )
            conn.commit()
            conn.close()

            self.message_input.clear()  # vẫn xoá được QTextEdit
            self.load_messages()
        except Exception as e:
            messagebox.showerror("Lỗi", f"\n{str(e)}")
    def back_up(self):
        dest = backup_db()
        QMessageBox.information(self, "Thành công", f"Đã backup DB đến:\n{dest}")


    def closeEvent(self, event):
        """Đóng tất cả các cửa sổ con khi main window đóng"""
        windows = [
            self.staff_window,
            self.assets_window,
            self.asset_assignments_window,
            self.handover_documents_window,
            self.user_manual_window,
            self.inventory_qr_window,
        ]
        for w in windows:
            if w is not None:
                w.close()
        event.accept()


class staff_window(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Quản lý cán bộ")
        self.setGeometry(200, 200, 1080, 600)
        self.setWindowIcon(QIcon("avatar_app.png"))
        self.setStyleSheet("""
            /* === NỀN CỬA SỔ === */
            QWidget {
                background-color: #f0f0f0;
                color: black;
            }

            /* === NÚT BẤM === */
            QPushButton {
                background: qlineargradient(
                    spread:pad, x1:0, y1:0, x2:1, y2:1,
                    stop:0 #016b67, stop:1 #0b9f87, stop:2 #00c3cc
                );
                color: white;
                border: none;
                border-radius: 6px;
                padding: 6px;
            }
            QPushButton:hover {
                background-color: #016b67;
                color: white;
            }

            /* === NHÃN === */
            QLabel {
                color: black;
                font-weight: bold;
            }

            /* === Ô COMBOBOX / LINEEDIT / TEXTEDIT / TABLE === */
            QComboBox,
            QLineEdit,
            QTextEdit,
            QTableWidget {
                background-color: white;
                border: 1px solid #016b67;
                border-radius: 4px;
                color: black;
            }

            /* === SCROLLBAR VERTICAL (5px) === */
            QScrollBar:vertical {
                border: none;
                background: #f0f0f0;
                width: 5px;             /* mỏng 5px */
                margin: 0px;
                border-radius: 2px;
            }
            QScrollBar::handle:vertical {
                background-color: #c0c0c0;   /* xám nhạt */
                min-height: 20px;
                border-radius: 2px;
            }
            QScrollBar::handle:vertical:hover {
                background-color: #a0a0a0;   /* xám đậm hơn khi hover */
            }
            QScrollBar::add-line:vertical,
            QScrollBar::sub-line:vertical {
                height: 0px;
                background: none;
            }

            /* === SCROLLBAR HORIZONTAL (5px) === */
            QScrollBar:horizontal {
                border: none;
                background: #f0f0f0;
                height: 5px;            /* mỏng 5px */
                margin: 0px;
                border-radius: 2px;
            }
            QScrollBar::handle:horizontal {
                background-color: #c0c0c0;   /* xám nhạt */
                min-width: 20px;
                border-radius: 2px;
            }
            QScrollBar::handle:horizontal:hover {
                background-color: #a0a0a0;   /* xám đậm hơn khi hover */
            }
            QScrollBar::add-line:horizontal,
            QScrollBar::sub-line:horizontal {
                width: 0px;
                background: none;
            }
        """)


        grid = QGridLayout(self)

        # Phần thêm cán bộ
        self.label1 = QLabel("Họ và tên:", self)
        grid.addWidget(self.label1, 1, 0)
        self.lineEdit1 = QLineEdit(self)
        self.lineEdit1.setFixedWidth(150)
        grid.addWidget(self.lineEdit1, 1, 1)

        self.label2 = QLabel("Nhóm:", self)
        grid.addWidget(self.label2, 2, 0)
        self.comboBox1 = QComboBox(self)
        self.comboBox1.addItems(["NK1", "NK2", "NK3", "NK4", "LCND", "XK1", "XK2"])
        self.comboBox1.setFixedWidth(150)
        grid.addWidget(self.comboBox1, 2, 1)

        self.label3 = QLabel("Chức vụ:", self)
        grid.addWidget(self.label3, 3, 0)
        self.comboBox2 = QComboBox(self)
        self.comboBox2.addItems(["Chuyên viên", "Phó Phòng", "Trưởng Phòng", "Phó giám đốc", "Giám đốc"])
        self.comboBox2.setFixedWidth(150)
        grid.addWidget(self.comboBox2, 3, 1)

        self.label4 = QLabel("Email:", self)
        grid.addWidget(self.label4, 4, 0)
        self.lineEdit2 = QLineEdit(self)
        self.lineEdit2.setFixedWidth(150)
        grid.addWidget(self.lineEdit2, 4, 1)

        self.label5 = QLabel("Mã cán bộ", self)
        grid.addWidget(self.label5, 5, 0)
        self.lineEdit3 = QLineEdit(self)
        self.lineEdit3.setFixedWidth(150)
        grid.addWidget(self.lineEdit3, 5, 1)

        self.label6 = QLabel("Số điện thoại:", self)
        grid.addWidget(self.label6, 6, 0)
        self.lineEdit4 = QLineEdit(self)
        self.lineEdit4.setFixedWidth(150)
        grid.addWidget(self.lineEdit4, 6, 1)

        self.label7 = QLabel("Ghi chú:", self)
        grid.addWidget(self.label7, 7, 0)
        self.lineEdit5 = QLineEdit(self)
        self.lineEdit5.setFixedWidth(150)
        grid.addWidget(self.lineEdit5, 7, 1)

        self.label8 = QLabel("Đây là cửa sổ để quản lý danh sách cán bộ", self)
        self.label8.setFixedWidth(300)
        grid.addWidget(self.label8,1, 6)

        # self.label9 = QLabel("Thêm cán bộ:", self)
        # grid.addWidget(self.label9, 0, 0)

        self.button1 = QPushButton("Thêm cán bộ", self)
        self.button1.clicked.connect(self.button1_clicked)
        grid.addWidget(self.button1, 8, 1)

        # Phần chỉnh sửa thông tin cán bộ
        self.label10 = QLabel("Cán bộ cần sửa đổi thông tin :", self)
        grid.addWidget(self.label10, 0, 3)
        self.lineEdit6 = QLineEdit(self)
        self.lineEdit6.setFixedWidth(150)
        self.lineEdit6.setPlaceholderText("MCB, Email hoặc Tên")
        grid.addWidget(self.lineEdit6, 0, 4)
        self.update_completer()

        self.label12 = QLabel("Họ và tên:", self)
        grid.addWidget(self.label12, 1, 3)
        self.lineEdit7 = QLineEdit(self)
        self.lineEdit7.setFixedWidth(150)
        grid.addWidget(self.lineEdit7, 1, 4)

        self.label13 = QLabel("Nhóm:", self)
        grid.addWidget(self.label13, 2, 3)
        self.comboBox3 = QComboBox(self)
        self.comboBox3.addItems(["NK1", "NK2", "NK3", "NK4", "LCND", "XK1", "XK2"])
        self.comboBox3.setFixedWidth(150)
        grid.addWidget(self.comboBox3, 2, 4)

        self.label14 = QLabel("Chức vụ:", self)
        grid.addWidget(self.label14, 3, 3)
        self.comboBox4 = QComboBox(self)
        self.comboBox4.addItems(["Chuyên viên", "Phó Phòng", "Trưởng Phòng", "Phó giám đốc", "Giám đốc"])
        self.comboBox4.setFixedWidth(150)
        grid.addWidget(self.comboBox4, 3, 4)

        self.label15 = QLabel("Email:", self)
        grid.addWidget(self.label15, 4, 3)
        self.lineEdit8 = QLineEdit(self)
        self.lineEdit8.setFixedWidth(150)
        grid.addWidget(self.lineEdit8, 4, 4)

        self.label16 = QLabel("Mã cán bộ", self)
        grid.addWidget(self.label16, 5, 3)
        self.lineEdit9 = QLineEdit(self)
        self.lineEdit9.setFixedWidth(150)
        grid.addWidget(self.lineEdit9, 5, 4)

        self.label17 = QLabel("Số điện thoại:", self)
        grid.addWidget(self.label17, 6, 3)
        self.lineEdit10 = QLineEdit(self)
        self.lineEdit10.setFixedWidth(150)
        grid.addWidget(self.lineEdit10, 6, 4)

        self.label18 = QLabel("Ghi chú:", self)
        grid.addWidget(self.label18, 7, 3)
        self.lineEdit11 = QLineEdit(self)
        self.lineEdit11.setFixedWidth(150)
        grid.addWidget(self.lineEdit11, 7, 4)

        self.label19 = QLabel("Trạng thái cán bộ:", self)
        grid.addWidget(self.label19, 8, 3)
        self.lineEdit12 = QLineEdit(self)
        self.lineEdit12.setFixedWidth(150)
        self.lineEdit12.setReadOnly(True)
        self.lineEdit12.setStyleSheet("background-color: #cccccc;")
        grid.addWidget(self.lineEdit12, 8, 4)

        self.button2 = QPushButton("Sửa đổi thông tin", self)
        self.button2.clicked.connect(self.button2_clicked)
        grid.addWidget(self.button2, 9, 4)

        # Xóa cán bộ
        self.button3 = QPushButton("Kích hoạt/Vô hiệu hóa", self)
        self.button3.clicked.connect(self.button3_clicked)
        grid.addWidget(self.button3, 9, 3)

        # Xuất excel
        self.button4 = QPushButton("Xuất excel", self)
        self.button4.clicked.connect(self.button4_clicked)
        self.button4.setFixedWidth(150)
        grid.addWidget(self.button4, 5, 6)

        # Nhập excel
        self.button5 = QPushButton("Nhập excel", self)
        self.button5.clicked.connect(self.button5_clicked)
        self.button5.setFixedWidth(150)
        grid.addWidget(self.button5, 4, 6)

        # Nút đóng
        self.function = QPushButton("Bật chức năng")
        self.function.setFixedWidth(150)
        self.function.clicked.connect(self.show_buttons)
        grid.addWidget(self.function, 3, 6)
        for i, btn in enumerate([self.button1, self.button2, self.button3, self.button4, self.button5], start=1):
            btn.setVisible(False)

        # Tìm kiếm
        self.lineEdit13 = QLineEdit(self)
        self.lineEdit13.setFixedWidth(150)
        self.lineEdit13.setPlaceholderText("Tìm kiếm...")
        self.lineEdit13.textChanged.connect(self.search_table)
        grid.addWidget(self.lineEdit13, 10, 0)
        self.button6 = QPushButton("🔍 Tìm kiếm nâng cao")
        self.button6.setFixedWidth(160)
        self.button6.clicked.connect(self.button6_clicked)
        grid.addWidget(self.button6, 10, 1)

        # Bảng
        self.table = QTableWidget()
        self.table.setColumnCount(9)
        self.table.setHorizontalHeaderLabels([
            "ID", "Họ và tên", "Nhóm", "Chức vụ", "Email", "Mã CB", "SĐT", "Trạng thái cán bộ", "Ghi chú"
        ])
        grid.addWidget(self.table, 11, 0, 1, 8)
        self.table.setColumnWidth(0, 40)
        self.table.setColumnWidth(8, 250)
        self.table.setColumnWidth(8, 300)
        self.table.verticalHeader().setVisible(False)
        self.table.setSortingEnabled(True)

        # Thêm tab bảng
        self.tab_widget1 = QTabWidget(self)
        self.tab_widget1.setStyleSheet("""
                        QTabBar::tab {
                            background: #cccccc;      /* màu tab chưa active */
                            color: black;
                            padding: 5px;
                            border: 1px solid #016b67;
                            border-top-left-radius: 4px;
                            border-top-right-radius: 4px;
                            margin-right: 3px;
                        }

                        QTabBar::tab:selected {
                            background: white;          /* màu tab active */
                            color: black;
                        }

                        QTabBar::tab:pressed {
                        background: #d0d0d0;        /* màu tab khi nhấn, tạo hiệu ứng chìm */
                        padding-top: 6px;           /* đẩy chữ xuống tạo cảm giác nhấn */
                        padding-bottom: 4px;
                        }

                        QTabWidget::pane {
                        border: 1px solid gray;          /* kẻ khung quanh toàn bộ tab */
                        top: -1px;                       /* kéo pane lên ăn khớp với tab */
                        }

                        """)
        grid.addWidget(self.tab_widget1, 0, 0, 9, 3)

        tab1 = QWidget()
        tab1_layout = QGridLayout(tab1)
        tab1_layout.addWidget(self.label1, 0, 0)
        tab1_layout.addWidget(self.lineEdit1, 0, 1)
        tab1_layout.addWidget(self.label2, 1, 0)
        tab1_layout.addWidget(self.comboBox1,1, 1)
        tab1_layout.addWidget(self.label3, 2, 0)
        tab1_layout.addWidget(self.comboBox2, 2, 1)
        tab1_layout.addWidget(self.label4, 3, 0)
        tab1_layout.addWidget(self.lineEdit2, 3, 1)
        tab1_layout.addWidget(self.label5, 4, 0)
        tab1_layout.addWidget(self.lineEdit3, 4, 1)
        tab1_layout.addWidget(self.label6, 5, 0)
        tab1_layout.addWidget(self.lineEdit4, 5, 1)
        tab1_layout.addWidget(self.label7, 6, 0)
        tab1_layout.addWidget(self.lineEdit5, 6, 1)
        tab1_layout.addWidget(self.button1, 0, 2)
        for col in range(3, 4):
            spacer = QSpacerItem(100, 0, QSizePolicy.Fixed, QSizePolicy.Minimum)
            tab1_layout.addItem(spacer, 0, col)
        self.tab_widget1.addTab(tab1, "Thêm cán bộ")

        tab2 = QWidget()
        tab2_layout = QGridLayout(tab2)
        tab2_layout.addWidget(self.label10, 0, 0)
        tab2_layout.addWidget(self.lineEdit6, 0, 1)
        tab2_layout.addWidget(self.label12, 1, 0)
        tab2_layout.addWidget(self.lineEdit7, 1, 1)
        tab2_layout.addWidget(self.label13, 2, 0)
        tab2_layout.addWidget(self.comboBox3, 2, 1)
        tab2_layout.addWidget(self.label14, 3, 0)
        tab2_layout.addWidget(self.comboBox4, 3, 1)
        tab2_layout.addWidget(self.label15, 4, 0)
        tab2_layout.addWidget(self.lineEdit8, 4, 1)
        tab2_layout.addWidget(self.label16, 5, 0)
        tab2_layout.addWidget(self.lineEdit9, 5, 1)
        tab2_layout.addWidget(self.label17, 6, 0)
        tab2_layout.addWidget(self.lineEdit10, 6, 1)
        tab2_layout.addWidget(self.label18, 7, 0)
        tab2_layout.addWidget(self.lineEdit11, 7, 1)
        tab2_layout.addWidget(self.label19, 8, 0)
        tab2_layout.addWidget(self.lineEdit12, 8, 1)
        tab2_layout.addWidget(self.button2, 0, 2)
        tab2_layout.addWidget(self.button3, 1, 2)
        for col in range(3, 4):
            spacer = QSpacerItem(100, 0, QSizePolicy.Fixed, QSizePolicy.Minimum)
            tab2_layout.addItem(spacer, 0, col)

        self.tab_widget1.addTab(tab2, "Sửa đổi thông tin cán bộ")

        for col in range(0, 8):
            spacer = QSpacerItem(25, 0, QSizePolicy.Minimum, QSizePolicy.Minimum)
            grid.addItem(spacer, 0, col)

        self.menu = search_menu_staff(self.table)
        self.load_staff()
        self.setLayout(grid)

    def showEvent(self, event):
        """Tự động load lại dữ liệu mỗi khi cửa sổ được show"""
        super().showEvent(event)
        self.update_completer()
        self.load_staff()

    def load_staff(self):
        conn = sqlite3.connect(db_path_default)
        cursor = conn.cursor()
        cursor.execute(
            "SELECT id, name, department, position, email, employee_code, phone,active_staff,notes FROM staff")
        rows = cursor.fetchall()
        conn.close()

        self.table.setRowCount(0)

        for row_idx, row_data in enumerate(rows):
            self.table.insertRow(row_idx)
            for col_idx, value in enumerate(row_data):
                item = QTableWidgetItem()
                if col_idx == 7:
                    if value == 1:
                        item.setText("Còn làm ở TFC")
                    else:
                        item.setText("Không còn làm ở TFC")
                else:
                    if isinstance(value, int):
                        item.setData(Qt.ItemDataRole.DisplayRole, value)
                    else:
                        item.setText(str(value) if value else "")
                self.table.setItem(row_idx, col_idx, item)

    def search_table(self, text):
        text = text.strip().lower()
        for row in range(self.table.rowCount()):
            match = False
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                if item and text in item.text().lower():
                    match = True
                    break
            self.table.setRowHidden(row, not match)

    def update_completer(self):
        conn = sqlite3.connect(db_path_default)
        cursor = conn.cursor()
        cursor.execute("SELECT name, employee_code, email FROM staff")
        rows = cursor.fetchall()
        conn.close()

        suggestions = []
        for row in rows:
            for value in row:
                if value:
                    suggestions.append(str(value))
        suggestions = list(set(suggestions))

        self.completer = QCompleter(suggestions, self)
        self.completer.setCaseSensitivity(Qt.CaseInsensitive)
        self.completer.setFilterMode(Qt.MatchContains)
        self.lineEdit6.setCompleter(self.completer)
        self.completer.activated[str].connect(self.Search_selected)
        self.lineEdit6.returnPressed.connect(self.Search_selected)

    def Search_selected(self):
        search_text = self.lineEdit6.text()
        result = search_staff(search_text)
        if result == "Thông tin vừa rồi không khớp cán bộ nào!":
            self.label8.setText(result)
            self.lineEdit7.setText("")
            self.comboBox3.setCurrentText("NK1")
            self.comboBox4.setCurrentText("Chuyên viên")
            self.lineEdit8.setText("")
            self.lineEdit9.setText("")
            self.lineEdit10.setText("")
            self.lineEdit11.setText("")
            self.lineEdit12.setText("")
            return

        self.label8.setText("Đã tìm thấy kết quả")
        self.lineEdit7.setText(result[1])
        self.comboBox3.setCurrentText(result[2])
        self.comboBox4.setCurrentText(result[3])
        self.lineEdit8.setText(result[4])
        self.lineEdit9.setText(result[5])
        self.lineEdit10.setText(result[6])
        self.lineEdit11.setText(result[8])
        if result[7]:
            self.lineEdit12.setText("Còn làm ở TFC")
        else:
            self.lineEdit12.setText("Không còn làm ở TFC")
        self.update_completer()
        self.load_staff()

    def check_password(self, action_name="thao tác"):
        password, ok = QInputDialog.getText(
            self,
            "Xác nhận mật khẩu",
            f"Nhập mật khẩu để {action_name}:",
            echo=QLineEdit.Password
        )
        if not ok:
            return False

        if password == "123456":
            return True
        else:
            QMessageBox.warning(self, "Sai mật khẩu", "Mật khẩu không đúng!")
            return False

    def show_buttons(self):
        if not self.check_password("Bật chức năng"):
            self.label8.setText("Đã huỷ thao tác bật chức năng.")
            return

        # Lấy trạng thái hiện/ẩn dựa vào nút 1
        currently_visible = self.button5.isVisible()

        # Chuyển trạng thái ngược lại cho tất cả nút
        for btn in [self.button1, self.button2, self.button3, self.button4, self.button5]:
            btn.setVisible(not currently_visible)
        if currently_visible:
            self.function.setText("Bật chức năng")
        else:
            self.function.setText("Tắt chức năng")

    def button1_clicked(self):

        name_add = self.lineEdit1.text()
        department_add = self.comboBox1.currentText()
        position_add = self.comboBox2.currentText()
        email_add = self.lineEdit2.text()
        employee_code_add = self.lineEdit3.text()
        phone_add = self.lineEdit4.text()
        notes_add = self.lineEdit5.text()
        result = add_staff(name_add, department_add, position_add, email_add, employee_code_add, phone_add, notes_add)
        self.label8.setText(result)

        self.lineEdit1.setText("")
        self.comboBox1.setCurrentText("NK1")
        self.comboBox2.setCurrentText("Chuyên viên")
        self.lineEdit2.setText("")
        self.lineEdit3.setText("")
        self.lineEdit4.setText("")
        self.lineEdit5.setText("")
        self.update_completer()
        self.load_staff()

    def button2_clicked(self):

        name_add = self.lineEdit7.text()
        department_add = self.comboBox3.currentText()
        position_add = self.comboBox4.currentText()
        email_add = self.lineEdit8.text()
        employee_code_add = self.lineEdit9.text()
        phone_add = self.lineEdit10.text()
        notes_add = self.lineEdit11.text()
        search = self.lineEdit6.text()

        result = modify_staff(name_add, department_add, position_add, email_add, employee_code_add, phone_add,
                              notes_add, search)
        self.label8.setText(result)

        self.lineEdit6.setText("")
        self.lineEdit7.setText("")
        self.comboBox3.setCurrentText("NK1")
        self.comboBox4.setCurrentText("Chuyên viên")
        self.lineEdit8.setText("")
        self.lineEdit9.setText("")
        self.lineEdit10.setText("")
        self.lineEdit11.setText("")
        self.lineEdit12.setText("")
        self.update_completer()
        self.load_staff()

    def button3_clicked(self):

        search = self.lineEdit6.text()
        result = delete_staff(search)
        self.label8.setText(result)
        self.lineEdit6.setText("")
        self.lineEdit7.setText("")
        self.comboBox3.setCurrentText("NK1")
        self.comboBox4.setCurrentText("Chuyên viên")
        self.lineEdit8.setText("")
        self.lineEdit9.setText("")
        self.lineEdit10.setText("")
        self.lineEdit11.setText("")
        self.lineEdit12.setText("")
        self.update_completer()
        self.load_staff()

    def button4_clicked(self):
        export_staff_to_excel()

    def button5_clicked(self):
        import_staff_from_excel()
        self.update_completer()
        self.load_staff()

    def button6_clicked(self):
        pos = self.button6.mapToGlobal(QPoint(160, -175))
        self.menu.move(pos)
        self.menu.setWindowOpacity(0)
        self.menu.show()

        # Giữ animation lại
        self.anim = QPropertyAnimation(self.menu, b"windowOpacity", self)
        self.anim.setDuration(200)
        self.anim.setStartValue(0)
        self.anim.setEndValue(1)
        self.anim.start()

class search_menu_staff(QFrame):
    def __init__(self, staff_table):
        super().__init__()
        self.setWindowFlags(Qt.Popup)
        self.setStyleSheet("""
                    /* === NỀN CỬA SỔ === */
                    QWidget {
                        background-color: white;
                    }
                    QFrame {
                        background-color: white;
                        border-radius: 6px;         /* bo góc nếu muốn */
                        border: 1px solid #016b67;
                        
                    }
                    
                    /* === NÚT BẤM === */
                    QPushButton {
                        background: qlineargradient(
                            spread:pad, x1:0, y1:0, x2:1, y2:1,
                            stop:0 #016b67, stop:1 #0b9f87, stop:2 #00c3cc
                        );
                        color: white;
                        border: none;
                        border-radius: 6px;
                        padding: 6px;
                    }
                    QPushButton:hover {
                        background-color: #016b67;
                        color: white;
                    }

                    /* === NHÃN === */
                    QLabel {
                        color: black;
                        font-weight: bold;
                        border: none;
                    }

                    /* === Ô COMBOBOX / LINEEDIT / TEXTEDIT / TABLE === */
                    QComboBox,
                    QLineEdit,
                    QTextEdit,
                    QTableWidget {
                        background-color: white;
                        border: 1px solid #016b67;
                        border-radius: 4px;
                        color: black;
                    }

                    /* === SCROLLBAR VERTICAL (5px) === */
                    QScrollBar:vertical {
                        border: none;
                        background: #f0f0f0;
                        width: 5px;             /* mỏng 5px */
                        margin: 0px;
                        border-radius: 2px;
                    }
                    QScrollBar::handle:vertical {
                        background-color: #c0c0c0;   /* xám nhạt */
                        min-height: 20px;
                        border-radius: 2px;
                    }
                    QScrollBar::handle:vertical:hover {
                        background-color: #a0a0a0;   /* xám đậm hơn khi hover */
                    }
                    QScrollBar::add-line:vertical,
                    QScrollBar::sub-line:vertical {
                        height: 0px;
                        background: none;
                    }

                    /* === SCROLLBAR HORIZONTAL (5px) === */
                    QScrollBar:horizontal {
                        border: none;
                        background: #f0f0f0;
                        height: 5px;            /* mỏng 5px */
                        margin: 0px;
                        border-radius: 2px;
                    }
                    QScrollBar::handle:horizontal {
                        background-color: #c0c0c0;   /* xám nhạt */
                        min-width: 20px;
                        border-radius: 2px;
                    }
                    QScrollBar::handle:horizontal:hover {
                        background-color: #a0a0a0;   /* xám đậm hơn khi hover */
                    }
                    QScrollBar::add-line:horizontal,
                    QScrollBar::sub-line:horizontal {
                        width: 0px;
                        background: none;
                    }
                """)


        # đổ bóng
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(20)
        shadow.setOffset(0, 2)
        shadow.setColor(QColor(0, 0, 0, 80))
        self.setGraphicsEffect(shadow)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)

        # cuộn nội dung
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        layout.addWidget(scroll)

        content = QWidget()
        grid = QGridLayout(content)
        grid.setSpacing(15)
        scroll.setWidget(content)

        # thêm các nút (icon + text)
        self.label1 = QLabel("Họ và tên:", self)
        grid.addWidget(self.label1, 0, 0)
        self.lineEdit1 = QLineEdit(self)
        self.lineEdit1.setFixedWidth(150)
        grid.addWidget(self.lineEdit1, 0, 1)

        self.label2 = QLabel("Nhóm:", self)
        grid.addWidget(self.label2, 1, 0)
        self.comboBox1 = QComboBox(self)
        self.comboBox1.addItem("All")
        self.comboBox1.addItems(["NK1", "NK2", "NK3", "NK4", "LCND", "XK1", "XK2"])
        self.comboBox1.setFixedWidth(150)
        grid.addWidget(self.comboBox1, 1, 1)

        self.label3 = QLabel("Chức vụ:", self)
        grid.addWidget(self.label3, 2, 0)
        self.comboBox2 = QComboBox(self)
        self.comboBox2.addItem("All")
        self.comboBox2.addItems(["Chuyên viên", "Phó Phòng", "Trưởng Phòng", "Phó giám đốc", "Giám đốc"])
        self.comboBox2.setFixedWidth(150)
        grid.addWidget(self.comboBox2, 2, 1)

        self.label4 = QLabel("Email:", self)
        grid.addWidget(self.label4, 0, 2)
        self.lineEdit2 = QLineEdit(self)
        self.lineEdit2.setFixedWidth(150)
        grid.addWidget(self.lineEdit2, 0, 3)

        self.label5 = QLabel("Mã cán bộ", self)
        grid.addWidget(self.label5, 1, 2)
        self.lineEdit3 = QLineEdit(self)
        self.lineEdit3.setFixedWidth(150)
        grid.addWidget(self.lineEdit3, 1, 3)

        self.label6 = QLabel("Số điện thoại:", self)
        grid.addWidget(self.label6, 2, 2)
        self.lineEdit4 = QLineEdit(self)
        self.lineEdit4.setFixedWidth(150)
        grid.addWidget(self.lineEdit4, 2, 3)

        self.label7 = QLabel("Ghi chú:", self)
        grid.addWidget(self.label7, 3, 0)
        self.lineEdit5 = QLineEdit(self)
        self.lineEdit5.setFixedWidth(150)
        grid.addWidget(self.lineEdit5, 3, 1)

        self.button1 = QPushButton("Tìm kiếm", self)
        self.button1.setFixedWidth(150)
        grid.addWidget(self.button1, 4, 1)
        self.button1.clicked.connect(self.search_table)

        self.staff_table = staff_table



    def search_table(self):
        table = self.staff_table
        name = self.lineEdit1.text().strip().lower()
        group = self.comboBox1.currentText()
        position = self.comboBox2.currentText()
        email = self.lineEdit2.text().strip().lower()
        staff_id = self.lineEdit3.text().strip().lower()
        phone = self.lineEdit4.text().strip().lower()
        note = self.lineEdit5.text().strip().lower()

        # Duyệt từng hàng trong table
        for row in range(table.rowCount()):
            match = True  # mặc định là đúng
            # Kiểm tra từng cột
            if name and name not in table.item(row, 1).text().lower():
                match = False
            if group != "All" and group != table.item(row, 2).text():
                match = False
            if position != "All" and position != table.item(row, 3).text():
                match = False
            if email and email not in table.item(row, 4).text().lower():
                match = False
            if staff_id and staff_id not in table.item(row, 5).text().lower():
                match = False
            if phone and phone not in table.item(row, 6).text().lower():
                match = False
            if note and note not in table.item(row, 8).text().lower():
                match = False

            # Hiển thị hoặc ẩn hàng tùy vào kết quả tìm kiếm
            table.setRowHidden(row, not match)

class assets_window(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Quản lý tài sản")
        self.setGeometry(200, 200, 1080, 700)
        self.setWindowIcon(QIcon("avatar_app.png"))
        self.setStyleSheet("""
            /* === NỀN CỬA SỔ === */
            QWidget {
                background-color: #f0f0f0;
                color: black;
            }

            /* === NÚT BẤM === */
            QPushButton {
                background: qlineargradient(
                    spread:pad, x1:0, y1:0, x2:1, y2:1,
                    stop:0 #016b67, stop:1 #0b9f87, stop:2 #00c3cc
                );
                color: white;
                border: none;
                border-radius: 6px;
                padding: 6px;
            }
            QPushButton:hover {
                background-color: #016b67;
                color: white;
            }

            /* === NHÃN === */
            QLabel {
                color: black;
                font-weight: bold;
            }

            /* === Ô COMBOBOX / LINEEDIT / TEXTEDIT / TABLE === */
            QComboBox,
            QLineEdit,
            QTextEdit,
            QTableWidget {
                background-color: white;
                border: 1px solid #016b67;
                border-radius: 4px;
                color: black;
            }

            /* === SCROLLBAR VERTICAL (5px) === */
            QScrollBar:vertical {
                border: none;
                background: #f0f0f0;
                width: 5px;             /* mỏng 5px */
                margin: 0px;
                border-radius: 2px;
            }
            QScrollBar::handle:vertical {
                background-color: #c0c0c0;   /* xám nhạt */
                min-height: 20px;
                border-radius: 2px;
            }
            QScrollBar::handle:vertical:hover {
                background-color: #a0a0a0;   /* xám đậm hơn khi hover */
            }
            QScrollBar::add-line:vertical,
            QScrollBar::sub-line:vertical {
                height: 0px;
                background: none;
            }

            /* === SCROLLBAR HORIZONTAL (5px) === */
            QScrollBar:horizontal {
                border: none;
                background: #f0f0f0;
                height: 5px;            /* mỏng 5px */
                margin: 0px;
                border-radius: 2px;
            }
            QScrollBar::handle:horizontal {
                background-color: #c0c0c0;   /* xám nhạt */
                min-width: 20px;
                border-radius: 2px;
            }
            QScrollBar::handle:horizontal:hover {
                background-color: #a0a0a0;   /* xám đậm hơn khi hover */
            }
            QScrollBar::add-line:horizontal,
            QScrollBar::sub-line:horizontal {
                width: 0px;
                background: none;
            }
        """)
        grid = QGridLayout(self)

        # Phần thêm tài sản
        #self.label9 = QLabel("Thêm tài sản:", self)
        #grid.addWidget(self.label9, 0, 0)

        self.label1 = QLabel("Loại tài sản:", self)
        grid.addWidget(self.label1, 1, 0)
        self.comboBox1 = QComboBox(self)
        self.comboBox1.addItems(["CPU", "Laptop", "Monitor", "Printer", "Scanner"])
        self.comboBox1.setFixedWidth(150)
        grid.addWidget(self.comboBox1, 1, 1)

        self.label2 = QLabel("Mã hiệu (dùng cho laptop):", self)
        grid.addWidget(self.label2, 2, 0)
        self.lineEdit1 = QLineEdit(self)
        self.lineEdit1.setFixedWidth(150)
        grid.addWidget(self.lineEdit1, 2, 1)

        self.label3 = QLabel("Số seri:", self)
        grid.addWidget(self.label3, 3, 0)
        self.lineEdit2 = QLineEdit(self)
        self.lineEdit2.setFixedWidth(150)
        grid.addWidget(self.lineEdit2, 3, 1)

        self.label4 = QLabel("Ngày nhận tài sản:", self)
        grid.addWidget(self.label4, 4, 0)
        self.lineEdit3 = QLineEdit(self)
        self.lineEdit3.setFixedWidth(150)
        today = datetime.today().strftime("%Y-%m-%d")
        self.lineEdit3.setText(today)
        grid.addWidget(self.lineEdit3, 4, 1)

        self.label5 = QLabel("Tình trạng", self)
        grid.addWidget(self.label5, 5, 0)
        self.comboBox2 = QComboBox(self)
        self.comboBox2.addItems(["Sử dụng bình thường", "Hỏng"])
        self.comboBox2.setFixedWidth(150)
        grid.addWidget(self.comboBox2, 5, 1)

        self.label8 = QLabel("Giá trị tài sản:", self)
        grid.addWidget(self.label8, 6, 0)
        self.lineEdit12 = QLineEdit(self)
        self.lineEdit12.setFixedWidth(150)
        validator = QIntValidator()
        self.lineEdit12.setValidator(validator)
        grid.addWidget(self.lineEdit12, 6, 1)

        self.label_bams1 = QLabel("Mã BAMS:", self)
        grid.addWidget(self.label_bams1, 7, 0)
        self.lineEdit_bams1 = QLineEdit(self)
        self.lineEdit_bams1.setFixedWidth(150)
        grid.addWidget(self.lineEdit_bams1, 7, 1)

        self.label6 = QLabel("Ghi chú:", self)
        grid.addWidget(self.label6, 8, 0)
        self.lineEdit4 = QLineEdit(self)
        self.lineEdit4.setFixedWidth(150)
        grid.addWidget(self.lineEdit4, 8, 1)

        self.label7 = QLabel("Đây là cửa sổ để quản lý danh sách tài sản", self)
        self.label7.setFixedWidth(300)
        grid.addWidget(self.label7, 1, 6)

        self.button1 = QPushButton("Thêm tài sản", self)
        self.button1.clicked.connect(self.button1_clicked)
        grid.addWidget(self.button1, 9, 1)

        # Phần chỉnh sửa thông tin tài sản
        self.label10 = QLabel("Sửa đổi thông tin tài sản :", self)
        grid.addWidget(self.label10, 0, 3)
        self.lineEdit5 = QLineEdit(self)
        self.lineEdit5.setFixedWidth(150)
        self.lineEdit5.setPlaceholderText("Số Seri")
        grid.addWidget(self.lineEdit5, 0, 4)
        self.update_completer()

        self.label11 = QLabel("Loại tài sản:", self)
        grid.addWidget(self.label11, 1, 3)
        self.comboBox3 = QComboBox(self)
        self.comboBox3.addItems(["CPU", "Laptop", "Monitor", "Printer", "Scanner"])
        self.comboBox3.setFixedWidth(150)
        grid.addWidget(self.comboBox3, 1, 4)

        self.label12 = QLabel("Mã hiệu (dùng cho laptop):", self)
        grid.addWidget(self.label12, 2, 3)
        self.lineEdit6 = QLineEdit(self)
        self.lineEdit6.setFixedWidth(150)
        grid.addWidget(self.lineEdit6, 2, 4)

        self.label13 = QLabel("Số seri:", self)
        grid.addWidget(self.label13, 3, 3)
        self.lineEdit7 = QLineEdit(self)
        self.lineEdit7.setFixedWidth(150)
        grid.addWidget(self.lineEdit7, 3, 4)

        self.label14 = QLabel("Ngày nhận tài sản:", self)
        grid.addWidget(self.label14, 4, 3)
        self.lineEdit8 = QLineEdit(self)
        self.lineEdit8.setFixedWidth(150)
        grid.addWidget(self.lineEdit8, 4, 4)

        self.label15 = QLabel("Tình trạng", self)
        grid.addWidget(self.label15, 5, 3)
        self.comboBox4 = QComboBox(self)
        self.comboBox4.addItems(["Sử dụng bình thường", "Hỏng"])
        self.comboBox4.setFixedWidth(150)
        grid.addWidget(self.comboBox4, 5, 4)

        self.label9 = QLabel("Giá trị tài sản:", self)
        grid.addWidget(self.label9, 6, 3)
        self.lineEdit13 = QLineEdit(self)
        self.lineEdit13.setFixedWidth(150)
        validator = QIntValidator()
        self.lineEdit13.setValidator(validator)
        grid.addWidget(self.lineEdit13, 6, 4)

        self.label_bams2 = QLabel("Mã BAMS:", self)
        self.lineEdit_bams2 = QLineEdit(self)
        self.lineEdit_bams2.setFixedWidth(150)

        self.label16 = QLabel("Ghi chú:", self)
        grid.addWidget(self.label16, 7, 3)
        self.lineEdit9 = QLineEdit(self)
        self.lineEdit9.setFixedWidth(150)
        grid.addWidget(self.lineEdit9, 7, 4)

        self.label17 = QLabel("Tình trạng sử dụng:", self)
        grid.addWidget(self.label17, 7, 3)
        self.lineEdit10 = QLineEdit(self)
        self.lineEdit10.setFixedWidth(150)
        self.lineEdit10.setReadOnly(True)
        self.lineEdit10.setStyleSheet("background-color: #cccccc;")
        grid.addWidget(self.lineEdit10, 7, 4)

        self.button2 = QPushButton("Sửa đổi thông tin", self)
        self.button2.clicked.connect(self.button2_clicked)
        grid.addWidget(self.button2, 8, 4)

        # Nút thu hồi
        self.button3 = QPushButton("Thu hồi", self)
        self.button3.clicked.connect(self.button3_clicked)
        grid.addWidget(self.button3, 8, 3)


        # Xuất excel
        self.button4 = QPushButton("Xuất excel", self)
        self.button4.clicked.connect(self.button4_clicked)
        self.button4.setFixedWidth(150)
        grid.addWidget(self.button4, 5, 6)

        # Nhập excel
        self.button5 = QPushButton("Nhập excel", self)
        self.button5.clicked.connect(self.button5_clicked)
        self.button5.setFixedWidth(150)
        grid.addWidget(self.button5, 4, 6)

        # Nút chức năng
        self.function = QPushButton("Bật chức năng")
        self.function.setFixedWidth(150)
        grid.addWidget(self.function, 3, 6)
        self.function.clicked.connect(self.show_buttons)
        for i, btn in enumerate([self.button1, self.button2, self.button3, self.button4, self.button5], start=1):
            btn.setVisible(False)

        # Tìm kiếm
        self.lineEdit11 = QLineEdit(self)
        self.lineEdit11.setFixedWidth(150)
        self.lineEdit11.setPlaceholderText("Tìm kiếm...")
        self.lineEdit11.textChanged.connect(self.search_table)
        grid.addWidget(self.lineEdit11, 9, 0)
        self.button6 = QPushButton("🔍 Tìm kiếm nâng cao")
        self.button6.setFixedWidth(160)
        self.button6.clicked.connect(self.button6_clicked)
        grid.addWidget(self.button6, 9, 1)

        self.button_qr_inventory = QPushButton("📷 Kiểm kê tài sản (QR)", self)
        self.button_qr_inventory.setFixedWidth(185)
        self.button_qr_inventory.clicked.connect(self.open_inventory_qr)
        self.inventory_qr_window = None
        grid.addWidget(self.button_qr_inventory, 9, 2)

        # Bảng chi tiết tài sản
        self.table = QTableWidget()
        self.table.setColumnCount(13)
        self.table.setHorizontalHeaderLabels([
            "ID", "Loại tài sản", "Mã hiệu", "Số seri", "Ngày nhận tài sản", "Tình trạng", "Tình trạng sử dụng",
            "Ngày thu hồi trả văn phòng", "Trị giá tài sản", "Cán bộ giữ tài sản", "Phòng ban", "Mã BAMS", "Ghi chú"
        ])
        grid.addWidget(self.table, 10, 0, 1, 8)
        self.table.setColumnWidth(0, 40)
        self.table.setColumnWidth(2, 50)
        self.table.setColumnWidth(3, 90)
        self.table.setColumnWidth(4, 120)
        self.table.setColumnWidth(5, 150)
        self.table.setColumnWidth(6, 150)
        self.table.setColumnWidth(7, 160)
        self.table.setColumnWidth(8, 150)
        self.table.setColumnWidth(9, 150)
        self.table.setColumnWidth(11, 100)
        self.table.setColumnWidth(12, 300)
        self.table.verticalHeader().setVisible(False)
        self.table.setSortingEnabled(True)

        #Bảng tài sản cán bộ
        self.table2 = QTableWidget()
        self.table2.setColumnCount(10)
        self.table2.setHorizontalHeaderLabels([
            "ID", "Loại tài sản", "Mã hiệu", "Số seri", "Cán bộ giữ tài sản", "Ngày nhận tài sản", "Tình trạng", "Trị giá tài sản", "Phòng ban", "Ghi chú"
        ])
        self.table2.setColumnWidth(0, 40)
        self.table2.setColumnWidth(4, 150)
        self.table2.setColumnWidth(5, 150)
        self.table2.setColumnWidth(6, 150)
        self.table2.setColumnWidth(9, 300)
        self.table2.verticalHeader().setVisible(False)
        self.table2.setSortingEnabled(True)

        # Bảng tài sản trong kho
        self.table3 = QTableWidget()
        self.table3.setColumnCount(8)
        self.table3.setHorizontalHeaderLabels([
            "ID", "Loại tài sản", "Mã hiệu", "Số seri", "Ngày nhận tài sản", "Tình trạng", "Trị giá tài sản", "Ghi chú"
        ])
        self.table3.setColumnWidth(0, 40)
        self.table3.setColumnWidth(3, 200)
        self.table3.setColumnWidth(4, 150)
        self.table3.setColumnWidth(5, 150)
        self.table3.setColumnWidth(6, 150)
        self.table3.setColumnWidth(7, 300)
        self.table3.verticalHeader().setVisible(False)
        self.table3.setSortingEnabled(True)

        # Bảng tài sản thu hồi trả VP
        self.table4 = QTableWidget()
        self.table4.setColumnCount(9)
        self.table4.setHorizontalHeaderLabels([
            "ID", "Loại tài sản", "Mã hiệu", "Số seri", "Ngày nhận tài sản", "Tình trạng", "Ngày thu hồi trả văn phòng", "Trị giá tài sản", "Ghi chú"
        ])
        self.table4.setColumnWidth(0, 40)
        self.table4.setColumnWidth(3, 150)
        self.table4.setColumnWidth(4, 150)
        self.table4.setColumnWidth(5, 150)
        self.table4.setColumnWidth(6, 180)
        self.table4.setColumnWidth(7, 150)
        self.table4.setColumnWidth(8, 300)
        self.table4.verticalHeader().setVisible(False)
        self.table4.setSortingEnabled(True)

        # Thêm tab chức năng
        self.tab_widget1 = QTabWidget(self)
        self.tab_widget1.setStyleSheet("""
                                QTabBar::tab {
                                    background: #cccccc;      /* màu tab chưa active */
                                    color: black;
                                    padding: 5px;
                                    border: 1px solid #016b67;
                                    border-top-left-radius: 4px;
                                    border-top-right-radius: 4px;
                                    margin-right: 3px;
                                }

                                QTabBar::tab:selected {
                                    background: white;          /* màu tab active */
                                    color: black;
                                }

                                QTabBar::tab:pressed {
                                background: #d0d0d0;        /* màu tab khi nhấn, tạo hiệu ứng chìm */
                                padding-top: 6px;           /* đẩy chữ xuống tạo cảm giác nhấn */
                                padding-bottom: 4px;
                                }

                                QTabWidget::pane {
                                border: 1px solid gray;          /* kẻ khung quanh toàn bộ tab */
                                top: -1px;                       /* kéo pane lên ăn khớp với tab */
                                }

                                """)
        grid.addWidget(self.tab_widget1, 0, 0, 9, 3)

        tab1 = QWidget()
        tab1_layout = QGridLayout(tab1)
        tab1_layout.addWidget(self.label1, 0, 0)
        tab1_layout.addWidget(self.comboBox1, 0, 1)
        tab1_layout.addWidget(self.label2, 1, 0)
        tab1_layout.addWidget(self.lineEdit1, 1, 1)
        tab1_layout.addWidget(self.label3, 2, 0)
        tab1_layout.addWidget(self.lineEdit2, 2, 1)
        tab1_layout.addWidget(self.label4, 3, 0)
        tab1_layout.addWidget(self.lineEdit3, 3, 1)
        tab1_layout.addWidget(self.label5, 4, 0)
        tab1_layout.addWidget(self.comboBox2, 4, 1)
        tab1_layout.addWidget(self.label8, 5, 0)
        tab1_layout.addWidget(self.lineEdit12, 5, 1)
        tab1_layout.addWidget(self.label_bams1, 6, 0)
        tab1_layout.addWidget(self.lineEdit_bams1, 6, 1)
        tab1_layout.addWidget(self.label6, 7, 0)
        tab1_layout.addWidget(self.lineEdit4, 7, 1)
        tab1_layout.addWidget(self.button1, 0, 2)
        for col in range(3, 4):
            spacer = QSpacerItem(100, 0, QSizePolicy.Fixed, QSizePolicy.Minimum)
            tab1_layout.addItem(spacer, 0, col)
        self.tab_widget1.addTab(tab1, "Thêm tài sản")

        tab2 = QWidget()
        tab2_layout = QGridLayout(tab2)
        tab2_layout.addWidget(self.label10, 0, 0)
        tab2_layout.addWidget(self.lineEdit5, 0, 1)
        tab2_layout.addWidget(self.label11, 1, 0)
        tab2_layout.addWidget(self.comboBox3, 1, 1)
        tab2_layout.addWidget(self.label12, 2, 0)
        tab2_layout.addWidget(self.lineEdit6, 2, 1)
        tab2_layout.addWidget(self.label13, 3, 0)
        tab2_layout.addWidget(self.lineEdit7, 3, 1)
        tab2_layout.addWidget(self.label14, 4, 0)
        tab2_layout.addWidget(self.lineEdit8, 4, 1)
        tab2_layout.addWidget(self.label15, 5, 0)
        tab2_layout.addWidget(self.comboBox4, 5, 1)
        tab2_layout.addWidget(self.label9, 6, 0)
        tab2_layout.addWidget(self.lineEdit13, 6, 1)
        tab2_layout.addWidget(self.label_bams2, 7, 0)
        tab2_layout.addWidget(self.lineEdit_bams2, 7, 1)
        tab2_layout.addWidget(self.label16, 8, 0)
        tab2_layout.addWidget(self.lineEdit9, 8, 1)
        tab2_layout.addWidget(self.label17, 9, 0)
        tab2_layout.addWidget(self.lineEdit10, 9, 1)
        tab2_layout.addWidget(self.button2, 0, 2)
        tab2_layout.addWidget(self.button3, 1, 2)
        for col in range(3, 4):
            spacer = QSpacerItem(100, 0, QSizePolicy.Fixed, QSizePolicy.Minimum)
            tab2_layout.addItem(spacer, 0, col)

        self.tab_widget1.addTab(tab2, "Sửa đổi thông tin tài sản")
        for col in range(0, 8):
            spacer = QSpacerItem(25, 0, QSizePolicy.Minimum, QSizePolicy.Minimum)
            grid.addItem(spacer, 0, col)

        #Thêm tab bảng
        self.tab_widget2 = QTabWidget(self)
        self.tab_widget2.setStyleSheet("""
                                        QTabBar::tab {
                                            background: #cccccc;      /* màu tab chưa active */
                                            color: black;
                                            padding: 5px;
                                            border: 1px solid #016b67;
                                            border-top-left-radius: 4px;
                                            border-top-right-radius: 4px;
                                            margin-right: 3px;
                                        }

                                        QTabBar::tab:selected {
                                            background: white;          /* màu tab active */
                                            color: black;
                                        }

                                        QTabBar::tab:pressed {
                                        background: #d0d0d0;        /* màu tab khi nhấn, tạo hiệu ứng chìm */
                                        padding-top: 6px;           /* đẩy chữ xuống tạo cảm giác nhấn */
                                        padding-bottom: 4px;
                                        }

                                        QTabWidget::pane {
                                        border: 1px solid gray;          /* kẻ khung quanh toàn bộ tab */
                                        top: -1px;                       /* kéo pane lên ăn khớp với tab */
                                        }

                                        """)
        grid.addWidget(self.tab_widget2, 10, 0, 1, 8)
        self.tab_widget2.addTab(self.table2, "Tài sản cán bộ")
        self.tab_widget2.addTab(self.table3, "Tài sản trong kho")
        self.tab_widget2.addTab(self.table4, "Tài sản thu hồi trả VP")
        self.tab_widget2.addTab(self.table, "Chi tiết tất cả tài sản")

        self.menu = search_menu_assets(self.table, self.table2, self.table3, self.table4)
        self.load_assets()
        self.setLayout(grid)

    def showEvent(self, event):
        """Tự động load lại dữ liệu mỗi khi cửa sổ được show"""
        super().showEvent(event)
        self.update_completer()
        self.load_assets()

    def load_assets(self):
        conn = sqlite3.connect(db_path_default)
        cursor = conn.cursor()
        cursor.execute("""
                       SELECT a.id   AS asset_id,
                              a.asset_name,
                              a.code,
                              a.serial_number,
                              a.purchase_date,
                              a.status,
                              a.active_asset,
                              a.return_date,
                              a.value,
                              s.name AS staff_name,
                              s.department,
                              a.bams,
                              a.notes
                       FROM assets a
                                LEFT JOIN asset_assignments aa ON aa.asset_id = a.id AND aa.returned_date IS NULL
                                LEFT JOIN staff s ON aa.staff_id = s.id
                       """)
        rows = cursor.fetchall()
        conn.close()

        self.table.setRowCount(0)

        for row_idx, row_data in enumerate(rows):
            self.table.insertRow(row_idx)
            for col_idx, value in enumerate(row_data):
                item = QTableWidgetItem()
                if col_idx == 6:
                    if value:
                        item.setText("Đang sử dụng")
                    else:
                        item.setText("Đã thu hồi trả văn phòng")
                elif col_idx == 8:
                    if value is not None:
                        formatted_value = f"{value:,}"
                        item.setText(formatted_value)
                        item.setData(Qt.ItemDataRole.UserRole, value)
                    else:
                        item.setText("0")
                elif col_idx == 9:
                    if value:
                        item.setText(str(value))
                    else:
                        item.setText("Trong kho")
                else:
                    if isinstance(value, int):
                        item.setData(Qt.ItemDataRole.DisplayRole, value)
                    else:
                        item.setText(str(value) if value else "")
                self.table.setItem(row_idx, col_idx, item)

        conn = sqlite3.connect(db_path_default)
        cursor = conn.cursor()
        cursor.execute("""
                               SELECT a.id   AS asset_id,
                                      a.asset_name,
                                      a.code,
                                      a.serial_number,
                                      s.name AS staff_name,
                                      a.purchase_date,
                                      a.status,
                                      a.value,
                                      s.department,
                                      a.notes
                               FROM assets a
                                        LEFT JOIN asset_assignments aa ON aa.asset_id = a.id AND aa.returned_date IS NULL
                                        LEFT JOIN staff s ON aa.staff_id = s.id
                               WHERE aa.id IS NOT NULL AND a.active_asset = 1
                               """)
        rows = cursor.fetchall()
        conn.close()

        self.table2.setRowCount(0)

        for row_idx, row_data in enumerate(rows):
            self.table2.insertRow(row_idx)
            for col_idx, value in enumerate(row_data):
                item = QTableWidgetItem()
                if col_idx == 7:
                    if value is not None:
                        formatted_value = f"{value:,}"
                        item.setText(formatted_value)
                        item.setData(Qt.ItemDataRole.UserRole, value)
                    else:
                        item.setText("0")
                elif col_idx == 4:
                    if value:
                        item.setText(str(value))
                    else:
                        item.setText("Trong kho")
                else:
                    if isinstance(value, int):
                        item.setData(Qt.ItemDataRole.DisplayRole, value)
                    else:
                        item.setText(str(value) if value else "")
                self.table2.setItem(row_idx, col_idx, item)

        conn = sqlite3.connect(db_path_default)
        cursor = conn.cursor()
        cursor.execute("""
                                       SELECT a.id   AS asset_id,
                                              a.asset_name,
                                              a.code,
                                              a.serial_number,
                                              a.purchase_date,
                                              a.status,
                                              a.value,
                                              a.notes
                                       FROM assets a
                                                LEFT JOIN asset_assignments aa ON aa.asset_id = a.id AND aa.returned_date IS NULL
                                                LEFT JOIN staff s ON aa.staff_id = s.id
                                       WHERE aa.id IS NULL AND a.active_asset = 1
                                       """)
        rows = cursor.fetchall()
        conn.close()

        self.table3.setRowCount(0)

        for row_idx, row_data in enumerate(rows):
            self.table3.insertRow(row_idx)
            for col_idx, value in enumerate(row_data):
                item = QTableWidgetItem()
                if col_idx == 6:
                    if value is not None:
                        formatted_value = f"{value:,}"
                        item.setText(formatted_value)
                        item.setData(Qt.ItemDataRole.UserRole, value)
                    else:
                        item.setText("0")
                else:
                    if isinstance(value, int):
                        item.setData(Qt.ItemDataRole.DisplayRole, value)
                    else:
                        item.setText(str(value) if value else "")
                self.table3.setItem(row_idx, col_idx, item)

        conn = sqlite3.connect(db_path_default)
        cursor = conn.cursor()
        cursor.execute("""
                                               SELECT a.id   AS asset_id,
                                                      a.asset_name,
                                                      a.code,
                                                      a.serial_number,
                                                      a.purchase_date,
                                                      a.status,
                                                      a.return_date,
                                                      a.value,
                                                      a.notes
                                               FROM assets a
                                                        LEFT JOIN asset_assignments aa ON aa.asset_id = a.id AND aa.returned_date IS NULL
                                                        LEFT JOIN staff s ON aa.staff_id = s.id
                                               WHERE aa.id IS NULL AND a.active_asset = 0
                                               """)
        rows = cursor.fetchall()
        conn.close()

        self.table4.setRowCount(0)

        for row_idx, row_data in enumerate(rows):
            self.table4.insertRow(row_idx)
            for col_idx, value in enumerate(row_data):
                item = QTableWidgetItem()
                if col_idx == 7:
                    if value is not None:
                        formatted_value = f"{value:,}"
                        item.setText(formatted_value)
                        item.setData(Qt.ItemDataRole.UserRole, value)
                    else:
                        item.setText("0")
                else:
                    if isinstance(value, int):
                        item.setData(Qt.ItemDataRole.DisplayRole, value)
                    else:
                        item.setText(str(value) if value else "")
                self.table4.setItem(row_idx, col_idx, item)

    def search_table(self, text):
        text = text.strip().lower()
        for row in range(self.table.rowCount()):
            match = False
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                if item and text in item.text().lower():
                    match = True
                    break
            self.table.setRowHidden(row, not match)

        for row in range(self.table2.rowCount()):
            match = False
            for col in range(self.table2.columnCount()):
                item = self.table2.item(row, col)
                if item and text in item.text().lower():
                    match = True
                    break
            self.table2.setRowHidden(row, not match)

        for row in range(self.table3.rowCount()):
            match = False
            for col in range(self.table3.columnCount()):
                item = self.table3.item(row, col)
                if item and text in item.text().lower():
                    match = True
                    break
            self.table3.setRowHidden(row, not match)

        for row in range(self.table4.rowCount()):
            match = False
            for col in range(self.table4.columnCount()):
                item = self.table4.item(row, col)
                if item and text in item.text().lower():
                    match = True
                    break
            self.table4.setRowHidden(row, not match)

    def update_completer(self):
        conn = sqlite3.connect(db_path_default)
        cursor = conn.cursor()
        cursor.execute("SELECT serial_number FROM assets")
        rows = cursor.fetchall()
        conn.close()

        suggestions = []
        for row in rows:
            for value in row:
                if value:
                    suggestions.append(str(value))
        suggestions = list(set(suggestions))

        self.completer = QCompleter(suggestions, self)
        self.completer.setCaseSensitivity(Qt.CaseInsensitive)
        self.completer.setFilterMode(Qt.MatchContains)
        self.lineEdit5.setCompleter(self.completer)
        self.completer.activated[str].connect(self.search_selected)
        self.lineEdit5.returnPressed.connect(self.search_selected)

    def search_selected(self):
        search_text = self.lineEdit5.text()
        result = search_assets(search_text)
        if result == "Thông tin vừa rồi không khớp tài sản nào!":
            self.label7.setText(result)
            self.comboBox3.setCurrentText("CPU")
            self.lineEdit6.setText("")
            self.lineEdit7.setText("")
            self.lineEdit8.setText("")
            self.comboBox4.setCurrentText("Sử dụng bình thường")
            self.lineEdit13.setText("")
            self.lineEdit9.setText("")
            self.lineEdit_bams2.setText("")
            self.lineEdit10.setText("")
            return

        self.label7.setText("Đã tìm thấy kết quả")
        self.comboBox3.setCurrentText(result[1])
        self.lineEdit6.setText(result[2])
        self.lineEdit7.setText(result[3])
        self.lineEdit8.setText(result[4])
        self.comboBox4.setCurrentText(result[5])
        self.lineEdit13.setText(f"{result[8]:,}")
        self.lineEdit_bams2.setText(result[9])
        self.lineEdit9.setText(result[10])
        if result[6]:
            self.lineEdit10.setText("Đang sử dụng")
        else:
            self.lineEdit10.setText("Đã thu hồi trả văn phòng")
        self.update_completer()
        self.load_assets()

    def check_password(self, action_name="thao tác"):
        password, ok = QInputDialog.getText(
            self,
            "Xác nhận mật khẩu",
            f"Nhập mật khẩu để {action_name}:",
            echo=QLineEdit.Password
        )
        if not ok:
            return False

        if password == "123456":
            return True
        else:
            QMessageBox.warning(self, "Sai mật khẩu", "Mật khẩu không đúng!")
            return False

    def show_buttons(self):
        if not self.check_password("Bật chức năng"):
            self.label7.setText("Đã huỷ thao tác bật chức năng.")
            return

        # Lấy trạng thái hiện/ẩn dựa vào nút 1
        currently_visible = self.button5.isVisible()

        # Chuyển trạng thái ngược lại cho tất cả nút
        for btn in [self.button1, self.button2, self.button3, self.button4, self.button5]:
            btn.setVisible(not currently_visible)
        if currently_visible:
            self.function.setText("Bật chức năng")
        else:
            self.function.setText("Tắt chức năng")

    def button1_clicked(self):

        asset_name_add = self.comboBox1.currentText()
        code_add = self.lineEdit1.text()
        serial_number_add = self.lineEdit2.text()
        purchase_date_add = self.lineEdit3.text()
        status_add = self.comboBox2.currentText()
        value_add = int(self.lineEdit12.text().replace(",", "") or 0)
        bams_add = self.lineEdit_bams1.text()
        notes_add = self.lineEdit4.text()

        result = add_assets(asset_name_add, code_add, serial_number_add, purchase_date_add, status_add, value_add, bams_add, notes_add)
        self.label7.setText(result)

        self.comboBox1.setCurrentText("CPU")
        self.lineEdit1.setText("")
        self.lineEdit2.setText("")
        today = datetime.today().strftime("%Y-%m-%d")
        self.lineEdit3.setText(today)
        self.comboBox2.setCurrentText("Sử dụng bình thường")
        self.lineEdit12.setText("")
        self.lineEdit_bams1.setText("")
        self.lineEdit4.setText("")

        self.update_completer()
        self.load_assets()

    def button2_clicked(self):

        asset_name_add = self.comboBox3.currentText()
        code_add = self.lineEdit6.text()
        serial_number_add = self.lineEdit7.text()
        purchase_date_add = self.lineEdit8.text()
        status_add = self.comboBox4.currentText()
        value_add = int(self.lineEdit13.text().replace(",", "") or 0)
        notes_add = self.lineEdit9.text()
        bams_add = self.lineEdit_bams2.text()
        search = self.lineEdit5.text()

        result = modify_assets(asset_name_add, code_add, serial_number_add, purchase_date_add, status_add,value_add, bams_add, notes_add,
                               search)

        self.label7.setText(result)
        self.lineEdit5.setText("")
        self.comboBox3.setCurrentText("CPU")
        self.lineEdit6.setText("")
        self.lineEdit7.setText("")
        self.lineEdit8.setText("")
        self.comboBox4.setCurrentText("Sử dụng bình thường")
        self.lineEdit13.setText("")
        self.lineEdit_bams2.setText("")
        self.lineEdit9.setText("")
        self.lineEdit10.setText("")
        self.update_completer()
        self.load_assets()

    def button3_clicked(self):

        search = self.lineEdit5.text()
        result = remove_assets(search)

        self.label7.setText(result)
        self.lineEdit5.setText("")
        self.comboBox3.setCurrentText("CPU")
        self.lineEdit6.setText("")
        self.lineEdit7.setText("")
        self.lineEdit8.setText("")
        self.comboBox4.setCurrentText("Sử dụng bình thường")
        self.lineEdit13.setText("")
        self.lineEdit_bams2.setText("")
        self.lineEdit9.setText("")
        self.lineEdit10.setText("")
        self.update_completer()
        self.load_assets()

    def button4_clicked(self):
        export_assets_to_excel()

    def button5_clicked(self):
        import_assets_from_excel()
        self.update_completer()
        self.load_assets()
    def button6_clicked(self):
        pos = self.button6.mapToGlobal(QPoint(160, -250))
        self.menu.move(pos)
        self.menu.setWindowOpacity(0)
        self.menu.show()

        # Giữ animation lại
        self.anim = QPropertyAnimation(self.menu, b"windowOpacity", self)
        self.anim.setDuration(200)
        self.anim.setStartValue(0)
        self.anim.setEndValue(1)
        self.anim.start()

    def open_inventory_qr(self):
        if self.inventory_qr_window is None:
            self.inventory_qr_window = inventory_qr_window()
        else:
            self.inventory_qr_window.reload_from_db(preserve_scanned=True)
        self.inventory_qr_window.show()
        self.inventory_qr_window.raise_()
        self.inventory_qr_window.activateWindow()

class search_menu_assets(QFrame):
    def __init__(self, assets_table, assets_table2, assets_table3, assets_table4):
        super().__init__()
        self.setWindowFlags(Qt.Popup)
        self.setStyleSheet("""
                            /* === NỀN CỬA SỔ === */
                            QWidget {
                                background-color: white;
                            }
                            QFrame {
                                background-color: white;
                                border-radius: 6px;         /* bo góc nếu muốn */
                                border: 1px solid #016b67;

                            }

                            /* === NÚT BẤM === */
                            QPushButton {
                                background: qlineargradient(
                                    spread:pad, x1:0, y1:0, x2:1, y2:1,
                                    stop:0 #016b67, stop:1 #0b9f87, stop:2 #00c3cc
                                );
                                color: white;
                                border: none;
                                border-radius: 6px;
                                padding: 6px;
                            }
                            QPushButton:hover {
                                background-color: #016b67;
                                color: white;
                            }

                            /* === NHÃN === */
                            QLabel {
                                color: black;
                                font-weight: bold;
                                border: none;
                            }

                            /* === Ô COMBOBOX / LINEEDIT / TEXTEDIT / TABLE === */
                            QComboBox,
                            QLineEdit,
                            QTextEdit,
                            QTableWidget {
                                background-color: white;
                                border: 1px solid #016b67;
                                border-radius: 4px;
                                color: black;
                            }

                            /* === SCROLLBAR VERTICAL (5px) === */
                            QScrollBar:vertical {
                                border: none;
                                background: #f0f0f0;
                                width: 5px;             /* mỏng 5px */
                                margin: 0px;
                                border-radius: 2px;
                            }
                            QScrollBar::handle:vertical {
                                background-color: #c0c0c0;   /* xám nhạt */
                                min-height: 20px;
                                border-radius: 2px;
                            }
                            QScrollBar::handle:vertical:hover {
                                background-color: #a0a0a0;   /* xám đậm hơn khi hover */
                            }
                            QScrollBar::add-line:vertical,
                            QScrollBar::sub-line:vertical {
                                height: 0px;
                                background: none;
                            }

                            /* === SCROLLBAR HORIZONTAL (5px) === */
                            QScrollBar:horizontal {
                                border: none;
                                background: #f0f0f0;
                                height: 5px;            /* mỏng 5px */
                                margin: 0px;
                                border-radius: 2px;
                            }
                            QScrollBar::handle:horizontal {
                                background-color: #c0c0c0;   /* xám nhạt */
                                min-width: 20px;
                                border-radius: 2px;
                            }
                            QScrollBar::handle:horizontal:hover {
                                background-color: #a0a0a0;   /* xám đậm hơn khi hover */
                            }
                            QScrollBar::add-line:horizontal,
                            QScrollBar::sub-line:horizontal {
                                width: 0px;
                                background: none;
                            }
                        """)

        # đổ bóng
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(20)
        shadow.setOffset(0, 2)
        shadow.setColor(QColor(0, 0, 0, 80))
        self.setGraphicsEffect(shadow)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)

        # cuộn nội dung
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        layout.addWidget(scroll)

        content = QWidget()
        grid = QGridLayout(content)
        grid.setSpacing(15)
        scroll.setWidget(content)

        # thêm các nút (icon + text)
        self.label1 = QLabel("Số Seri:", self)
        grid.addWidget(self.label1, 0, 0)
        self.lineEdit1 = QLineEdit(self)
        self.lineEdit1.setFixedWidth(150)
        grid.addWidget(self.lineEdit1, 0, 1)

        self.label2 = QLabel("Loại tài sản:", self)
        grid.addWidget(self.label2, 1, 0)
        self.comboBox1 = QComboBox(self)
        self.comboBox1.addItem("All")
        self.comboBox1.addItems(["CPU", "Laptop", "Monitor", "Printer", "Scanner"])
        self.comboBox1.setFixedWidth(150)
        grid.addWidget(self.comboBox1, 1, 1)

        self.label3 = QLabel("Tình trạng:", self)
        grid.addWidget(self.label3, 2, 0)
        self.comboBox2 = QComboBox(self)
        self.comboBox2.addItem("All")
        self.comboBox2.addItems(["Sử dụng bình thường", "Hỏng"])
        self.comboBox2.setFixedWidth(150)
        grid.addWidget(self.comboBox2, 2, 1)

        self.label4 = QLabel("Mã hiệu:", self)
        grid.addWidget(self.label4, 0, 2)
        self.lineEdit2 = QLineEdit(self)
        self.lineEdit2.setFixedWidth(150)
        grid.addWidget(self.lineEdit2, 0, 3)

        self.label5 = QLabel("Ngày nhận tài sản", self)
        grid.addWidget(self.label5, 1, 2)
        self.lineEdit3 = QLineEdit(self)
        self.lineEdit3.setFixedWidth(150)
        grid.addWidget(self.lineEdit3, 1, 3)

        self.label6 = QLabel("Tình trạng sử dụng:", self)
        grid.addWidget(self.label6, 2, 2)
        self.comboBox3 = QComboBox(self)
        self.comboBox3.addItem("All")
        self.comboBox3.addItems(["Đang sử dụng", "Đã thu hồi trả văn phòng"])
        self.comboBox3.setFixedWidth(150)
        grid.addWidget(self.comboBox3, 2, 3)

        self.label8 = QLabel("Ngày nhận trả sản", self)
        grid.addWidget(self.label8, 3, 0)
        self.lineEdit6 = QLineEdit(self)
        self.lineEdit6.setFixedWidth(150)
        grid.addWidget(self.lineEdit6, 3, 1)

        self.label9 = QLabel("Trị giá:", self)
        grid.addWidget(self.label9, 3, 2)
        self.lineEdit7 = QLineEdit(self)
        self.lineEdit7.setFixedWidth(150)
        grid.addWidget(self.lineEdit7, 3, 3)

        self.label10 = QLabel("Cán bộ giữ tài sản:", self)
        grid.addWidget(self.label10, 4, 0)
        self.lineEdit8 = QLineEdit(self)
        self.lineEdit8.setFixedWidth(150)
        grid.addWidget(self.lineEdit8, 4, 1)

        self.label11 = QLabel("Phòng ban:", self)
        grid.addWidget(self.label11, 4, 2)
        self.comboBox4 = QComboBox(self)
        self.comboBox4.addItem("All")
        self.comboBox4.addItems(["NK1", "NK2", "NK3", "NK4", "LCND", "XK1", "XK2"])
        self.comboBox4.setFixedWidth(150)
        grid.addWidget(self.comboBox4, 4, 3)

        self.label7 = QLabel("Ghi chú:", self)
        grid.addWidget(self.label7, 5, 0)
        self.lineEdit5 = QLineEdit(self)
        self.lineEdit5.setFixedWidth(150)
        grid.addWidget(self.lineEdit5, 5, 1)

        self.button1 = QPushButton("Tìm kiếm", self)
        self.button1.setFixedWidth(150)
        grid.addWidget(self.button1, 6, 1)
        self.button1.clicked.connect(self.search_table)

        self.assets_table = assets_table
        self.assets_table2 = assets_table2
        self.assets_table3 = assets_table3
        self.assets_table4 = assets_table4

    def search_table(self):
        table = self.assets_table
        table2 = self.assets_table2
        table3 = self.assets_table3
        table4 = self.assets_table4

        serial_number = self.lineEdit1.text().strip().lower()
        asset_name  = self.comboBox1.currentText()
        status = self.comboBox2.currentText()
        code = self.lineEdit2.text().strip().lower()
        purchase_date = self.lineEdit3.text().strip().lower()
        active_asset = self.comboBox3.currentText()
        return_date = self.lineEdit6.text().strip().lower()
        value  = self.lineEdit7.text().strip().lower()
        name = self.lineEdit8.text().strip().lower()
        department = self.comboBox4.currentText()
        notes = self.lineEdit5.text().strip().lower()

        # Duyệt từng hàng trong table
        for row in range(table.rowCount()):
            match = True  # mặc định là đúng
            # Kiểm tra từng cột
            if serial_number and serial_number not in table.item(row, 3).text().lower():
                match = False
            if asset_name != "All" and asset_name != table.item(row, 1).text():
                match = False
            if status != "All" and status != table.item(row, 5).text():
                match = False
            if code and code not in table.item(row, 2).text().lower():
                match = False
            if purchase_date and purchase_date not in table.item(row, 4).text().lower():
                match = False
            if active_asset != "All" and active_asset != table.item(row, 6).text():
                match = False
            if return_date and return_date not in table.item(row, 7).text().lower():
                match = False
            if value and value not in table.item(row, 8).text().lower():
                match = False
            if name and name not in table.item(row, 9).text().lower():
                match = False
            if department != "All" and department != table.item(row, 10).text():
                match = False
            if notes and notes not in table.item(row, 12).text().lower():
                match = False

            # Hiển thị hoặc ẩn hàng tùy vào kết quả tìm kiếm
            table.setRowHidden(row, not match)

        for row in range(table2.rowCount()):
            match = True  # mặc định là đúng
            # Kiểm tra từng cột
            if serial_number and serial_number not in table2.item(row, 3).text().lower():
                match = False
            if asset_name != "All" and asset_name != table2.item(row, 1).text():
                match = False
            if status != "All" and status != table2.item(row, 6).text():
                match = False
            if code and code not in table2.item(row, 2).text().lower():
                match = False
            if purchase_date and purchase_date not in table2.item(row, 5).text().lower():
                match = False
            if value and value not in table2.item(row, 7).text().lower():
                match = False
            if name and name not in table2.item(row, 4).text().lower():
                match = False
            if department != "All" and department != table2.item(row, 8).text():
                match = False
            if notes and notes not in table2.item(row, 9).text().lower():
                match = False
            if active_asset != "All" and active_asset != "Đang sử dụng":
                match = False
            if return_date:
                match = False

            # Hiển thị hoặc ẩn hàng tùy vào kết quả tìm kiếm
            table2.setRowHidden(row, not match)

        for row in range(table3.rowCount()):
            match = True  # mặc định là đúng
            # Kiểm tra từng cột
            if serial_number and serial_number not in table3.item(row, 3).text().lower():
                match = False
            if asset_name != "All" and asset_name != table3.item(row, 1).text():
                match = False
            if status != "All" and status != table3.item(row, 5).text():
                match = False
            if code and code not in table3.item(row, 2).text().lower():
                match = False
            if purchase_date and purchase_date not in table3.item(row, 4).text().lower():
                match = False
            if value and value not in table3.item(row, 6).text().lower():
                match = False
            if name:
                match = False
            if department != "All":
                match = False
            if notes and notes not in table3.item(row, 7).text().lower():
                match = False
            if active_asset != "All" and active_asset != "Đang sử dụng":
                match = False
            if return_date:
                match = False

            # Hiển thị hoặc ẩn hàng tùy vào kết quả tìm kiếm
            table3.setRowHidden(row, not match)

        for row in range(table4.rowCount()):
            match = True  # mặc định là đúng
            # Kiểm tra từng cột
            if serial_number and serial_number not in table4.item(row, 3).text().lower():
                match = False
            if asset_name != "All" and asset_name != table4.item(row, 1).text():
                match = False
            if status != "All" and status != table4.item(row, 5).text():
                match = False
            if code and code not in table4.item(row, 2).text().lower():
                match = False
            if purchase_date and purchase_date not in table4.item(row, 4).text().lower():
                match = False
            if value and value not in table4.item(row, 7).text().lower():
                match = False
            if name:
                match = False
            if department != "All":
                match = False
            if notes and notes not in table4.item(row, 8).text().lower():
                match = False
            if active_asset != "All" and active_asset != "Đã thu hồi trả văn phòng":
                match = False
            if return_date and return_date not in table4.item(row, 6).text().lower():
                match = False

            # Hiển thị hoặc ẩn hàng tùy vào kết quả tìm kiếm
            table4.setRowHidden(row, not match)

class asset_assignments_window(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Bàn giao tài sản")
        self.setGeometry(300, 80, 1410, 850)
        self.setWindowIcon(QIcon("avatar_app.png"))
        self.setStyleSheet("""
            /* === NỀN CỬA SỔ === */
            QWidget {
                background-color: #f0f0f0;
                color: black;
            }

            /* === NÚT BẤM === */
            QPushButton {
                background: qlineargradient(
                    spread:pad, x1:0, y1:0, x2:1, y2:1,
                    stop:0 #016b67, stop:1 #0b9f87, stop:2 #00c3cc
                );
                color: white;
                border: none;
                border-radius: 6px;
                padding: 6px;
            }
            QPushButton:hover {
                background-color: #016b67;
                color: white;
            }

            /* === NHÃN === */
            QLabel {
                color: black;
                font-weight: bold;
            }

            /* === Ô COMBOBOX / LINEEDIT / TEXTEDIT / TABLE === */
            QComboBox,
            QLineEdit,
            QTextEdit,
            QTableWidget {
                background-color: white;
                border: 1px solid #016b67;
                border-radius: 4px;
                color: black;
            }

            /* === SCROLLBAR VERTICAL (5px) === */
            QScrollBar:vertical {
                border: none;
                background: #f0f0f0;
                width: 5px;             /* mỏng 5px */
                margin: 0px;
                border-radius: 2px;
            }
            QScrollBar::handle:vertical {
                background-color: #c0c0c0;   /* xám nhạt */
                min-height: 20px;
                border-radius: 2px;
            }
            QScrollBar::handle:vertical:hover {
                background-color: #a0a0a0;   /* xám đậm hơn khi hover */
            }
            QScrollBar::add-line:vertical,
            QScrollBar::sub-line:vertical {
                height: 0px;
                background: none;
            }

            /* === SCROLLBAR HORIZONTAL (5px) === */
            QScrollBar:horizontal {
                border: none;
                background: #f0f0f0;
                height: 5px;            /* mỏng 5px */
                margin: 0px;
                border-radius: 2px;
            }
            QScrollBar::handle:horizontal {
                background-color: #c0c0c0;   /* xám nhạt */
                min-width: 20px;
                border-radius: 2px;
            }
            QScrollBar::handle:horizontal:hover {
                background-color: #a0a0a0;   /* xám đậm hơn khi hover */
            }
            QScrollBar::add-line:horizontal,
            QScrollBar::sub-line:horizontal {
                width: 0px;
                background: none;
            }
        """)
        # grid = QGridLayout()

        # Giao tài sản
        self.label1 = QLabel("Giao tài sản:", self)
        self.label1.setGeometry(200, 20, 150, 22)
        # grid.addWidget(self.label1, 0, 1)

        self.label2 = QLabel("Cán bộ:", self)
        self.label2.setGeometry(40, 50, 150, 22)
        # grid.addWidget(self.label2, 1, 0)
        self.lineEdit1 = QLineEdit(self)
        self.lineEdit1.setFixedWidth(150)
        self.lineEdit1.setGeometry(200, 50, 150, 22)
        self.lineEdit1.setPlaceholderText("MCB, Email hoặc Tên")
        # grid.addWidget(self.lineEdit1, 1, 1)

        self.label3 = QLabel("Họ và Tên:", self)
        self.label3.setGeometry(40, 80, 150, 22)
        # grid.addWidget(self.label3, 2, 0)
        self.lineEdit2 = QLineEdit(self)
        self.lineEdit2.setFixedWidth(150)
        self.lineEdit2.setGeometry(200, 80, 150, 22)
        self.lineEdit2.setReadOnly(True)
        self.lineEdit2.setStyleSheet("background-color: #cccccc;")
        # grid.addWidget(self.lineEdit2, 2, 1)

        self.label4 = QLabel("Ghi chú:", self)
        self.label4.setGeometry(40, 110, 150, 22)
        # grid.addWidget(self.label4, 3, 0)
        self.lineEdit3 = QLineEdit(self)
        self.lineEdit3.setFixedWidth(230)
        self.lineEdit3.setGeometry(200, 110, 230, 22)
        # self.lineEdit3.setText("Trung tâm Tác nghiệp tài trợ thương mại")
        # self.lineEdit3.setReadOnly(True)
        # self.lineEdit3.setStyleSheet("background-color: #cccccc;")
        # grid.addWidget(self.lineEdit3, 3, 1)

        self.label5 = QLabel("Thêm tài sản bàn giao:", self)
        self.label5.setGeometry(40, 140, 150, 22)
        # grid.addWidget(self.label5, 4, 0)
        self.lineEdit4 = QLineEdit(self)
        self.lineEdit4.setFixedWidth(150)
        self.lineEdit4.setGeometry(200, 140, 150, 22)
        self.lineEdit4.setPlaceholderText("Số Seri")
        # grid.addWidget(self.lineEdit4, 4, 1)

        self.button1 = QPushButton("+", self)
        self.button1.clicked.connect(self.button1_clicked)
        self.button1.setFixedWidth(50)
        self.button1.setGeometry(360, 140, 50, 25)
        # grid.addWidget(self.button1, 4, 2)

        self.button2 = QPushButton("-", self)
        self.button2.clicked.connect(self.button2_clicked)
        self.button2.setFixedWidth(50)
        self.button2.setGeometry(420, 140, 50, 25)
        # grid.addWidget(self.button2, 4, 3)

        self.table1 = QTableWidget()
        self.table1.setColumnCount(8)
        self.table1.setHorizontalHeaderLabels([
            "ID tài sản", "Tên cán bộ nhận tài sản", "ID cán bộ", "Loại tài sản", "Mã hiệu", "Số seri", "Tình trạng",
            "Ghi chú"
        ])
        self.table1.setParent(self)
        self.table1.setGeometry(40, 170, 620, 120)
        self.table1.setColumnWidth(0, 60)
        self.table1.setColumnWidth(1, 150)
        self.table1.setColumnWidth(5, 150)
        self.table1.setColumnWidth(6, 120)
        self.table1.setColumnWidth(7, 300)
        self.table1.verticalHeader().setVisible(False)
        self.table1.setSortingEnabled(False)
        self.table1.setEditTriggers(QAbstractItemView.NoEditTriggers)

        # Trả tài sản
        self.label6 = QLabel("Trả tài sản:", self)
        self.label6.setGeometry(910, 20, 150, 22)
        # grid.addWidget(self.label6, 0, 7)

        self.label7 = QLabel("Cán bộ:", self)
        self.label7.setGeometry(750, 50, 150, 22)
        # grid.addWidget(self.label7, 1, 6)
        self.lineEdit5 = QLineEdit(self)
        self.lineEdit5.setFixedWidth(150)
        self.lineEdit5.setGeometry(910, 50, 150, 22)
        self.lineEdit5.setPlaceholderText("MCB, Email hoặc Tên")
        # grid.addWidget(self.lineEdit5, 1, 7)

        self.label8 = QLabel("Họ và Tên:", self)
        self.label8.setGeometry(750, 80, 150, 22)
        # grid.addWidget(self.label8, 2, 6)
        self.lineEdit6 = QLineEdit(self)
        self.lineEdit6.setFixedWidth(150)
        self.lineEdit6.setGeometry(910, 80, 150, 22)
        self.lineEdit6.setReadOnly(True)
        self.lineEdit6.setStyleSheet("background-color: #cccccc;")
        # grid.addWidget(self.lineEdit6, 2, 7)

        self.label9 = QLabel("Ghi chú:", self)
        self.label9.setGeometry(750, 110, 150, 22)
        # grid.addWidget(self.label9, 3, 6)
        self.lineEdit7 = QLineEdit(self)
        self.lineEdit7.setFixedWidth(230)
        self.lineEdit7.setGeometry(910, 110, 230, 22)
        # grid.addWidget(self.lineEdit7, 3, 7)

        self.label10 = QLabel("Thêm tài sản bàn giao:", self)
        self.label10.setGeometry(750, 140, 150, 22)
        # grid.addWidget(self.label10, 4, 6)
        self.lineEdit8 = QLineEdit(self)
        self.lineEdit8.setFixedWidth(150)
        self.lineEdit8.setGeometry(910, 140, 230, 22)
        self.lineEdit8.setPlaceholderText("Số Seri")
        # grid.addWidget(self.lineEdit8, 4, 7)

        self.button3 = QPushButton("+", self)
        self.button3.clicked.connect(self.button3_clicked)
        self.button3.setFixedWidth(50)
        self.button3.setGeometry(1070, 140, 50, 25)
        # grid.addWidget(self.button3, 4, 8)

        self.button4 = QPushButton("-", self)
        self.button4.clicked.connect(self.button4_clicked)
        self.button4.setFixedWidth(50)
        self.button4.setGeometry(1130, 140, 50, 25)
        # grid.addWidget(self.button4, 4, 9)

        self.table2 = QTableWidget()
        self.table2.setColumnCount(9)
        self.table2.setHorizontalHeaderLabels([
            "ID tài sản", "Tên cán bộ trả tài sản", "ID cán bộ", "Loại tài sản", "Mã hiệu", "Số seri", "Tình trạng",
            "Ngày giao tài sản", "Ghi chú"
        ])
        # grid.addWidget(self.table2, 5, 7, 3, 4)
        self.table2.setParent(self)
        self.table2.setGeometry(750, 170, 620, 120)
        self.table2.setColumnWidth(0, 60)
        self.table2.setColumnWidth(1, 150)
        self.table2.setColumnWidth(5, 150)
        self.table2.setColumnWidth(6, 120)
        self.table2.setColumnWidth(8, 300)
        self.table2.verticalHeader().setVisible(False)
        self.table2.setSortingEnabled(False)
        self.table2.setEditTriggers(QAbstractItemView.NoEditTriggers)

        # Bảng tài sản trong kho
        self.table3 = QTableWidget()
        self.table3.setColumnCount(7)
        self.table3.setHorizontalHeaderLabels([
            "ID", "Loại tài sản", "Mã hiệu", "Số seri", "Tình trạng", "Năm nhận tài sản", "Ghi chú"
        ])
        # grid.addWidget(self.table3, 10, 0, 3,4)
        self.table3.setParent(self)
        self.table3.setGeometry(40, 330, 620, 120)
        self.table3.setColumnWidth(0, 40)
        self.table3.setColumnWidth(3, 150)
        self.table3.setColumnWidth(4, 120)
        self.table3.setColumnWidth(6, 300)
        self.table3.verticalHeader().setVisible(False)
        self.table3.setSortingEnabled(True)

        # Tìm kiếm tài sản trong kho
        self.lineEdit9 = QLineEdit(self)
        self.lineEdit9.setFixedWidth(150)
        self.lineEdit9.setGeometry(40, 300, 150, 22)
        self.lineEdit9.setPlaceholderText("Tìm kiếm trong kho")
        self.lineEdit9.textChanged.connect(self.search_table3)

        # Bảng tài sản của cán bộ
        self.table4 = QTableWidget()
        self.table4.setColumnCount(8)
        self.table4.setHorizontalHeaderLabels([
            "ID", "Loại tài sản", "Mã hiệu", "Số seri", "Tình trạng", "Năm nhận tài sản", "Cán bộ giữ tài sản",
            "Ghi chú"
        ])
        # grid.addWidget(self.table4, 10, 7, 3, 4)
        self.table4.setParent(self)
        self.table4.setGeometry(750, 330, 620, 120)
        self.table4.setColumnWidth(0, 40)
        self.table4.setColumnWidth(3, 150)
        self.table4.setColumnWidth(4, 120)
        self.table4.setColumnWidth(6, 120)
        self.table4.setColumnWidth(7, 300)
        self.table4.verticalHeader().setVisible(False)
        self.table4.setSortingEnabled(True)

        # Tìm kiếm tài sản cán bộ
        self.lineEdit10 = QLineEdit(self)
        self.lineEdit10.setFixedWidth(150)
        self.lineEdit10.setPlaceholderText("Tìm kiếm tài sản cán bộ")
        self.lineEdit10.setGeometry(750, 300, 150, 22)
        self.lineEdit10.textChanged.connect(self.search_table4)

        # Nút bàn giao
        self.button5 = QPushButton("Thực hiện bàn giao", self)
        self.button5.clicked.connect(self.button5_clicked)
        self.button5.setFixedWidth(200)
        self.button5.setFixedHeight(50)
        self.button5.setGeometry(40, 455, 200, 50)
        # grid.addWidget(self.button5, 14, 0)

        # Nút xuất excel
        self.button6 = QPushButton("Xuất file để nhập excel", self)
        self.button6.clicked.connect(self.button6_clicked)
        self.button6.setGeometry(1150, 525, 200, 30)
        # grid.addWidget(self.button6, 15, 0)

        # Nút nhập excel
        self.button7 = QPushButton("Nhập bàn giao theo excel", self)
        self.button7.clicked.connect(self.button7_clicked)
        self.button7.setGeometry(940, 525, 200, 30)
        # grid.addWidget(self.button6, 15, 1)

        # Tìm kiếm lịch sử bàn giao
        self.lineEdit11 = QLineEdit(self)
        self.lineEdit11.setFixedWidth(150)
        self.lineEdit11.setGeometry(40, 550, 150, 22)
        self.lineEdit11.setPlaceholderText("Tìm kiếm lịch sử bàn giao")
        self.lineEdit11.textChanged.connect(self.search_table5)
        self.lineEdit11.textChanged.connect(self.search_table6)
        # grid.addWidget(self.lineEdit11, 16, 0)

        # Bảng lịch sử bàn giao
        self.table5 = QTableWidget()
        self.table5.setColumnCount(9)
        self.table5.setHorizontalHeaderLabels([
            "ID", "ID tài sản", "Loại tài sản", "Số Seri", "ID cán bộ", "Tên cán bộ", "Ngày bàn giao",
            "Ngày trả về kho", "Ghi chú bàn giao"
        ])
        self.table5.setParent(self)
        self.table5.setGeometry(40, 570, 1330, 150)
        self.table5.setColumnWidth(0, 40)
        self.table5.setColumnWidth(2, 150)
        self.table5.setColumnWidth(3, 200)
        self.table5.setColumnWidth(5, 200)
        self.table5.setColumnWidth(8, 320)
        self.table5.verticalHeader().setVisible(False)
        self.table5.setSortingEnabled(True)

        # Bảng cán bộ nắm giữ tài sản
        self.table6 = QTableWidget()
        self.table6.setColumnCount(11)
        self.table6.setHorizontalHeaderLabels([
            "ID cán bộ", "Tên cán bộ", "Nhóm", "Chức vụ", "ID tài sản", "Loại tài sản", "Mã hiệu", "Số Seri",
            "Tình trạng tài sản", "Ngày bàn giao", "Ghi chú bàn giao tài sản"
        ])
        self.table6.setParent(self)
        self.table6.setGeometry(40, 570, 1330, 150)
        self.table6.setColumnWidth(0, 60)
        self.table6.setColumnWidth(2, 70)
        self.table6.setColumnWidth(5, 80)
        self.table6.setColumnWidth(7, 150)
        self.table6.setColumnWidth(8, 150)
        self.table6.setColumnWidth(10, 312)
        self.table6.verticalHeader().setVisible(False)
        self.table6.setSortingEnabled(True)

        # Thêm tab bảng
        self.tab_widget = QTabWidget(self)
        self.tab_widget.addTab(self.table6, "Cán bộ nắm giữ tài sản")
        self.tab_widget.addTab(self.table5, "Lịch sử bàn giao")
        self.tab_widget.setGeometry(40, 580, 1330, 260)
        self.tab_widget.setStyleSheet("""
        QTabBar::tab {
            background: #cccccc;      /* màu tab chưa active */
            color: black;
            padding: 5px;
            border: 1px solid #016b67;
            border-top-left-radius: 4px;
            border-top-right-radius: 4px;
            margin-right: 3px;
        }

        QTabBar::tab:selected {
            background: white;          /* màu tab active */
            color: black;
        }

        QTabBar::tab:pressed {
        background: #d0d0d0;        /* màu tab khi nhấn, tạo hiệu ứng chìm */
        padding-top: 6px;           /* đẩy chữ xuống tạo cảm giác nhấn */
        padding-bottom: 4px;
        }

        QTabWidget::pane {
        border: 1px solid gray;          /* kẻ khung quanh toàn bộ tab */
        top: -1px;                       /* kéo pane lên ăn khớp với tab */
        }

        """)

        # label để thông báo
        self.label11 = QLabel("Đây là cửa sổ để bàn giao tài sản", self)
        self.label11.setGeometry(480, 455, 500, 50)

        # Nút xuất excel bảng assets_assignment
        self.button8 = QPushButton("Xuất bảng asset_assignments", self)
        self.button8.clicked.connect(self.button8_clicked)
        self.button8.setGeometry(1150, 565, 200, 30)

        # Nút nhập excel bảng assets_assignment
        self.button9 = QPushButton("Nhập bảng asset_assignments", self)
        self.button9.clicked.connect(self.button9_clicked)
        self.button9.setGeometry(940, 565, 200, 30)

        # Nút xuất excel bảng tài sản theo cán bộ
        self.button10 = QPushButton("Xuất bảng tài sản theo cán bộ", self)
        self.button10.clicked.connect(self.button10_clicked)
        self.button10.setGeometry(730, 565, 200, 30)

        # Nút xuất biên bản bàn giao
        self.button11 = QPushButton("Xuất biên bản bàn giao", self)
        self.button11.clicked.connect(self.button11_clicked)
        self.button11.setGeometry(260, 455, 200, 50)
        self.button11.setFixedWidth(200)
        self.button11.setFixedHeight(50)

        # Nút chức năng
        self.function = QPushButton("Bật chức năng", self)
        self.function.clicked.connect(self.show_buttons)
        self.function.setGeometry(40, 510, 150, 30)
        self.function.setFixedWidth(150)

        for i, btn in enumerate([self.button6, self.button7, self.button8, self.button9, self.button10], start=1):
            btn.setVisible(False)


        self.load_assets_warehouse()
        self.update_completer_staff()
        self.update_completer_assign_warehouse_asset()
        self.update_completer_staff_2()
        self.load_assets_assignment_history()
        self.load_staff_holding_assets()

    def showEvent(self, event):
        """Hàm này sẽ chạy mỗi khi window được show"""
        super().showEvent(event)  # gọi hàm gốc
        self.load_assets_warehouse()
        self.update_completer_staff()
        self.update_completer_assign_warehouse_asset()
        self.update_completer_staff_2()
        self.load_assets_assignment_history()
        self.load_staff_holding_assets()

    def load_assets_assignment_history(self):
        conn = sqlite3.connect(db_path_default)
        cursor = conn.cursor()

        cursor.execute("""
                       SELECT aa.id,
                              aa.asset_id,
                              a.asset_name,
                              a.serial_number,
                              aa.staff_id,
                              s.name AS staff_name,
                              aa.assigned_date,
                              aa.returned_date,
                              aa.notes
                       FROM asset_assignments aa
                                LEFT JOIN assets a ON aa.asset_id = a.id
                                LEFT JOIN staff s ON aa.staff_id = s.id
                       """)

        rows = cursor.fetchall()
        conn.close()

        self.table5.setRowCount(0)

        for row_idx, row_data in enumerate(rows):
            self.table5.insertRow(row_idx)
            for col_idx, value in enumerate(row_data):
                item = QTableWidgetItem()
                if isinstance(value, int):
                    item.setData(Qt.ItemDataRole.DisplayRole, value)
                else:
                    item.setText(str(value) if value else "")
                self.table5.setItem(row_idx, col_idx, item)

    def load_staff_holding_assets(self):
        conn = sqlite3.connect(db_path_default)
        cursor = conn.cursor()
        cursor.execute("""
                       SELECT s.id     AS staff_id,
                              s.name   AS staff_name,
                              s.department,
                              s.position,
                              a.id     AS asset_id,
                              a.asset_name,
                              a.code,
                              a.serial_number,
                              a.status,
                              aa.assigned_date,
                              aa.notes AS assignment_notes
                       FROM asset_assignments aa
                                JOIN staff s ON aa.staff_id = s.id
                                JOIN assets a ON aa.asset_id = a.id
                       WHERE aa.returned_date IS NULL
                       ORDER BY s.id, s.name, a.asset_name;
                       """)
        rows = cursor.fetchall()
        conn.close()

        self.table6.setRowCount(0)

        for row_idx, row_data in enumerate(rows):
            self.table6.insertRow(row_idx)
            for col_idx, value in enumerate(row_data):
                item = QTableWidgetItem()
                if isinstance(value, int):
                    item.setData(Qt.ItemDataRole.DisplayRole, value)
                else:
                    item.setText(str(value) if value else "")
                self.table6.setItem(row_idx, col_idx, item)

    def load_assets_warehouse(self):
        conn = sqlite3.connect(db_path_default)
        cursor = conn.cursor()
        cursor.execute("""
                       SELECT a.id AS asset_id,
                              a.asset_name,
                              a.code,
                              a.serial_number,
                              a.status,
                              a.purchase_date,
                              a.notes
                       FROM assets a
                                LEFT JOIN asset_assignments aa ON aa.asset_id = a.id AND aa.returned_date IS NULL
                                LEFT JOIN staff s ON aa.staff_id = s.id
                       WHERE s.name IS NULL
                         AND a.active_asset = 1
                       """)
        rows = cursor.fetchall()
        conn.close()

        self.table3.setRowCount(0)

        for row_idx, row_data in enumerate(rows):
            self.table3.insertRow(row_idx)
            for col_idx, value in enumerate(row_data):
                item = QTableWidgetItem()
                if isinstance(value, int):
                    item.setData(Qt.ItemDataRole.DisplayRole, value)
                else:
                    item.setText(str(value) if value else "")
                self.table3.setItem(row_idx, col_idx, item)

    def search_table3(self, text):
        text = text.strip().lower()
        for row in range(self.table3.rowCount()):
            match = False
            for col in range(self.table3.columnCount()):
                item = self.table3.item(row, col)
                if item and text in item.text().lower():
                    match = True
                    break
            self.table3.setRowHidden(row, not match)

    def search_table4(self, text):
        text = text.strip().lower()
        for row in range(self.table4.rowCount()):
            match = False
            for col in range(self.table4.columnCount()):
                item = self.table4.item(row, col)
                if item and text in item.text().lower():
                    match = True
                    break
            self.table4.setRowHidden(row, not match)

    def search_table5(self, text):
        text = text.strip().lower()
        for row in range(self.table5.rowCount()):
            match = False
            for col in range(self.table5.columnCount()):
                item = self.table5.item(row, col)
                if item and text in item.text().lower():
                    match = True
                    break
            self.table5.setRowHidden(row, not match)

    def search_table6(self, text):
        text = text.strip().lower()
        for row in range(self.table6.rowCount()):
            match = False
            for col in range(self.table6.columnCount()):
                item = self.table6.item(row, col)
                if item and text in item.text().lower():
                    match = True
                    break
            self.table6.setRowHidden(row, not match)

    def update_completer_staff(self):
        conn = sqlite3.connect(db_path_default)
        cursor = conn.cursor()
        cursor.execute("SELECT name, employee_code, email FROM staff WHERE active_staff = 1")
        rows = cursor.fetchall()
        conn.close()

        suggestions = []
        for row in rows:
            for value in row:
                if value:
                    suggestions.append(str(value))
        suggestions = list(set(suggestions))

        self.completer = QCompleter(suggestions, self)
        self.completer.setCaseSensitivity(Qt.CaseInsensitive)
        self.completer.setFilterMode(Qt.MatchContains)
        self.lineEdit5.setCompleter(self.completer)
        self.lineEdit5.returnPressed.connect(self.staff_search_selected)
        self.completer.activated[str].connect(self.staff_search_selected)

    def update_completer_staff_2(self):
        conn = sqlite3.connect(db_path_default)
        cursor = conn.cursor()
        cursor.execute("SELECT name, employee_code, email FROM staff WHERE active_staff = 1")
        rows = cursor.fetchall()
        conn.close()

        suggestions = []
        for row in rows:
            for value in row:
                if value:
                    suggestions.append(str(value))
        suggestions = list(set(suggestions))

        self.completer = QCompleter(suggestions, self)
        self.completer.setCaseSensitivity(Qt.CaseInsensitive)
        self.completer.setFilterMode(Qt.MatchContains)
        self.lineEdit1.setCompleter(self.completer)
        self.lineEdit1.returnPressed.connect(self.staff_search_selected_2)
        self.completer.activated[str].connect(self.staff_search_selected_2)

    def update_completer_assign_warehouse_asset(self):
        conn = sqlite3.connect(db_path_default)
        cursor = conn.cursor()
        cursor.execute("SELECT serial_number FROM assets WHERE active_asset = 1")
        rows = cursor.fetchall()
        conn.close()

        suggestions = []
        for row in rows:
            for value in row:
                if value:
                    suggestions.append(str(value))
        suggestions = list(set(suggestions))

        self.completer = QCompleter(suggestions, self)
        self.completer.setCaseSensitivity(Qt.CaseInsensitive)
        self.completer.setFilterMode(Qt.MatchContains)
        self.lineEdit4.setCompleter(self.completer)

    def staff_search_selected(self):
        search_text = self.lineEdit5.text()
        result = search_assets_of_staff(search_text)
        if isinstance(result, str):  # chuỗi thông báo
            self.label11.setText(result)
            self.lineEdit6.setText("")
            self.table4.setRowCount(0)
            return
        if not result:
            self.label11.setText("Không tìm thấy dữ liệu")
            self.lineEdit6.setText("")
            self.table4.setRowCount(0)
            return

        self.label11.setText("Đã tìm thấy kết quả")
        self.lineEdit6.setText(result[0][6])

        rows = result
        self.table4.setRowCount(0)

        for row_idx, row_data in enumerate(rows):
            self.table4.insertRow(row_idx)
            for col_idx, value in enumerate(row_data):
                item = QTableWidgetItem()
                if isinstance(value, int):
                    item.setData(Qt.ItemDataRole.DisplayRole, value)
                else:
                    item.setText(str(value) if value else "")
                self.table4.setItem(row_idx, col_idx, item)

        self.update_completer_staff()
        self.update_completer_staff_2()

    def staff_search_selected_2(self):
        search_text = self.lineEdit1.text()
        result = search_staff(search_text)

        if isinstance(result, str):  # chuỗi thông báo
            self.label11.setText(result)
            return

        self.lineEdit2.setText(result[1])
        self.label11.setText("Đã tìm thấy cán bộ")



    def check_password(self, action_name="thao tác"):
        password, ok = QInputDialog.getText(
            self,
            "Xác nhận mật khẩu",
            f"Nhập mật khẩu để {action_name}:",
            echo=QLineEdit.Password
        )
        if not ok:
            return False

        if password == "123456":
            return True
        else:
            QMessageBox.warning(self, "Sai mật khẩu", "Mật khẩu không đúng!")
            return False

    def show_buttons(self):
        if not self.check_password("Bật chức năng"):
            self.label11.setText("Đã huỷ thao tác bật chức năng.")
            return

        # Lấy trạng thái hiện/ẩn dựa vào nút 1
        currently_visible = self.button10.isVisible()

        # Chuyển trạng thái ngược lại cho tất cả nút
        for btn in [self.button6, self.button7, self.button8, self.button9, self.button10]:
            btn.setVisible(not currently_visible)
        if currently_visible:
            self.function.setText("Bật chức năng")
        else:
            self.function.setText("Tắt chức năng")

    def button1_clicked(self):
        search_text = self.lineEdit4.text()
        result = search_serial_number_in_warehouse(search_text)
        search = self.lineEdit1.text()
        search = search.strip()

        conn = sqlite3.connect(db_path_default)
        cursor = conn.cursor()

        cursor.execute("""
                       SELECT *
                       FROM staff
                       WHERE name = ?
                          OR employee_code = ?
                          OR email = ?
                       """, (search, search, search))
        row = cursor.fetchone()
        conn.close()

        if not row:
            self.label11.setText("Thông tin cán bộ không đúng")
            return
        staff_assign_name = row[1]
        staff_assign_id = row[0]

        if isinstance(result, str):  # chuỗi thông báo
            self.label11.setText(result)
            return
        if not result:
            self.label11.setText("Không tìm thấy dữ liệu")
            return

        rows = result
        self.label11.setText("Đã tìm thấy kết quả")

        for row_data in rows:
            id_value = row_data[0]  # cột 0 (id)

            # kiểm tra trùng id ở cột 0 của table1
            duplicate = False
            for row_index in range(self.table1.rowCount()):
                item = self.table1.item(row_index, 0)
                if item and str(item.text()) == str(id_value):
                    duplicate = True
                    break
            if duplicate:
                self.label11.setText("Tài sản vừa rồi đã có trong bảng giao tài sản!")
                continue

            insert_row_index = self.table1.rowCount()
            self.table1.insertRow(insert_row_index)

            # cột 0: id
            item_id = QTableWidgetItem()
            item_id.setData(Qt.ItemDataRole.DisplayRole, id_value)
            self.table1.setItem(insert_row_index, 0, item_id)

            # cột 1: name (thêm mới)
            item_name = QTableWidgetItem(str(staff_assign_name))
            self.table1.setItem(insert_row_index, 1, item_name)

            # cột 2: id (thêm mới)
            item_name = QTableWidgetItem(str(staff_assign_id))
            self.table1.setItem(insert_row_index, 2, item_name)

            # các cột còn lại từ cột 2 trở đi
            for col_idx, value in enumerate(row_data[1:], start=3):
                item = QTableWidgetItem()
                if isinstance(value, int):
                    item.setData(Qt.ItemDataRole.DisplayRole, value)
                else:
                    item.setText(str(value) if value else "")
                self.table1.setItem(insert_row_index, col_idx, item)

            item_note = QTableWidgetItem(str(self.lineEdit3.text()))
            self.table1.setItem(insert_row_index, 7, item_note)

        self.update_completer_assign_warehouse_asset()

    def button2_clicked(self):
        search_text = self.lineEdit4.text()
        result = search_serial_number_in_warehouse(search_text)

        if isinstance(result, str):  # chuỗi thông báo
            self.label11.setText(result)
            return
        if not result:
            self.label11.setText("Không tìm thấy dữ liệu")
            return

        delete_asset_id = str(result[0][0])
        deleted = False
        for row_index in reversed(range(self.table1.rowCount())):
            item = self.table1.item(row_index, 0)  # lấy ô cột 0
            if item and item.text() == delete_asset_id:
                self.table1.removeRow(row_index)
                deleted = True

        if deleted:
            self.label11.setText("Xoá giao tài sản thành công")
        else:
            self.label11.setText("Không tìm thấy dữ liệu để xoá ở bảng giao tài sản")

    def button3_clicked(self):
        search_serial_number = self.lineEdit8.text()
        search_staff = self.lineEdit5.text()
        result = search_serial_number_of_staff(search_serial_number, search_staff)

        if isinstance(result, str):  # chuỗi thông báo
            self.label11.setText(result)
            return
        if not result:
            self.label11.setText("Không tìm thấy dữ liệu")
            return

        rows = result
        self.label11.setText("Đã tìm thấy kết quả")

        for row_data in rows:
            id_value = row_data[0]  # cột 0 (id)

            # kiểm tra trùng id ở cột 0 của table1
            duplicate = False
            for row_index in range(self.table2.rowCount()):
                item = self.table2.item(row_index, 0)
                if item and str(item.text()) == str(id_value):
                    duplicate = True
                    break
            if duplicate:
                self.label11.setText("Tài sản vừa rồi đã có trong bảng trả tài sản!")
                continue

            insert_row_index = self.table2.rowCount()
            self.table2.insertRow(insert_row_index)

            # cột 0: id
            item_id = QTableWidgetItem()
            item_id.setData(Qt.ItemDataRole.DisplayRole, id_value)
            self.table2.setItem(insert_row_index, 0, item_id)

            # các cột còn lại từ cột 1 trở đi
            for col_idx, value in enumerate(row_data[1:], start=1):
                item = QTableWidgetItem()
                if isinstance(value, int):
                    item.setData(Qt.ItemDataRole.DisplayRole, value)
                else:
                    item.setText(str(value) if value else "")
                self.table2.setItem(insert_row_index, col_idx, item)

            item_note = QTableWidgetItem(str(self.lineEdit7.text()))
            self.table2.setItem(insert_row_index, 8, item_note)

    def button4_clicked(self):
        search_serial_number = self.lineEdit8.text()
        search_staff = self.lineEdit5.text()
        result = search_serial_number_of_staff(search_serial_number, search_staff)

        if isinstance(result, str):  # chuỗi thông báo
            self.label11.setText(result)
            return
        if not result:
            self.label11.setText("Không tìm thấy dữ liệu")
            return

        delete_asset_id = str(result[0][0])
        deleted = False
        for row_index in reversed(range(self.table2.rowCount())):
            item = self.table2.item(row_index, 0)  # lấy ô cột 0
            if item and item.text() == delete_asset_id:
                self.table2.removeRow(row_index)
                deleted = True

        if deleted:
            self.label11.setText("Xoá trả tài sản thành công")
        else:
            self.label11.setText("Không tìm thấy dữ liệu để xoá ở bảng trả tài sản")

    def button5_clicked(self):
        if not self.check_password("Bàn giao"):
            self.label11.setText("Đã huỷ thao tác bàn giao.")
            return

        assign_assets(self.table1, self.label11)
        msg = self.label11.text()
        return_assets(self.table2, self.label11)
        msg += "\n" + self.label11.text()

        self.lineEdit1.setText("")
        self.lineEdit2.setText("")
        self.lineEdit3.setText("")
        self.lineEdit4.setText("")
        self.lineEdit5.setText("")
        self.lineEdit6.setText("")
        self.lineEdit7.setText("")
        self.lineEdit8.setText("")
        self.table1.setRowCount(0)
        self.table2.setRowCount(0)
        self.table4.setRowCount(0)

        self.label11.setText(msg)
        self.load_assets_warehouse()
        self.update_completer_staff()
        self.update_completer_assign_warehouse_asset()
        self.update_completer_staff_2()
        self.load_assets_assignment_history()
        self.load_staff_holding_assets()

    def button6_clicked(self):
        export_two_tables_to_excel(self.table1, self.table2)

    def button7_clicked(self):
        import_two_tables_from_excel_to_widgets(self.table1, self.table2)

    def button8_clicked(self):
        export_assets_assignments_to_excel()

    def button9_clicked(self):
        import_assets_assignments_from_excel()
        self.load_assets_warehouse()
        self.update_completer_staff()
        self.update_completer_assign_warehouse_asset()
        self.update_completer_staff_2()
        self.load_assets_assignment_history()
        self.load_staff_holding_assets()

    def button10_clicked(self):
        export_assets_of_staffs_to_excel(self.table6)

    def button11_clicked(self):
        try:
            doc = Document("mau_bienban.docx")

            # Tự động điền ngày tháng năm hiện tại vào {{DAY}}, {{MONTH}}, {{YEAR}}
            fill_date_placeholders(doc)

            # tìm placeholder {{TABLE1}}, {{TABLE2}}, {{SIGN}}
            for para in doc.paragraphs:
                if "{{TABLE1}}" in para.text:
                    para.text = ""  # xóa placeholder
                    add_table_from_qtable(doc, self.table1, after_paragraph=para)

                if "{{TABLE2}}" in para.text:
                    para.text = ""
                    add_table_from_qtable(doc, self.table2, after_paragraph=para)

                if "{{SIGN}}" in para.text:
                    para.text = ""
                    # chèn bảng chữ ký
                    add_sign_table(doc, self.table1, self.table2, after_paragraph=para)

            # Điền tên người nhận/trả vào ô {{SIGN}} trong bảng chữ ký và chèn thêm dòng nếu có nhiều người
            fill_sign_placeholder_in_tables(doc, self.table1, self.table2)

            # --- tạo tên file mặc định ---
            today = datetime.now().strftime("%Y-%m-%d")
            default_name = f"Biên bản bàn giao {today}.docx"

            # --- tìm thư mục Downloads ---
            downloads_path = QStandardPaths.writableLocation(QStandardPaths.DownloadLocation)
            if not downloads_path or not os.path.exists(downloads_path):
                downloads_path = os.getcwd()  # thư mục hiện tại của app

            default_path = os.path.join(downloads_path, default_name)

            # --- hộp thoại lưu ---
            options = QFileDialog.Options()
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "Lưu file biên bản bàn giao",
                default_path,
                "Word Files (*.docx);;All Files (*)",
                options=options
            )
            if not file_path:
                self.label11.setText("Đã huỷ thao tác xuất biên bản bàn giao")
                return
            if not file_path.endswith(".docx"):
                file_path += ".docx"

            doc.save(file_path)
            self.label11.setText(f"Đã lưu biên bản bàn giao: {file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"\n{str(e)}")

class handover_documents_window(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Biên bản bàn giao")
        self.setGeometry(350, 120, 1200, 750)
        self.setWindowIcon(QIcon("avatar_app.png"))
        self.setStyleSheet("""
            /* === NỀN CỬA SỔ === */
            QWidget {
                background-color: #f0f0f0;
                color: black;
            }

            /* === NÚT BẤM === */
            QPushButton {
                background: qlineargradient(
                    spread:pad, x1:0, y1:0, x2:1, y2:1,
                    stop:0 #016b67, stop:1 #0b9f87, stop:2 #00c3cc
                );
                color: white;
                border: none;
                border-radius: 6px;
                padding: 6px;
            }
            QPushButton:hover {
                background-color: #016b67;
                color: white;
            }

            /* === NHÃN === */
            QLabel {
                color: black;
                font-weight: bold;
            }

            /* === Ô COMBOBOX / LINEEDIT / TEXTEDIT / TABLE === */
            QComboBox,
            QLineEdit,
            QTextEdit,
            QTableWidget {
                background-color: white;
                border: 1px solid #016b67;
                border-radius: 4px;
                color: black;
            }

            /* === SCROLLBAR VERTICAL (5px) === */
            QScrollBar:vertical {
                border: none;
                background: #f0f0f0;
                width: 5px;             /* mỏng 5px */
                margin: 0px;
                border-radius: 2px;
            }
            QScrollBar::handle:vertical {
                background-color: #c0c0c0;   /* xám nhạt */
                min-height: 20px;
                border-radius: 2px;
            }
            QScrollBar::handle:vertical:hover {
                background-color: #a0a0a0;   /* xám đậm hơn khi hover */
            }
            QScrollBar::add-line:vertical,
            QScrollBar::sub-line:vertical {
                height: 0px;
                background: none;
            }

            /* === SCROLLBAR HORIZONTAL (5px) === */
            QScrollBar:horizontal {
                border: none;
                background: #f0f0f0;
                height: 5px;            /* mỏng 5px */
                margin: 0px;
                border-radius: 2px;
            }
            QScrollBar::handle:horizontal {
                background-color: #c0c0c0;   /* xám nhạt */
                min-width: 20px;
                border-radius: 2px;
            }
            QScrollBar::handle:horizontal:hover {
                background-color: #a0a0a0;   /* xám đậm hơn khi hover */
            }
            QScrollBar::add-line:horizontal,
            QScrollBar::sub-line:horizontal {
                width: 0px;
                background: none;
            }
        """)

        grid = QGridLayout(self)
        self.pixmap_item = None
        self.view = ImageViewer()
        self.scene = QGraphicsScene(self)
        self.view.setScene(self.scene)
        grid.addWidget(self.view, 0, 0, 8, 4)
        self.view.setStyleSheet("""
            border: 2px solid gray;
            background-color: white;
        """)
        self.placeholder = self.scene.addText("Ảnh sẽ hiển thị ở đây…")

        self.button1 = QPushButton("Import ảnh", self)
        self.button1.clicked.connect(self.load_image)
        self.button1.setFixedWidth(150)
        grid.addWidget(self.button1, 10, 1)

        self.label1 = QLabel("Lưu biên bản đã ký", self)
        grid.addWidget(self.label1, 10, 0)

        self.label2 = QLabel("Ngày ký bàn giao:", self)
        grid.addWidget(self.label2, 11, 0)
        self.dateEdit = QDateEdit(self)
        self.dateEdit.setFixedWidth(150)
        self.dateEdit.setDisplayFormat("yyyy-MM-dd")
        self.dateEdit.setDate(QDate.currentDate())
        self.dateEdit.setCalendarPopup(True)
        grid.addWidget(self.dateEdit, 11, 1)

        self.label3 = QLabel("Ghi chú:", self)
        grid.addWidget(self.label3, 12, 0)
        self.lineEdit1 = QLineEdit(self)
        self.lineEdit1.setFixedWidth(150)
        grid.addWidget(self.lineEdit1, 12, 1)

        self.label4 = QLabel("Loại bàn giao:", self)
        grid.addWidget(self.label4, 13, 0)
        self.comboBox1 = QComboBox(self)
        self.comboBox1.addItems(["Bàn giao cho cán bộ TFC", "Bàn giao với văn phòng", "Bàn giao khác"])
        self.comboBox1.setFixedWidth(170)
        grid.addWidget(self.comboBox1, 13, 1)

        self.button4 = QPushButton("Lưu biên bản bàn giao", self)
        grid.addWidget(self.button4, 14, 1)
        self.button4.setFixedWidth(150)
        self.button4.clicked.connect(self.button4_clicked)

        self.button5 = QPushButton("Xem ảnh biên bản với ID:", self)
        grid.addWidget(self.button5, 9, 2)
        self.button5.setFixedWidth(150)
        self.button5.clicked.connect(self.button5_clicked)
        self.lineEdit3 = QLineEdit(self)
        self.lineEdit3.setFixedWidth(50)
        self.lineEdit3.setPlaceholderText("...")
        grid.addWidget(self.lineEdit3, 9, 3)

        self.label5 = QLabel("Gắn với giao dịch bàn giao:", self)
        grid.addWidget(self.label5, 0, 5)

        self.label6 = QLabel("ID biên bản bàn giao:", self)
        grid.addWidget(self.label6, 1, 4)
        self.lineEdit2 = QLineEdit(self)
        self.lineEdit2.setFixedWidth(150)
        grid.addWidget(self.lineEdit2, 1, 5)

        self.label7 = QLabel("ID assignment:", self)
        grid.addWidget(self.label7, 2, 4)
        self.lineEdit4 = QLineEdit(self)
        self.lineEdit4.setFixedWidth(150)
        grid.addWidget(self.lineEdit4, 2, 5)

        self.label8 = QLabel("Giao tài sản/Trả tài sản:", self)
        grid.addWidget(self.label8, 3, 4)
        self.comboBox2 = QComboBox(self)
        self.comboBox2.addItems(["Giao tài sản", "Trả tài sản"])
        self.comboBox2.setFixedWidth(150)
        grid.addWidget(self.comboBox2, 3, 5)

        self.button2 = QPushButton("+", self)
        grid.addWidget(self.button2, 3, 6)
        self.button2.setFixedWidth(40)
        self.button2.clicked.connect(self.button2_clicked)
        self.button3 = QPushButton("-", self)
        grid.addWidget(self.button3, 3, 7)
        self.button3.setFixedWidth(40)
        self.button3.clicked.connect(self.button3_clicked)

        self.button6 = QPushButton("Update vào ID assignment", self)
        grid.addWidget(self.button6, 4, 5)
        self.button6.setFixedWidth(150)
        self.button6.clicked.connect(self.button6_clicked)

        self.table1 = QTableWidget()
        self.table1.setColumnCount(8)
        self.table1.setHorizontalHeaderLabels([
            "ID assignment", "ID bàn giao nhận", "ID bàn giao trả", "Ngày giao tài sản", "Ngày trả tài sản",
            "Tên cán bộ", "Loại tài sản", "Số seri"
        ])
        grid.addWidget(self.table1, 5, 4, 1, 7)
        self.table1.verticalHeader().setVisible(False)
        self.table1.setSortingEnabled(True)

        self.lineEdit6 = QLineEdit(self)
        self.lineEdit6.setFixedWidth(150)
        self.lineEdit6.setPlaceholderText("Tìm kiếm lịch sử bàn giao")
        self.lineEdit6.textChanged.connect(self.search_table2)
        grid.addWidget(self.lineEdit6, 6, 4)

        self.table2 = QTableWidget()
        self.table2.setColumnCount(9)
        self.table2.setHorizontalHeaderLabels([
            "ID assignment", "Họ và tên", "Loại tài sản", "Số seri", "Ngày giao tài sản", "ID biên bản giao",
            "Ngày trả tài sản", "ID biên bản trả", "Ghi chú"
        ])
        grid.addWidget(self.table2, 7, 4, 1, 7)
        self.table2.verticalHeader().setVisible(False)
        self.table2.setSortingEnabled(True)

        self.lineEdit7 = QLineEdit(self)
        self.lineEdit7.setFixedWidth(150)
        self.lineEdit7.setPlaceholderText("Tìm kiếm bảng biên bản")
        self.lineEdit7.textChanged.connect(self.search_table3)
        grid.addWidget(self.lineEdit7, 9, 10)

        self.table3 = QTableWidget()
        self.table3.setColumnCount(4)
        self.table3.setHorizontalHeaderLabels([
            "ID", "Ngày biên bản", "Loại biên bản", "Ghi chú"
        ])
        grid.addWidget(self.table3, 10, 2, 5, 9)
        self.table3.verticalHeader().setVisible(False)
        self.table3.setSortingEnabled(True)
        self.table3.setColumnWidth(1, 150)
        self.table3.setColumnWidth(2, 200)
        self.table3.setColumnWidth(3, 400)

        self.label = QLabel("Đây là cửa sổ để lưu biên bản bàn giao đã ký", self)
        self.label.setFixedWidth(150)
        self.label.setFixedHeight(100)
        self.label.setWordWrap(True)
        grid.addWidget(self.label, 0, 10)

        self.load_assets_assignment_history()
        self.load_handover_documents()

    def showEvent(self, event):
        """Hàm này sẽ chạy mỗi khi window được show"""
        super().showEvent(event)  # gọi hàm gốc
        self.load_assets_assignment_history()
        self.load_handover_documents()

    def load_image(self):
        # Lấy đường dẫn thư mục Downloads mặc định của hệ điều hành
        downloads_dir = QStandardPaths.writableLocation(QStandardPaths.DownloadLocation)
        if not downloads_dir:  # fallback nếu không lấy được
            downloads_dir = str(Path.home() / "Downloads")

        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn ảnh",
            downloads_dir,  # <--- đường dẫn mặc định
            "Images (*.png *.jpg *.jpeg)"  # filter
        )
        if file_path:
            if self.placeholder:
                self.scene.removeItem(self.placeholder)
                self.placeholder = None
            pixmap = QPixmap(file_path)
            self.scene.clear()  # xóa scene cũ
            self.pixmap_item = QGraphicsPixmapItem(pixmap)
            self.scene.addItem(self.pixmap_item)

            # Đặt scene vừa khung nhìn ban đầu
            self.view.fitInView(self.pixmap_item, Qt.KeepAspectRatio)
            self.view.scale_factor = 1.0

    def load_handover_documents(self):
        conn = sqlite3.connect(db_path_default)
        cursor = conn.cursor()

        cursor.execute("""
                       SELECT id, document_date, type_handover_documents, notes
                       FROM handover_documents
                       """)
        rows = cursor.fetchall()
        conn.close()

        self.table3.setRowCount(0)
        self.table3.setRowCount(len(rows))
        for row_idx, row_data in enumerate(rows):
            for col_idx, value in enumerate(row_data):
                item = QTableWidgetItem(str(value) if value is not None else "")
                self.table3.setItem(row_idx, col_idx, item)

    def load_assets_assignment_history(self):
        conn = sqlite3.connect(db_path_default)
        cursor = conn.cursor()

        cursor.execute("""
                       SELECT aa.id,
                              s.name AS staff_name,
                              a.asset_name,
                              a.serial_number,
                              aa.assigned_date,
                              aa.handover_documents_assign_id,
                              aa.returned_date,
                              aa.handover_documents_return_id,
                              aa.notes
                       FROM asset_assignments aa
                                LEFT JOIN assets a ON aa.asset_id = a.id
                                LEFT JOIN staff s ON aa.staff_id = s.id
                       """)

        rows = cursor.fetchall()
        conn.close()

        self.table2.setRowCount(0)

        for row_idx, row_data in enumerate(rows):
            self.table2.insertRow(row_idx)
            for col_idx, value in enumerate(row_data):
                item = QTableWidgetItem()
                if isinstance(value, int):
                    item.setData(Qt.ItemDataRole.DisplayRole, value)
                else:
                    item.setText(str(value) if value else "")
                self.table2.setItem(row_idx, col_idx, item)

    def check_password(self, action_name="thao tác"):
        password, ok = QInputDialog.getText(
            self,
            "Xác nhận mật khẩu",
            f"Nhập mật khẩu để {action_name}:",
            echo=QLineEdit.Password
        )
        if not ok:
            return False

        if password == "123456":
            return True
        else:
            QMessageBox.warning(self, "Sai mật khẩu", "Mật khẩu không đúng!")
            return False

    def button2_clicked(self):
        handover_documents_id = self.lineEdit2.text()
        asset_assignments_id = self.lineEdit4.text()
        type = self.comboBox2.currentText()

        result = search_asset_assignments_id_and_handover_documents_id(handover_documents_id, asset_assignments_id)

        if isinstance(result, str):  # chuỗi thông báo
            self.label.setText(result)
            return

        self.label.setText("Đã tìm thấy kết quả")

        if result == 0:
            pass  # không có gì gắn, cứ tiếp tục
        else:
            # chỉ check nếu là list
            if isinstance(result, list):
                if 1 in result and type == "Giao tài sản":
                    reply = QMessageBox.question(
                        self,
                        "Xác nhận",
                        "Biên bản giao tài sản đã được gắn với ID này. Bạn có muốn tiếp tục không?",
                        QMessageBox.Yes | QMessageBox.No,
                        QMessageBox.No
                    )
                    if reply == QMessageBox.No:
                        self.label.setText("Đã hủy thao tác theo yêu cầu người dùng.")
                        return

                if 2 in result and type == "Trả tài sản":
                    reply = QMessageBox.question(
                        self,
                        "Xác nhận",
                        "Biên bản trả tài sản đã được gắn với ID này. Bạn có muốn tiếp tục không?",
                        QMessageBox.Yes | QMessageBox.No,
                        QMessageBox.No
                    )
                    if reply == QMessageBox.No:
                        self.label.setText("Đã hủy thao tác theo yêu cầu người dùng.")
                        return

        # Lấy ID assignment/handover
        assignment_id = asset_assignments_id
        handover_id = handover_documents_id

        found_duplicate = False
        for r in range(self.table1.rowCount()):
            existing_assignment_id = self.table1.item(r, 0).text() if self.table1.item(r, 0) else ""
            if (existing_assignment_id == assignment_id):
                found_duplicate = True
                break

        if found_duplicate:
            self.label.setText("Bản ghi ID assignment đã có trong bảng, không thêm.")
            return

        # Thêm một dòng mới vào table1
        row = self.table1.rowCount()
        self.table1.insertRow(row)

        # cột 0: ID assignment (luôn lấy từ lineEdit3)
        item0 = QTableWidgetItem(assignment_id)
        self.table1.setItem(row, 0, item0)

        # cột 1 hoặc 2 tùy giá trị combobox
        if self.comboBox2.currentText() == "Giao tài sản":
            # cột 1: ID bàn giao nhận
            self.table1.setItem(row, 1, QTableWidgetItem(handover_id))
            # cột 2 để trống
            self.table1.setItem(row, 2, QTableWidgetItem(""))
        elif self.comboBox2.currentText() == "Trả tài sản":
            # cột 2: ID bàn giao trả
            self.table1.setItem(row, 2, QTableWidgetItem(handover_id))
            # cột 1 để trống
            self.table1.setItem(row, 1, QTableWidgetItem(""))

        # Lấy dữ liệu từ DB cho asset_assignments_id (các cột 3–7)
        conn = sqlite3.connect(db_path_default)
        cursor = conn.cursor()
        cursor.execute("""
                       SELECT aa.assigned_date,
                              aa.returned_date,
                              s.name AS staff_name,
                              a.asset_name,
                              a.serial_number
                       FROM asset_assignments aa
                                LEFT JOIN staff s ON aa.staff_id = s.id
                                LEFT JOIN assets a ON aa.asset_id = a.id
                       WHERE aa.id = ?
                       """, (asset_assignments_id,))
        row_data = cursor.fetchone()
        conn.close()

        # Đổ dữ liệu cột 3–7 vào dòng vừa thêm
        if row_data:
            for offset, value in enumerate(row_data, start=3):  # cột 3→7
                item = QTableWidgetItem(str(value) if value is not None else "")
                self.table1.setItem(row, offset, item)

        self.lineEdit2.setText("")
        self.lineEdit4.setText("")

    def button3_clicked(self):
        assignment_id = self.lineEdit4.text()
        assignment_id = assignment_id.strip()

        if not assignment_id:
            self.label.setText("Chưa nhập ID assignment để xóa.")
            return

        # tìm dòng có ID này ở cột 0
        row_to_delete = -1
        for r in range(self.table1.rowCount()):
            item = self.table1.item(r, 0)
            if item and item.text() == assignment_id:
                row_to_delete = r
                break

        if row_to_delete == -1:
            self.label.setText(f"Không tìm thấy ID {assignment_id} trong bảng.")
            return

        self.table1.removeRow(row_to_delete)
        self.label.setText(f"Đã xóa ID {assignment_id} khỏi bảng.")
        self.lineEdit2.setText("")
        self.lineEdit4.setText("")

    def button4_clicked(self):
        if not self.check_password("Lưu biên bản bàn giao"):
            self.label.setText("Đã huỷ thao tác lưu biên bản bàn giao.")
            return

        result = add_handover_documents(
            self.pixmap_item,
            self.dateEdit.date().toString("yyyy-MM-dd"),
            self.comboBox1.currentText(),
            self.lineEdit1.text()
        )

        if isinstance(result, str):  # chuỗi thông báo
            self.label.setText(result)
            return

        self.label.setText("Đã lưu biên bản bàn giao")

        if self.pixmap_item:
            self.scene.removeItem(self.pixmap_item)
            self.pixmap_item = None
        self.lineEdit1.setText("")
        self.comboBox1.setCurrentText("Bàn giao cho cán bộ TFC")

        # Thêm placeholder căn giữa và font lớn
        text_item = self.scene.addText("Ảnh sẽ hiển thị ở đây…")

        # Đặt font lớn hơn
        font = QFont()
        font.setPointSize(50)  # ví dụ 18pt
        text_item.setFont(font)
        # text_item.setTextWidth(self.view.width())

        # căn giữa text bên trong item
        text_option = text_item.document().defaultTextOption()
        text_option.setAlignment(Qt.AlignCenter)
        text_item.document().setDefaultTextOption(text_option)

        # đặt item vào giữa scene
        scene_rect = self.scene.sceneRect()
        text_rect = text_item.boundingRect()
        text_item.setPos(
            (scene_rect.width() - text_rect.width()) / 2,
            (scene_rect.height() - text_rect.height()) / 2
        )
        self.placeholder = text_item

        self.load_handover_documents()
        self.load_assets_assignment_history()

    def button5_clicked(self):

        id_text = self.lineEdit3.text().strip()
        if not id_text.isdigit():
            self.label.setText("ID biên bản không hợp lệ")
            return

        doc_id = int(id_text)

        import sqlite3
        conn = sqlite3.connect(db_path_default)
        cursor = conn.cursor()
        cursor.execute("SELECT document_file FROM handover_documents WHERE id = ?", (doc_id,))
        row = cursor.fetchone()
        conn.close()

        if not row or not row[0]:
            self.label.setText("Không tìm thấy ảnh trong DB")
            return

        image_bytes = row[0]

        # Xóa placeholder cũ
        if self.placeholder:
            self.scene.removeItem(self.placeholder)
            self.placeholder = None

        # Tạo QPixmap từ dữ liệu bytes
        pixmap = QPixmap()
        pixmap.loadFromData(image_bytes)

        # Xóa scene cũ và add pixmap mới
        self.scene.clear()
        self.pixmap_item = QGraphicsPixmapItem(pixmap)
        self.scene.addItem(self.pixmap_item)

        # Fit vào view
        self.view.fitInView(self.pixmap_item, Qt.KeepAspectRatio)
        self.view.scale_factor = 1.0

        self.label.setText(f"Đã load ảnh cho biên bản ID {doc_id}")

    def button6_clicked(self):
        if not self.check_password("Gán biên bản với assignment"):
            self.label.setText("Đã huỷ thao tác gán biên bản với assignment.")
            return
        add_handover_document_to_assignment(self.table1, self.label)
        self.table1.setRowCount(0)
        self.load_assets_assignment_history()
        self.load_handover_documents()

    def search_table2(self, text):
        text = text.strip().lower()
        for row in range(self.table2.rowCount()):
            match = False
            for col in range(self.table2.columnCount()):
                item = self.table2.item(row, col)
                if item and text in item.text().lower():
                    match = True
                    break
            self.table2.setRowHidden(row, not match)

    def search_table3(self, text):
        text = text.strip().lower()
        for row in range(self.table3.rowCount()):
            match = False
            for col in range(self.table3.columnCount()):
                item = self.table3.item(row, col)
                if item and text in item.text().lower():
                    match = True
                    break
            self.table3.setRowHidden(row, not match)


class ImageViewer(QGraphicsView):
    def __init__(self):
        super().__init__()
        # dùng QPainter.Antialiasing thay vì self.viewport().paintEngine()
        self.setRenderHint(QPainter.Antialiasing)
        self.setDragMode(QGraphicsView.ScrollHandDrag)  # kéo ảnh bằng chuột trái
        self.scale_factor = 1.0

    def wheelEvent(self, event):
        """Zoom bằng con lăn chuột"""
        if event.angleDelta().y() > 0:
            zoom_in_factor = 1.25
            self.scale_factor *= zoom_in_factor
            self.scale(zoom_in_factor, zoom_in_factor)
        else:
            zoom_out_factor = 0.8
            self.scale_factor *= zoom_out_factor
            self.scale(zoom_out_factor, zoom_out_factor)


class user_manual_window(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Hướng dẫn sử dụng")
        self.setGeometry(200, 200, 500, 500)
        self.setWindowIcon(QIcon("avatar_app.png"))
        layout = QGridLayout(self)
        self.setStyleSheet("""
            /* === NỀN CỬA SỔ === */
            QWidget {
                background-color: #f0f0f0;
                color: black;
            }

            /* === NÚT BẤM === */
            QPushButton {
                background: qlineargradient(
                    spread:pad, x1:0, y1:0, x2:1, y2:1,
                    stop:0 #016b67, stop:1 #0b9f87, stop:2 #00c3cc
                );
                color: white;
                border: none;
                border-radius: 6px;
                padding: 6px;
            }
            QPushButton:hover {
                background-color: #016b67;
                color: white;
            }

            /* === NHÃN === */
            QLabel {
                color: black;
                font-weight: bold;
            }

            /* === Ô COMBOBOX / LINEEDIT / TEXTEDIT / TABLE === */
            QComboBox,
            QLineEdit,
            QTextEdit,
            QTableWidget {
                background-color: white;
                border: 1px solid #016b67;
                border-radius: 4px;
                color: black;
            }

            /* === SCROLLBAR VERTICAL (5px) === */
            QScrollBar:vertical {
                border: none;
                background: #f0f0f0;
                width: 5px;             /* mỏng 5px */
                margin: 0px;
                border-radius: 2px;
            }
            QScrollBar::handle:vertical {
                background-color: #c0c0c0;   /* xám nhạt */
                min-height: 20px;
                border-radius: 2px;
            }
            QScrollBar::handle:vertical:hover {
                background-color: #a0a0a0;   /* xám đậm hơn khi hover */
            }
            QScrollBar::add-line:vertical,
            QScrollBar::sub-line:vertical {
                height: 0px;
                background: none;
            }

            /* === SCROLLBAR HORIZONTAL (5px) === */
            QScrollBar:horizontal {
                border: none;
                background: #f0f0f0;
                height: 5px;            /* mỏng 5px */
                margin: 0px;
                border-radius: 2px;
            }
            QScrollBar::handle:horizontal {
                background-color: #c0c0c0;   /* xám nhạt */
                min-width: 20px;
                border-radius: 2px;
            }
            QScrollBar::handle:horizontal:hover {
                background-color: #a0a0a0;   /* xám đậm hơn khi hover */
            }
            QScrollBar::add-line:horizontal,
            QScrollBar::sub-line:horizontal {
                width: 0px;
                background: none;
            }
        """)

        # Tạo QTextEdit chỉ đọc để hiển thị hướng dẫn
        self.manual_text = QTextEdit(self)
        self.manual_text.setReadOnly(True)

        # Nội dung hướng dẫn
        huong_dan = """
        <h2>Cấu trúc cơ sở dữ liệu</h2>

        <h3>1. Bảng staff</h3>
        <table border="1" cellspacing="0" cellpadding="4">
          <tr><th>Tên cột</th><th>Kiểu dữ liệu</th><th>Ràng buộc</th></tr>
          <tr><td>id</td><td>INTEGER</td><td>PRIMARY KEY AUTOINCREMENT</td></tr>
          <tr><td>name</td><td>TEXT</td><td>NOT NULL</td></tr>
          <tr><td>department</td><td>TEXT</td><td>NOT NULL</td></tr>
          <tr><td>position</td><td>TEXT</td><td></td></tr>
          <tr><td>email</td><td>TEXT</td><td>NOT NULL UNIQUE</td></tr>
          <tr><td>employee_code</td><td>TEXT</td><td>NOT NULL UNIQUE</td></tr>
          <tr><td>phone</td><td>TEXT</td><td></td></tr>
          <tr><td>active_staff</td><td>BOOLEAN</td><td>NOT NULL</td></tr>
          <tr><td>notes</td><td>TEXT</td><td></td></tr>
        </table>

        <h3>2. Bảng assets</h3>
        <table border="1" cellspacing="0" cellpadding="4">
          <tr><th>Tên cột</th><th>Kiểu dữ liệu</th><th>Ràng buộc</th></tr>
          <tr><td>id</td><td>INTEGER</td><td>PRIMARY KEY AUTOINCREMENT</td></tr>
          <tr><td>asset_name</td><td>TEXT</td><td>NOT NULL</td></tr>
          <tr><td>code</td><td>TEXT</td><td></td></tr>
          <tr><td>serial_number</td><td>TEXT</td><td>NOT NULL UNIQUE</td></tr>
          <tr><td>purchase_date</td><td>TEXT</td><td></td></tr>
          <tr><td>status</td><td>TEXT</td><td></td></tr>
          <tr><td>active_asset</td><td>BOOLEAN</td><td>NOT NULL</td></tr>
          <tr><td>return_date</td><td>TEXT</td><td></td></tr>
          <tr><td>notes</td><td>TEXT</td><td></td></tr>
        </table>

        <h3>3. Bảng asset_assignments</h3>
        <table border="1" cellspacing="0" cellpadding="4">
          <tr><th>Tên cột</th><th>Kiểu dữ liệu</th><th>Ràng buộc</th></tr>
          <tr><td>id</td><td>INTEGER</td><td>PRIMARY KEY AUTOINCREMENT</td></tr>
          <tr><td>asset_id</td><td>INTEGER</td><td>NOT NULL, FK → assets(id)</td></tr>
          <tr><td>staff_id</td><td>INTEGER</td><td>NOT NULL, FK → staff(id)</td></tr>
          <tr><td>assigned_date</td><td>TEXT</td><td>NOT NULL</td></tr>
          <tr><td>returned_date</td><td>TEXT</td><td></td></tr>
          <tr><td>notes</td><td>TEXT</td><td></td></tr>
          <tr><td>handover_documents_assign_id</td><td>INTEGER</td><td>FK → handover_documents(id)</td></tr>
          <tr><td>handover_documents_return_id</td><td>INTEGER</td><td>FK → handover_documents(id)</td></tr>
        </table>

        <h3>4. Bảng handover_documents</h3>
        <table border="1" cellspacing="0" cellpadding="4">
          <tr><th>Tên cột</th><th>Kiểu dữ liệu</th><th>Ràng buộc</th></tr>
          <tr><td>id</td><td>INTEGER</td><td>PRIMARY KEY AUTOINCREMENT</td></tr>
          <tr><td>document_date</td><td>TEXT</td><td>NOT NULL</td></tr>
          <tr><td>type_handover_documents</td><td>TEXT</td><td></td></tr>
          <tr><td>document_file</td><td>BLOB</td><td></td></tr>
          <tr><td>notes</td><td>TEXT</td><td></td></tr>
        </table>

        <h3>5. Bảng chat_box</h3>
        <table border="1" cellspacing="0" cellpadding="4">
          <tr><th>Tên cột</th><th>Kiểu dữ liệu</th><th>Ràng buộc</th></tr>
          <tr><td>id</td><td>INTEGER</td><td>PRIMARY KEY AUTOINCREMENT</td></tr>
          <tr><td>chat</td><td>TEXT</td><td>NOT NULL</td></tr>
          <tr><td>date_and_time</td><td>TEXT</td><td>NOT NULL</td></tr>
          <tr><td>user_chat</td><td>TEXT</td><td>NOT NULL</td></tr>
        </table>
        """
        self.manual_text.setHtml(huong_dan)

        layout.addWidget(self.manual_text, 0, 0)
        self.setLayout(layout)


# === Bộ tạo mã QR chuẩn ISO/IEC 18004 (Byte Mode, ECC Level M, Version 1-6) ===
_QR_EXP = [0] * 512
_QR_LOG = [0] * 256
_qr_x = 1
for _qr_i in range(255):
    _QR_EXP[_qr_i] = _qr_x
    _QR_LOG[_qr_x] = _qr_i
    _qr_x <<= 1
    if _qr_x & 0x100:
        _qr_x ^= 0x11D
for _qr_i in range(255, 512):
    _QR_EXP[_qr_i] = _QR_EXP[_qr_i - 255]

_QR_MASK_FUNCS = (
    lambda r, c: (r + c) % 2 == 0,
    lambda r, c: r % 2 == 0,
    lambda r, c: c % 3 == 0,
    lambda r, c: (r + c) % 3 == 0,
    lambda r, c: (r // 2 + c // 3) % 2 == 0,
    lambda r, c: (r * c) % 2 + (r * c) % 3 == 0,
    lambda r, c: ((r * c) % 2 + (r * c) % 3) % 2 == 0,
    lambda r, c: ((r + c) % 2 + (r * c) % 3) % 2 == 0,
)


def _qr_gf_mul(a: int, b: int) -> int:
    if a == 0 or b == 0:
        return 0
    return _QR_EXP[_QR_LOG[a] + _QR_LOG[b]]


def _qr_rs_encode(msg: list[int], ec_len: int) -> list[int]:
    gen = [1]
    for i in range(ec_len):
        nxt = [0] * (len(gen) + 1)
        factor = _QR_EXP[i]
        for j, coef in enumerate(gen):
            nxt[j] ^= coef
            nxt[j + 1] ^= _qr_gf_mul(coef, factor)
        gen = nxt
    rem = list(msg) + [0] * ec_len
    for i in range(len(msg)):
        coef = rem[i]
        if coef != 0:
            for j in range(1, len(gen)):
                rem[i + j] ^= _qr_gf_mul(gen[j], coef)
    return rem[len(msg):]


def _qr_penalty(mat: list[list[int]]) -> int:
    n = len(mat)
    score = 0
    for r in range(n):
        run_color = mat[r][0]
        run_len = 1
        for c in range(1, n):
            if mat[r][c] == run_color:
                run_len += 1
            else:
                if run_len >= 5:
                    score += 3 + (run_len - 5)
                run_color = mat[r][c]
                run_len = 1
        if run_len >= 5:
            score += 3 + (run_len - 5)
    for c in range(n):
        run_color = mat[0][c]
        run_len = 1
        for r in range(1, n):
            if mat[r][c] == run_color:
                run_len += 1
            else:
                if run_len >= 5:
                    score += 3 + (run_len - 5)
                run_color = mat[r][c]
                run_len = 1
        if run_len >= 5:
            score += 3 + (run_len - 5)
    for r in range(n - 1):
        for c in range(n - 1):
            v = mat[r][c]
            if mat[r][c + 1] == v and mat[r + 1][c] == v and mat[r + 1][c + 1] == v:
                score += 3
    p1 = [1, 0, 1, 1, 1, 0, 1, 0, 0, 0, 0]
    p2 = [0, 0, 0, 0, 1, 0, 1, 1, 1, 0, 1]
    for r in range(n):
        row = mat[r]
        for c in range(n - 10):
            seg = row[c:c + 11]
            if seg == p1 or seg == p2:
                score += 40
    for c in range(n):
        for r in range(n - 10):
            seg = [mat[r + k][c] for k in range(11)]
            if seg == p1 or seg == p2:
                score += 40
    dark = sum(sum(row) for row in mat)
    pct = (dark * 100.0) / (n * n)
    prev5 = int(pct // 5) * 5
    next5 = prev5 + 5
    score += min(abs(prev5 - 50), abs(next5 - 50)) // 5 * 10
    return score


def encode_qr_matrix(text: str) -> list[list[int]]:
    """Mã hóa chuỗi thành ma trận QR chuẩn ISO/IEC 18004 (Byte mode, ECC Level M)."""
    data = text.encode("utf-8")
    n = len(data)
    version_params = (
        (1, 16, 10, 1, (), 0),
        (2, 28, 16, 1, (6, 18), 7),
        (3, 44, 26, 1, (6, 22), 7),
        (4, 64, 18, 2, (6, 26), 7),
        (5, 86, 24, 2, (6, 30), 7),
        (6, 108, 16, 4, (6, 34), 7),
    )
    chosen = None
    for params in version_params:
        if n + 2 <= params[1]:
            chosen = params
            break
    if chosen is None:
        raise ValueError("Nội dung QR vượt quá độ dài hỗ trợ (tối đa 106 byte)")

    ver, cap, ec_len, blocks, align, rem_bits = chosen

    bits = [0, 1, 0, 0]
    for i in range(7, -1, -1):
        bits.append((n >> i) & 1)
    for byte in data:
        for i in range(7, -1, -1):
            bits.append((byte >> i) & 1)
    max_bits = cap * 8
    bits.extend([0] * min(4, max_bits - len(bits)))
    while len(bits) % 8 != 0:
        bits.append(0)

    codewords = []
    for i in range(0, len(bits), 8):
        val = 0
        for bit in bits[i:i + 8]:
            val = (val << 1) | bit
        codewords.append(val)
    pad_bytes = (0xEC, 0x11)
    pad_idx = 0
    while len(codewords) < cap:
        codewords.append(pad_bytes[pad_idx % 2])
        pad_idx += 1

    block_size = cap // blocks
    data_blocks = [codewords[i * block_size:(i + 1) * block_size] for i in range(blocks)]
    ec_blocks = [_qr_rs_encode(b, ec_len) for b in data_blocks]

    final_bytes = []
    for i in range(block_size):
        for b in data_blocks:
            final_bytes.append(b[i])
    for i in range(ec_len):
        for b in ec_blocks:
            final_bytes.append(b[i])

    final_bits = []
    for byte in final_bytes:
        for i in range(7, -1, -1):
            final_bits.append((byte >> i) & 1)
    final_bits.extend([0] * rem_bits)

    size = 17 + 4 * ver
    base = [[0] * size for _ in range(size)]
    reserved = [[False] * size for _ in range(size)]

    def set_mod(r, c, val):
        if 0 <= r < size and 0 <= c < size:
            base[r][c] = 1 if val else 0
            reserved[r][c] = True

    def place_finder(row, col):
        for dr in range(-1, 8):
            for dc in range(-1, 8):
                r, c = row + dr, col + dc
                if 0 <= r < size and 0 <= c < size:
                    if 0 <= dr <= 6 and 0 <= dc <= 6:
                        is_dark = dr in (0, 6) or dc in (0, 6) or (2 <= dr <= 4 and 2 <= dc <= 4)
                        set_mod(r, c, is_dark)
                    else:
                        set_mod(r, c, 0)

    place_finder(0, 0)
    place_finder(0, size - 7)
    place_finder(size - 7, 0)

    for ar in align:
        for ac in align:
            if reserved[ar][ac]:
                continue
            for dr in range(-2, 3):
                for dc in range(-2, 3):
                    set_mod(ar + dr, ac + dc, abs(dr) == 2 or abs(dc) == 2 or (dr == 0 and dc == 0))

    for i in range(8, size - 8):
        if not reserved[6][i]:
            set_mod(6, i, i % 2 == 0)
        if not reserved[i][6]:
            set_mod(i, 6, i % 2 == 0)

    set_mod(4 * ver + 9, 8, 1)
    for i in range(9):
        reserved[8][i] = True
        reserved[i][8] = True
    for i in range(8):
        reserved[8][size - 1 - i] = True
        reserved[size - 1 - i][8] = True

    bit_idx = 0
    upward = True
    col = size - 1
    while col > 0:
        if col == 6:
            col -= 1
        rows = range(size - 1, -1, -1) if upward else range(size)
        for r in rows:
            for dc in (0, -1):
                c = col + dc
                if not reserved[r][c]:
                    base[r][c] = final_bits[bit_idx] if bit_idx < len(final_bits) else 0
                    bit_idx += 1
        upward = not upward
        col -= 2

    coords_tl = (
        (8, 0), (8, 1), (8, 2), (8, 3), (8, 4), (8, 5), (8, 7), (8, 8),
        (7, 8), (5, 8), (4, 8), (3, 8), (2, 8), (1, 8), (0, 8),
    )

    best_mat = None
    best_score = None
    for mask_id, mask_fn in enumerate(_QR_MASK_FUNCS):
        cand = [row[:] for row in base]
        for r in range(size):
            for c in range(size):
                if not reserved[r][c] and mask_fn(r, c):
                    cand[r][c] ^= 1
        fmt_data = (0b00 << 3) | mask_id
        rem = fmt_data << 10
        for i in range(4, -1, -1):
            if rem & (1 << (i + 10)):
                rem ^= (0x537 << i)
        fmt = ((fmt_data << 10) | rem) ^ 0x5412
        fbits = [(fmt >> i) & 1 for i in range(15)]
        for i, (r, c) in enumerate(coords_tl):
            cand[r][c] = fbits[14 - i]
        for i in range(7):
            cand[size - 1 - i][8] = fbits[14 - i]
        for i in range(8):
            cand[8][size - 8 + i] = fbits[7 - i]

        sc = _qr_penalty(cand)
        if best_score is None or sc < best_score:
            best_score = sc
            best_mat = cand

    return best_mat


def generate_qr_pixmap(text: str, size: int = 320, quiet_zone: int = 4) -> QPixmap:
    """Render chuỗi thành QPixmap mã QR sắc nét có viền trắng quiet zone."""
    text = (text or "").strip()
    if not text:
        return QPixmap()
    mat = encode_qr_matrix(text)
    mat_size = len(mat)
    total = mat_size + 2 * quiet_zone
    img = QImage(total, total, QImage.Format_RGB32)
    img.fill(0xFFFFFFFF)
    for r in range(mat_size):
        for c in range(mat_size):
            if mat[r][c]:
                img.setPixel(c + quiet_zone, r + quiet_zone, 0xFF000000)
    return QPixmap.fromImage(img).scaled(
        size, size, Qt.KeepAspectRatio, Qt.FastTransformation
    )


class inventory_qr_window(QWidget):
    """Cửa sổ kiểm kê tài sản: render mã BAMS thành mã QR để quét từng mã một."""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Kiểm kê tài sản - Quét mã QR BAMS")
        self.setGeometry(100, 60, 1240, 740)
        self.setWindowIcon(QIcon("avatar_app.png"))
        self.setStyleSheet("""
            QWidget {
                background-color: #f0f0f0;
                color: black;
            }
            QPushButton {
                background: qlineargradient(
                    spread:pad, x1:0, y1:0, x2:1, y2:1,
                    stop:0 #016b67, stop:1 #0b9f87, stop:2 #00c3cc
                );
                color: white;
                border: none;
                border-radius: 6px;
                padding: 6px 12px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #016b67;
                color: white;
            }
            QLabel {
                color: black;
                font-weight: bold;
            }
            QComboBox, QLineEdit, QTextEdit, QTableWidget {
                background-color: white;
                border: 1px solid #016b67;
                border-radius: 4px;
                color: black;
            }
            QGroupBox {
                border: 1px solid #016b67;
                border-radius: 6px;
                margin-top: 8px;
                padding-top: 8px;
                font-weight: bold;
                color: #016b67;
                background-color: #fafafa;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 4px;
            }
            QScrollBar:vertical {
                border: none;
                background: #f0f0f0;
                width: 6px;
                margin: 0px;
                border-radius: 3px;
            }
            QScrollBar::handle:vertical {
                background-color: #c0c0c0;
                min-height: 20px;
                border-radius: 3px;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0px;
            }
            QScrollBar:horizontal {
                border: none;
                background: #f0f0f0;
                height: 6px;
                margin: 0px;
                border-radius: 3px;
            }
            QScrollBar::handle:horizontal {
                background-color: #c0c0c0;
                min-width: 20px;
                border-radius: 3px;
            }
            QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
                width: 0px;
            }
        """)

        self.raw_records = []
        self.view_entries = []
        self.filtered_entries = []
        self.scanned_keys = set()
        self.current_row = -1
        self.source_description = ""

        self.auto_timer = QTimer(self)
        self.auto_timer.timeout.connect(self._on_auto_tick)

        self._build_ui()
        self.reload_from_db(preserve_scanned=False)

    def _build_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(8)

        # Hàng 1: Nguồn dữ liệu & Bộ lọc
        top_bar = QHBoxLayout()
        top_bar.setSpacing(8)

        self.btn_reload_db = QPushButton("🔄 Tải lại từ DB", self)
        self.btn_reload_db.clicked.connect(lambda: self.reload_from_db(preserve_scanned=True))
        top_bar.addWidget(self.btn_reload_db)

        self.btn_load_excel = QPushButton("📂 Nạp từ file Excel", self)
        self.btn_load_excel.clicked.connect(self.load_from_excel_dialog)
        top_bar.addWidget(self.btn_load_excel)

        top_bar.addWidget(QLabel("Chế độ:", self))
        self.combo_mode = QComboBox(self)
        self.combo_mode.addItems([
            "Gộp theo mã BAMS duy nhất",
            "Hiển thị từng tài sản riêng lẻ",
        ])
        self.combo_mode.currentIndexChanged.connect(self._on_mode_changed)
        top_bar.addWidget(self.combo_mode)

        top_bar.addWidget(QLabel("Loại:", self))
        self.combo_asset_type = QComboBox(self)
        self.combo_asset_type.addItems(["Tất cả", "CPU", "Monitor", "Laptop", "Printer", "Scanner", "Máy chiếu"])
        self.combo_asset_type.currentIndexChanged.connect(lambda: self.apply_filters())
        top_bar.addWidget(self.combo_asset_type)

        top_bar.addWidget(QLabel("Sử dụng:", self))
        self.combo_active = QComboBox(self)
        self.combo_active.addItems(["Tất cả", "Đang sử dụng", "Đã thu hồi trả văn phòng"])
        self.combo_active.currentIndexChanged.connect(lambda: self.apply_filters())
        top_bar.addWidget(self.combo_active)

        top_bar.addWidget(QLabel("Quét:", self))
        self.combo_scan_status = QComboBox(self)
        self.combo_scan_status.addItems(["Tất cả", "Chưa quét", "Đã quét"])
        self.combo_scan_status.currentIndexChanged.connect(lambda: self.apply_filters())
        top_bar.addWidget(self.combo_scan_status)

        self.search_input = QLineEdit(self)
        self.search_input.setPlaceholderText("Tìm mã BAMS, số seri, cán bộ, ghi chú...")
        self.search_input.setMinimumWidth(200)
        self.search_input.textChanged.connect(lambda: self.apply_filters())
        self.search_input.returnPressed.connect(lambda: self.table.setFocus())
        top_bar.addWidget(self.search_input, 1)

        main_layout.addLayout(top_bar)

        # Thanh thông báo nguồn dữ liệu
        self.source_label = QLabel("", self)
        self.source_label.setStyleSheet("color: #016b67; font-weight: normal;")
        main_layout.addWidget(self.source_label)

        # Chia đôi: Trái = Bảng danh sách | Phải = Khung QR & Điều khiển
        splitter = QSplitter(Qt.Horizontal, self)

        # === Bảng danh sách bên trái ===
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 4, 0)

        self.table = QTableWidget(self)
        self.table.setColumnCount(8)
        self.table.setHorizontalHeaderLabels([
            "STT", "Trạng thái", "Mã BAMS", "Loại tài sản",
            "Số seri", "Tình trạng sử dụng", "Cán bộ / Kho", "Ghi chú"
        ])
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setColumnWidth(0, 45)
        self.table.setColumnWidth(1, 95)
        self.table.setColumnWidth(2, 130)
        self.table.setColumnWidth(3, 110)
        self.table.setColumnWidth(4, 155)
        self.table.setColumnWidth(5, 135)
        self.table.setColumnWidth(6, 150)
        self.table.setColumnWidth(7, 220)
        self.table.currentCellChanged.connect(self._on_table_cell_changed)
        left_layout.addWidget(self.table)

        splitter.addWidget(left_widget)

        # === Khung hiển thị QR bên phải ===
        right_panel = QGroupBox("Mã QR Kiểm Kê (BAMS)", self)
        right_layout = QVBoxLayout(right_panel)
        right_layout.setSpacing(10)
        right_layout.setAlignment(Qt.AlignTop)

        self.progress_label = QLabel("Mã 0 / 0   |   Đã quét: 0   |   Còn lại: 0", self)
        self.progress_label.setAlignment(Qt.AlignCenter)
        self.progress_label.setStyleSheet("font-size: 13pt; color: #016b67; font-weight: bold;")
        right_layout.addWidget(self.progress_label)

        # Ảnh QR đặt giữa
        qr_holder = QHBoxLayout()
        qr_holder.addStretch()
        self.qr_label = QLabel(self)
        self.qr_label.setFixedSize(340, 340)
        self.qr_label.setAlignment(Qt.AlignCenter)
        self.qr_label.setStyleSheet(
            "background-color: white; border: 2px solid #016b67; border-radius: 8px; padding: 6px;"
        )
        qr_holder.addWidget(self.qr_label)
        qr_holder.addStretch()
        right_layout.addLayout(qr_holder)

        # Mã BAMS cỡ lớn + nút Copy
        bams_row = QHBoxLayout()
        bams_row.addStretch()
        self.bams_label = QLabel("---", self)
        self.bams_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.bams_label.setStyleSheet("font-size: 22pt; font-weight: bold; color: #016b67;")
        bams_row.addWidget(self.bams_label)

        self.btn_copy = QPushButton("📋 Copy mã BAMS", self)
        self.btn_copy.clicked.connect(self.copy_current_bams)
        bams_row.addWidget(self.btn_copy)
        bams_row.addStretch()
        right_layout.addLayout(bams_row)

        # Chi tiết thiết bị
        self.detail_text = QTextEdit(self)
        self.detail_text.setReadOnly(True)
        self.detail_text.setMinimumHeight(130)
        self.detail_text.setMaximumHeight(180)
        right_layout.addWidget(self.detail_text)

        # Nút điều hướng từng mã
        nav_row = QHBoxLayout()
        nav_row.setSpacing(8)
        self.btn_prev = QPushButton("⏮ Trước (←)", self)
        self.btn_prev.clicked.connect(self.prev_item)
        nav_row.addWidget(self.btn_prev)

        self.btn_scan_next = QPushButton("✅ Đã quét & Tiếp theo (Space)", self)
        self.btn_scan_next.setStyleSheet("""
            QPushButton {
                background-color: #198754;
                color: white;
                border-radius: 6px;
                padding: 8px 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #146c43;
            }
        """)
        self.btn_scan_next.clicked.connect(self.scan_and_next)
        nav_row.addWidget(self.btn_scan_next, 1)

        self.btn_next = QPushButton("⏭ Tiếp theo (→)", self)
        self.btn_next.clicked.connect(self.next_item)
        nav_row.addWidget(self.btn_next)
        right_layout.addLayout(nav_row)

        # Hàng tự động chuyển rảnh tay & đặt lại trạng thái
        auto_row = QHBoxLayout()
        auto_row.setSpacing(8)
        self.chk_auto = QCheckBox("⏱ Tự động chuyển sau:", self)
        self.chk_auto.toggled.connect(self._toggle_auto_timer)
        auto_row.addWidget(self.chk_auto)

        self.combo_interval = QComboBox(self)
        self.combo_interval.addItems(["2s", "3s", "5s", "10s"])
        self.combo_interval.setCurrentText("3s")
        self.combo_interval.currentIndexChanged.connect(self._update_auto_interval)
        auto_row.addWidget(self.combo_interval)

        auto_row.addStretch()

        self.btn_reset_scan = QPushButton("🔄 Đặt lại trạng thái quét", self)
        self.btn_reset_scan.clicked.connect(self.reset_scanned_status)
        auto_row.addWidget(self.btn_reset_scan)
        right_layout.addLayout(auto_row)

        self.status_hint_label = QLabel(
            "Mẹo: Dùng Space/Enter để đánh dấu đã quét & qua mã tiếp theo, ←/→ để lùi/tiến.",
            self,
        )
        self.status_hint_label.setWordWrap(True)
        self.status_hint_label.setStyleSheet("color: #555555; font-weight: normal;")
        right_layout.addWidget(self.status_hint_label)

        splitter.addWidget(right_panel)
        splitter.setSizes([720, 500])
        main_layout.addWidget(splitter, 1)

        # Cài eventFilter để bắt phím Space/Enter/Left/Right/Up/Down mà không cản ô tìm kiếm
        for widget in (
            self.table, self.btn_prev, self.btn_scan_next, self.btn_next,
            self.btn_reset_scan, self.btn_copy, self.detail_text,
            self.combo_mode, self.combo_asset_type, self.combo_active,
            self.combo_scan_status, self.chk_auto, self.combo_interval,
        ):
            widget.installEventFilter(self)

    def eventFilter(self, watched, event):
        if event.type() == QEvent.KeyPress:
            key = event.key()
            if key in (Qt.Key_Space, Qt.Key_Return, Qt.Key_Enter):
                self.scan_and_next()
                return True
            if key in (Qt.Key_Left, Qt.Key_Up):
                self.prev_item()
                return True
            if key in (Qt.Key_Right, Qt.Key_Down):
                self.next_item()
                return True
        return super().eventFilter(watched, event)

    def _normalize_excel_records(self, df):
        records = []
        for idx, row in df.iterrows():
            bams_raw = row.get("bams")
            bams_val = "" if pd.isna(bams_raw) else str(bams_raw).strip()
            if not bams_val or bams_val.lower() in ("nan", "none"):
                continue

            def _txt(col):
                v = row.get(col)
                if v is None or pd.isna(v):
                    return ""
                s = str(v).strip()
                return "" if s.lower() in ("nan", "none") else s

            active_raw = _txt("active_asset").lower()
            is_active = active_raw in ("1", "true", "yes")
            id_raw = _txt("id")
            try:
                asset_id = int(float(id_raw)) if id_raw else (idx + 1)
            except ValueError:
                asset_id = idx + 1

            records.append({
                "id": asset_id,
                "asset_name": _txt("asset_name"),
                "code": _txt("code"),
                "serial_number": _txt("serial_number"),
                "purchase_date": _txt("purchase_date"),
                "status": _txt("status"),
                "active_asset": 1 if is_active else 0,
                "return_date": _txt("return_date"),
                "value": _txt("value"),
                "staff_name": _txt("staff_name"),
                "department": _txt("department"),
                "bams": bams_val,
                "notes": _txt("notes"),
            })
        return records

    def reload_from_db(self, preserve_scanned=True):
        if not preserve_scanned:
            self.scanned_keys.clear()

        conn = sqlite3.connect(db_path_default)
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM assets")
        total_in_db = cursor.fetchone()[0]

        records = []
        if total_in_db > 0:
            cursor.execute("""
                SELECT a.id, a.asset_name, a.code, a.serial_number, a.purchase_date,
                       a.status, a.active_asset, a.return_date, a.value,
                       s.name AS staff_name, s.department, a.bams, a.notes
                FROM assets a
                LEFT JOIN asset_assignments aa ON aa.asset_id = a.id AND aa.returned_date IS NULL
                LEFT JOIN staff s ON aa.staff_id = s.id
                WHERE a.bams IS NOT NULL AND TRIM(a.bams) != ''
                ORDER BY a.id ASC
            """)
            rows = cursor.fetchall()
            for r in rows:
                records.append({
                    "id": r[0],
                    "asset_name": str(r[1] or "").strip(),
                    "code": str(r[2] or "").strip(),
                    "serial_number": str(r[3] or "").strip(),
                    "purchase_date": str(r[4] or "").strip(),
                    "status": str(r[5] or "").strip(),
                    "active_asset": 1 if r[6] else 0,
                    "return_date": str(r[7] or "").strip(),
                    "value": "" if r[8] is None else str(r[8]),
                    "staff_name": str(r[9] or "").strip(),
                    "department": str(r[10] or "").strip(),
                    "bams": str(r[11] or "").strip(),
                    "notes": str(r[12] or "").strip(),
                })
        conn.close()

        if total_in_db == 0 and os.path.exists("assets.xlsx"):
            try:
                df = pd.read_excel("assets.xlsx", dtype=str)
                records = self._normalize_excel_records(df)
                self.source_description = f"Nguồn: assets.xlsx (DB trống, tự động nạp {len(records)} tài sản có BAMS)"
            except Exception as e:
                self.source_description = f"Không đọc được assets.xlsx: {e}"
        else:
            self.source_description = f"Nguồn: SQLite ({db_path_default}) — {len(records)} tài sản có mã BAMS"

        self.raw_records = records
        self._refresh_asset_type_options()
        self.rebuild_view_entries()
        self.apply_filters()

    def load_from_excel_dialog(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn file Excel tài sản",
            "",
            "Excel Files (*.xlsx)",
        )
        if not file_path:
            return
        try:
            df = pd.read_excel(file_path, dtype=str)
            if "bams" not in df.columns:
                QMessageBox.warning(self, "Thiếu cột", "File Excel cần có cột 'bams'!")
                return
            self.raw_records = self._normalize_excel_records(df)
            self.scanned_keys.clear()
            self.source_description = f"Nguồn: {file_path} — {len(self.raw_records)} tài sản có mã BAMS"
            self._refresh_asset_type_options()
            self.rebuild_view_entries()
            self.apply_filters()
        except Exception as e:
            QMessageBox.critical(self, "Lỗi đọc Excel", f"Không thể đọc file Excel:\n{str(e)}")

    def _refresh_asset_type_options(self):
        current = self.combo_asset_type.currentText()
        base_types = ["CPU", "Monitor", "Laptop", "Printer", "Scanner", "Máy chiếu"]
        extras = sorted({
            r["asset_name"] for r in self.raw_records
            if r["asset_name"] and r["asset_name"] not in base_types
        })
        self.combo_asset_type.blockSignals(True)
        self.combo_asset_type.clear()
        self.combo_asset_type.addItems(["Tất cả"] + base_types + extras)
        idx = self.combo_asset_type.findText(current)
        self.combo_asset_type.setCurrentIndex(idx if idx >= 0 else 0)
        self.combo_asset_type.blockSignals(False)

    def _on_mode_changed(self):
        self.rebuild_view_entries()
        self.apply_filters()

    def rebuild_view_entries(self):
        grouped_mode = self.combo_mode.currentIndex() == 0
        entries = []

        if grouped_mode:
            groups = {}
            order = []
            for r in self.raw_records:
                bams = r["bams"]
                if bams not in groups:
                    groups[bams] = []
                    order.append(bams)
                groups[bams].append(r)

            for bams in order:
                items = groups[bams]
                asset_types = []
                active_labels = []
                holders = []
                statuses = []
                notes = []
                serial_short_parts = []
                serial_detail_parts = []

                for it in items:
                    aname = it["asset_name"]
                    if aname and aname not in asset_types:
                        asset_types.append(aname)
                    act_lbl = "Đang sử dụng" if it["active_asset"] else "Đã thu hồi trả văn phòng"
                    if act_lbl not in active_labels:
                        active_labels.append(act_lbl)
                    if it["staff_name"]:
                        h = f"{it['staff_name']} ({it['department']})" if it["department"] else it["staff_name"]
                    else:
                        h = "Trong kho"
                    if h not in holders:
                        holders.append(h)
                    if it["status"] and it["status"] not in statuses:
                        statuses.append(it["status"])
                    if it["notes"] and it["notes"] not in notes:
                        notes.append(it["notes"])

                    sn_text = it["serial_number"] or "(trống)"
                    if len(items) > 1:
                        serial_short_parts.append(f"{aname}: {sn_text}")
                        serial_detail_parts.append(f"  • {aname}: {sn_text} (ID: {it['id']}, {act_lbl})")
                    else:
                        serial_short_parts.append(sn_text)
                        serial_detail_parts.append(f"{sn_text} (ID: {it['id']})")

                entries.append({
                    "key": ("GROUP", bams),
                    "bams": bams,
                    "asset_types_set": set(asset_types),
                    "asset_names": " + ".join(asset_types),
                    "serial_short": " | ".join(serial_short_parts),
                    "serial_detail": "\n".join(serial_detail_parts),
                    "active_set": set(active_labels),
                    "active_label": " / ".join(active_labels),
                    "holder_display": "; ".join(holders),
                    "status_display": "; ".join(statuses),
                    "notes_display": " | ".join(notes),
                    "items": items,
                })
        else:
            for it in self.raw_records:
                act_lbl = "Đang sử dụng" if it["active_asset"] else "Đã thu hồi trả văn phòng"
                if it["staff_name"]:
                    h = f"{it['staff_name']} ({it['department']})" if it["department"] else it["staff_name"]
                else:
                    h = "Trong kho"
                sn_text = it["serial_number"] or "(trống)"
                entries.append({
                    "key": ("SINGLE", it["id"], it["bams"], it["serial_number"]),
                    "bams": it["bams"],
                    "asset_types_set": {it["asset_name"]} if it["asset_name"] else set(),
                    "asset_names": it["asset_name"],
                    "serial_short": sn_text,
                    "serial_detail": f"{sn_text} (ID: {it['id']})",
                    "active_set": {act_lbl},
                    "active_label": act_lbl,
                    "holder_display": h,
                    "status_display": it["status"],
                    "notes_display": it["notes"],
                    "items": [it],
                })

        self.view_entries = entries
        mode_hint = (
            f"{len(entries)} mã BAMS duy nhất"
            if grouped_mode
            else f"{len(entries)} tài sản riêng lẻ"
        )
        self.source_label.setText(f"{self.source_description}  |  Hiển thị: {mode_hint}")

    def apply_filters(self, preserve_key=None, preferred_row=None):
        type_filter = self.combo_asset_type.currentText()
        active_filter = self.combo_active.currentText()
        scan_filter = self.combo_scan_status.currentText()
        query = self.search_input.text().strip().lower()

        if preserve_key is None and preferred_row is None and 0 <= self.current_row < len(self.filtered_entries):
            preserve_key = self.filtered_entries[self.current_row]["key"]

        filtered = []
        for entry in self.view_entries:
            if type_filter != "Tất cả" and type_filter not in entry["asset_types_set"]:
                continue
            if active_filter != "Tất cả" and active_filter not in entry["active_set"]:
                continue
            is_scanned = entry["key"] in self.scanned_keys
            if scan_filter == "Chưa quét" and is_scanned:
                continue
            if scan_filter == "Đã quét" and not is_scanned:
                continue
            if query:
                haystack = " ".join([
                    entry["bams"],
                    entry["asset_names"],
                    entry["serial_short"],
                    entry["holder_display"],
                    entry["notes_display"],
                    entry["status_display"],
                ]).lower()
                if query not in haystack:
                    continue
            filtered.append(entry)

        self.filtered_entries = filtered
        self._populate_table()

        if not filtered:
            self.current_row = -1
            self.update_progress_label()
            self.display_current_entry()
            return

        target_row = 0
        if preserve_key is not None:
            for i, entry in enumerate(filtered):
                if entry["key"] == preserve_key:
                    target_row = i
                    break
        elif preferred_row is not None:
            target_row = max(0, min(preferred_row, len(filtered) - 1))

        self.select_row(target_row)

    def _populate_table(self):
        self.table.blockSignals(True)
        self.table.setRowCount(len(self.filtered_entries))
        scanned_bg = QColor("#d4edda")
        normal_bg = QColor("white")

        for row_idx, entry in enumerate(self.filtered_entries):
            is_scanned = entry["key"] in self.scanned_keys
            bg = scanned_bg if is_scanned else normal_bg
            vals = [
                str(row_idx + 1),
                "✅ Đã quét" if is_scanned else "⬜ Chưa quét",
                entry["bams"],
                entry["asset_names"],
                entry["serial_short"],
                entry["active_label"],
                entry["holder_display"],
                entry["notes_display"],
            ]
            for col_idx, val in enumerate(vals):
                item = QTableWidgetItem(val)
                item.setBackground(bg)
                if col_idx == 0:
                    item.setTextAlignment(Qt.AlignCenter)
                self.table.setItem(row_idx, col_idx, item)
        self.table.blockSignals(False)

    def _refresh_table_row_status(self, row_idx):
        if row_idx < 0 or row_idx >= len(self.filtered_entries):
            return
        entry = self.filtered_entries[row_idx]
        is_scanned = entry["key"] in self.scanned_keys
        bg = QColor("#d4edda") if is_scanned else QColor("white")
        status_item = self.table.item(row_idx, 1)
        if status_item:
            status_item.setText("✅ Đã quét" if is_scanned else "⬜ Chưa quét")
        for col_idx in range(self.table.columnCount()):
            cell_item = self.table.item(row_idx, col_idx)
            if cell_item:
                cell_item.setBackground(bg)

    def _on_table_cell_changed(self, row, _col, _prev_row, _prev_col):
        if 0 <= row < len(self.filtered_entries):
            self.current_row = row
            self.update_progress_label()
            self.display_current_entry()

    def select_row(self, row_idx):
        if not self.filtered_entries:
            self.current_row = -1
            self.update_progress_label()
            self.display_current_entry()
            return
        row_idx = max(0, min(row_idx, len(self.filtered_entries) - 1))
        self.current_row = row_idx
        self.table.blockSignals(True)
        self.table.selectRow(row_idx)
        item = self.table.item(row_idx, 0)
        if item:
            self.table.scrollToItem(item)
        self.table.blockSignals(False)
        self.update_progress_label()
        self.display_current_entry()

    def update_progress_label(self):
        total = len(self.filtered_entries)
        scanned_in_view = sum(1 for e in self.filtered_entries if e["key"] in self.scanned_keys)
        remaining = total - scanned_in_view
        pos = (self.current_row + 1) if 0 <= self.current_row < total else 0
        self.progress_label.setText(
            f"Mã {pos} / {total}   |   Đã quét: {scanned_in_view}   |   Còn lại: {remaining}"
        )

    def display_current_entry(self):
        if self.current_row < 0 or self.current_row >= len(self.filtered_entries):
            self.qr_label.clear()
            self.qr_label.setText("Không có dữ liệu phù hợp bộ lọc")
            self.bams_label.setText("---")
            self.detail_text.setPlainText("")
            return

        entry = self.filtered_entries[self.current_row]
        bams = entry["bams"]
        try:
            pix = generate_qr_pixmap(bams, size=320, quiet_zone=4)
            self.qr_label.setPixmap(pix)
        except Exception as e:
            self.qr_label.clear()
            self.qr_label.setText(f"Lỗi tạo QR:\n{e}")

        self.bams_label.setText(bams)
        serial_prefix = "\n" if len(entry["items"]) > 1 else " "
        detail_lines = [
            f"Loại tài sản: {entry['asset_names']}",
            f"Số seri:{serial_prefix}{entry['serial_detail']}",
            f"Tình trạng thiết bị: {entry['status_display'] or '(không rõ)'}",
            f"Tình trạng sử dụng: {entry['active_label']}",
            f"Cán bộ giữ / Kho: {entry['holder_display']}",
            f"Ghi chú: {entry['notes_display'] or '(không có)'}",
        ]
        self.detail_text.setPlainText("\n".join(detail_lines))

    def copy_current_bams(self):
        if 0 <= self.current_row < len(self.filtered_entries):
            bams = self.filtered_entries[self.current_row]["bams"]
            QApplication.clipboard().setText(bams)
            self.status_hint_label.setText(f"📋 Đã copy mã BAMS: {bams}")

    def prev_item(self):
        if not self.filtered_entries:
            return
        if self.current_row > 0:
            self.select_row(self.current_row - 1)
            self.status_hint_label.setText("Mẹo: Dùng Space/Enter để đánh dấu đã quét & qua mã tiếp theo, ←/→ để lùi/tiến.")

    def next_item(self):
        if not self.filtered_entries:
            return
        if self.current_row + 1 < len(self.filtered_entries):
            self.select_row(self.current_row + 1)
            self.status_hint_label.setText("Mẹo: Dùng Space/Enter để đánh dấu đã quét & qua mã tiếp theo, ←/→ để lùi/tiến.")
        else:
            self.status_hint_label.setText("Đã ở mã cuối cùng trong danh sách.")

    def scan_and_next(self):
        if self.current_row < 0 or self.current_row >= len(self.filtered_entries):
            return
        old_row = self.current_row
        entry = self.filtered_entries[old_row]
        self.scanned_keys.add(entry["key"])

        if self.combo_scan_status.currentText() == "Chưa quét":
            self.apply_filters(preserve_key=None, preferred_row=old_row)
            if not self.filtered_entries:
                self._stop_auto_timer()
                self.status_hint_label.setText("🎉 Đã quét xong toàn bộ các mã chưa quét!")
        else:
            self._refresh_table_row_status(old_row)
            if old_row + 1 < len(self.filtered_entries):
                self.select_row(old_row + 1)
                self.status_hint_label.setText(f"✅ Đã đánh dấu quét: {entry['bams']}")
            else:
                self.update_progress_label()
                self._stop_auto_timer()
                self.status_hint_label.setText("🎉 Đã quét tới mã cuối cùng trong danh sách!")

    def reset_scanned_status(self):
        self._stop_auto_timer()
        self.scanned_keys.clear()
        self.apply_filters(preferred_row=0)
        self.status_hint_label.setText("🔄 Đã đặt lại toàn bộ trạng thái quét về chưa quét.")

    def _get_interval_ms(self):
        text = self.combo_interval.currentText().rstrip("s")
        try:
            return int(text) * 1000
        except ValueError:
            return 3000

    def _toggle_auto_timer(self, checked):
        if checked:
            if not self.filtered_entries:
                self.chk_auto.setChecked(False)
                return
            self.auto_timer.start(self._get_interval_ms())
            self.status_hint_label.setText(f"⏱ Đang tự động quét & chuyển mỗi {self.combo_interval.currentText()}...")
        else:
            self.auto_timer.stop()

    def _update_auto_interval(self):
        if self.auto_timer.isActive():
            self.auto_timer.start(self._get_interval_ms())
            self.status_hint_label.setText(f"⏱ Đã đổi chu kỳ tự động chuyển thành {self.combo_interval.currentText()}.")

    def _stop_auto_timer(self):
        self.auto_timer.stop()
        if self.chk_auto.isChecked():
            self.chk_auto.blockSignals(True)
            self.chk_auto.setChecked(False)
            self.chk_auto.blockSignals(False)

    def _on_auto_tick(self):
        if self.current_row < 0 or self.current_row >= len(self.filtered_entries):
            self._stop_auto_timer()
            return
        self.scan_and_next()

    def closeEvent(self, event):
        self._stop_auto_timer()
        super().closeEvent(event)


def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()