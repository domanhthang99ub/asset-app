
import sqlite3
from datetime import date, datetime
import shutil
import os
import sys
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from PySide6 import QtWidgets
from PySide6.QtCore import QByteArray, QBuffer, QIODevice


db_path_default = r"\\h17des02071\mInput\NHÀ MỚI\NHẬP KHẨU 1\IX.KHAC\Back up Thắng\Data_base\assets.db"
#db_path_default = r"assets.db"
#db_path_default = r"\\DESKTOP-1S9FC0J\data_base\assets.db"
#pyinstaller --onefile --noconsole --icon=avatar_app.ico --name "Asset Management" main.py

def delete_db():
    if os.path.exists(db_path_default):
        os.remove(db_path_default)

# Kết nối (tạo file database nếu chưa có)

conn = sqlite3.connect(db_path_default)
cursor = conn.cursor()

cursor.execute("""
CREATE TABLE IF NOT EXISTS staff (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    department TEXT NOT NULL,
    position TEXT,
    email TEXT NOT NULL UNIQUE,
    employee_code TEXT NOT NULL UNIQUE,
    phone TEXT,
    active_staff BOOLEAN NOT NULL,
    notes TEXT
)
""")


cursor.execute("""
CREATE TABLE IF NOT EXISTS assets (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    asset_name TEXT NOT NULL,
    code TEXT ,
    serial_number TEXT NOT NULL UNIQUE,
    purchase_date TEXT,
    status TEXT,
    active_asset BOOLEAN NOT NULL,
    return_date TEXT,
    value INTEGER,
    bams TEXT,
    notes TEXT
)
""")

cursor.execute("""
CREATE TABLE IF NOT EXISTS asset_assignments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    asset_id INTEGER NOT NULL,
    staff_id INTEGER NOT NULL,
    assigned_date TEXT NOT NULL,
    returned_date TEXT,
    notes TEXT,
    handover_documents_assign_id INTEGER,
    handover_documents_return_id INTEGER,
    FOREIGN KEY (asset_id) REFERENCES assets(id),
    FOREIGN KEY (staff_id) REFERENCES staff(id),
    FOREIGN KEY (handover_documents_assign_id) REFERENCES handover_documents(id),
    FOREIGN KEY (handover_documents_return_id) REFERENCES handover_documents(id)
)
""")


cursor.execute("""
CREATE TABLE IF NOT EXISTS handover_documents (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    document_date TEXT NOT NULL,
    type_handover_documents TEXT,
    document_file BLOB,
    notes TEXT
)
""")

cursor.execute("""
CREATE TABLE IF NOT EXISTS chat_box (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    chat TEXT NOT NULL,
    date_and_time TEXT NOT NULL,
    user_chat TEXT NOT NULL
)
""")

cursor.execute("PRAGMA journal_mode=WAL;")

conn.commit()
conn.close()

def add_staff(name_add, department_add, position_add,email_add,employee_code_add, phone_add, notes_add):
    try:
        conn = sqlite3.connect(db_path_default)
        cursor = conn.cursor()

        name_add = name_add.strip()
        department_add = department_add.strip()
        position_add = position_add.strip()
        email_add = email_add.strip()
        employee_code_add = employee_code_add.strip()
        phone_add = phone_add.strip()
        notes_add = notes_add.strip()

        if name_add == "":
            name_add = None
        if department_add == "":
            department_add = None
        if position_add == "":
            position_add = None
        if email_add == "":
            email_add = None
        if employee_code_add == "":
            employee_code_add = None
        if phone_add == "":
            phone_add = None
        if notes_add == "":
            notes_add = None

        if not name_add:
            conn.close()
            return "Lỗi: Tên cán bộ không được để trống!"
        if not email_add:
            conn.close()
            return "Lỗi: Email không được để trống!"
        if not employee_code_add:
            conn.close()
            return "Lỗi: Mã cán bộ không được để trống!"


        cursor.execute("SELECT * FROM staff WHERE employee_code = ?", (employee_code_add,))
        results = cursor.fetchone()
        if results:
            conn.close()
            return "Lỗi: Mã cán bộ đã bị trùng!"
        cursor.execute("SELECT * FROM staff WHERE email = ?", (email_add,))
        results = cursor.fetchone()
        if results:
            conn.close()
            return "Lỗi: Email đã bị trùng!"

        cursor.execute(
            "INSERT INTO staff (name, department, position,email, employee_code,phone,active_staff,notes) VALUES (?, ?, ?, ?, ?, ?, ?,?)",
            (name_add, department_add, position_add, email_add, employee_code_add, phone_add, True, notes_add))
        conn.commit()
        conn.close()
        return "Thêm cán bộ thành công!"
    except Exception as e:
        messagebox.showerror("Lỗi", f"\n{str(e)}")


def modify_staff(name_add, department_add, position_add,email_add,employee_code_add, phone_add, notes_add, search):
    try:
        conn = sqlite3.connect(db_path_default)
        cursor = conn.cursor()

        name_add = name_add.strip()
        department_add = department_add.strip()
        position_add = position_add.strip()
        email_add = email_add.strip()
        employee_code_add = employee_code_add.strip()
        phone_add = phone_add.strip()
        notes_add = notes_add.strip()
        search = search.strip()

        if name_add == "":
            name_add = None
        if department_add == "":
            department_add = None
        if position_add == "":
            position_add = None
        if email_add == "":
            email_add = None
        if employee_code_add == "":
            employee_code_add = None
        if phone_add == "":
            phone_add = None
        if notes_add == "":
            notes_add = None
        if search == "":
            conn.close()
            return "Lỗi: Cán bộ cần sửa đổi không được để trống!"

        if not name_add:
            conn.close()
            return "Lỗi: Tên cán bộ không được để trống!"
        if not email_add:
            conn.close()
            return "Lỗi: Email không được để trống!"
        if not employee_code_add:
            conn.close()
            return "Lỗi: Mã cán bộ không được để trống!"


        cursor.execute("SELECT * FROM staff WHERE name = ? OR employee_code = ? OR email = ?", (search, search, search))
        results = cursor.fetchone()
        if not results:
            conn.close()
            return "Không tìm thấy cán bộ cần sửa!"

        employee_code_modify = results[5]
        email_modify = results[4]
        staff_id_modify = results[0]

        if employee_code_add == employee_code_modify:
            pass
        elif employee_code_add != employee_code_modify:
            cursor.execute("SELECT * FROM staff WHERE employee_code = ?", (employee_code_add,))
            results = cursor.fetchone()
            if results:
                conn.close()
                return "Lỗi: Mã cán bộ đã bị trùng!"

        if email_add == email_modify:
            pass
        elif email_add != email_modify:
            cursor.execute("SELECT * FROM staff WHERE email = ?", (email_add,))
            results = cursor.fetchone()
            if results:
                conn.close()
                return "Lỗi: Email đã bị trùng!"

        cursor.execute("""
                       UPDATE staff
                       SET name          = ?,
                           department    = ?,
                           position      = ?,
                           email         = ?,
                           employee_code = ?,
                           phone         = ?,
                           notes         = ?
                       WHERE id = ?
                       """, (name_add, department_add, position_add, email_add, employee_code_add, phone_add, notes_add,
                             staff_id_modify))
        conn.commit()
        conn.close()
        return "Sửa thông tin cán bộ thành công!"

    except Exception as e:
        messagebox.showerror("Lỗi", f"\n{str(e)}")




def search_staff(search):
    try:
        search = search.strip()
        if search == "":
            return "Thông tin vừa rồi không khớp cán bộ nào!"

        conn = sqlite3.connect(db_path_default)
        cursor = conn.cursor()

        cursor.execute("""
                       SELECT *
                       FROM staff
                       WHERE name = ?
                          OR employee_code = ?
                          OR email = ?
                       """, (search, search, search))
        result = cursor.fetchone()
        conn.close()
        if result:
            return result
        else:
            return "Thông tin vừa rồi không khớp cán bộ nào!"
    except Exception as e:
        messagebox.showerror("Lỗi", f"\n{str(e)}")


def delete_staff(search):
    try:
        search = search.strip()
        if search == "":
            return "Không tìm thấy cán bộ để kích hoạt/vô hiệu hóa."

        conn = sqlite3.connect(db_path_default)
        cursor = conn.cursor()
        cursor.execute("""
                       SELECT *
                       FROM staff
                       WHERE name = ?
                          OR employee_code = ?
                          OR email = ?
                       """, (search, search, search))
        result = cursor.fetchone()

        if not result:
            conn.close()
            return "Không tìm thấy cán bộ để kích hoạt/vô hiệu hóa."

        staff_id_delete = result[0]
        active_staff_delete = result[7]

        if active_staff_delete:
            cursor.execute("""
                           SELECT *
                           FROM asset_assignments
                           WHERE staff_id = ?
                             AND returned_date IS NULL
                           """, (staff_id_delete,))
            holding_asset = cursor.fetchone()

            if holding_asset:
                conn.close()
                return f"Không thể vô hiệu hóa\nCán bộ {search} đang giữ tài sản."

            cursor.execute("""
                           UPDATE staff
                           SET active_staff = ?
                           WHERE id = ?
                           """, (False, staff_id_delete))
            conn.commit()
            conn.close()
            return f"Đã vô hiệu hóa cán bộ {search} thành công!"
        else:
            cursor.execute("""
                           UPDATE staff
                           SET active_staff = ?
                           WHERE id = ?
                           """, (True, staff_id_delete))
            conn.commit()
            conn.close()
            return f"Đã kích hoạt cán bộ {search} thành công!"

    except Exception as e:
        messagebox.showerror("Lỗi", f"\n{str(e)}")




def export_staff_to_excel(db_path=db_path_default):
    try:
        root = tk.Tk()
        root.withdraw()
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            title="Chọn nơi lưu file Excel",
            initialfile="staff.xlsx"
        )

        if not file_path:
            print("Đã huỷ xuất file.")
            return

        conn = sqlite3.connect(db_path)
        df = pd.read_sql_query("SELECT * FROM staff", conn)

        df.to_excel(file_path, index=False)
        conn.close()
        messagebox.showinfo("Thành công", f"Đã xuất dữ liệu staff ra file:\n{file_path}")
    except Exception as e:
        messagebox.showerror("Lỗi", f"Không thể xuất file:\n{str(e)}")

def import_staff_from_excel(db_path=db_path_default):
    try:
        root = tk.Tk()
        root.withdraw()

        file_path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx")],
            title="Chọn file Excel để import"
        )

        if not file_path:
            print("Đã huỷ import file.")
            return

        # Đọc Excel, giữ NaN để SQLite ghi NULL
        df = pd.read_excel(file_path, dtype=str)

        # Bỏ cột id nếu có
        if "id" in df.columns:
            df = df.drop(columns=["id"])

        # Ép active_staff về kiểu boolean/integer
        if "active_staff" in df.columns:
            # chuyển các giá trị "1", "True", "true"… thành 1, còn lại 0
            df["active_staff"] = df["active_staff"].apply(
                lambda x: 1 if str(x).strip().lower() in ("1", "true", "yes") else 0
            )
        else:
            raise ValueError("File Excel thiếu cột 'active_staff'")

        # NaN thành None để SQLite ghi NULL
        df = df.where(df.notnull(), None)
        df = df.replace({'nan': None, 'NaN': None, '': None})



        conn = sqlite3.connect(db_path)
        df.to_sql("staff", conn, if_exists="append", index=False)
        conn.commit()
        conn.close()

        messagebox.showinfo("Thành công", f"Đã import {len(df)} dòng từ:\n{file_path}")
    except Exception as e:
        messagebox.showerror("Lỗi", f"Không thể import file:\n{str(e)}")








def add_assets(asset_name_add, code_add, serial_number_add, purchase_date_add, status_add,value_add, bams_add, notes_add):
    try:
        conn = sqlite3.connect(db_path_default)
        cursor = conn.cursor()

        asset_name_add = asset_name_add.strip()
        code_add = code_add.strip()
        serial_number_add = serial_number_add.strip()
        purchase_date_add = purchase_date_add.strip()
        status_add = status_add.strip()
        value_add = int(value_add)
        bams_add = bams_add.strip()
        notes_add = notes_add.strip()

        if asset_name_add == "":
            asset_name_add = None
        if code_add == "":
            code_add = None
        if serial_number_add == "":
            serial_number_add = None
        if purchase_date_add == "":
            purchase_date_add = None
        if status_add == "":
            status_add = None
        if bams_add == "":
            bams_add = None
        if notes_add == "":
            notes_add = None


        if not asset_name_add:
            conn.close()
            return "Lỗi: Loại tài sản không được để trống!"
        if not serial_number_add:
            conn.close()
            return "Lỗi: Số seri không được để trống!"
        if not purchase_date_add:
            conn.close()
            return "Lỗi: Năm nhận tài sản không được để trống!"

        cursor.execute("SELECT * FROM assets WHERE serial_number = ?", (serial_number_add,))
        results = cursor.fetchone()
        if results:
            conn.close()
            return "Lỗi: Số seri đã bị trùng!"

        cursor.execute("SELECT * FROM assets WHERE code = ?", (code_add,))
        results = cursor.fetchone()

        if results:
            conn.close()
            return "Lỗi: Mã hiệu đã bị trùng!"

        cursor.execute(
            "INSERT INTO assets (asset_name, code, serial_number, purchase_date, status, bams, active_asset, value, notes) VALUES (?, ?, ?, ?, ?, ?, ?, ?,?)",
            (asset_name_add, code_add, serial_number_add, purchase_date_add, status_add, bams_add, True, value_add,notes_add))
        conn.commit()
        conn.close()
        return "Thêm tài sản thành công!"
    except Exception as e:
        messagebox.showerror("Lỗi", f"\n{str(e)}")

def modify_assets(asset_name_add, code_add, serial_number_add, purchase_date_add, status_add, value_add, bams_add, notes_add, search):
    try:
        conn = sqlite3.connect(db_path_default)
        cursor = conn.cursor()

        asset_name_add = asset_name_add.strip()
        code_add = code_add.strip()
        serial_number_add = serial_number_add.strip()
        purchase_date_add = purchase_date_add.strip()
        status_add = status_add.strip()
        value_add = int(value_add)
        notes_add = notes_add.strip()
        bams_add = bams_add.strip()
        search = search.strip()

        if asset_name_add == "":
            asset_name_add = None
        if code_add == "":
            code_add = None
        if serial_number_add == "":
            serial_number_add = None
        if purchase_date_add == "":
            purchase_date_add = None
        if status_add == "":
            status_add = None
        if bams_add == "":
            bams_add = None
        if notes_add == "":
            notes_add = None
        if search == "":
            conn.close()
            return "Lỗi: Tài sản cần sửa đổi không được để trống!"

        if not asset_name_add:
            conn.close()
            return "Lỗi: Loại tài sản không được để trống!"
        if not serial_number_add:
            conn.close()
            return "Lỗi: Số seri không được để trống!"
        if not purchase_date_add:
            conn.close()
            return "Lỗi: Năm nhận tài sản không được để trống!"

        cursor.execute("SELECT * FROM assets WHERE serial_number = ?", (search,))
        results = cursor.fetchone()
        if not results:
            conn.close()
            return "Không tìm thấy tài sản cần sửa!"

        serial_number_modify = results[3]
        code_modify = results[2]
        asset_id_modify = results[0]

        if serial_number_add == serial_number_modify:
            pass
        elif serial_number_add != serial_number_modify:
            cursor.execute("SELECT * FROM assets WHERE serial_number = ?", (serial_number_add,))
            results = cursor.fetchone()
            if results:
                conn.close()
                return "Lỗi: Số seri đã bị trùng!"

        if code_add == code_modify:
            pass
        elif code_add != code_modify:
            cursor.execute("SELECT * FROM assets WHERE code = ?", (code_add,))
            results = cursor.fetchone()
            if results:
                if results[2] == "":
                    pass
                else:
                    conn.close()
                    return "Lỗi: Mã hiệu đã bị trùng!"

        cursor.execute("""
                       UPDATE assets
                       SET asset_name    = ?,
                           code          = ?,
                           serial_number = ?,
                           purchase_date = ?,
                           status        = ?,
                           value         = ?,
                           bams          = ?,
                           notes         = ?
                       WHERE id = ?
                       """, (asset_name_add, code_add, serial_number_add, purchase_date_add, status_add,value_add, bams_add, notes_add,
                             asset_id_modify))
        conn.commit()
        conn.close()
        return "Sửa thông tin tài sản thành công!"
    except Exception as e:
        messagebox.showerror("Lỗi", f"\n{str(e)}")



def search_assets(search):
    try:
        search = search.strip()
        if search == "":
            return "Thông tin vừa rồi không khớp tài sản nào!"
        conn = sqlite3.connect(db_path_default)
        cursor = conn.cursor()

        cursor.execute("""
                       SELECT *
                       FROM assets
                       WHERE serial_number = ?
                       """, (search,))
        result = cursor.fetchone()
        conn.close()
        if result:
            return result
        else:
            return "Thông tin vừa rồi không khớp tài sản nào!"
    except Exception as e:
        messagebox.showerror("Lỗi", f"\n{str(e)}")

def remove_assets(search):
    try:
        search = search.strip()
        if search == "":
            return "Không tìm thấy tài sản để thu hồi."

        conn = sqlite3.connect(db_path_default)
        cursor = conn.cursor()
        cursor.execute("""
                       SELECT *
                       FROM assets
                       WHERE serial_number = ?
                       """, (search,))
        result = cursor.fetchone()

        if not result:
            conn.close()
            return "Không tìm thấy tài sản để thu hồi."

        asset_id_remove = result[0]

        cursor.execute("""
                       SELECT *
                       FROM asset_assignments
                       WHERE asset_id = ?
                         AND returned_date IS NULL
                       """, (asset_id_remove,))
        holding_asset = cursor.fetchone()

        if holding_asset:
            conn.close()
            return f"Không thể trả VP tài sản {search} vì có cán bộ đang giữ tài sản."

        if result[6]:
            cursor.execute("""
                           UPDATE assets
                           SET active_asset = ?,
                               return_date  = ?
                           WHERE id = ?
                           """, (False, date.today(), asset_id_remove))

            conn.commit()
            conn.close()
            return f"Đã thu hồi tài sản {search} thành công!"
        else:
            return f"Không thể thu hồi vì:\nTài sản {search} đã được thu hồi từ trước!"

    except Exception as e:
        messagebox.showerror("Lỗi", f"\n{str(e)}")



def export_assets_to_excel(db_path=db_path_default):
    try:
        root = tk.Tk()
        root.withdraw()
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            title="Chọn nơi lưu file Excel",
            initialfile="assets.xlsx"
        )

        if not file_path:
            print("Đã huỷ xuất file.")
            return

        conn = sqlite3.connect(db_path)
        df = pd.read_sql_query("SELECT * FROM assets", conn)

        df.to_excel(file_path, index=False)
        conn.close()
        messagebox.showinfo("Thành công", f"Đã xuất dữ liệu assets ra file:\n{file_path}")
    except Exception as e:
        messagebox.showerror("Lỗi", f"Không thể xuất file:\n{str(e)}")

ASSETS_IMPORT_COLUMNS = [
    "asset_name", "code", "serial_number", "purchase_date",
    "status", "active_asset", "return_date", "value", "bams", "notes",
]


def upsert_assets_dataframe(df, db_path=db_path_default):
    """Chuẩn hóa DataFrame tài sản và nạp/cập nhật vào bảng assets theo serial_number."""
    if "active_asset" not in df.columns:
        raise ValueError("File Excel thiếu cột 'active_asset'")
    if "asset_name" not in df.columns or "serial_number" not in df.columns:
        raise ValueError("File Excel thiếu cột bắt buộc 'asset_name' hoặc 'serial_number'")

    cols = [c for c in ASSETS_IMPORT_COLUMNS if c in df.columns]
    df_clean = df[cols].copy()

    def _clean_text(val):
        if val is None or pd.isna(val):
            return None
        s = str(val).strip()
        if s in ("", "nan", "NaN", "None"):
            return None
        return s

    def _clean_int(val):
        s = _clean_text(val)
        if s is None:
            return None
        try:
            return int(float(s.replace(",", "")))
        except ValueError:
            return 0

    df_clean["active_asset"] = df_clean["active_asset"].apply(
        lambda x: 1 if str(x).strip().lower() in ("1", "true", "yes") else 0
    )
    if "value" in df_clean.columns:
        df_clean["value"] = df_clean["value"].apply(_clean_int)
    for col in cols:
        if col not in ("active_asset", "value"):
            df_clean[col] = df_clean[col].apply(_clean_text)

    df_clean = df_clean[df_clean["serial_number"].notna() & df_clean["asset_name"].notna()]

    placeholders = ", ".join(["?"] * len(cols))
    update_cols = [c for c in cols if c != "serial_number"]
    update_clause = ", ".join(f"{c} = excluded.{c}" for c in update_cols)
    sql = (
        f"INSERT INTO assets ({', '.join(cols)}) VALUES ({placeholders}) "
        f"ON CONFLICT(serial_number) DO UPDATE SET {update_clause}"
    )

    records = [
        tuple(None if pd.isna(row[c]) else row[c] for c in cols)
        for _, row in df_clean.iterrows()
    ]

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.executemany(sql, records)
    conn.commit()
    conn.close()
    return len(records)


def import_assets_from_excel(db_path=db_path_default):
    try:
        root = tk.Tk()
        root.withdraw()

        file_path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx")],
            title="Chọn file Excel để import"
        )

        if not file_path:
            print("Đã huỷ import file.")
            return

        # Đọc Excel dạng chuỗi để giữ số 0 đầu
        df = pd.read_excel(file_path, dtype=str)
        count = upsert_assets_dataframe(df, db_path=db_path)

        messagebox.showinfo("Thành công", f"Đã import/cập nhật {count} dòng từ:\n{file_path}")

    except Exception as e:
        messagebox.showerror("Lỗi", f"Không thể import file:\n{str(e)}")






def assign_assets(qtablewidget, label_widget=None):
    try:
        conn = sqlite3.connect(db_path_default)
        cursor = conn.cursor()
        today = date.today()

        asset_staff_pairs = []

        for row_index in range(qtablewidget.rowCount()):
            asset_item = qtablewidget.item(row_index, 0)
            staff_item = qtablewidget.item(row_index, 2)
            notes_item = qtablewidget.item(row_index, 7)

            if not asset_item or not staff_item:
                continue

            try:
                asset_id = int(asset_item.text())
                staff_id = int(staff_item.text())
            except ValueError:
                continue

            notes = notes_item.text() if notes_item else None
            if notes == "":
                notes = None

            asset_staff_pairs.append((asset_id, staff_id, notes))

        if not asset_staff_pairs:
            if label_widget:
                label_widget.setText("Không có dữ liệu trong bảng để giao")
            conn.close()
            return

        asset_ids = [pair[0] for pair in asset_staff_pairs]
        staff_ids = [pair[1] for pair in asset_staff_pairs]

        # Kiểm tra các tài sản đang được giao
        placeholders = ",".join(["?"] * len(asset_ids))
        cursor.execute(f"""
                SELECT asset_id FROM asset_assignments
                WHERE asset_id IN ({placeholders}) AND returned_date IS NULL
            """, asset_ids)
        rows_conflict = cursor.fetchall()
        if rows_conflict:
            conflicted_ids = [str(r[0]) for r in rows_conflict]
            conn.close()
            if label_widget:
                label_widget.setText(f"Không thể ghi. Các tài sản đang có người giữ: {', '.join(conflicted_ids)}")
            return

        # Kiểm tra trạng thái active của tất cả asset và staff
        cursor.execute(f"SELECT id FROM assets WHERE id IN ({placeholders}) AND active_asset=0", asset_ids)
        inactive_assets = [str(r[0]) for r in cursor.fetchall()]

        cursor.execute(f"SELECT id FROM staff WHERE id IN ({','.join(['?'] * len(staff_ids))} ) AND active_staff=0",
                       staff_ids)
        inactive_staff = [str(r[0]) for r in cursor.fetchall()]

        if inactive_assets or inactive_staff:
            msg = ""
            if inactive_assets:
                msg += f"Tài sản không active: {', '.join(inactive_assets)}. "
            if inactive_staff:
                msg += f"Nhân viên không active: {', '.join(inactive_staff)}."
            if label_widget:
                label_widget.setText(msg)
            conn.close()
            return

        # Insert các bản ghi hợp lệ
        inserted_count = 0
        for asset_id, staff_id, notes in asset_staff_pairs:
            cursor.execute("""
                           INSERT INTO asset_assignments (asset_id, staff_id, assigned_date, notes)
                           VALUES (?, ?, ?, ?)
                           """, (asset_id, staff_id, today, notes))
            inserted_count += 1

        conn.commit()
        conn.close()

        if label_widget:
            label_widget.setText(f"Đã ghi {inserted_count} bản ghi giao tài sản thành công")

    except Exception as e:
        messagebox.showerror("Lỗi", f"\n{str(e)}")




def return_assets(qtablewidget, label_widget=None):
    try:
        today = date.today().strftime("%Y-%m-%d")

        # Lấy toàn bộ dữ liệu trong table2
        rows_data = []
        for row_index in range(qtablewidget.rowCount()):
            asset_id_item = qtablewidget.item(row_index, 0)  # ID tài sản
            staff_id_item = qtablewidget.item(row_index, 2)  # ID cán bộ
            note_item     = qtablewidget.item(row_index, 8)  # Ghi chú

            if asset_id_item and staff_id_item:
                try:
                    asset_id = int(asset_id_item.text())
                    staff_id = int(staff_id_item.text())
                except ValueError:
                    continue  # bỏ qua dòng lỗi

                note = note_item.text() if note_item else None
                if note == "":
                    note = None
                rows_data.append((asset_id, staff_id, note))

        if not rows_data:
            if label_widget is not None:
                label_widget.setText("Không có dữ liệu trong bảng để trả")
            return

        # Cập nhật vào DB
        conn = sqlite3.connect(db_path_default)
        cursor = conn.cursor()

        count = 0
        for asset_id, staff_id, note in rows_data:
            cursor.execute("""
                           UPDATE asset_assignments
                           SET returned_date = ?,
                               notes         = CASE
                                                   WHEN ? IS NOT NULL AND notes IS NULL THEN ?
                                                   WHEN ? IS NOT NULL AND notes IS NOT NULL THEN notes || CHAR (10) || ?
                               ELSE notes
                           END
                           WHERE asset_id = ?
                              AND staff_id = ?
                              AND returned_date IS NULL
                           """, (today, note, note, note, note, asset_id, staff_id))
            count += cursor.rowcount

        conn.commit()
        conn.close()

        if label_widget is not None:
            if count == 0:
                label_widget.setText(f"Không có bản ghi trả tài sản nào được cập nhật")
            else:
                label_widget.setText(f"Đã cập nhật {count} bản ghi trả tài sản thành công")

    except Exception as e:
        messagebox.critical(None, "Lỗi", f"\n{str(e)}")


def search_serial_number_in_warehouse(search):
    try:
        search = search.strip()
        if search == "":
            return "Thông tin vừa rồi không khớp tài sản nào trong kho!"
        conn = sqlite3.connect(db_path_default)
        cursor = conn.cursor()

        cursor.execute("""
                       SELECT a.id,
                              a.asset_name,
                              a.code,
                              a.serial_number,
                              a.status
                       FROM assets a
                                LEFT JOIN asset_assignments aa
                                          ON a.id = aa.asset_id
                                              AND aa.returned_date IS NULL
                       WHERE a.serial_number = ?
                         AND aa.asset_id IS NULL
                         AND a.active_asset = 1
                       """, (search,))
        result = cursor.fetchall()
        conn.close()
        if result:
            return result
        else:
            return "Thông tin vừa rồi không khớp tài sản nào trong kho!"

    except Exception as e:
        messagebox.showerror("Lỗi", f"\n{str(e)}")

def search_assets_of_staff(search):
    try:
        search = search.strip()
        if search == "":
            return "Thông tin vừa rồi không khớp cán bộ nào!"

        conn = sqlite3.connect(db_path_default)
        cursor = conn.cursor()

        cursor.execute("""
                       SELECT *
                       FROM staff
                       WHERE name = ?
                          OR employee_code = ?
                          OR email = ?
                       """, (search, search, search))
        result = cursor.fetchone()
        if result:
            pass
        else:
            conn.close()
            return "Thông tin vừa rồi không khớp cán bộ nào!"

        staff_search_id = result[0]
        cursor.execute("""
                       SELECT a.id   AS asset_id,
                              a.asset_name,
                              a.code,
                              a.serial_number,
                              a.status,
                              a.purchase_date,
                              s.name AS staff_name,
                              aa.notes
                       FROM assets a
                                LEFT JOIN asset_assignments aa ON aa.asset_id = a.id
                                LEFT JOIN staff s ON aa.staff_id = s.id
                       WHERE s.id = ?
                         AND aa.returned_date IS NULL
                       """, (staff_search_id,))
        result = cursor.fetchall()
        conn.close()
        if result:
            return result
        else:
            return "Cán bộ không có tài sản nào!"


    except Exception as e:
        messagebox.showerror("Lỗi", f"\n{str(e)}")

def search_serial_number_of_staff(search_serial_number, search_staff):
    try:
        search_serial_number = search_serial_number.strip()
        search_staff = search_staff.strip()

        if search_serial_number == "":
            return "Thông tin vừa rồi không khớp tài sản nào của cán bộ!"
        elif search_staff == "":
            return "Thông tin vừa rồi không khớp tài sản nào của cán bộ!"


        conn = sqlite3.connect(db_path_default)
        cursor = conn.cursor()

        cursor.execute("""
                       SELECT *
                       FROM staff
                       WHERE name = ?
                          OR employee_code = ?
                          OR email = ?
                       """, (search_staff, search_staff, search_staff))
        result = cursor.fetchone()
        if not result:
            conn.close()
            return "Thông tin vừa rồi không khớp tài sản nào của cán bộ!"
        staff_search_id = result[0]


        cursor.execute("""
                       SELECT *
                       FROM assets
                       WHERE serial_number = ?
                       """, (search_serial_number,))
        result = cursor.fetchone()
        if not result:
            conn.close()
            return "Thông tin vừa rồi không khớp tài sản nào của cán bộ!"
        asset_search_id = result[0]

        cursor.execute("""
                       SELECT a.id   AS asset_id,
                              s.name AS staff_name,
                              s.id   AS staff_id,
                              a.asset_name,
                              a.code,
                              a.serial_number,
                              a.status,
                              aa.assigned_date
                       FROM asset_assignments aa
                                JOIN assets a ON aa.asset_id = a.id
                                JOIN staff s ON aa.staff_id = s.id
                       WHERE aa.asset_id = ?
                         AND aa.staff_id = ?
                         AND aa.returned_date IS NULL
                       """, (asset_search_id, staff_search_id))
        results = cursor.fetchall()
        conn.close()
        if results:
            return results
        else:
            return "Thông tin vừa rồi không khớp tài sản nào của cán bộ!"

    except Exception as e:
        messagebox.showerror("Lỗi", f"\n{str(e)}")

def export_assets_assignments_to_excel(db_path=db_path_default):
    try:
        root = tk.Tk()
        root.withdraw()
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            title="Chọn nơi lưu file Excel",
            initialfile="asset_assignments.xlsx"
        )

        if not file_path:
            print("Đã huỷ xuất file.")
            return

        conn = sqlite3.connect(db_path)
        df = pd.read_sql_query("SELECT * FROM asset_assignments", conn)

        df.to_excel(file_path, index=False)
        conn.close()
        messagebox.showinfo("Thành công", f"Đã xuất dữ liệu asset_assignments ra file:\n{file_path}")
    except Exception as e:
        messagebox.showerror("Lỗi", f"Không thể xuất file:\n{str(e)}")


def import_assets_assignments_from_excel(db_path=db_path_default):
    try:
        root = tk.Tk()
        root.withdraw()

        file_path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx")],
            title="Chọn file Excel để import"
        )

        if not file_path:
            print("Đã huỷ import file.")
            return

        # 1. Đọc Excel dạng chuỗi để giữ nguyên giá trị (số 0 đầu…)
        df = pd.read_excel(file_path, dtype=str)

        # 2. Bỏ cột id nếu có
        if "id" in df.columns:
            df = df.drop(columns=["id"])

        # 3. Ép asset_id & staff_id về kiểu số nguyên (nullable)
        if "asset_id" in df.columns:
            df["asset_id"] = pd.to_numeric(df["asset_id"], errors="coerce").astype("Int64")
        else:
            raise ValueError("File Excel thiếu cột 'asset_id'")

        if "staff_id" in df.columns:
            df["staff_id"] = pd.to_numeric(df["staff_id"], errors="coerce").astype("Int64")
        else:
            raise ValueError("File Excel thiếu cột 'staff_id'")

        # 5. Thay NaN/'nan'/'NaN'/rỗng thành None để SQLite ghi NULL
        df = df.where(df.notnull(), None)
        df = df.replace({'nan': None, 'NaN': None, '': None})

        # 6. Ghi vào SQLite
        conn = sqlite3.connect(db_path)
        df.to_sql("asset_assignments", conn, if_exists="append", index=False)
        conn.commit()
        conn.close()

        messagebox.showinfo("Thành công", f"Đã import {len(df)} dòng từ:\n{file_path}")
    except Exception as e:
        messagebox.showerror("Lỗi", f"Không thể import file:\n{str(e)}")


def qtablewidget_to_df(table):
    """Chuyển QTableWidget thành pandas DataFrame"""
    headers = [table.horizontalHeaderItem(col).text()
               for col in range(table.columnCount())]

    data = []
    for row in range(table.rowCount()):
        row_data = []
        for col in range(table.columnCount()):
            item = table.item(row, col)
            row_data.append(item.text() if item else "")
        data.append(row_data)

    return pd.DataFrame(data, columns=headers)

def export_two_tables_to_excel(table1, table2):
    """Xuất 2 QTableWidget ra 1 file Excel với Table1 trước, Table2 sau"""
    file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx")],
        title="Chọn nơi lưu file Excel",
        initialfile="asset_assignments_in_lot.xlsx"
    )

    if not file_path:
        return  # người dùng bấm hủy

    try:
        # Table1 trước
        df1 = qtablewidget_to_df(table1)
        df2 = qtablewidget_to_df(table2)

        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            df1.to_excel(writer, sheet_name="Danh sách nhận TS", index=False)
            df2.to_excel(writer, sheet_name="Danh sách trả TS", index=False)

        messagebox.showinfo("Thành công", f"Đã xuất file Excel:\n{file_path}")
    except Exception as e:
        messagebox.showerror("Lỗi", f"Không thể xuất file Excel:\n{str(e)}")


def import_two_tables_from_excel_to_widgets(table1, table2):
    """
    Đọc file Excel chứa 2 sheet:
    - Sheet1: Danh sách nhận TS
    - Sheet2: Danh sách trả TS
    và đổ vào QTableWidget table1 và table2
    """
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel Files", "*.xlsx")],
        title="Chọn file Excel để import"
    )

    if not file_path:
        return  # Người dùng hủy

    try:
        # Đọc 2 sheet
        sheets = pd.read_excel(file_path, sheet_name=None)
        df1 = sheets.get("Danh sách nhận TS")
        df2 = sheets.get("Danh sách trả TS")

        if df1 is None or df2 is None:
            raise ValueError("Không tìm thấy 1 hoặc cả 2 sheet trong file Excel")

        # Đổ dữ liệu vào QTableWidget
        df_to_qtablewidget(df1, table1)
        df_to_qtablewidget(df2, table2)

        messagebox.showinfo("Thành công", f"Đã import dữ liệu từ file Excel:\n{file_path}")
    except Exception as e:
        messagebox.showerror("Lỗi", f"Không thể import file Excel:\n{str(e)}")

def df_to_qtablewidget(df, qtablewidget):
    """
    Đổ dữ liệu từ pandas DataFrame vào QTableWidget.
    """
    qtablewidget.setRowCount(len(df))
    qtablewidget.setColumnCount(len(df.columns))
    qtablewidget.setHorizontalHeaderLabels([str(col) for col in df.columns])

    for row in range(len(df)):
        for col in range(len(df.columns)):
            value = df.iat[row, col]
            if pd.isna(value):  # nếu ô trống
                value = ""
            qtablewidget.setItem(row, col, QtWidgets.QTableWidgetItem(str(value)))

def export_assets_of_staffs_to_excel(table):
    """Xuất 2 QTableWidget ra 1 file Excel với Table1 trước, Table2 sau"""
    file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx")],
        title="Chọn nơi lưu file Excel",
        initialfile="assets_by_staff.xlsx"
    )

    if not file_path:
        return  # người dùng bấm hủy

    try:
        # Table1 trước
        df = qtablewidget_to_df(table)

        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name="Danh sách tài sản theo cán bộ", index=False)

        messagebox.showinfo("Thành công", f"Đã xuất file Excel:\n{file_path}")
    except Exception as e:
        messagebox.showerror("Lỗi", f"Không thể xuất file Excel:\n{str(e)}")

def add_table_from_qtable(doc, qtable, after_paragraph=None):
    try:
        row_count = qtable.rowCount()
        col_count = qtable.columnCount()

        # không có dữ liệu thì bỏ qua
        if row_count == 0:
            return None

        headers = [qtable.horizontalHeaderItem(col).text() for col in range(col_count)]

        table = doc.add_table(rows=row_count + 1, cols=col_count)
        try:
            table.style = 'Table Grid'
        except:
            pass

        # header
        for j, header in enumerate(headers):
            table.cell(0, j).text = header

        # dữ liệu
        for i in range(row_count):
            for j in range(col_count):
                item = qtable.item(i, j)
                table.cell(i + 1, j).text = item.text() if item else ""

        # di chuyển bảng vào ngay sau paragraph nếu được chỉ định
        if after_paragraph is not None:
            after_paragraph._p.addnext(table._tbl)

        return table
    except Exception as e:
        messagebox.showerror("Lỗi", f"\n{str(e)}")

def _collect_unique_staff_names(qtable1, qtable2):
    names = []
    seen = set()
    for qt in (qtable1, qtable2):
        if qt is None:
            continue
        for i in range(qt.rowCount()):
            item = qt.item(i, 1)
            name = item.text().strip() if item and item.text() else ""
            if name and name not in seen:
                seen.add(name)
                names.append(name)
    return names


def _replace_markers_in_paragraph(para, replacements):
    if not any(key in para.text for key in replacements):
        return
    for run in para.runs:
        for key, val in replacements.items():
            if key in run.text:
                run.text = run.text.replace(key, val)
    # Fallback nếu Word tách marker qua nhiều run
    if any(key in para.text for key in replacements):
        text = para.text
        for key, val in replacements.items():
            text = text.replace(key, val)
        para.text = text


def fill_date_placeholders(doc, dt=None):
    if dt is None:
        dt = datetime.now()
    replacements = {
        "{{DAY}}": f"{dt.day:02d}",
        "{{MONTH}}": f"{dt.month:02d}",
        "{{YEAR}}": f"{dt.year}",
    }

    for para in doc.paragraphs:
        _replace_markers_in_paragraph(para, replacements)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    _replace_markers_in_paragraph(para, replacements)


def _format_sign_cell(cell, text, font_name="Times New Roman", font_pt=13, bold=True, italic=True):
    # Xóa các paragraph thừa trong ô, giữ lại paragraph đầu tiên
    for p in list(cell.paragraphs)[1:]:
        p._element.getparent().remove(p._element)
    para = cell.paragraphs[0]
    para.clear()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = para.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.name = font_name
    run.font.size = Pt(font_pt)
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is not None:
        rFonts.set(qn("w:cs"), font_name)
        rFonts.set(qn("w:eastAsia"), font_name)


def fill_sign_placeholder_in_tables(doc, qtable1, qtable2):
    names = _collect_unique_staff_names(qtable1, qtable2)

    for table in doc.tables:
        target = None
        for r_idx, row in enumerate(table.rows):
            for c_idx, cell in enumerate(row.cells):
                if "{{SIGN}}" in cell.text:
                    target = (r_idx, c_idx)
                    break
            if target is not None:
                break

        if target is None:
            continue

        row_idx, col_idx = target
        base_row = table.rows[row_idx]
        col_widths = [c.width for c in base_row.cells]
        space_height = Pt(57.5)
        if row_idx > 0 and table.rows[row_idx - 1].height is not None:
            space_height = table.rows[row_idx - 1].height

        first_cell = base_row.cells[col_idx]
        _format_sign_cell(first_cell, names[0] if names else "")

        anchor_tr = base_row._tr
        for name in names[1:]:
            space_row = table.add_row()
            space_row.height = space_height
            for j, w in enumerate(col_widths):
                if w is not None:
                    space_row.cells[j].width = w
                space_row.cells[j].text = ""

            name_row = table.add_row()
            if base_row.height is not None:
                name_row.height = base_row.height
            for j, w in enumerate(col_widths):
                if w is not None:
                    name_row.cells[j].width = w
                if j == col_idx:
                    _format_sign_cell(name_row.cells[j], name)
                else:
                    name_row.cells[j].text = ""

            # Chèn ngay sau dòng tên trước đó (giữ nguyên thứ tự các dòng sau trong bảng)
            anchor_tr.addnext(space_row._tr)
            space_row._tr.addnext(name_row._tr)
            anchor_tr = name_row._tr

    return names


def add_sign_table(doc, qtable1, qtable2, after_paragraph=None):
    try:
        names = _collect_unique_staff_names(qtable1, qtable2)

        if not names:
            return None

        # +1 hàng để chứa header
        total_rows = len(names) + 1
        table = doc.add_table(rows=total_rows, cols=2)
        try:
            table.style = 'Table Grid'
        except:
            pass

        # Header (hàng 0)
        table.cell(0, 0).text = "                                Bên bàn giao"
        table.cell(0, 1).text = "                                            Ký"

        # Dữ liệu (từ hàng 1 trở đi)
        for i, name in enumerate(names, start=1):
            table.cell(i, 0).text = name
            table.cell(i, 1).text = " " * 20  # để trống chỗ ký

        # di chuyển bảng vào ngay sau paragraph
        if after_paragraph is not None:
            after_paragraph._p.addnext(table._tbl)

        return table

    except Exception as e:
        messagebox.showerror("Lỗi", f"\n{str(e)}")



def add_handover_documents(pixmap_item, document_date, type_handover_documents, notes):
    try:
        if pixmap_item is None:
            raise ValueError("pixmap_item is None – chưa có ảnh để lưu")

        document_date = document_date.strip()
        type_handover_documents = type_handover_documents.strip()
        notes = notes.strip()

        if notes == "":
            notes = None
        if type_handover_documents == "":
            type_handover_documents = None
        if document_date == "":
            document_date = None

        if not type_handover_documents:
            return "Lỗi: Loại biên bản bàn giao không được để trống!"
        if not document_date:
            return "Lỗi: Ngày biên bản bàn giao không được để trống!"

        # Lấy QPixmap từ pixmap_item
        pixmap = pixmap_item.pixmap()

        # Chuyển QPixmap thành bytes
        ba = QByteArray()
        buffer = QBuffer(ba)
        buffer.open(QIODevice.WriteOnly)
        pixmap.save(buffer, "PNG")
        image_bytes = ba.data()  # bytes để lưu DB

        # Lưu vào DB
        conn = sqlite3.connect(db_path_default)
        cursor = conn.cursor()
        cursor.execute("""
                       INSERT INTO handover_documents (document_date, type_handover_documents, document_file, notes)
                       VALUES (?, ?, ?, ?)
                       """, (document_date, type_handover_documents, image_bytes, notes))
        conn.commit()
        conn.close()
    except Exception as e:
        messagebox.showerror("Lỗi", f"\n{str(e)}")


def delete_handover_document(handover_documents_id):
    try:
        conn = sqlite3.connect(db_path_default)
        cursor = conn.cursor()
        cursor.execute("DELETE FROM handover_documents WHERE id = ?", (handover_documents_id,))
        conn.commit()
        conn.close()

        #self.label.setText(f"Đã xóa biên bản bàn giao ID {doc_id}")

    except Exception as e:
        self.label.setText(f"Lỗi xóa biên bản: {e}")

def search_asset_assignments_id_and_handover_documents_id(handover_documents_id,asset_assignments_id):
    try:
        handover_documents_id = handover_documents_id.strip()
        asset_assignments_id = asset_assignments_id.strip()

        if not handover_documents_id:
            return "Thông tin vừa rồi không khớp với ID biên bản nào!"
        if not asset_assignments_id:
            return "Thông tin vừa rồi không khớp với ID assignment nào!"


        conn = sqlite3.connect(db_path_default)
        cursor = conn.cursor()

        cursor.execute("""
                       SELECT *
                       FROM handover_documents
                       WHERE id = ?
                       """, (handover_documents_id,))
        result = cursor.fetchone()
        if result:
            pass
        else:
            conn.close()
            return "Thông tin vừa rồi không khớp với ID biên bản nào!"


        cursor.execute("""
                       SELECT *
                       FROM asset_assignments
                       WHERE id = ?
                       """, (asset_assignments_id,))
        result = cursor.fetchone()
        if result:
            pass
        else:
            conn.close()
            return "Thông tin vừa rồi không khớp với ID assignment nào!"

        cursor.execute("""
                       SELECT handover_documents_assign_id,
                              handover_documents_return_id
                       FROM asset_assignments
                       WHERE id = ?
                       """, (asset_assignments_id,))
        row = cursor.fetchone()
        conn.close()

        if row:
            assign_id, return_id = row
            flags = []
            if assign_id:
                flags.append(1)
            if return_id:
                flags.append(2)

            if not flags:  # cả hai rỗng
                return 0
            return flags  # trả về [1], [2] hoặc [1,2]
        else:
            return 0


    except Exception as e:
        messagebox.showerror("Lỗi", f"\n{str(e)}")


def add_handover_document_to_assignment(qtablewidget, label_widget=None):
    try:
        rows_data = []

        for row_index in range(qtablewidget.rowCount()):
            asset_assignment_id_item = qtablewidget.item(row_index, 0)  # ID assignment
            handover_documents_assign_id_item = qtablewidget.item(row_index, 1)  # ID bàn giao nhận
            handover_documents_return_id_item = qtablewidget.item(row_index, 2)  # ID bàn giao trả

            if asset_assignment_id_item:
                asset_assignment_id_text = asset_assignment_id_item.text().strip()

                # lấy text, có thể rỗng
                assign_text = handover_documents_assign_id_item.text().strip() if handover_documents_assign_id_item else ""
                return_text = handover_documents_return_id_item.text().strip() if handover_documents_return_id_item else ""

                # chỉ cần 1 trong 2 cột có dữ liệu
                if assign_text or return_text:
                    try:
                        asset_assignment_id = int(asset_assignment_id_text)
                        # nếu rỗng thì để None để update null
                        handover_documents_assign_id = int(assign_text) if assign_text else None
                        handover_documents_return_id = int(return_text) if return_text else None
                    except ValueError:
                        continue  # bỏ qua dòng lỗi

                    rows_data.append((asset_assignment_id,
                                      handover_documents_assign_id,
                                      handover_documents_return_id))

        if not rows_data:
            if label_widget is not None:
                label_widget.setText("Không có dữ liệu trong bảng để ghi")
            return

        # cập nhật DB
        conn = sqlite3.connect(db_path_default)
        cursor = conn.cursor()

        count = 0
        for asset_assignment_id, handover_documents_assign_id, handover_documents_return_id in rows_data:
            # tạo danh sách SET động
            set_parts = []
            params = []

            if handover_documents_assign_id is not None:  # chỉ update khi có giá trị
                set_parts.append("handover_documents_assign_id = ?")
                params.append(handover_documents_assign_id)

            if handover_documents_return_id is not None:  # chỉ update khi có giá trị
                set_parts.append("handover_documents_return_id = ?")
                params.append(handover_documents_return_id)

            # nếu không có gì để update thì bỏ qua
            if not set_parts:
                continue

            sql = f"""
                UPDATE asset_assignments
                SET {', '.join(set_parts)}
                WHERE id = ?
            """
            params.append(asset_assignment_id)

            cursor.execute(sql, params)
            count += cursor.rowcount

        conn.commit()
        conn.close()

        if label_widget is not None:
            if count == 0:
                label_widget.setText(f"Không có bản ghi nào được cập nhật")
            else:
                label_widget.setText(f"Đã cập nhật {count} bản ghi biên bản thành công")
    except Exception as e:
        messagebox.showerror("Lỗi", f"\n{str(e)}")

def backup_db():
    try:
        if not os.path.exists(db_path_default):
            raise FileNotFoundError(f"Không tìm thấy file DB: {db_path_default}")

            # Lấy thư mục chứa exe (khi đóng gói) hoặc file .py hiện tại
        if getattr(sys, 'frozen', False):  # đang chạy dạng .exe
            app_dir = os.path.dirname(sys.executable)
        else:  # đang chạy bằng Python
            app_dir = os.path.dirname(os.path.abspath(__file__))

            # Tạo thư mục "back_up" cạnh file exe/py nếu chưa có
        backup_dir = os.path.join(app_dir, "back_up")
        os.makedirs(backup_dir, exist_ok=True)  # tạo nếu chưa có

        # Đường dẫn file backup
        dest = os.path.join(backup_dir, "back_up_assets.db")

        shutil.copy2(db_path_default, dest)
        return dest  # trả về đường dẫn file backup
    except Exception as e:
        messagebox.showerror("Lỗi", f"\n{str(e)}")



# conn = sqlite3.connect(db_path_default)
# cursor = conn.cursor()
# cursor.execute("""
#                        SELECT *
#                        FROM assets
#                        """)
# rows = cursor.fetchall()
# for row in rows:
#     print (row)
# conn.close()

