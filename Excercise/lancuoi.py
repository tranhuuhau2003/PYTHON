from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
import os
import smtplib
import sqlite3
from email import encoders
import pandas as pd
import tkinter as tk
from tkinter import *
from tkinter import filedialog, messagebox, Toplevel, ttk
from datetime import datetime
from tkinter import Frame, Tk
from PIL import Image, ImageTk
from tkinter import Label, Entry, Button,  Radiobutton, IntVar
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# Biến toàn cục lưu trữ dữ liệu sinh viên
global df_sinh_vien, ma_lop, ten_mon_hoc
df_sinh_vien, ma_lop, ten_mon_hoc = None, None, None

chart_frame = None

# Biến toàn cục để lưu tên file tóm tắt
summary_file = 'TongHopSinhVienVangCacLop.xlsx'

# Đăng ký adapter datetime
sqlite3.register_adapter(datetime, lambda d: d.timestamp())
sqlite3.register_converter("timestamp", lambda t: datetime.fromtimestamp(t))

def load_data():
    # Mở hộp thoại chọn file
    Tk().withdraw()  # Ẩn cửa sổ chính
    file_path = filedialog.askopenfilename(title="Chọn file Excel", filetypes=[("Excel files", "*.xlsx;*.xls")])

    if not file_path:
        print("Không có file nào được chọn.")
        return None, None, None, None

    try:
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Không tìm thấy file tại: {file_path}")

        # Đọc file excel
        df = pd.read_excel(file_path, header=None)
        df = df.fillna('')

        # Lấy thông tin "Đợt", "Mã lớp học phần", "Tên môn học"
        dot = df.iloc[5, 2]  # Đợt nằm ở hàng 5, cột 2
        ma_lop = df.iloc[7, 2]  # Mã lớp học phần nằm ở hàng 7, cột 2
        ten_mon_hoc = df.iloc[8, 2]  # Tên môn học nằm ở hàng 8, cột 2

        # Lấy dữ liệu sinh viên từ hàng 13 trở đi
        df_sinh_vien = df.iloc[13:, [1, 2, 3, 4, 5, 24, 25, 26, 27]]  # Chỉ lấy các cột cần thiết
        df_sinh_vien.columns = ['MSSV', 'Họ đệm', 'Tên', 'Giới tính', 'Ngày sinh', 'Vắng có phép', 'Vắng không phép', 'Tổng số tiết', '(%) vắng']

        # Chuyển đổi các cột phần trăm vắng từ ',', sang '.'
        if '(%) vắng' in df_sinh_vien.columns:
            df_sinh_vien['(%) vắng'] = df_sinh_vien['(%) vắng'].apply(lambda x: str(x).replace(',', '.') if isinstance(x, str) else x)
            
        # Xử lý các cột "Vắng có phép" và "Vắng không phép" để đảm bảo chúng là số
        df_sinh_vien['Vắng có phép'] = pd.to_numeric(df_sinh_vien['Vắng có phép'], errors='coerce').fillna(0)
        df_sinh_vien['Vắng không phép'] = pd.to_numeric(df_sinh_vien['Vắng không phép'], errors='coerce').fillna(0)

        # Thêm cột "Tổng buổi vắng" bằng cách cộng Vắng có phép và Vắng không phép
        df_sinh_vien['Tổng buổi vắng'] = df_sinh_vien['Vắng có phép'] + df_sinh_vien['Vắng không phép']
        
        # Lấy danh sách mssv để thêm email test vào db
        mssv_list = df_sinh_vien['MSSV'].tolist()

        return df_sinh_vien, dot, ma_lop, ten_mon_hoc, mssv_list
    except Exception as e:
        print(f"Lỗi khi đọc dữ liệu từ Excel: {e}")
        return None, None, None, None
    
def add_data_to_sqlite(df_sinh_vien, dot, ma_lop, ten_mon_hoc, mssv_list):
    try:
        conn = sqlite3.connect('students.db')
        cursor = conn.cursor()

        # Tạo bảng mới với các cột chính xác
        cursor.execute("""CREATE TABLE IF NOT EXISTS students (
                            mssv TEXT PRIMARY KEY,
                            ho_dem TEXT,
                            ten TEXT,
                            gioi_tinh TEXT,
                            ngay_sinh TEXT,
                            vang_co_phep INTEGER,
                            vang_khong_phep INTEGER,
                            tong_so_tiet INTEGER,
                            ty_le_vang REAL,
                            tong_buoi_vang INTEGER,
                            dot TEXT,
                            ma_lop TEXT,
                            ten_mon_hoc TEXT,
                            email_student TEXT
                        )""")
        
             # Tạo bảng parents
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS parents (
                mssv TEXT PRIMARY KEY,
                email_ph TEXT  -- Email của phụ huynh
            )
        ''')

        # Tạo bảng teachers
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS teachers (
                mssv TEXT PRIMARY KEY,
                email_gvcn TEXT  -- Email của giáo viên chủ nhiệm
            )
        ''')
        
        # Tạo bảng TBM
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS tbm (
                mssv TEXT PRIMARY KEY,
                email_tbm TEXT  -- Email của tổng bộ môn
            )
        ''')
        
        conn.commit()

        # Xóa dữ liệu cũ trước khi thêm dữ liệu mới
        cursor.execute("DELETE FROM students")
        conn.commit()
        
        # Thêm dữ liệu mới vào bảng 
        for _, row in df_sinh_vien.iterrows():
            try:
                values_to_insert = (
                    str(row['MSSV']),
                    str(row['Họ đệm']),
                    str(row['Tên']),
                    str(row['Giới tính']),
                    str(row['Ngày sinh']),
                    int(float(row['Vắng có phép'])),
                    int(float(row['Vắng không phép'])),
                    int(float(row['Tổng số tiết'])),
                    float(row['(%) vắng']),
                    int(row['Tổng buổi vắng']),  # Tổng buổi vắng
                    dot,
                    ma_lop,
                    ten_mon_hoc
                )

                cursor.execute("""INSERT INTO students (
                                    mssv, ho_dem, ten, gioi_tinh, ngay_sinh, 
                                    vang_co_phep, vang_khong_phep, tong_so_tiet, ty_le_vang, tong_buoi_vang,
                                    dot, ma_lop, ten_mon_hoc) 
                                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""", values_to_insert)

            except Exception as e:
                print(f"Lỗi khi thêm sinh viên {row['MSSV']}: {e}")


        
        # Thêm dữ liệu vào bảng students
        for mssv in mssv_list:
            email_student = f"tranhuuhauthh@gmail.com"  # Tạo email mẫu cho phụ huynh
            cursor.execute('UPDATE students SET email_student = ? WHERE mssv = ?', (email_student, mssv))
            
     
        # Thêm dữ liệu vào bảng parents
        for mssv in mssv_list:
            email_ph = f"tranhuuhauthh@gmail.com"  # Tạo email mẫu cho phụ huynh
            cursor.execute('INSERT OR IGNORE INTO parents (mssv, email_ph) VALUES (?, ?)', (mssv, email_ph))

        # Thêm dữ liệu vào bảng teachers
        for mssv in mssv_list:
            email_gvcn = f"tranhuuhauthh@gmail.com"  # Tạo email mẫu cho giáo viên chủ nhiệm
            cursor.execute('INSERT OR IGNORE INTO teachers (mssv, email_gvcn) VALUES (?, ?)', (mssv, email_gvcn))

        # Thêm dữ liệu vào bảng tbm
        for mssv in mssv_list:
            email_tbm = f"tranhuuhauthh@gmail.com"  # Tạo email mẫu cho giáo viên chủ nhiệm
            cursor.execute('INSERT OR IGNORE INTO tbm (mssv, email_tbm) VALUES (?, ?)', (mssv, email_tbm))

        

        conn.commit()
        conn.close()
    except Exception as e:
        print(f"Lỗi khi thêm dữ liệu vào SQLite: {e}")

def load_from_excel_to_treeview(tree):
    df_sinh_vien, dot, ma_lop, ten_mon_hoc, mssv_list = load_data()
    
    if df_sinh_vien is not None:
        add_data_to_sqlite(df_sinh_vien, dot, ma_lop, ten_mon_hoc, mssv_list)  # Thêm dữ liệu vào SQLite
        
        # Xóa dữ liệu hiện tại trong Treeview
        for row in tree.get_children():
            tree.delete(row)
        
        # Thêm cột Đợt, Mã lớp và Tên môn học vào DataFrame
        df_sinh_vien['Đợt'] = dot
        df_sinh_vien['Mã lớp'] = ma_lop
        df_sinh_vien['Tên môn học'] = ten_mon_hoc
        
        # Loại bỏ cột email nếu tồn tại
        if 'email_student' in df_sinh_vien.columns:
            df_sinh_vien = df_sinh_vien.drop(columns=['email_student'])

        # Hiển thị dữ liệu đã tải vào Treeview với cột STT
        for index, row in df_sinh_vien.iterrows():
            # Tính STT độc lập với df_sinh_vien
            stt = len(tree.get_children()) + 1  # Lấy số lượng hàng hiện tại trong Treeview và cộng thêm 1
            tree.insert('', 'end', values=[stt] + list(row))

def refresh_treeview(tree):
    # Xóa dữ liệu hiện tại trong treeview
    for item in tree.get_children():
        tree.delete(item)

    # Kết nối đến SQLite và lấy dữ liệu
    conn = sqlite3.connect('students.db')
    cursor = conn.cursor()
     # Chỉ lấy các cột cần thiết
    cursor.execute("""
        SELECT MSSV, ho_dem, ten, gioi_tinh, ngay_sinh, 
               vang_co_phep, vang_khong_phep, tong_so_tiet, 
               ty_le_vang, tong_buoi_vang, dot, ma_lop, ten_mon_hoc 
        FROM students
    """)
    rows = cursor.fetchall()
    
    for index, row in enumerate(rows):
        # Chèn dữ liệu vào TreeView với cột STT
        stt = index + 1  # Tính STT, bắt đầu từ 1
        tree.insert('', 'end', values=(stt,  # STT
            row[0],  # MSSV
            row[1],  # Họ đệm
            row[2],  # Tên
            row[3],  # Giới tính
            row[4],  # Ngày sinh
            row[5],  # Vắng có phép
            row[6],  # Vắng không phép
            row[7],  # Tổng số tiết
            row[8],  # (%) vắng
            row[9],  # Tổng buổi vắng
            row[10],  # Đợt
            row[11],  # Mã lớp
            row[12]   # Tên môn học
        ))
    
    conn.close()

def add_student(tree):
    # Tạo một cửa sổ mới để thêm sinh viên
    window = Toplevel()
    window.title("Thêm Sinh Viên")
    
    # Đặt kích thước cho cửa sổ
    window.geometry("250x450")  # Điều chỉnh kích thước cửa sổ

    labels = ["MSSV", "Họ đệm", "Tên", "Giới tính", "Ngày sinh", 
              "Vắng có phép", "Vắng không phép", "Tổng số tiết", 
              'Đợt', 'Mã lớp', 'Tên môn học']
    entries = []

    # Thay đổi kiểu chữ cho các nhãn và ô nhập
    font_style = ("Times New Roman", 9)

    # Sử dụng grid để đặt nhãn bên trái ô nhập
    for i, label in enumerate(labels):
        Label(window, text=label, font=font_style).grid(row=i, column=0, padx=10, pady=5, sticky='w')
        
        if label == "Giới tính":
            # Tạo biến để lưu giá trị của giới tính
            gender_var = IntVar()
            # Thêm nút radio cho giới tính
            Radiobutton(window, text="Nam", variable=gender_var, value=1, font=font_style).grid(row=i, column=1, padx=10, pady=5, sticky='w')
            Radiobutton(window, text="Nữ", variable=gender_var, value=2, font=font_style).grid(row=i, column=1, columnspan=2, padx=10, pady=5, sticky='e')
            entries.append(gender_var)
        else:
            entry = Entry(window, font=font_style)
            entry.grid(row=i, column=1, padx=10, pady=5)
            entries.append(entry)

    def save_student():
        # Lưu sinh viên vào SQLite
        conn = sqlite3.connect('students.db')
        cursor = conn.cursor()

        # Lấy các giá trị từ các trường nhập liệu
        try:
            mssv = entries[0].get()  # MSSV
            ho_dem = entries[1].get()  # Họ đệm
            ten = entries[2].get()  # Tên
            gioi_tinh = "Nam" if entries[3].get() == 1 else "Nữ"  # Lấy giá trị giới tính từ radio button
            ngay_sinh = entries[4].get()  # Ngày sinh
            vang_co_phep = int(entries[5].get())  # Vắng có phép
            vang_khong_phep = int(entries[6].get())  # Vắng không phép
            tong_so_tiet = int(entries[7].get())  # Tổng số tiết
            tong_buoi_vang = vang_co_phep + vang_khong_phep  # Tính tổng buổi vắng
            
            # Tính % vắng
            ty_le_vang = round((tong_buoi_vang / tong_so_tiet) * 100, 1) if tong_so_tiet > 0 else 0

            dot = entries[8].get()  # Đợt
            ma_lop = entries[9].get()  # Mã lớp
            ten_mon_hoc = entries[10].get()  # Tên môn học

            # Chuẩn bị giá trị để chèn vào cơ sở dữ liệu
            values_to_insert = (
                mssv, ho_dem, ten, gioi_tinh, ngay_sinh, 
                vang_co_phep, vang_khong_phep, tong_so_tiet, 
                ty_le_vang, tong_buoi_vang, dot, ma_lop, ten_mon_hoc
            )
            
            cursor.execute("""INSERT INTO students 
                              (mssv, ho_dem, ten, gioi_tinh, ngay_sinh, 
                               vang_co_phep, vang_khong_phep, tong_so_tiet, 
                               ty_le_vang, tong_buoi_vang, dot, ma_lop, ten_mon_hoc) 
                              VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""", 
                           values_to_insert)
            conn.commit()
            messagebox.showinfo("Thành công", "Đã thêm sinh viên thành công.")
            refresh_treeview(tree)  # Cập nhật Treeview
            window.destroy()
        except sqlite3.IntegrityError:
            messagebox.showerror("Lỗi", "MSSV đã tồn tại.")
        except ValueError:
            messagebox.showerror("Lỗi", "Vui lòng nhập số hợp lệ cho các trường vắng.")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Có lỗi xảy ra: {e}")
        finally:
            conn.close()

    # Nút lưu với kiểu chữ
    Button(window, text="Lưu", command=save_student, font=font_style).grid(row=len(labels), column=0, columnspan=2, pady=20)

def edit_student(tree):
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("Chọn sinh viên", "Vui lòng chọn sinh viên để chỉnh sửa.")
        return

    # Tạo một cửa sổ mới để chỉnh sửa sinh viên
    window = Toplevel()
    window.title("Chỉnh Sửa Sinh Viên")
    
    # Đặt kích thước cho cửa sổ
    window.geometry("400x400")  # Tăng kích thước để vừa với các ô nhập

    # Lấy dữ liệu của sinh viên đã chọn
    item_values = tree.item(selected_item, "values")
    
    # Danh sách các nhãn cho trường cần hiển thị
    labels = ["MSSV", "Họ đệm", "Tên", "Giới tính", "Ngày sinh", 
              "Vắng có phép", "Vắng không phép", "Tổng số tiết", 
              "Đợt", "Mã lớp", "Tên môn học"]
    
    # Chỉ lấy các giá trị cần thiết từ item_values
    values_to_display = item_values[1:9] + item_values[11:14]  
    
    entries = []
    gender_var = IntVar()  # Biến để lưu giá trị giới tính

    # Thay đổi kiểu chữ cho các nhãn và ô nhập
    font_style = ("Times New Roman", 9) 

    # Hiển thị thông tin hiện có của sinh viên vào các trường thông tin để chỉnh sửa
    for i, (label, value) in enumerate(zip(labels, values_to_display)):  
        Label(window, text=label, font=font_style).grid(row=i, column=0, padx=10, pady=5, sticky='w')  # Căn trái
        
        if label == "Giới tính":
            # Đặt giá trị cho radio button dựa vào dữ liệu
            gioi_tinh = item_values[4].strip()
            if gioi_tinh == "Nam":
                gender_var.set(1)  # Đặt giá trị 1 cho Nam
            elif gioi_tinh == "Nữ":
                gender_var.set(2)  # Đặt giá trị 2 cho Nữ
            
            # Thêm nút radio cho giới tính
            Radiobutton(window, text="Nam", variable=gender_var, value=1, font=font_style).grid(row=i, column=1, padx=10, pady=5, sticky='w')
            Radiobutton(window, text="Nữ", variable=gender_var, value=2, font=font_style).grid(row=i, column=1, columnspan=2, padx=10, pady=5, sticky='e')
        elif label in ["MSSV", "Đợt", "Mã lớp", "Tên môn học"]:
            label_value = Label(window, text=value, font=font_style)  # Hiển thị dưới dạng Label
            label_value.grid(row=i, column=1, padx=10, pady=5)  # Đặt bên cạnh nhãn
        else:
            entry = Entry(window, font=font_style)  # Đặt kiểu chữ cho ô nhập
            entry.insert(0, value)  # Điền giá trị hiện tại vào ô nhập
            entry.grid(row=i, column=1, padx=10, pady=5)  # Đặt ô nhập bên cạnh nhãn
            entries.append(entry)

    def update_student():
        conn = sqlite3.connect('students.db')
        cursor = conn.cursor()

        # Lấy giá trị từ các trường đã nhập
        mssv_cu = item_values[1]  # MSSV cũ từ item_values
        ho_dem = entries[0].get()
        gioi_tinh = "Nam" if gender_var.get() == 1 else "Nữ"  # Lấy giá trị giới tính từ radio button
        ten = entries[1].get()
        ngay_sinh = entries[2].get()
        vang_co_phep = int(entries[3].get())
        vang_khong_phep = int(entries[4].get())
        tong_so_tiet = int(entries[5].get())
        dot = item_values[11]  # Đợt cũ
        ma_lop = item_values[12]  # Mã lớp cũ
        ten_mon_hoc = item_values[13]  # Tên môn cũ

        # Tính tổng buổi vắng
        tong_buoi_vang = vang_co_phep + vang_khong_phep
        # Tính % vắng
        if tong_so_tiet > 0:
            ty_le_vang = round((tong_buoi_vang / tong_so_tiet) * 100, 1)  # Làm tròn tới 1 chữ số thập phân
        else:
            ty_le_vang = 0

        # Cập nhật thông tin sinh viên
        try:
            cursor.execute("""UPDATE students SET 
                              ho_dem = ?, ten = ?, gioi_tinh = ?, ngay_sinh = ?, 
                              vang_co_phep = ?, vang_khong_phep = ?, tong_so_tiet = ?, 
                              ty_le_vang = ?, tong_buoi_vang = ?, dot = ?, ma_lop = ?, ten_mon_hoc = ?
                              WHERE mssv = ?""", 
                           (ho_dem, ten, gioi_tinh, ngay_sinh, vang_co_phep, vang_khong_phep, 
                            tong_so_tiet, ty_le_vang, tong_buoi_vang, dot, ma_lop, ten_mon_hoc, mssv_cu))
            conn.commit()
            messagebox.showinfo("Thành công", "Đã cập nhật sinh viên thành công.")
            refresh_treeview(tree)  # Cập nhật Treeview
            window.destroy()
        except Exception as e:
            messagebox.showerror("Lỗi", f"Có lỗi xảy ra: {e}")
        finally:
            conn.close()

    Button(window, text="Cập nhật", command=update_student, font=font_style).grid(row=len(values_to_display), column=0, columnspan=2, pady=20)

def delete_student(tree):
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("Chọn sinh viên", "Vui lòng chọn sinh viên để xóa.")
        return

    item_values = tree.item(selected_item, "values")
    mssv = item_values[1]  # Lấy MSSV của sinh viên được chọn

    # Hiển thị hộp thoại xác nhận
    confirm = messagebox.askyesno("Xác nhận xóa", f"Bạn có chắc chắn muốn xóa sinh viên có MSSV: {mssv}?")
    if not confirm:
        return  # Nếu người dùng không xác nhận, thoát khỏi hàm

    conn = sqlite3.connect('students.db')
    cursor = conn.cursor()
    cursor.execute("DELETE FROM students WHERE mssv=?", (mssv,))
    conn.commit()
    conn.close()

    refresh_treeview(tree)  # Cập nhật Treeview
    messagebox.showinfo("Thành công", "Đã xóa sinh viên thành công.")

def sort_students_by_absences(tree):
    # Xóa dữ liệu hiện tại trong treeview
    for item in tree.get_children():
        tree.delete(item)

    # Kết nối đến SQLite và lấy dữ liệu
    conn = sqlite3.connect('students.db')
    cursor = conn.cursor()

    # Sắp xếp theo tổng buổi vắng giảm dần
     # Chỉ lấy các cột cần thiết
    cursor.execute("""
        SELECT MSSV, ho_dem, ten, gioi_tinh, ngay_sinh, 
               vang_co_phep, vang_khong_phep, tong_so_tiet, 
               ty_le_vang, tong_buoi_vang, dot, ma_lop, ten_mon_hoc 
        FROM students
        ORDER BY tong_buoi_vang DESC
    """)
    rows = cursor.fetchall()

    # Biến đếm để đánh số thứ tự
    stt = 1

    for row in rows:
        # Chèn dữ liệu vào TreeView với STT
        item_id = tree.insert('', 'end', values=(
            stt,  # STT - cột đầu tiên
            row[0],  # MSSV
            row[1],  # Họ đệm
            row[2],  # Tên
            row[3],  # Giới tính
            row[4],  # Ngày sinh
            row[5],  # Vắng có phép
            row[6],  # Vắng không phép
            row[7],  # Tổng số tiết
            row[8],  # (%) vắng
            row[9],  # Tổng buổi vắng
            row[10], # Đợt
            row[11], # Mã lớp
            row[12]  # Tên môn học
        ))

        # Kiểm tra tỷ lệ vắng và bôi đỏ nếu >= 50.0
        if row[8] >= 50.0:  # Giả sử cột 8 là tỷ lệ vắng
            tree.item(item_id, tags=('highlight',))

        # Tăng số thứ tự cho lần lặp tiếp theo
        stt += 1
    # Định nghĩa kiểu bôi đỏ cho tag
    tree.tag_configure('highlight', foreground='red')

    conn.close()
    
# Hàm tìm kiếm sinh viên theo nhiều tiêu chí
def search_students(tree, search_by, search_value):
    # Xóa dữ liệu hiện tại trong treeview
    for item in tree.get_children():
        tree.delete(item)

    # Kết nối tới CSDL SQLite
    conn = sqlite3.connect('students.db')
    cursor = conn.cursor()

    # Câu truy vấn dựa trên tiêu chí tìm kiếm
    query = "SELECT mssv, ho_dem, ten, gioi_tinh, ngay_sinh, vang_co_phep, vang_khong_phep, tong_so_tiet, ty_le_vang, tong_buoi_vang, dot, ma_lop, ten_mon_hoc FROM students WHERE "

    if search_by == "MSSV":
        query += "mssv LIKE ?"
        search_value = '%' + search_value.strip() + '%'
    elif search_by == "Tên":
        query += "(ho_dem || ' ' || ten) LIKE ?"
        search_value = '%' + search_value.strip() + '%'
    elif search_by == "Lớp":
        query += "ma_lop LIKE ?"
        search_value = '%' + search_value.strip() + '%'
    elif search_by == "Tỷ lệ vắng":
        try:
            search_value = float(search_value)  # Chuyển thành float để so sánh số liệu
            query += "ty_le_vang >= ?"
        except ValueError:
            print("Giá trị tìm kiếm tỷ lệ vắng phải là số.")
            conn.close()
            return
    else:
        conn.close()
        return

    # Thực thi câu truy vấn
    cursor.execute(query, (search_value,))
    rows = cursor.fetchall()

    # Biến đếm để đánh số thứ tự
    stt = 1
    for row in rows:
        # Chèn dữ liệu vào TreeView với STT
        tree.insert('', 'end', values=(
            stt,  # STT
            row[0],  # MSSV
            row[1],  # Họ đệm
            row[2],  # Tên
            row[3],  # Giới tính
            row[4],  # Ngày sinh
            row[5],  # Vắng có phép
            row[6],  # Vắng không phép
            row[7],  # Tổng số tiết
            row[8],  # (%) vắng
            row[9],  # Tổng buổi vắng
            row[10], # Đợt
            row[11], # Mã lớp
            row[12]  # Tên môn học
        ))

        stt += 1

    conn.close()
# Thêm giao diện tìm kiếm vào hệ thống chính

def add_search_interface(center_frame, tree):
    search_frame = Frame(center_frame, bg="#B6CFB6", bd=1)  # Giảm chiều cao bằng cách giảm bd
    search_frame.pack(side='top', fill='x', padx=5, pady=3)

    # Dropdown chọn tiêu chí tìm kiếm
    # search_by_label = Label(search_frame, text="Tìm kiếm theo:", bg="#B6CFB6", font=("Times New Roman", 14))  
    # search_by_label.pack(side='left', padx=2)

    search_by_var = StringVar(value="MSSV")
    search_by_menu = OptionMenu(search_frame, search_by_var, "MSSV", "Tên", "Mã Lớp", "Tỷ lệ vắng")
    search_by_menu.config(bg="#A9EDE9")
    search_by_menu.pack(side='left', padx=2)

    # Entry tìm kiếm
    search_entry = Entry(search_frame, font=("Times New Roman", 12), bd=3, width=17)
    search_entry.pack(side='left', padx=2)

    # Nút tìm kiếm
    search_button = Button(search_frame, text="Tìm", command=lambda: search_students(tree, search_by_var.get(), search_entry.get()), bg="#A9EDE9", font=("Times New Roman", 10))
    search_button.pack(side='left', padx=2)

# Hàm khởi tạo cơ sở dữ liệu tonghopsv
def initialize_database():
    conn = sqlite3.connect('tonghopsv.db', detect_types=sqlite3.PARSE_DECLTYPES)
    cursor = conn.cursor()

    # Tạo bảng nếu chưa tồn tại
    cursor.execute("""CREATE TABLE IF NOT EXISTS tonghopsv (
                        mssv TEXT PRIMARY KEY,
                        ho_dem TEXT,
                        ten TEXT,
                        gioi_tinh TEXT,
                        ngay_sinh TIMESTAMP,
                        vang_co_phep INTEGER,
                        vang_khong_phep INTEGER,
                        tong_so_tiet INTEGER,
                        ty_le_vang REAL,
                        tong_buoi_vang INTEGER,
                        dot TEXT,
                        ma_lop TEXT,
                        ten_mon_hoc TEXT
                    )""")
    
    # Xóa dữ liệu trong bảng khi khởi động
    cursor.execute("DELETE FROM tonghopsv")
    conn.commit()
    conn.close()
    
# Hàm lưu sinh viên vào SQLite dànhcho tonghopsv
def save_students_to_sqlite(df):
    print("Đang lưu sinh viên vào SQLite...")
    conn = sqlite3.connect('tonghopsv.db', detect_types=sqlite3.PARSE_DECLTYPES)
    cursor = conn.cursor()

    for _, row in df.iterrows():
        try:
            # Kiểm tra nếu 'ngay_sinh' là datetime, chuyển đổi thành chuỗi
            ngay_sinh_value = row['Ngày sinh']
            if isinstance(ngay_sinh_value, pd.Timestamp):  # Nếu là kiểu pandas Timestamp
                ngay_sinh_value = ngay_sinh_value.strftime('%Y-%m-%d')  # Chuyển đổi thành chuỗi

            cursor.execute("""INSERT OR IGNORE INTO tonghopsv (
                                mssv, ho_dem, ten, gioi_tinh, ngay_sinh, 
                                vang_co_phep, vang_khong_phep, tong_so_tiet, 
                                ty_le_vang, tong_buoi_vang, dot, ma_lop, ten_mon_hoc
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                           (row['MSSV'], row['Họ đệm'], row['Tên'], row['Giới tính'],
                            ngay_sinh_value,  # Dùng giá trị đã chuyển đổi
                            row['Vắng có phép'], row['Vắng không phép'],
                            row['Tổng số tiết'], row['(%) vắng'], row['Tổng buổi vắng'],
                            row['Đợt'], row['Mã lớp'], row['Tên môn học']))
        except Exception as e:
            print(f"Lỗi khi thêm sinh viên {row['MSSV']}: {e}")

    conn.commit()
    conn.close()

def send_email_with_ssl(summary_file):
    sender_email = "carotneee4@gmail.com"
    app_password = "bgjx tavb oxba ickr"  # Thay bằng mật khẩu ứng dụng của bạn
    receiver_email = "vokhanhlinh04112k3@gmail.com"
    subject = "Tổng hợp sinh viên vắng nhiều"
    body = "Đính kèm là danh sách sinh viên vắng >= 50%."

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(sender_email, app_password)
            server.sendmail(sender_email, receiver_email, msg.as_string())
        print("Email đã được gửi thành công.")
    except Exception as e:
        print(f"Lỗi khi gửi email: {e}")
        
def send_email(to_address, subject, message):
    """Send email to the recipient."""
    from_address = "carotneee4@gmail.com"
    password = "bgjx tavb oxba ickr"  # Make sure to use an app-specific password for Gmail if 2FA is enabled.
    
    # Initialize email
    msg = MIMEMultipart()
    msg['From'] = from_address
    msg['To'] = to_address
    msg['Subject'] = subject

    # Email content
    msg.attach(MIMEText(message, 'plain'))

    # Configure SMTP server to send email
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(from_address, password)
        text = msg.as_string()
        server.sendmail(from_address, to_address, text)
        server.quit()
        print(f"Email sent to {to_address}")
    except Exception as e:
        print(f"Failed to send email to {to_address}: {e}")

def send_warning_emails():
    """Check and send warning emails for students."""
    # Connect to SQLite database
    connection = sqlite3.connect('students.db')
    cursor = connection.cursor()

    try:
        query = """
        SELECT mssv, ho_dem, ten, vang_co_phep, vang_khong_phep, tong_so_tiet, ty_le_vang
        FROM students
        """
        cursor.execute(query)
        records = cursor.fetchall()

        for row in records:
            mssv, ho_dem, ten, vang_co_phep, vang_khong_phep, tong_so_tiet, ty_le_vang = row
            
            # Get student and related emails
            student_email = get_student_email(cursor, mssv)
            parent_email = get_parent_email(cursor, mssv)
            homeroom_teacher_email = get_teacher_email(cursor, mssv)
            tbm_email = get_tbm_email(cursor, mssv)

            # Check and send warnings based on absence rate
            if ty_le_vang >= 50:
                subject = "Cảnh báo học vụ: Vắng học quá 50%"
                message = f"Sinh viên {ho_dem} {ten} đã vắng hơn 50% số buổi học."
                send_email(student_email, subject, message)
                send_email(parent_email, subject, message)
                send_email(homeroom_teacher_email, subject, message)
                send_email(tbm_email, subject, message)
            elif ty_le_vang >= 20:
                subject = "Cảnh báo học vụ: Vắng học quá 20%"
                message = f"Sinh viên {ho_dem} {ten} đã vắng hơn 20% số buổi học."
                send_email(student_email, subject, message)

    except Exception as e:
        print(f"Lỗi khi gửi email cảnh báo: {e}")
    finally:
        connection.close()  # Ensure the database connection is closed

def get_student_email(cursor, mssv):
    """Retrieve student email from the database based on MSSV."""
    query = f"SELECT email_student FROM students WHERE mssv = ?"
    cursor.execute(query, (mssv,))
    result = cursor.fetchone()
    return result[0] if result else None

def get_parent_email(cursor, mssv):
    """Retrieve parent email from the database based on MSSV."""
    query = f"SELECT email_ph FROM parents WHERE mssv = ?"
    cursor.execute(query, (mssv,))
    result = cursor.fetchone()
    return result[0] if result else None

def get_teacher_email(cursor, mssv):
    """Retrieve homeroom teacher email from the database based on MSSV."""
    query = f"SELECT email_gvcn FROM teachers WHERE mssv = ?"
    cursor.execute(query, (mssv,))
    result = cursor.fetchone()
    return result[0] if result else None

def get_tbm_email(cursor, mssv):
    """Retrieve TBM email from the database based on MSSV."""
    query = f"SELECT email_tbm FROM tbm WHERE mssv = ?"
    cursor.execute(query, (mssv,))
    result = cursor.fetchone()
    return result[0] if result else None

def send_email_with_attachment(summary_file):
    # Cấu hình thông tin email
    sender_email = "carotneee4@gmail.com" # Địa chỉ email của bạn
    sender_password ="bgjx tavb oxba ickr" # Mật khẩu email của bạn
    recipient_email = "tranhuuhauthh@gmail.com"  # Địa chỉ email của người nhận

    # Tạo một đối tượng MIMEMultipart
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = "Báo cáo sinh viên vắng nhiều"

    # Tạo phần thân email
    body = "Xin chào,\n\nĐây là báo cáo tổng hợp sinh viên vắng nhiều.\n\nTrân trọng!"
    msg.attach(MIMEText(body, 'plain'))

    # Đính kèm tệp Excel
    attachment = open(summary_file, "rb")
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename= {summary_file}')
    msg.attach(part)

    # Gửi email
    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as server:  # Thay 'smtp.example.com' với máy chủ SMTP của bạn
            server.starttls()  # Bật chế độ bảo mật
            server.login(sender_email, sender_password)
            server.send_message(msg)
            print("Email đã được gửi thành công!")
    except Exception as e:
        print(f"Có lỗi xảy ra khi gửi email: {e}")
    finally:
        attachment.close()
        
# Hàm tải và tổng hợp sinh viên
def load_and_summarize_students(tree):
    global summary_file
    # Xóa tất cả mục trong giao diện (tree) ngay từ đầu
    for item in tree.get_children():
        tree.delete(item)

    excel_files = filedialog.askopenfilenames(title='Chọn các file Excel', filetypes=[("Excel files", "*.xlsx;*.xls")])
    
    if not excel_files:
        return

    all_data = []

    for file in excel_files:
        df = pd.read_excel(file, header=None)
        df = df.fillna('')

        dot = df.iloc[5, 2]
        ma_lop = df.iloc[7, 2]
        ten_mon_hoc = df.iloc[8, 2]

        df_sinh_vien = df.iloc[13:, [1, 2, 3, 4, 5, 24, 25, 26, 27]]
        df_sinh_vien.columns = ['MSSV', 'Họ đệm', 'Tên', 'Giới tính', 'Ngày sinh', 'Vắng có phép', 'Vắng không phép', 'Tổng số tiết', '(%) vắng']

        df_sinh_vien['(%) vắng'] = df_sinh_vien['(%) vắng'].apply(lambda x: str(x).replace(',', '.') if isinstance(x, str) else x)
        df_sinh_vien['Vắng có phép'] = pd.to_numeric(df_sinh_vien['Vắng có phép'], errors='coerce').fillna(0)
        df_sinh_vien['Vắng không phép'] = pd.to_numeric(df_sinh_vien['Vắng không phép'], errors='coerce').fillna(0)
        df_sinh_vien['Tổng buổi vắng'] = df_sinh_vien['Vắng có phép'] + df_sinh_vien['Vắng không phép']

        df_sinh_vien['Đợt'] = dot
        df_sinh_vien['Mã lớp'] = ma_lop
        df_sinh_vien['Tên môn học'] = ten_mon_hoc

        # Xử lý cột Ngày sinh
        df_sinh_vien['Ngày sinh'] = pd.to_datetime(df_sinh_vien['Ngày sinh'], errors='coerce')

        all_data.append(df_sinh_vien)

        save_students_to_sqlite(df_sinh_vien)  # Gọi hàm lưu sinh viên vào SQLite

    combined_data = pd.concat(all_data, ignore_index=True)
    combined_data['(%) vắng'] = pd.to_numeric(combined_data['(%) vắng'], errors='coerce').fillna(0)

    combined_data.rename(columns={
        'Họ đệm': 'ho_dem',
        'Tên': 'ten',
        'Giới tính': 'gioi_tinh',
        'Ngày sinh': 'ngay_sinh',
        'Vắng có phép': 'vang_co_phep',
        'Vắng không phép': 'vang_khong_phep',
        'Tổng số tiết': 'tong_so_tiet',
        '(%) vắng': 'ty_le_vang',
        'Tổng buổi vắng': 'tong_buoi_vang'
    }, inplace=True)

    # Xóa tất cả mục trong giao diện (tree) sau khi tổng hợp
    for item in tree.get_children():
        tree.delete(item)

    conn = sqlite3.connect('tonghopsv.db')
    cursor = conn.cursor()

    cursor.execute("SELECT * FROM tonghopsv")
    rows = cursor.fetchall()
        
    # Hiển thị dữ liệu đã tải vào Treeview với cột STT
    for row in rows:
        # Tính STT độc lập với df_sinh_vien
        stt = len(tree.get_children()) + 1  # Lấy số lượng hàng hiện tại trong Treeview và cộng thêm 1
        tree.insert('', 'end', values=[stt] + list(row))

    conn.close()

    # Lưu sinh viên vắng >= 50% vào file Excel mới
    absent_students = combined_data[combined_data['ty_le_vang'] >= 50.0]
    summary_file = 'TongHopSinhVienVangCacLop.xlsx'
    absent_students.to_excel(summary_file, index=False)
    print(f"Tệp tóm tắt đã được lưu tại: {summary_file}")
    
# dùng để vẽ biểu đồ tỷ lệ vắng
def get_data_from_sqlite():
    conn = sqlite3.connect('students.db')
    cursor = conn.cursor()
    
    # Lấy dữ liệu MSSV, Họ tên và tỷ lệ vắng của sinh viên
    cursor.execute("SELECT mssv, ho_dem || ' ' || ten as ho_ten, ty_le_vang FROM students")
    data = cursor.fetchall()
    
    # Lấy dữ liệu lớp và số buổi vắng
    cursor.execute("SELECT ma_lop, SUM(vang_co_phep + vang_khong_phep) AS tong_buoi_vang FROM students GROUP BY ma_lop")
    class_data = cursor.fetchall()

    conn.close()
    
    return data, class_data

def plot_student_absence_chart(student_data):
    fig, ax = plt.subplots(figsize=(8, 5))
    names = [row[1] for row in student_data]
    absence_rates = [row[2] for row in student_data]

    # Vẽ biểu đồ cột ngang
    ax.barh(names, absence_rates, color='orange')

    # Thiết lập nhãn cho trục x và tiêu đề
    ax.set_xlabel('Tỷ lệ vắng (%)', fontsize=10)  # Kích thước chữ nhãn trục x
    ax.set_title('Tỷ lệ vắng của sinh viên', fontsize=12)  # Kích thước chữ tiêu đề

    # Thiết lập kích thước chữ cho nhãn cột
    ax.tick_params(axis='y', labelsize=6)  # Kích thước chữ cho nhãn trên trục y

    return fig

def show_student_chart():
    global chart_frame  # Sử dụng biến toàn cục

    student_data, _ = get_data_from_sqlite()
    fig = plot_student_absence_chart(student_data)

    # Xóa các widget hiện tại trong chart_frame
    for widget in chart_frame.winfo_children():
        widget.destroy()

    canvas = FigureCanvasTkAgg(fig, master=chart_frame)
    canvas.draw()
    canvas.get_tk_widget().pack(fill='both', expand=True)

#dùng đê vẽ biểu đồ vắng có phếp / không phép
def get_absence_types_data():
    conn = sqlite3.connect('students.db')
    cursor = conn.cursor()
    
    # Lấy tổng số vắng có phép và vắng không phép
    cursor.execute("""
        SELECT SUM(vang_co_phep) AS tong_vang_co_phep,
               SUM(vang_khong_phep) AS tong_vang_khong_phep
        FROM students
    """)
    absence_data = cursor.fetchone()
    conn.close()
    
    return absence_data

def plot_absence_types_chart(absence_data):
    fig, ax = plt.subplots(figsize=(8, 5))

    labels = ['Vắng có phép', 'Vắng không phép']
    values = [absence_data[0], absence_data[1]]

    # Vẽ biểu đồ cột
    ax.bar(labels, values, color=['green', 'red'])

    # Thiết lập nhãn cho trục y và tiêu đề
    ax.set_ylabel('Số lượng', fontsize=10)  # Kích thước chữ nhãn trục y
    ax.set_title('Số lượng vắng có phép và vắng không phép', fontsize=12)  # Kích thước chữ tiêu đề

    return fig

def show_absence_types_chart():
    global chart_frame  # Sử dụng biến toàn cục

    absence_data = get_absence_types_data()
    fig = plot_absence_types_chart(absence_data)

    # Xóa các widget hiện tại trong chart_frame
    for widget in chart_frame.winfo_children():
        widget.destroy()

    canvas = FigureCanvasTkAgg(fig, master=chart_frame)
    canvas.draw()
    canvas.get_tk_widget().pack(fill='both', expand=True)
    
def main():
    global df_sinh_vien, ma_lop, ten_mon_hoc, summary_file
    global chart_frame 
    root = Tk()
    root.title("Quản Lý Sinh Viên")
    
    # Thay đổi màu nền cho cửa sổ chính
    root.configure(bg="#B6CFB6")

    # Thêm logo vào tiêu đề của ứng dụng
    logo_icon = Image.open("Excercise\logoSGu.png")
    logo_icon = logo_icon.resize((32, 32), Image.LANCZOS)
    logo_icon_photo = ImageTk.PhotoImage(logo_icon)
    root.iconphoto(False, logo_icon_photo)

    # Đặt kích thước và vị trí cho giao diện chính
    root.geometry("1500x750+10+20")
    
    # Tạo style cho các nút
    style = ttk.Style()
    style.configure("TButton", font=("Times New Roman", 10), padding=6)

    # Thêm logo vào giao diện
    logo_image = Image.open("Excercise\logocnttsgu.png")
    logo_image = logo_image.resize((240, 50), Image.LANCZOS)
    logo_photo = ImageTk.PhotoImage(logo_image)
    logo_label = Label(root, image=logo_photo, bg="#B6CFB6")
    logo_label.image = logo_photo
    logo_label.pack(side=TOP, pady=10)

    # Tạo Treeview để hiển thị dữ liệu sinh viên
    tree = ttk.Treeview(root, columns=('STT', 'MSSV', 'Họ đệm', 'Tên', 'Giới tính', 'Ngày sinh', 'Vắng có phép', 'Vắng không phép', 'Tổng số tiết', '(%) vắng', 'Tổng buổi vắng', 'Đợt', 'Mã lớp', 'Tên môn học'), show='headings')
    
    style.configure("Treeview", font=("Times New Roman", 10), rowheight=25)
    style.configure("Treeview.Heading", font=("Times New Roman", 11, "bold"), background="#4CAF50", foreground="black")
    style.map("Treeview", background=[("selected", "#A3C1DA")], foreground=[("selected", "black")])

    # Tùy chỉnh chiều rộng cột
    tree.column("STT", width=40, anchor="center")  
    tree.column("MSSV", width=100, anchor="center")
    tree.column("Họ đệm", width=150, anchor="center")
    tree.column("Tên", width=80, anchor="center")
    tree.column("Giới tính", width=80, anchor="center")
    tree.column("Ngày sinh", width=120, anchor="center")
    tree.column("Vắng có phép", width=120, anchor="center")
    tree.column("Vắng không phép", width=120, anchor="center")
    tree.column("Tổng số tiết", width=120, anchor="center")
    tree.column("(%) vắng", width=80, anchor="center")
    tree.column("Tổng buổi vắng", width=120, anchor="center")
    tree.column("Đợt", width=100, anchor="center")
    tree.column("Mã lớp", width=100, anchor="center")
    tree.column("Tên môn học", width=150, anchor="center")

    for col in tree['columns']:
        tree.heading(col, text=col)

    tree.pack(fill='both', expand=True, padx=10, pady=10)

    # Tạo frame chứa các nút bên trái
    left_frame = Frame(root, bg="#B6CFB6")
    left_frame.pack(side=LEFT, padx=10, pady=10, fill='y')

    # Tạo frame chứa các nút ở giữa
    center_frame = Frame(root, bg="#B6CFB6")
    center_frame.pack(side=LEFT, padx=10, pady=10, fill='y')

    # Thêm giao diện tìm kiếm vào center_frame trước các nút khác
    add_search_interface(center_frame, tree)  # Thêm interface tìm kiếm lên trên cùng

    # Các nút nằm bên trái, với độ rộng cố định
    button_width = 35
    button_color = "#A9EDE9"

    load_button = Button(left_frame, text="Tải dữ liệu từ Excel", command=lambda: load_from_excel_to_treeview(tree), width=button_width, bg=button_color, fg='black', font=("Times New Roman", 10))
    load_button.pack(anchor='w', pady=5, fill='x')

    add_button = Button(left_frame, text="Thêm sinh viên", command=lambda: add_student(tree), width=button_width, bg=button_color, fg='black', font=("Times New Roman", 10))
    add_button.pack(anchor='w', pady=5, fill='x')

    edit_button = Button(left_frame, text="Chỉnh sửa sinh viên", command=lambda: edit_student(tree), width=button_width, bg=button_color, fg='black', font=("Times New Roman", 10))
    edit_button.pack(anchor='w', pady=5, fill='x')

    delete_button = Button(left_frame, text="Xóa sinh viên", command=lambda: delete_student(tree), width=button_width, bg=button_color, fg='black', font=("Times New Roman", 10))
    delete_button.pack(anchor='w', pady=5, fill='x')

    sort_button = Button(left_frame, text="Sắp xếp sinh viên theo tổng buổi vắng", command=lambda: sort_students_by_absences(tree), width=button_width, bg=button_color, fg='black', font=("Times New Roman", 10))
    sort_button.pack(anchor='w', pady=5, fill='x')

    send_warning_email_button = Button(left_frame, text="Gửi Email cảnh báo học vụ", command=send_warning_emails, width=button_width, bg=button_color, fg='black', font=("Times New Roman", 10))
    send_warning_email_button.pack(anchor='w', pady=5, fill='x')

    refresh_button = Button(left_frame, text="Refresh", command=lambda: refresh_treeview(tree), width=button_width, bg=button_color, fg='black', font=("Times New Roman", 10))
    refresh_button.pack(anchor='w', pady=5, fill='x')

    # Các nút nằm ở giữa
    student_chart_button = Button(center_frame, text="Hiển thị biểu đồ sinh viên", command=show_student_chart, width=button_width, bg=button_color, fg='black', font=("Times New Roman", 10))
    student_chart_button.pack(anchor='center', pady=10)

    absence_types_chart_button = Button(center_frame, text="Hiển thị biểu đồ vắng", command=show_absence_types_chart, width=button_width, bg=button_color, fg='black', font=("Times New Roman", 10))
    absence_types_chart_button.pack(anchor='center', pady=10)

    summarize_button = Button(center_frame, text="Tổng hợp thông tin sinh viên", command=lambda: load_and_summarize_students(tree), width=button_width, bg=button_color, fg='black', font=("Times New Roman", 10))
    summarize_button.pack(anchor='center', pady=10)

    send_summary_email_button = Button(center_frame, text="Gửi Email tổng hợp thông tin", command=lambda: send_email_with_attachment(summary_file) if summary_file else print("Không có tệp tóm tắt để gửi!"), width=button_width, bg=button_color, fg='black', font=("Times New Roman", 10))
    send_summary_email_button.pack(anchor='center', pady=10)
    
    initialize_database()
    
    # Khởi tạo chart_frame
    chart_frame = Frame(root, bg="#B6CFB6")
    chart_frame.pack(fill='both', expand=True)
    
    root.mainloop()

if __name__ == "__main__":
    main()




