import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import sqlite3
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import time
import schedule
import threading
from datetime import datetime

class ExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Data Viewer")
        self.filepath = ''  # Để trống ban đầu, sẽ được cập nhật khi chọn file
        
        # Kết nối SQLite và tạo bảng nếu chưa có
        self.conn = sqlite3.connect('students.db')
        self.cursor = self.conn.cursor()
        
        print("Đang thiết lập lịch trình gửi email...")
       # Khởi động luồng cho lịch trình gửi email
        schedule.every().day.at("17:07").do(self.check_send_email)
        schedule.every().day.at("17:07").do(self.check_send_email)
        self.start_scheduler()

        # Giao diện chính
        self.main_menu()
        
    def load_data_from_excel(self):
        try:
            # Đọc file Excel một lần
            df = pd.read_excel(self.filepath, engine='xlrd', header=None)
            df = df.fillna('')
        
            # Lấy thông tin "Đợt", "Mã lớp học phần", "Tên môn học"
            dot = df.iloc[5, 2]  # Đợt nằm ở hàng 5, cột 2
            ma_lop = df.iloc[7, 2]  # Mã lớp học phần nằm ở hàng 7, cột 2
            ten_mon_hoc = df.iloc[8, 2]  # Tên môn học nằm ở hàng 8, cột 2

            df_sinh_vien = df.iloc[11:]  # Bắt đầu từ hàng 12 (hàng 11 trong lập chỉ số 0)
            
            # Kết hợp tên cột từ hai hàng đầu
            header1 = df_sinh_vien.iloc[0]  # Hàng đầu tiên
            header2 = df_sinh_vien.iloc[1]  # Hàng thứ hai
            
            df_sinh_vien.columns = [
                f"{str(header1[i]).strip()}_{str(header2[i]).strip()}" if header1[i] or header2[i] else ''
                for i in range(len(header1))
            ]

            df_sinh_vien = df_sinh_vien[2:]  # Loại bỏ hai hàng tiêu đề đã sử dụng

            # Xóa các cột không cần thiết
            df_sinh_vien = df_sinh_vien.loc[:, ~df_sinh_vien.columns.str.contains(
                r'\[Thứ hai\] - \[7->11\] - 29/07/2024|\[Thứ hai\] - \[7->11\] - 12/08/2024|'
                r'\[Thứ hai\] - \[7->11\] - 19/08/2024|\[Thứ ba\] - \[2->6\] - 20/08/2024|'
                r'\[Thứ hai\] - \[7->11\] - 26/08/2024|\[Thứ hai\] - \[7->11\] - 09/09/2024|'
                r'\(P/K\)|ST|LD'
            )]


            self.mssv_list = df_sinh_vien['Mã sinh viên_'].tolist()

            # Chuyển đổi các giá trị trong cột '%' vắng từ ',' thành '.'
            if '_(%) vắng' in df_sinh_vien.columns:
                df_sinh_vien['_(%) vắng'] = df_sinh_vien['_(%) vắng'].apply(
                    lambda x: str(x).replace(',', '.') if isinstance(x, str) else x)
                

            # Xóa bảng cũ nếu tồn tại
            self.cursor.execute("DROP TABLE IF EXISTS students")

            self.cursor.execute(""" 
            CREATE TABLE IF NOT EXISTS students (
                mssv TEXT PRIMARY KEY,
                ho_dem TEXT,
                ten TEXT,
                gioi_tinh TEXT,
                ngay_sinh TEXT,
                vắng_có_phép INTEGER,
                vắng_không_phép INTEGER,
                tong_so_tiet INTEGER,
                ty_le_vang REAL,
                dot TEXT,             -- Cột mới: Đợt
                ma_lop TEXT,          -- Cột mới: Mã lớp
                ten_mon_hoc TEXT,    -- Cột mới: Tên môn học
                email_student TEXT
            )
            """)
            
                # Tạo bảng parents
            self.cursor.execute('''
                CREATE TABLE IF NOT EXISTS parents (
                    mssv TEXT PRIMARY KEY,
                    email_ph TEXT  -- Email của phụ huynh
                )
            ''')

            # Tạo bảng teachers
            self.cursor.execute('''
                CREATE TABLE IF NOT EXISTS teachers (
                    mssv TEXT PRIMARY KEY,
                    email_gvcn TEXT  -- Email của giáo viên chủ nhiệm
                )
            ''')
            
            self.conn.commit()


            # Xóa dữ liệu cũ trước khi thêm dữ liệu mới
            self.cursor.execute("DELETE FROM students")
            self.conn.commit()
            # Lưu dữ liệu vào SQLite
            for index, row in df_sinh_vien.iterrows():
                vắng_có_phép = int(str(row['Tổng cộng_Vắng có phép']).strip() or 0)
                vắng_không_phép = int(str(row['_Vắng không phép']).strip() or 0)
                tong_so_tiet = int(str(row['_Tổng số tiết']).strip() or 0)
                ty_le_vang = float(str(row['_(%) vắng']).strip().replace(',', '.') or 0.0)

                self.cursor.execute(""" 
                    INSERT OR IGNORE INTO students (mssv, ho_dem, ten, gioi_tinh, ngay_sinh, vắng_có_phép, vắng_không_phép, tong_so_tiet, ty_le_vang, dot, ma_lop, ten_mon_hoc)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    row['Mã sinh viên_'],
                    row['Họ đệm_'],
                    row['Tên_'],
                    row['Giới tính_'],
                    str(row['Ngày sinh_']),
                    vắng_có_phép,
                    vắng_không_phép,
                    tong_so_tiet,
                    ty_le_vang,
                    dot,  # Đợt
                    ma_lop,  # Mã lớp học phần
                    ten_mon_hoc  # Tên môn học
                ))

            self.conn.commit()
            
   
            # Thêm dữ liệu vào bảng parents
            for mssv in self.mssv_list:
                email_student = f"tranhuuhauthh@gmail.com"  # Tạo email mẫu cho phụ huynh
                self.cursor.execute('UPDATE students SET email_student = ? WHERE mssv = ?', (email_student, mssv))
        

            # Thêm dữ liệu vào bảng parents
            for mssv in self.mssv_list:
                email_ph = f"tranhuuhauthh@gmail.com"  # Tạo email mẫu cho phụ huynh
                self.cursor.execute('INSERT OR IGNORE INTO parents (mssv, email_ph) VALUES (?, ?)', (mssv, email_ph))

            # Thêm dữ liệu vào bảng teachers
            for mssv in self.mssv_list:
                email_gvcn = f"tranhuuhauthh@gmail.com"  # Tạo email mẫu cho giáo viên chủ nhiệm
                self.cursor.execute('INSERT OR IGNORE INTO teachers (mssv, email_gvcn) VALUES (?, ?)', (mssv, email_gvcn))
                
                
            # messagebox.showinfo("Thông báo", "Dữ liệu đã được tải thành công!")
            self.load_students_to_treeview()  # Cập nhật Treeview với dữ liệu mới
        except Exception as e:
            messagebox.showerror("Lỗi", str(e))



    def load_students_to_treeview(self):
        # Xóa dữ liệu hiện có trong Treeview
        for row in self.tree.get_children():
            self.tree.delete(row)

        # Truy vấn lấy dữ liệu từ bảng students
        self.cursor.execute("""
            SELECT mssv, ho_dem, ten, ngay_sinh, dot, ma_lop, ten_mon_hoc, (vắng_có_phép + vắng_không_phép) as tong_vang
            FROM students
        """)
        rows = self.cursor.fetchall()
        for i, row in enumerate(rows, 1):
                self.tree.insert('', 'end', values=(i, *row))
        
        
        
        
    def main_menu(self):
        self.frame = tk.Frame(self.root, bg="#f0f0f0", width=800, height=600)
        self.frame.pack_propagate(False)
        self.frame.pack(fill='both', expand=True, padx=10, pady=10)

        self.tree_frame = tk.Frame(self.frame)
        self.tree_frame.pack(fill='both', expand=True)

        self.tree = ttk.Treeview(self.tree_frame, columns=['STT', 'MSSV', 'Họ đệm', 'Tên', 'Ngày Sinh', 'Đợt', 'Mã lớp', 'Tên môn học', 'Tổng vắng'], show="headings")
        self.tree.pack(side='left', fill='both', expand=True)

        self.tree.column('STT', width=30, anchor='center')
        self.tree.column('MSSV', width=70, anchor='center')
        self.tree.column('Họ đệm', width=100, anchor='w')
        self.tree.column('Tên', width=70, anchor='w')
        self.tree.column('Ngày Sinh', width=70, anchor='w')
        self.tree.column('Đợt', width=70, anchor='center')
        self.tree.column('Mã lớp', width=70, anchor='center')
        self.tree.column('Tên môn học', width=100, anchor='w')
        self.tree.column('Tổng vắng', width=70, anchor='center')

        for col in self.tree["columns"]:
            self.tree.heading(col, text=col, command=lambda _col=col: self.sort_column(_col, False))

        self.scrollbar = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        self.scrollbar.pack(side='right', fill='y')
        self.tree.configure(yscrollcommand=self.scrollbar.set)

   # Nút chọn file
        self.select_file_button = tk.Button(self.frame, text="Chọn File Excel", command=self.select_file, bg="#4CAF50", fg="white", font=("Arial", 10))
        self.select_file_button.pack(pady=10)
        
        self.view_button = tk.Button(self.frame, text="Xem chi tiết", command=self.view_details, bg="#4CAF50", fg="white", font=("Arial", 10))
        self.view_button.pack(pady=10)

        search_frame = tk.Frame(self.root, bg="#f0f0f0")
        search_frame.pack(fill='x', pady=5)

        tk.Label(search_frame, text="Search by MSSV or Name:", font=("Arial", 8)).pack(side="left", padx=5)
        self.search_entry = tk.Entry(search_frame, font=("Arial", 8), width=15)
        self.search_entry.pack(side="left", padx=5)

        search_button = tk.Button(search_frame, text="Search", command=self.search_student, bg="#4CAF50", fg="white", font=("Arial", 8), width=7)
        search_button.pack(side="left", padx=5)

        update_frame = tk.Frame(self.root, bg="#f0f0f0")
        update_frame.pack(fill='x', pady=5)

        add_button = tk.Button(update_frame, text="Add Student", command=self.add_student, bg="#4CAF50", fg="white", font=("Arial", 8), width=9)
        add_button.pack(side="left", padx=5)

        remove_button = tk.Button(update_frame, text="Remove Student", command=self.remove_student, bg="#f44336", fg="white", font=("Arial", 8), width=10)
        remove_button.pack(side="left", padx=5)

        load_button = tk.Button(update_frame, text="Load Data", command=self.load_data_from_excel, bg="#008CBA", fg="white", font=("Arial", 8), width=9)
        load_button.pack(side="left", padx=5)

        load_button = tk.Button(update_frame, text="Send Mail", command=self.send_warning_emails, bg="#008CBA", fg="white", font=("Arial", 8), width=9)
        load_button.pack(side="left", padx=5)
        
        
        # self.load_data()
        
    def select_file(self):
        self.filepath = filedialog.askopenfilename(title="Chọn File Excel", filetypes=[("Excel Files", "*.xls;*.xlsx")])
        if self.filepath:
            self.load_data_from_excel()

    def load_data(self):
        try:
            query = """
            SELECT mssv, ho_dem, ten, ngay_sinh, dot, ma_lop, ten_mon_hoc, (vắng_có_phép + vắng_không_phép) as tong_vang
            FROM students

            """
            self.cursor.execute(query)
            records = self.cursor.fetchall()

            for i, row in enumerate(records, 1):
                self.tree.insert('', 'end', values=(i, *row))

        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể tải dữ liệu: {e}")

    def sort_column(self, col, reverse):
        # Hàm để sắp xếp các cột trong treeview
        l = [(self.tree.set(k, col), k) for k in self.tree.get_children('')]
        l.sort(reverse=reverse)

        # Sắp xếp lại dữ liệu
        for index, (val, k) in enumerate(l):
            self.tree.move(k, '', index)

        # Sắp xếp theo thứ tự ngược lại khi nhấn tiếp vào tiêu đề cột
        self.tree.heading(col, command=lambda: self.sort_column(col, not reverse))

    # Các hàm view_details, search_student, add_student, remove_student sẽ được giữ nguyên

    def search_student(self):
        search_value = self.search_entry.get().lower()
        if search_value:
            query = f"SELECT * FROM students WHERE mssv LIKE ? OR ten LIKE ?"
            self.cursor.execute(query, ('%' + search_value + '%', '%' + search_value + '%'))
            result = self.cursor.fetchall()

            self.tree.delete(*self.tree.get_children())

            # for row in result:
            #     self.tree.insert("", "end", values=row)
            for index, row in enumerate(result, start=1):
                self.tree.insert("", "end", values=(index,) + row)
        else:
            self.load_data()

    def view_details(self):
        selected_item = self.tree.selection()
        if selected_item:
            item_data = self.tree.item(selected_item, 'values')
            mssv = item_data[1]  # Get MSSV
            query = f"SELECT * FROM students WHERE mssv = ?"
            self.cursor.execute(query, (mssv,))
            details_data = self.cursor.fetchone()
            
            if details_data:
                details = (
                    f"MSSV: {details_data[0]}\n"
                    f"Họ đệm: {details_data[1]}\n"
                    f"Tên: {details_data[2]}\n"
                    f"Giới tính: {details_data[3]}\n"
                    f"Ngày sinh: {details_data[4]}\n"
                    f"Vắng có phép: {details_data[5]}\n"
                    f"Vắng không phép: {details_data[6]}\n"
                    f"Tổng số tiết: {details_data[7]}\n"
                    f"Tỷ lệ vắng: {details_data[8]}%\n"
                    f"Đợt: {details_data[9]}\n"
                    f"Mã lớp: {details_data[10]}\n"
                    f"Tên môn học: {details_data[11]}"
                )
                messagebox.showinfo("Chi tiết thông tin sinh viên", details)
            else:
                messagebox.showerror("Lỗi", "Không tìm thấy thông tin sinh viên.")
        else:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn một sinh viên để xem chi tiết.")

    def add_student(self):
        def save_student():
            mssv = mssv_entry.get().strip()  # Xóa khoảng trắng đầu và cuối
            ho_dem = ho_dem_entry.get().strip()
            ten = ten_entry.get().strip()
            gioi_tinh = gioi_tinh_var.get()
            ngay_sinh = ngay_sinh_entry.get().strip()
            vắng_có_phép = int(vang_co_phap_entry.get() or 0)
            vắng_không_phép = int(vang_khong_phap_entry.get() or 0)
            tong_so_tiet = int(tong_so_tiet_entry.get() or 0)

            # Kiểm tra xem MSSV và Tên có rỗng hay không
            if not mssv or not ten:
                messagebox.showerror("Error", "MSSV và Tên sinh viên không được để trống!")
                return  # Ngừng thực hiện nếu MSSV hoặc Tên rỗng

            # Tính tỷ lệ vắng
            if tong_so_tiet > 0:
                ty_le_vang = (vắng_có_phép + vắng_không_phép) / tong_so_tiet * 100
            else:
                ty_le_vang = 0.0  # Nếu không có tiết nào, tỷ lệ vắng sẽ là 0%

            # Lưu thông tin sinh viên vào cơ sở dữ liệu
            self.cursor.execute("""
                INSERT INTO students (mssv, ho_dem, ten, gioi_tinh, ngay_sinh, vắng_có_phép, vắng_không_phép, tong_so_tiet, ty_le_vang)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (mssv, ho_dem, ten, gioi_tinh, ngay_sinh, vắng_có_phép, vắng_không_phép, tong_so_tiet, ty_le_vang))

            self.conn.commit()
            messagebox.showinfo("Info", "Student added successfully!")
            add_window.destroy()
            self.load_data()

        # Tạo cửa sổ thêm sinh viên
        add_window = tk.Toplevel(self.root)
        add_window.title("Add Student")

        tk.Label(add_window, text="MSSV:").grid(row=0, column=0, padx=10, pady=10)
        tk.Label(add_window, text="Họ đệm:").grid(row=1, column=0, padx=10, pady=10)
        tk.Label(add_window, text="Tên:").grid(row=2, column=0, padx=10, pady=10)
        tk.Label(add_window, text="Giới tính:").grid(row=3, column=0, padx=10, pady=10)
        tk.Label(add_window, text="Ngày sinh:").grid(row=4, column=0, padx=10, pady=10)
        tk.Label(add_window, text="Vắng có phép:").grid(row=5, column=0, padx=10, pady=10)
        tk.Label(add_window, text="Vắng không phép:").grid(row=6, column=0, padx=10, pady=10)
        tk.Label(add_window, text="Tổng số tiết:").grid(row=7, column=0, padx=10, pady=10)

        mssv_entry = tk.Entry(add_window)
        ho_dem_entry = tk.Entry(add_window)
        ten_entry = tk.Entry(add_window)
        gioi_tinh_var = tk.StringVar(value="Nam")  # Mặc định giới tính là Nam
        tk.Radiobutton(add_window, text="Nam", variable=gioi_tinh_var, value="Nam").grid(row=3, column=1)
        tk.Radiobutton(add_window, text="Nữ", variable=gioi_tinh_var, value="Nữ").grid(row=3, column=2)
        ngay_sinh_entry = tk.Entry(add_window)
        vang_co_phap_entry = tk.Entry(add_window)
        vang_khong_phap_entry = tk.Entry(add_window)
        tong_so_tiet_entry = tk.Entry(add_window)

        mssv_entry.grid(row=0, column=1)
        ho_dem_entry.grid(row=1, column=1)
        ten_entry.grid(row=2, column=1)
        ngay_sinh_entry.grid(row=4, column=1)
        vang_co_phap_entry.grid(row=5, column=1)
        vang_khong_phap_entry.grid(row=6, column=1)
        tong_so_tiet_entry.grid(row=7, column=1)

        tk.Button(add_window, text="Save student", command=save_student, bg="#4CAF50", fg="white").grid(row=8, column=0, columnspan=2, pady=10)



    def remove_student(self):
        selected_item = self.tree.selection()
        if selected_item:
            student = self.tree.item(selected_item, "values")[1]  # Lấy giá trị MSSV từ cột thứ 2
            confirm = messagebox.askyesno("Xóa sinh viên", f"Bạn có chắc muốn xóa sinh viên {student}?")
            if confirm: 
                self.cursor.execute("DELETE FROM students WHERE mssv = ?", (student,))
                self.conn.commit()
                self.load_data()
                messagebox.showinfo("Info", f"Student {student} removed successfully!")
        else:
            messagebox.showwarning("Cảnh báo", "Chọn sinh viên để xóa!")
            
    def send_email(self, to_address, subject, message):
        """Gửi email tới địa chỉ nhận"""
        from_address = "tranhuuhau2003@gmail.com"
        password = "frpq iken dpth pxku"
        
        # Khởi tạo email
        msg = MIMEMultipart()
        msg['From'] = from_address
        msg['To'] = to_address
        msg['Subject'] = subject

        # Nội dung email
        msg.attach(MIMEText(message, 'plain'))

        # Cấu hình server SMTP để gửi email
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

    def send_warning_emails(self):
        """Kiểm tra và gửi cảnh báo học vụ cho sinh viên"""
        try:
            query = """
            SELECT mssv, ho_dem, ten, vắng_có_phép, vắng_không_phép, tong_so_tiet, ty_le_vang
            FROM students
            """
            self.cursor.execute(query)
            records = self.cursor.fetchall()

            for row in records:
                mssv, ho_dem, ten, vắng_có_phép, vắng_không_phép, tong_so_tiet, ty_le_vang = row
                # ty_le_vang = self.calculate_absence_percentage(vắng_có_phép, vắng_không_phép, tong_so_tiet)

                # Lấy email của sinh viên, phụ huynh, GVCN từ cơ sở dữ liệu (giả sử có các cột email)
                student_email = self.get_student_email(mssv)
                parent_email = self.get_parent_email(mssv)
                homeroom_teacher_email = self.get_teacher_email(mssv)

                # Kiểm tra và gửi cảnh báo theo tỷ lệ vắng
                if ty_le_vang >= 50:
                    # Gửi email cho sinh viên, phụ huynh, GVCN
                    subject = "Cảnh báo học vụ: Vắng học quá 50%"
                    message = f"Sinh viên {ho_dem} {ten} đã vắng hơn 50% số buổi học."
                    self.send_email(student_email, subject, message)
                    self.send_email(parent_email, subject, message)
                    self.send_email(homeroom_teacher_email, subject, message)
                elif ty_le_vang >= 20:
                    # Gửi email chỉ cho sinh viên
                    subject = "Cảnh báo học vụ: Vắng học quá 20%"
                    message = f"Sinh viên {ho_dem} {ten} đã vắng hơn 20% số buổi học."
                    self.send_email(student_email, subject, message)

        except Exception as e:
            print(f"Lỗi khi gửi email cảnh báo: {e}")

    def get_student_email(self, mssv):
        """Lấy email sinh viên từ cơ sở dữ liệu dựa trên MSSV"""
        # Bạn cần có cột email trong bảng students hoặc một bảng riêng
        query = f"SELECT email_student FROM students WHERE mssv = '{mssv}'"
        self.cursor.execute(query)
        result = self.cursor.fetchone()
        return result[0] if result else None

    def get_parent_email(self, mssv):
        """Lấy email phụ huynh từ cơ sở dữ liệu dựa trên MSSV"""
        query = f"SELECT email_ph FROM parents WHERE mssv = '{mssv}'"
        self.cursor.execute(query)
        result = self.cursor.fetchone()
        return result[0] if result else None

    def get_teacher_email(self, mssv):
        """Lấy email GVCN từ cơ sở dữ liệu dựa trên MSSV"""
        query = f"SELECT email_gvcn FROM teachers WHERE mssv = '{mssv}'"
        self.cursor.execute(query)
        result = self.cursor.fetchone()
        return result[0] if result else None
    
    def print_emails(self):
        try:
            self.cursor.execute("SELECT mssv, email_student FROM students")
            parents_records = self.cursor.fetchall()

            print("Emails of student:")
            for record in parents_records:
                print(f"MSSV: {record[0]}, Email: {record[1]}")


            self.cursor.execute("SELECT mssv, email_ph FROM parents")
            parents_records = self.cursor.fetchall()

            print("Emails of Parents:")
            for record in parents_records:
                print(f"MSSV: {record[0]}, Email: {record[1]}")
                

            self.cursor.execute("SELECT mssv, email_gvcn FROM teachers")
            teachers_records = self.cursor.fetchall()

            print("Emails of Teachers:")
            for record in teachers_records:
                print(f"MSSV: {record[0]}, Email: {record[1]}")

        except Exception as e:
            messagebox.showerror("Error", f"Could not retrieve emails: {e}")
    
    
    def generate_report_and_send_email(self):
        # Tạo kết nối SQLite mới trong hàm này
        conn = sqlite3.connect('students.db')
        cursor = conn.cursor()

        try:
            # Lấy dữ liệu sinh viên vắng nhiều
            query = """
            SELECT mssv, ho_dem, ten, gioi_tinh, ngay_sinh, vắng_có_phép, vắng_không_phép, tong_so_tiet, ty_le_vang
            FROM students
            WHERE ty_le_vang >= 20
            """
            
            # Thực hiện truy vấn
            cursor.execute(query)
            results = cursor.fetchall()
            
            # Chuyển đổi kết quả thành DataFrame
            columns = [column[0] for column in cursor.description]  # Lấy tên cột
            df = pd.DataFrame(results, columns=columns)

            # Lưu dữ liệu vào file Excel
            file_path = 'tong_hop_sinh_vien_vang_nhieu.xlsx'
            df.to_excel(file_path, index=False)

            # Gửi email với tệp đính kèm
            self.send_email_with_attachment(file_path)  # Đảm bảo rằng phương thức này có sẵn trong lớp của bạn

        except Exception as e:
            print(f"Có lỗi xảy ra trong quá trình tạo báo cáo và gửi email: {e}")

        finally:
            conn.close()  # Đảm bảo đóng kết nối

    
    def send_email_with_attachment(self, file_path):
        # Cấu hình thông tin email
        sender_email = "tranhuuhau2003@gmail.com"  # Địa chỉ email của bạn
        sender_password = "frpq iken dpth pxku"  # Mật khẩu email của bạn
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
        attachment = open(file_path, "rb")
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename= {file_path}')
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
                
    def start_scheduler(self):
        print("Khởi động lịch trình gửi email...")
        def run_scheduler():
            while True:
                schedule.run_pending()
                time.sleep(1)

        threading.Thread(target=run_scheduler, daemon=True).start()

    def check_send_email(self):
        today = datetime.now()
        print(f"Kiểm tra thời gian gửi email: Ngày {today.day}, Giờ {today.strftime('%H:%M:%S')}")

        if today.day in [1, 23]:  # Kiểm tra nếu hôm nay là ngày 1 hoặc 23
            print(f"Gửi email vào ngày {today.day}...")
            self.generate_report_and_send_email()
        else:
            print("Không đến thời gian gửi email.")


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelApp(root)
    root.mainloop()
