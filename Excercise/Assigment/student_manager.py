class StudentManager:
    def __init__(self, db_manager):
        self.db_manager = db_manager

    def view_students(self, tree):
        records = self.db_manager.fetch_students()
        for i, row in enumerate(records, 1):
            tree.insert('', 'end', values=(i, *row))

    def search_student(self, tree, search_value):
        results = self.db_manager.search_student(search_value)
        for index, row in enumerate(results, start=1):
            tree.insert("", "end", values=(index,) + row)

    
    def load_data(self):
        try:
            query = """
            SELECT mssv, ho_dem, ten, dot, ma_lop, ten_mon_hoc, (vắng_có_phép + vắng_không_phép) as tong_vang
            FROM students
            ORDER BY tong_vang DESC  -- Sắp xếp theo tổng vắng giảm dần
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
            