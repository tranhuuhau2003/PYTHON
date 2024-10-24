import pandas as pd
import tkinter as tk
from tkinter import messagebox, ttk

# Đường dẫn đến file Excel
file_path = 'diem-danh-sinh-vien-04102024094447.xls'

# Đọc file excel
df = pd.read_excel(file_path, sheet_name='Sheet1', header=11)

# Loại bỏ các cột không cần thiết bằng cách sử dụng chỉ số cột
df_cleaned = df.drop(df.columns[6:24], axis=1)

print(df_cleaned.head(30))  # Hiển thị 10 dòng đầu tiên để kiểm tra

# Lớp quản lý sinh viên
class StudentManagementApp:
    def __init__(self, root, df_cleaned):
        self.root = root
        self.root.title("Quản lý Sinh Viên")
        self.root.geometry("800x400")

        # Lưu DataFrame đã sạch
        self.df_cleaned = df_cleaned

        # Khung nhập thông tin sinh viên
        self.frame = tk.Frame(self.root)
        self.frame.pack(pady=20)

        self.label_name = tk.Label(self.frame, text="Họ và Tên")
        self.label_name.grid(row=0, column=0)
        self.entry_name = tk.Entry(self.frame)
        self.entry_name.grid(row=0, column=1)

        self.label_id = tk.Label(self.frame, text="Mã Sinh Viên")
        self.label_id.grid(row=1, column=0)
        self.entry_id = tk.Entry(self.frame)
        self.entry_id.grid(row=1, column=1)

        self.label_class = tk.Label(self.frame, text="Lớp")
        self.label_class.grid(row=2, column=0)
        self.entry_class = tk.Entry(self.frame)
        self.entry_class.grid(row=2, column=1)

        # Nút Thêm sinh viên
        self.add_button = tk.Button(self.frame, text="Thêm Sinh Viên", command=self.add_student)
        self.add_button.grid(row=3, column=0, pady=10)

        # Nút Sửa sinh viên
        self.update_button = tk.Button(self.frame, text="Sửa Sinh Viên", command=self.update_student)
        self.update_button.grid(row=3, column=1, pady=10)

        # Nút Xóa sinh viên
        self.delete_button = tk.Button(self.frame, text="Xóa Sinh Viên", command=self.delete_student)
        self.delete_button.grid(row=3, column=2, pady=10)

        # Bảng hiển thị danh sách sinh viên
        self.tree = ttk.Treeview(self.root, columns=("ID", "Name", "Class"), show="headings")
        self.tree.heading("ID", text="Mã Sinh Viên")
        self.tree.heading("Name", text="Họ và Tên")
        self.tree.heading("Class", text="Lớp")
        self.tree.pack(pady=20)

        # Danh sách sinh viên
        self.student_list = self.load_students_from_dataframe()

        # Cập nhật Treeview ban đầu
        self.update_treeview()

    # Hàm tải sinh viên từ DataFrame
    def load_students_from_dataframe(self):
        students = []
        for index, row in self.df_cleaned.iterrows():
            students.append((row[1], row[0], ""))  # Sử dụng cột Mã Sinh Viên và Họ và Tên
        return students

    # Hàm thêm sinh viên
    def add_student(self):
        name = self.entry_name.get()
        student_id = self.entry_id.get()
        student_class = self.entry_class.get()

        if name and student_id and student_class:
            self.student_list.append((student_id, name, student_class))
            self.update_treeview()
            self.clear_entries()
        else:
            messagebox.showwarning("Thiếu thông tin", "Vui lòng nhập đầy đủ thông tin")

    # Hàm cập nhật sinh viên
    def update_student(self):
        selected_item = self.tree.selection()
        if selected_item:
            index = self.tree.index(selected_item)
            self.student_list[index] = (
                self.entry_id.get(),
                self.entry_name.get(),
                self.entry_class.get()
            )
            self.update_treeview()
            self.clear_entries()
        else:
            messagebox.showwarning("Chọn Sinh Viên", "Vui lòng chọn sinh viên để sửa")

    # Hàm xóa sinh viên
    def delete_student(self):
        selected_item = self.tree.selection()
        if selected_item:
            index = self.tree.index(selected_item)
            del self.student_list[index]
            self.update_treeview()
        else:
            messagebox.showwarning("Chọn Sinh Viên", "Vui lòng chọn sinh viên để xóa")

    # Hàm cập nhật danh sách hiển thị
    def update_treeview(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        for student in self.student_list:
            self.tree.insert("", "end", values=student)

    # Hàm xóa các ô nhập
    def clear_entries(self):
        self.entry_name.delete(0, tk.END)
        self.entry_id.delete(0, tk.END)
        self.entry_class.delete(0, tk.END)

# Chạy ứng dụng
if __name__ == "__main__":
    root = tk.Tk()
    app = StudentManagementApp(root, df_cleaned)
    root.mainloop()
