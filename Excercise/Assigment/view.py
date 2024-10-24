
import tkinter as tk
from tkinter import ttk, messagebox

class StudentView:
    def __init__(self, root):
        self.root = root
        self.root.title("Student Attendance Tracker")
        self.tree = None
        self.create_widgets()

    def create_widgets(self):
        self.tree = ttk.Treeview(self.root, columns=('MSSV', 'Họ đệm', 'Tên', 'Giới tính', 'Ngày sinh', 'Vắng có phép', 'Vắng không phép', 'Tổng số tiết', 'Tỷ lệ vắng', 'Đợt'), show='headings')
        for col in self.tree['columns']:
            self.tree.heading(col, text=col)
        self.tree.pack(expand=True, fill=tk.BOTH)

    def populate_table(self, data):
        for row in data:
            self.tree.insert("", tk.END, values=row)

    def show_error(self, message):
        messagebox.showerror("Error", message)
