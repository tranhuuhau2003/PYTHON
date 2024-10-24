import tkinter as tk
from tkinter import ttk
from database_manager import DatabaseManager
from excel_loader import ExcelLoader
from student_manager import StudentManager

class MainApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Data Viewer")

        self.db_manager = DatabaseManager()
        self.excel_loader = ExcelLoader(self.db_manager, 'Excel/diem-danh-sinh-vien-04102024094447.xls')
        self.student_manager = StudentManager(self.db_manager)

        self.main_menu()

    def main_menu(self):
        self.frame = tk.Frame(self.root, bg="#f0f0f0", width=800, height=600)
        self.frame.pack_propagate(False)
        self.frame.pack(fill='both', expand=True, padx=10, pady=10)

        self.tree_frame = tk.Frame(self.frame)
        self.tree_frame.pack(fill='both', expand=True)

        self.tree = ttk.Treeview(self.tree_frame, columns=['STT', 'MSSV', 'Họ đệm', 'Tên', 'Đợt', 'Mã lớp', 'Tên môn học', 'Tổng vắng'], show="headings")
        self.tree.pack(side='left', fill='both', expand=True)

        for col in self.tree["columns"]:
            self.tree.heading(col, text=col)

        self.scrollbar = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        self.scrollbar.pack(side='right', fill='y')
        self.tree.configure(yscrollcommand=self.scrollbar.set)

        self.student_manager.view_students(self.tree)

        search_frame = tk.Frame(self.root, bg="#f0f0f0")
        search_frame.pack(fill='x', pady=5)

        tk.Label(search_frame, text="Search by MSSV or Name:", font=("Arial", 8)).pack(side="left", padx=5)
        self.search_entry = tk.Entry(search_frame, font=("Arial", 8), width=15)
        self.search_entry.pack(side="left", padx=5)

        search_button = tk.Button(search_frame, text="Search", command=self.search_student, bg="#4CAF50", fg="white", font=("Arial", 8), width=7)
        search_button.pack(side="left", padx=5)

    def search_student(self):
        search_value = self.search_entry.get().lower()
        self.tree.delete(*self.tree.get_children())
        self.student_manager.search_student(self.tree, search_value)


if __name__ == "__main__":
    root = tk.Tk()
    app = MainApp(root)
    root.mainloop()
