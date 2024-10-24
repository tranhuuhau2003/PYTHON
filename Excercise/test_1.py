import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd

class ExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Data Viewer")
        self.filepath = 'Excercise\diem-danh-sinh-vien-04102024094447.xls'

        # Frame chứa Treeview và thanh cuộn
        frame = tk.Frame(root)
        frame.pack(fill='both', expand=True)

        # Tạo Treeview để hiển thị dữ liệu
        self.tree = ttk.Treeview(frame)
        self.tree.pack(side='left', fill='both', expand=True)

        # Thanh cuộn cho Treeview
        self.scrollbar = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        self.scrollbar.pack(side='right', fill='y')
        self.tree.configure(yscrollcommand=self.scrollbar.set)

        # Tải dữ liệu tự động khi khởi động ứng dụng
        self.load_data()    

    # def load_data(self):
    #     try:
    #         # Đọc file Excel
    #         df = pd.read_excel(self.filepath, engine='xlrd', header=[11,12])  # Đổi engine tùy thuộc vào file

    #         df = df.fillna('')
            
    #         df.columns = ['_'.join(map(str, col)) for col in df.columns]

    #         df = df.drop(columns=[col for col in df.columns if '29/07/2024' in col or 
    #                                              '12/08/2024' in col or
    #                                              '19/08/2024' in col or
    #                                              '20/08/2024' in col or
    #                                              '26/08/2024' in col or
    #                                              '09/09/2024' in col])
            
    #         # Xóa dữ liệu cũ trong Treeview nếu có
    #         self.tree.delete(*self.tree.get_children())

    #         # Hiển thị tiêu đề cột
    #         self.tree["columns"] = list(df.columns)
    #         self.tree["show"] = "headings"
    #         for col in df.columns:
    #             self.tree.heading(col, text=col)
    #             max_width = max(df[col].astype(str).apply(len).max(), len(col)) * 2  # Tính độ rộng tự động
    #             self.tree.column(col, anchor="center", width=max_width)

    #         # Hiển thị dữ liệu
    #         for index, row in df.iterrows():
    #             self.tree.insert("", "end", values=list(row))

    #     except Exception as e:
    #         messagebox.showerror("Lỗi", f"Không thể tải file: {e}")
    
    def load_data(self):
        try:
            # Đọc file Excel
            df = pd.read_excel(self.filepath, engine='xlrd', header=[11,12])            
            # Chỉ chọn các cột cần hiển thị
            columns_to_display = ['STT', 'Mã sinh viên', 'Họ đệm', 'Tên', 'Giới tính', 'Ngày sinh', 'Tổng cộng']
                                #   Vắng có phép', 'Vắng không phép', 'Tổng số tiết', '(%) vắng'
            df = df[columns_to_display]  # Lọc những cột cần thiết để hiển thị

            # Xóa dữ liệu cũ trong Treeview nếu có
            self.tree.delete(*self.tree.get_children())

            # Hiển thị tiêu đề cột
            self.tree["columns"] = list(df.columns)
            self.tree["show"] = "headings"
            for col in df.columns:
                self.tree.heading(col, text=col)
                max_width = max(df[col].astype(str).apply(len).max(), len(col)) * 2  # Tính độ rộng tự động
                self.tree.column(col, anchor="center", width=max_width)

            # Hiển thị dữ liệu
            for index, row in df.iterrows():
                self.tree.insert("", "end", values=list(row))

        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể tải file: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelApp(root)
    root.geometry("1400x700")  # Tăng kích thước cửa sổ để hiển thị nhiều cột hơn
    root.mainloop()
