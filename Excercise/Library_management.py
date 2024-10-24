import tkinter as tk
from tkinter import messagebox, Listbox
import sqlite3

# Kết nối đến SQLite và tạo bảng nếu chưa tồn tại
def connect_db():
    conn = sqlite3.connect("library.db")
    cursor = conn.cursor()
    cursor.execute('''CREATE TABLE IF NOT EXISTS books 
                      (id INTEGER PRIMARY KEY AUTOINCREMENT, 
                      title TEXT, 
                      author TEXT, 
                      year TEXT)''')
    conn.commit()
    conn.close()

# Hàm thêm sách vào cơ sở dữ liệu
def add_book():
    title = title_entry.get()
    author = author_entry.get()
    year = year_entry.get()

    if title and author and year:
        conn = sqlite3.connect("library.db")
        cursor = conn.cursor()
        cursor.execute("INSERT INTO books (title, author, year) VALUES (?, ?, ?)", (title, author, year))
        conn.commit()
        conn.close()

        # Cập nhật giao diện sau khi thêm sách
        show_books()
        title_entry.delete(0, tk.END)
        author_entry.delete(0, tk.END)
        year_entry.delete(0, tk.END)
    else:
        messagebox.showwarning("Input Error", "All fields must be filled out")

# Hàm hiển thị danh sách sách trong Listbox
def show_books():
    conn = sqlite3.connect("library.db")
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM books")
    records = cursor.fetchall()
    conn.close()

    book_listbox.delete(0, tk.END)
    for record in records:
        book_listbox.insert(tk.END, f"{record[1]}, by {record[2]} ({record[3]})")

# Hàm xóa sách khỏi cơ sở dữ liệu
def delete_book():
    try:
        selected_book = book_listbox.get(book_listbox.curselection())
        title_to_delete = selected_book.split(", by")[0]

        conn = sqlite3.connect("library.db")
        cursor = conn.cursor()
        cursor.execute("DELETE FROM books WHERE title=?", (title_to_delete,))
        conn.commit()
        conn.close()

        # Cập nhật giao diện sau khi xóa sách
        show_books()
    except:
        messagebox.showwarning("Selection Error", "No book selected")

# Tạo cửa sổ chính
root = tk.Tk()
root.title("Library Management System")
root.geometry("350x350")

# Label và Entry cho Title
tk.Label(root, text="Title:").grid(row=0, column=0, padx=10, pady=5)
title_entry = tk.Entry(root)
title_entry.grid(row=0, column=1)

# Label và Entry cho Author
tk.Label(root, text="Author:").grid(row=1, column=0, padx=10, pady=5)
author_entry = tk.Entry(root)
author_entry.grid(row=1, column=1)

# Label và Entry cho Year
tk.Label(root, text="Year:").grid(row=2, column=0, padx=10, pady=5)
year_entry = tk.Entry(root)
year_entry.grid(row=2, column=1)

# Nút Add Book
add_button = tk.Button(root, text="Add Book", command=add_book)
add_button.grid(row=3, column=0, columnspan=2, pady=10)

# Listbox để hiển thị danh sách sách
book_listbox = Listbox(root, width=50, height=8)
book_listbox.grid(row=4, column=0, columnspan=2, padx=10, pady=10)

# Nút Delete Book
delete_button = tk.Button(root, text="Delete Book", command=delete_book)
delete_button.grid(row=5, column=0, columnspan=2, pady=10)

# Kết nối đến cơ sở dữ liệu và hiển thị sách khi khởi động ứng dụng
connect_db()
show_books()

root.mainloop()
