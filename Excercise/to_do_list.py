import tkinter as tk
from tkinter import Listbox, messagebox
import sqlite3

# Kết nối đến SQLite và tạo bảng nếu chưa tồn tại
def connect_db():
    conn = sqlite3.connect("tasks.db")
    cursor = conn.cursor()
    cursor.execute('''CREATE TABLE IF NOT EXISTS tasks 
                      (id INTEGER PRIMARY KEY AUTOINCREMENT, 
                      task TEXT, 
                      status TEXT)''')
    conn.commit()
    conn.close()

# Hàm thêm nhiệm vụ vào cơ sở dữ liệu
def add_task():
    task = task_entry.get()
    if task != "":
        conn = sqlite3.connect("tasks.db")
        cursor = conn.cursor()
        cursor.execute("INSERT INTO tasks (task, status) VALUES (?, ?)", (task, "incomplete"))
        conn.commit()
        conn.close()

        # Cập nhật giao diện sau khi thêm nhiệm vụ
        show_tasks()
        task_entry.delete(0, tk.END)
    else:
        messagebox.showwarning("Input Error", "Please enter a task.")

# Hàm hiển thị danh sách nhiệm vụ trong Listbox
def show_tasks():
    conn = sqlite3.connect("tasks.db")
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM tasks")
    records = cursor.fetchall()
    conn.close()

    task_listbox.delete(0, tk.END)
    for record in records:
        task_listbox.insert(tk.END, f"{record[1]} - {record[2]}")

# Hàm đánh dấu nhiệm vụ là hoàn thành trong cơ sở dữ liệu
def complete_task():
    try:
        selected_task_index = task_listbox.curselection()[0]
        task = task_listbox.get(selected_task_index)
        task_text = task.split(" - ")[0]

        conn = sqlite3.connect("tasks.db")
        cursor = conn.cursor()
        cursor.execute("UPDATE tasks SET status=? WHERE task=?", ("complete", task_text))
        conn.commit()
        conn.close()

        # Cập nhật giao diện sau khi đánh dấu hoàn thành
        show_tasks()
    except:
        messagebox.showwarning("Selection Error", "No task selected.")

# Hàm xóa nhiệm vụ khỏi cơ sở dữ liệu
def delete_task():
    try:
        selected_task_index = task_listbox.curselection()[0]
        task = task_listbox.get(selected_task_index)
        task_text = task.split(" - ")[0]

        conn = sqlite3.connect("tasks.db")
        cursor = conn.cursor()
        cursor.execute("DELETE FROM tasks WHERE task=?", (task_text,))
        conn.commit()
        conn.close()

        # Cập nhật giao diện sau khi xóa nhiệm vụ
        show_tasks()
    except:
        messagebox.showwarning("Selection Error", "No task selected.")

# Tạo cửa sổ chính
root = tk.Tk()
root.title("To-Do List App")
root.geometry("400x300")

# Label và Entry cho việc nhập nhiệm vụ
tk.Label(root, text="Enter task:").grid(row=0, column=0, padx=10, pady=5)
task_entry = tk.Entry(root)
task_entry.grid(row=0, column=1)

# Nút Add Task
add_button = tk.Button(root, text="Add Task", command=add_task)
add_button.grid(row=1, column=0)

# Nút Complete Task
complete_button = tk.Button(root, text="Complete Task", command=complete_task)
complete_button.grid(row=1, column=1)

# Nút Delete Task
delete_button = tk.Button(root, text="Delete Task", command=delete_task)
delete_button.grid(row=1, column=2)

# Listbox để hiển thị danh sách nhiệm vụ
task_listbox = Listbox(root, width=50, height=10)
task_listbox.grid(row=2, column=0, columnspan=3, padx=10, pady=10)

# Kết nối đến cơ sở dữ liệu và hiển thị nhiệm vụ khi khởi động ứng dụng
connect_db()
show_tasks()

# Chạy ứng dụng
root.mainloop()
