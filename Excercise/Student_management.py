import tkinter as tk
from tkinter import messagebox
import sqlite3


# Function to connect to SQLite and create the table if not exists
def connect_db():
    conn = sqlite3.connect("students.db")
    cursor = conn.cursor()
    cursor.execute(
        '''CREATE TABLE IF NOT EXISTS students 
           (id INTEGER PRIMARY KEY AUTOINCREMENT, 
           name TEXT, age INTEGER, grade TEXT)'''
    )
    conn.commit()
    conn.close()


# Function to add a new student to the database
def add_student():
    conn = sqlite3.connect("students.db")
    cursor = conn.cursor()

    name = entry_name.get()
    age = entry_age.get()
    grade = entry_grade.get()

    if name == "" or age == "" or grade == "":
        messagebox.showwarning("Input Error", "Please fill in all fields")
        return

    cursor.execute("INSERT INTO students (name, age, grade) VALUES (?, ?, ?)", (name, age, grade))
    conn.commit()
    conn.close()

    messagebox.showinfo("Success", "Student added successfully")
    clear_entries()
    show_students()


# Function to display students in the Listbox
def show_students():
    conn = sqlite3.connect("students.db")
    cursor = conn.cursor()

    cursor.execute("SELECT * FROM students")
    records = cursor.fetchall()

    listbox_students.delete(0, tk.END)
    for record in records:
        listbox_students.insert(tk.END, f"{record[1]}, Age: {record[2]}, Grade: {record[3]}")

    conn.close()


# Function to clear input fields
def clear_entries():
    entry_name.delete(0, tk.END)
    entry_age.delete(0, tk.END)
    entry_grade.delete(0, tk.END)


# Function to delete selected student from the database
def delete_student():
    try:
        selected_item = listbox_students.get(listbox_students.curselection())
        name_to_delete = selected_item.split(",")[0]

        conn = sqlite3.connect("students.db")
        cursor = conn.cursor()
        cursor.execute("DELETE FROM students WHERE name=?", (name_to_delete,))
        conn.commit()
        conn.close()

        messagebox.showinfo("Success", "Student deleted successfully")
        show_students()
    except:
        messagebox.showwarning("Selection Error", "Please select a student to delete")


# Create the main window
root = tk.Tk()
root.title("Student Management System")

# Create labels and entry fields for student information
label_name = tk.Label(root, text="Name")
label_name.grid(row=0, column=0)

entry_name = tk.Entry(root)
entry_name.grid(row=0, column=1)

label_age = tk.Label(root, text="Age")
label_age.grid(row=1, column=0)

entry_age = tk.Entry(root)
entry_age.grid(row=1, column=1)

label_grade = tk.Label(root, text="Grade")
label_grade.grid(row=2, column=0)

entry_grade = tk.Entry(root)
entry_grade.grid(row=2, column=1)

# Add Student button
button_add = tk.Button(root, text="Add Student", command=add_student)
button_add.grid(row=3, column=0, columnspan=2, pady=10)

# Listbox to display students
listbox_students = tk.Listbox(root, height=10, width=50)
listbox_students.grid(row=4, column=0, columnspan=2, padx=10, pady=10)

# Delete Student button
button_delete = tk.Button(root, text="Delete Student", command=delete_student)
button_delete.grid(row=5, column=0, columnspan=2, pady=10)

# Connect to the database and show students on startup
connect_db()
show_students()

# Start the Tkinter event loop
root.mainloop()
