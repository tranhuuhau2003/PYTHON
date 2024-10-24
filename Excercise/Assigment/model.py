
import sqlite3
import pandas as pd

class StudentModel:
    def __init__(self, db_name="students.db"):
        self.conn = sqlite3.connect(db_name)
        self.cursor = self.conn.cursor()
        self.create_table()

    def create_table(self):
        self.cursor.execute("DROP TABLE IF EXISTS students")
        self.cursor.execute(
            ''' 
            CREATE TABLE IF NOT EXISTS students (
                mssv TEXT PRIMARY KEY,
                ho_dem TEXT,
                ten TEXT,
                gioi_tinh TEXT,
                ngay_sinh TEXT,
                vang_co_phep INTEGER,
                vang_khong_phep INTEGER,
                tong_so_tiet INTEGER,
                ty_le_vang REAL,
                dot TEXT
            )
            '''
        )
        self.conn.commit()

    def insert_student(self, student_data):
        self.cursor.execute(
            '''INSERT INTO students (mssv, ho_dem, ten, gioi_tinh, ngay_sinh, vang_co_phep, vang_khong_phep, tong_so_tiet, ty_le_vang, dot)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''', student_data)
        self.conn.commit()

    def fetch_all_students(self):
        self.cursor.execute("SELECT * FROM students")
        return self.cursor.fetchall()

    def close(self):
        self.conn.close()
    
    def load_from_excel(self, filepath):
        df = pd.read_excel(filepath)
        for index, row in df.iterrows():
            student_data = (
                row['Mã sinh viên'], row['Họ đệm'], row['Tên'], row['Giới tính'],
                row['Ngày sinh'], row['Vắng có phép'], row['Vắng không phép'],
                row['Tổng số tiết'], row['(%) vắng'], 'default'
            )
            self.insert_student(student_data)
