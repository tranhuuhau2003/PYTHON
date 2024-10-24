import sqlite3

class DatabaseManager:
    def __init__(self):
        self.conn = sqlite3.connect('students.db')
        self.cursor = self.conn.cursor()
        self.setup_database()

    def setup_database(self):
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
            dot TEXT,
            ma_lop TEXT,
            ten_mon_hoc TEXT
        )
        """)
        self.cursor.execute('''CREATE TABLE IF NOT EXISTS parents (mssv TEXT PRIMARY KEY, email_ph TEXT)''')
        self.cursor.execute('''CREATE TABLE IF NOT EXISTS teachers (mssv TEXT PRIMARY KEY, email_gvcn TEXT)''')
        self.conn.commit()

    def add_parent_and_teacher_emails(self, mssv_list):
        for mssv in mssv_list:
            email_ph = f"{mssv}@example.com"
            self.cursor.execute('INSERT OR IGNORE INTO parents (mssv, email_ph) VALUES (?, ?)', (mssv, email_ph))
            email_gvcn = f"gvcn_{mssv}@example.com"
            self.cursor.execute('INSERT OR IGNORE INTO teachers (mssv, email_gvcn) VALUES (?, ?)', (mssv, email_gvcn))
        self.conn.commit()

    def insert_student_data(self, student_data):
        self.cursor.execute(""" 
            INSERT OR IGNORE INTO students (mssv, ho_dem, ten, gioi_tinh, ngay_sinh, vắng_có_phép, vắng_không_phép, 
                                            tong_so_tiet, ty_le_vang, dot, ma_lop, ten_mon_hoc)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, student_data)
        self.conn.commit()

    def fetch_students(self):
        self.cursor.execute("""
            SELECT mssv, ho_dem, ten, dot, ma_lop, ten_mon_hoc, (vắng_có_phép + vắng_không_phép) as tong_vang
            FROM students
            ORDER BY tong_vang DESC
        """)
        return self.cursor.fetchall()

    def search_student(self, search_value):
        self.cursor.execute(f"SELECT * FROM students WHERE mssv LIKE ? OR ten LIKE ?", 
                            ('%' + search_value + '%', '%' + search_value + '%'))
        return self.cursor.fetchall()
