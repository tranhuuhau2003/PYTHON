import pandas as pd

class ExcelLoader:
    def __init__(self, db_manager, filepath):
        self.db_manager = db_manager
        self.filepath = filepath

    def load_data_from_excel(self):
        try:
            df = pd.read_excel(self.filepath, engine='xlrd', header=None).fillna('')
            dot = df.iloc[5, 2]
            ma_lop = df.iloc[7, 2]
            ten_mon_hoc = df.iloc[8, 2]
            df_sinh_vien = df.iloc[11:]

            header1 = df_sinh_vien.iloc[0]
            header2 = df_sinh_vien.iloc[1]
            df_sinh_vien.columns = [
                f"{str(header1[i]).strip()}_{str(header2[i]).strip()}" if header1[i] or header2[i] else ''
                for i in range(len(header1))
            ]
            df_sinh_vien = df_sinh_vien[2:]

            if '_(%) vắng' in df_sinh_vien.columns:
                df_sinh_vien['_(%) vắng'] = df_sinh_vien['_(%) vắng'].apply(
                    lambda x: str(x).replace(',', '.') if isinstance(x, str) else x)

            for index, row in df_sinh_vien.iterrows():
                student_data = (
                    row['Mã sinh viên_'], row['Họ đệm_'], row['Tên_'], row['Giới tính_'], str(row['Ngày sinh_']),
                    int(row['Tổng cộng_Vắng có phép'] or 0), int(row['_Vắng không phép'] or 0),
                    int(row['_Tổng số tiết'] or 0), float(row['_(%) vắng'].replace(',', '.') or 0.0),
                    dot, ma_lop, ten_mon_hoc
                )
                self.db_manager.insert_student_data(student_data)
        except Exception as e:
            print(f"Không thể tải file: {e}")
