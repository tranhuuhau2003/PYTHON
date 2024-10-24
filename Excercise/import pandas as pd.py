import pandas as pd

# Đọc file Excel
df = pd.read_excel('Excercise\diem-danh-sinh-vien-04102024094447.xls', header=[12,13])

# Gộp các tiêu đề cột lại thành một dòng duy nhất
df.columns = [' '.join([str(c) for c in col]).strip() for col in df.columns.values]

# Hiển thị dữ liệu
print(df.head(15))
