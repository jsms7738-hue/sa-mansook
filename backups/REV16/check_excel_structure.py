import pandas as pd
import os

base_dir = r'c:\Users\yoonh\Desktop\AI\PBA 생산 결과'
file_path = os.path.join(base_dir, "PBA 실적.xlsx")

try:
    xl = pd.ExcelFile(file_path)
    print(f"Sheets: {xl.sheet_names}")
    for sheet in xl.sheet_names:
        if '3월' in sheet:
            df = pd.read_excel(xl, sheet_name=sheet, header=1, nrows=5)
            print(f"\nSheet: {sheet}")
            print(f"Columns: {df.columns.tolist()}")
            print(f"Sample data:\n{df.head()}")
except Exception as e:
    print(f"Error: {e}")
