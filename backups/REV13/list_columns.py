import pandas as pd
import os

base_dir = r'c:\Users\yoonh\Desktop\AI\PBA 생산 결과'
file_path = os.path.join(base_dir, "PBA 실적.xlsx")

try:
    xl = pd.ExcelFile(file_path)
    march_sheet = [s for s in xl.sheet_names if '3월' in s][0]
    df = pd.read_excel(xl, sheet_name=march_sheet, header=1, nrows=1)
    print("Columns in March sheet:")
    for i, col in enumerate(df.columns):
        print(f"{i}: {col}")
except Exception as e:
    print(f"Error: {e}")
