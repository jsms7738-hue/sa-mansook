import pandas as pd
import os
import sys

# Ensure UTF-8 output
sys.stdout.reconfigure(encoding='utf-8')

base_dir = r'c:\Users\yoonh\Desktop\AI\PBA 생산 결과'
file_path = os.path.join(base_dir, "PBA 실적.xlsx")

def check_columns():
    try:
        xl = pd.ExcelFile(file_path)
        sheet_name = [s for s in xl.sheet_names if '3월' in s][0]
        df = pd.read_excel(xl, sheet_name=sheet_name, header=1, nrows=5)
        print(f"Sheet: {sheet_name}")
        for i, col in enumerate(df.columns):
            print(f"{i}: {col}")
        
        # Also print a sample row to see data
        print("\nSample Data (First 2 rows):")
        print(df.head(2).to_string())
        
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    check_columns()
