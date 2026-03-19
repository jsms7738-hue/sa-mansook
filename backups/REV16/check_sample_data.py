import pandas as pd
import os
import json

base_dir = r'c:\Users\yoonh\Desktop\AI\PBA 생산 결과'
file_path = os.path.join(base_dir, "PBA 실적.xlsx")

class CustomEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, pd.Timestamp):
            return obj.strftime('%Y-%m-%d %H:%M:%S')
        if pd.isna(obj):
            return None
        return super().default(obj)

def check_data():
    try:
        xl = pd.ExcelFile(file_path)
        march_sheet = [s for s in xl.sheet_names if '3월' in s][0]
        df = pd.read_excel(xl, sheet_name=march_sheet, header=1, nrows=10)
        
        data_subset = df.iloc[:, :25].copy()
        
        results = []
        # Add column names as the first row
        results.append(data_subset.columns.to_list())
        for i, row in data_subset.iterrows():
            results.append(row.to_list())
            
        with open('pba_sample_rows.json', 'w', encoding='utf-8') as f:
            json.dump(results, f, ensure_ascii=False, indent=2, cls=CustomEncoder)
            
        print("Sample data saved to pba_sample_rows.json")
        
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    check_data()
