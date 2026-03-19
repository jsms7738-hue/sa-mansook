import pandas as pd
import os
import json

base_dir = r'c:\Users\yoonh\Desktop\AI\PBA 생산 결과'
file_path = os.path.join(base_dir, "PBA 실적.xlsx")

def analyze():
    try:
        xl = pd.ExcelFile(file_path)
        march_sheet = [s for s in xl.sheet_names if '3월' in s][0]
        df = pd.read_excel(xl, sheet_name=march_sheet, header=1)
        
        # Clean up columns (remove special characters if any)
        df.columns = [str(c).replace(' ◇', '').strip() for c in df.columns]
        
        analysis = {}
        
        # 1. Total Production
        total_qty = int(df['생산수량'].sum())
        analysis['total_production'] = total_qty
        
        # 2. Production by Line (Top 5)
        line_stats = df.groupby('실제생산라인')['생산수량'].sum().sort_values(ascending=False)
        analysis['top_lines'] = line_stats.head(5).to_dict()
        
        # 3. Production by Model (Top 5)
        model_stats = df.groupby('모델명')['생산수량'].sum().sort_values(ascending=False)
        analysis['top_models'] = model_stats.head(5).to_dict()
        
        # 4. Daily Trend
        df['작업일자'] = pd.to_datetime(df['작업일자'])
        daily_stats = df.groupby(df['작업일자'].dt.strftime('%m-%d'))['생산수량'].sum()
        analysis['daily_trend'] = daily_stats.to_dict()
        
        # 5. Hourly Trend (Peak Times)
        df['hour'] = pd.to_datetime(df['JobStrat'], errors='coerce').dt.hour
        hourly_stats = df.groupby('hour')['생산수량'].sum().sort_values(ascending=False)
        analysis['peak_hours'] = hourly_stats.to_dict()
        
        # 6. Check for Defect/Yield columns
        # Looking for columns like '불량', 'NG', '실적', '양품'
        potential_defect_cols = [c for c in df.columns if any(keyword in c for keyword in ['불량', 'NG', '실적', '양품', '수율', '직행'])]
        analysis['potential_kpi_columns'] = potential_defect_cols
        
        if '불량수' in df.columns:
            analysis['total_defects'] = int(df['불량수'].sum())
            analysis['overall_yield'] = (1 - (analysis['total_defects'] / total_qty)) * 100 if total_qty > 0 else 100
        
        with open('pba_analysis_results.json', 'w', encoding='utf-8') as f:
            json.dump(analysis, f, ensure_ascii=False, indent=2)
            
        print("Analysis complete. Results saved to pba_analysis_results.json")
        
    except Exception as e:
        print(f"Error during analysis: {e}")

if __name__ == "__main__":
    analyze()
