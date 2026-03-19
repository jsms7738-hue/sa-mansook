import pandas as pd
import os
import json

base_dir = r'c:\Users\yoonh\Desktop\AI\PBA 생산 결과'
file_path = os.path.join(base_dir, "PBA 실적.xlsx")

def final_analyze():
    try:
        xl = pd.ExcelFile(file_path)
        march_sheet = [s for s in xl.sheet_names if '3월' in s][0]
        df = pd.read_excel(xl, sheet_name=march_sheet, header=1)
        
        # Clean column names
        df.columns = [str(c).replace(' ◇', '').strip() for c in df.columns]
        
        # 0. Basic cleaning (remove rows with no production)
        df = df[df['생산수량'] > 0]
        
        # Filter out excluded lines
        excluded_lines = [
            'MFVB', 'MFVC', 'MPVA', 'MPVB', 'MPVD', 'MPVE', 'MPVH', 'MPVM', 'MPVN',
            'MPJ2', 'MPJ3', 'MPVC', 'MPVF', 'MPVG', 'MPVI', 'MPVK', 'MPVO'
        ]
        df = df[~df['조장라인명'].isin(excluded_lines)]
        
        analysis = {}
        
        # 1. Overall Summary
        total_qty = int(df['생산수량'].sum())
        avg_efficiency = float(df['작업공수효율'].mean())
        avg_prod_per_person = float(df['인당생산수'].mean())
        
        analysis['summary'] = {
            "total_production": total_qty,
            "avg_job_efficiency": round(avg_efficiency, 2),
            "avg_prod_per_person": round(avg_prod_per_person, 2)
        }
        
        # 2. Top 5 Lines by Production
        line_prod = df.groupby('조장라인명')['생산수량'].sum().sort_values(ascending=False).head(5)
        analysis['top_lines_by_qty'] = line_prod.to_dict()
        
        # 3. Top 5 Lines by Efficiency (with min production threshold to avoid outliers)
        line_eff = df.groupby('조장라인명').agg({'작업공수효율': 'mean', '생산수량': 'sum'})
        line_eff = line_eff[line_eff['생산수량'] > 1000].sort_values(by='작업공수효율', ascending=False).head(5)
        analysis['top_lines_by_efficiency'] = line_eff['작업공수효율'].to_dict()
        
        # 4. Top 5 Models by Production
        model_prod = df.groupby('모델명')['생산수량'].sum().sort_values(ascending=False).head(5)
        analysis['top_models_by_qty'] = model_prod.to_dict()
        
        # 5. Production Trend (Weekly)
        df['date'] = pd.to_datetime(df['작업일자'])
        df['week'] = df['date'].dt.isocalendar().week
        weekly_prod = df.groupby('week')['생산수량'].sum().to_dict()
        # Convert keys to strings to avoid uint32 error
        analysis['weekly_production'] = {str(k): int(v) for k, v in weekly_prod.items()}
        
        # 6. Peak Production Hours
        df['hour'] = pd.to_datetime(df['JobStrat'], errors='coerce').dt.hour
        hourly_prod = df.groupby('hour')['생산수량'].sum().sort_values(ascending=False).head(5)
        analysis['peak_hours'] = {str(k): int(v) for k, v in hourly_prod.to_dict().items()}
        
        with open('pba_final_analysis.json', 'w', encoding='utf-8') as f:
            json.dump(analysis, f, ensure_ascii=False, indent=2)
            
        print("Final analysis complete. Results saved to pba_final_analysis.json")
        
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    final_analyze()
