import pandas as pd
import json
import os
import sys
from datetime import datetime

# Ensure UTF-8 output
sys.stdout.reconfigure(encoding='utf-8')

# Get the directory where the script is located
base_dir = r'c:\Users\yoonh\Desktop\AI\PBA 생산 결과'
file_path = os.path.join(base_dir, "PBA 실적.xlsx")
output_dir = os.path.join(base_dir, "data")
if not os.path.exists(output_dir):
    os.makedirs(output_dir)
json_path = os.path.join(output_dir, "data.json")

def process_file():
    try:
        xl = pd.ExcelFile(file_path)
        print(f"DEBUG: All Excel sheets: {xl.sheet_names}")
        
        # Include all sheets that look like monthly data (e.g. 3월, 3월데이타, 4월, 4월데이타)
        import re
        sheet_names = [s for s in xl.sheet_names if '월' in s or re.search(r'\d+', s)]
        print(f"DEBUG: Target sheet_names: {sheet_names}")
        
        overall_total = 0
        lines_summary = {} 
        daily_averages_map = {} 

        # We'll use a map to keep track of global model stats for the summary table
        global_model_stats = {}

        for sheet in sheet_names:
            print(f"Processing sheet: {sheet}")
            df = pd.read_excel(xl, sheet_name=sheet, header=1)
            
            # Drop rows where essential data is missing
            df = df.dropna(subset=['작업일자 ◇', '생산수량 ◇'])

            for _, row in df.iterrows():
                try:
                    # Date formatting
                    date_val = row['작업일자 ◇']
                    if isinstance(date_val, str):
                        date_obj = pd.to_datetime(date_val)
                    else:
                        date_obj = date_val
                    
                    display_date = f"{date_obj.month}월 {date_obj.day}일"
                    
                    line_id = str(row['조장라인명 ◇']) if pd.notna(row['조장라인명 ◇']) else "Unknown"
                    
                    excluded_lines = [
                        'MFVB', 'MFVC', 'MPVA', 'MPVB', 'MPVD', 'MPVE', 'MPVH', 'MPVM', 'MPVN',
                        'MPJ2', 'MPJ3', 'MPVC', 'MPVF', 'MPVG', 'MPVI', 'MPVK', 'MPVO'
                    ]
                    if line_id in excluded_lines:
                        continue
                        
                    model_name = str(row['모델명 ◇']) if pd.notna(row['모델명 ◇']) else "Unknown"
                    product_code = str(row['완제품코드']) if pd.notna(row['완제품코드']) else ""
                    model_st = float(row['모델st ◇']) if pd.notna(row['모델st ◇']) else 0.0
                    prod_qty = int(row['생산수량 ◇'])
                    
                    overall_total += prod_qty
                    
                    # Global model aggregation
                    if model_name not in global_model_stats:
                        global_model_stats[model_name] = {
                            "name": model_name,
                            "code": product_code,
                            "st": model_st,
                            "total": 0
                        }
                    global_model_stats[model_name]["total"] += prod_qty
                    # Ensure code and st are updated if they were blank before
                    if not global_model_stats[model_name]["code"] and product_code:
                        global_model_stats[model_name]["code"] = product_code
                    if global_model_stats[model_name]["st"] == 0 and model_st:
                        global_model_stats[model_name]["st"] = model_st

                    # Daily stats for overview
                    if display_date not in daily_averages_map:
                        daily_averages_map[display_date] = {
                            "total_qty": 0,
                            "efficiencies": [],
                            "manpower_list": [],
                            "prod_per_person_list": [],
                            "loss_rates": []
                        }
                    day_map = daily_averages_map[display_date]
                    day_map["total_qty"] += prod_qty
                    
                    # Store values for averaging
                    if pd.notna(row['작업공수효율 ◇']):
                        day_map["efficiencies"].append(float(row['작업공수효율 ◇']))
                    if pd.notna(row['투입인원 ◇']):
                        day_map["manpower_list"].append(int(row['투입인원 ◇']))
                    if pd.notna(row['인당생산수 ◇']):
                        day_map["prod_per_person_list"].append(float(row['인당생산수 ◇']))
                    if pd.notna(row['유실율 ◇']):
                        day_map["loss_rates"].append(float(row['유실율 ◇']))
                    
                    # Line stats
                    if line_id not in lines_summary:
                        lines_summary[line_id] = {
                            "line_id": line_id,
                            "total": 0,
                            "pass": 0, # PBA data doesn't seem to have pass/fail, defaulting to total for dashboard compatibility
                            "fail": 0,
                            "pass_rate": 100.0,
                            "efficiencies": [],
                            "models": set(),
                            "model_stats": {},
                            "trend": {}, # {date: {total, pass, fail, efficiencies}}
                            "hourly_trend": {} # {date_slot: {total, pass, fail, efficiencies}}
                        }
                    
                    ls = lines_summary[line_id]
                    ls["total"] += prod_qty
                    ls["pass"] += prod_qty
                    ls["models"].add(model_name)
                    if pd.notna(row['작업공수효율 ◇']):
                        ls["efficiencies"].append(float(row['작업공수효율 ◇']))
                    
                    # Model stats per line
                    if model_name not in ls["model_stats"]:
                        ls["model_stats"][model_name] = {"total": 0, "pass": 0, "fail": 0, " efficiencies": [], "code": product_code, "st": model_st, "manpowers": []}
                    ls["model_stats"][model_name]["total"] += prod_qty
                    ls["model_stats"][model_name]["pass"] += prod_qty
                    if pd.notna(row['작업공수효율 ◇']):
                        ls["model_stats"][model_name][" efficiencies"].append(float(row['작업공수효율 ◇']))
                    if pd.notna(row['투입인원 ◇']):
                        ls["model_stats"][model_name]["manpowers"].append(int(row['투입인원 ◇']))
                    if not ls["model_stats"][model_name]["code"] and product_code:
                        ls["model_stats"][model_name]["code"] = product_code
                    if ls["model_stats"][model_name]["st"] == 0 and model_st:
                        ls["model_stats"][model_name]["st"] = model_st
                    
                    # Trend
                    if display_date not in ls["trend"]:
                        ls["trend"][display_date] = {"total": 0, "pass": 0, "fail": 0, "efficiencies": [], "model_stats": {}, "std_labor": 0.0, "actual_labor": 0.0}
                    ls["trend"][display_date]["total"] += prod_qty
                    ls["trend"][display_date]["pass"] += prod_qty
                    if pd.notna(row['작업공수효율 ◇']):
                        ls["trend"][display_date]["efficiencies"].append(float(row['작업공수효율 ◇']))
                    if pd.notna(row['표준공수 ◇']):
                        ls["trend"][display_date]["std_labor"] += float(row['표준공수 ◇'])
                    if pd.notna(row['작업공수 ◇']):
                        ls["trend"][display_date]["actual_labor"] += float(row['작업공수 ◇'])
                    
                    # Model stats per date
                    if model_name not in ls["trend"][display_date]["model_stats"]:
                        ls["trend"][display_date]["model_stats"][model_name] = {"total": 0, "pass": 0, "fail": 0, " efficiencies": [], "code": product_code, "st": model_st, "manpowers": []}
                    dm_stats = ls["trend"][display_date]["model_stats"][model_name]
                    dm_stats["total"] += prod_qty
                    dm_stats["pass"] += prod_qty
                    if pd.notna(row['작업공수효율 ◇']):
                        dm_stats[" efficiencies"].append(float(row['작업공수효율 ◇']))
                    if pd.notna(row['투입인원 ◇']):
                        dm_stats["manpowers"].append(int(row['투입인원 ◇']))
                    if not dm_stats["code"] and product_code: dm_stats["code"] = product_code
                    if dm_stats["st"] == 0 and model_st: dm_stats["st"] = model_st
                    
                    # Hourly Trend (Using JobEnd ◇)
                    end_time = row['JobEnd ◇']
                    if pd.notna(end_time):
                        hour = end_time.hour
                        # Map to standard slots
                        slot = "기타"
                        if 6 <= hour < 8: slot = "A-1(06:00~07:59)"
                        elif 8 <= hour < 10: slot = "A(08:00~10:00)"
                        elif 10 <= hour < 12: slot = "B(10:10~12:00)"
                        elif 13 <= hour < 15: slot = "C(13:00~15:00)"
                        elif 15 <= hour < 17: slot = "D(15:10~17:00)"
                        elif 17 <= hour < 20: slot = "E(17:30~20:00)"
                        else: slot = f"기타({hour:02d}:00)"
                        
                        date_slot = f"{display_date} {slot}"
                        if date_slot not in ls["hourly_trend"]:
                            ls["hourly_trend"][date_slot] = {"total": 0, "pass": 0, "fail": 0, "efficiencies": []}
                        ls["hourly_trend"][date_slot]["total"] += prod_qty
                        ls["hourly_trend"][date_slot]["pass"] += prod_qty
                        if pd.notna(row['작업공수효율 ◇']):
                            ls["hourly_trend"][date_slot]["efficiencies"].append(float(row['작업공수효율 ◇']))
                        
                except Exception as e:
                    # print(f"Row error: {e}")
                    continue

        # Convert global_model_stats to list and sort by total production
        summary_models = []
        for mn in sorted(global_model_stats.keys(), key=lambda x: global_model_stats[x]["total"], reverse=True):
            summary_models.append(global_model_stats[mn])

        # Finalize and Sort
        def parse_date_key(d_str):
            parts = d_str.replace("월", "").replace("일", "").split(" ")
            return (int(parts[0]), int(parts[1]))

        daily_averages = []
        for d in sorted(daily_averages_map.keys(), key=parse_date_key):
            dm = daily_averages_map[d]
            daily_averages.append({
                "date": d,
                "qty": int(dm["total_qty"]),
                "efficiency": round(sum(dm["efficiencies"]) / len(dm["efficiencies"]), 2) if dm["efficiencies"] else 0,
                "manpower": round(sum(dm["manpower_list"]) / len(dm["manpower_list"]), 1) if dm["manpower_list"] else 0,
                "prod_per_person": round(sum(dm["prod_per_person_list"]) / len(dm["prod_per_person_list"]), 2) if dm["prod_per_person_list"] else 0,
                "loss_rate": round(sum(dm["loss_rates"]) / len(dm["loss_rates"]), 2) if dm["loss_rates"] else 0
            })

        final_lines = []
        for lid in sorted(lines_summary.keys()):
            data = lines_summary[lid]
            data["models"] = sorted(list(data["models"]))
            data["efficiency"] = round(sum(data["efficiencies"]) / len(data["efficiencies"]), 2) if data["efficiencies"] else 0
            del data["efficiencies"] # Clean up internal array
            
            # Finalize model_stats
            ms_raw = data["model_stats"]
            data["model_stats"] = []
            for mname in sorted(ms_raw.keys()):
                m = ms_raw[mname]
                data["model_stats"].append({
                    "name": mname,
                    "code": m.get("code", ""),
                    "st": m.get("st", 0.0),
                    "total": m["total"],
                    "pass": m["pass"],
                    "fail": m["fail"],
                    "pass_rate": 100.0,
                    "manpower": round(sum(m["manpowers"]) / len(m["manpowers"]), 1) if m.get("manpowers") else 0,
                    "efficiency": round(sum(m[" efficiencies"]) / len(m[" efficiencies"]), 2) if m.get(" efficiencies") else 0,
                    "worst1_step": ""
                })
            
            # Finalize trend
            trend_raw = data["trend"]
            trend_list = []
            for d in sorted(trend_raw.keys(), key=parse_date_key):
                t = trend_raw[d]
                
                # Finalize daily model stats
                d_ms_raw = t.get("model_stats", {})
                d_ms_list = []
                for mname in sorted(d_ms_raw.keys()):
                    m = d_ms_raw[mname]
                    d_ms_list.append({
                        "name": mname,
                        "code": m.get("code", ""),
                        "st": m.get("st", 0.0),
                        "total": m["total"],
                        "pass": m["pass"],
                        "fail": m["fail"],
                        "pass_rate": 100.0,
                        "manpower": round(sum(m["manpowers"]) / len(m["manpowers"]), 1) if m.get("manpowers") else 0,
                        "efficiency": round(sum(m[" efficiencies"]) / len(m[" efficiencies"]), 2) if m.get(" efficiencies") else 0,
                    })

                trend_list.append({
                    "date": d,
                    "total": t["total"],
                    "pass": t["pass"],
                    "fail": t["fail"],
                    "pass_rate": 100.0,
                    "std_labor": round(t.get("std_labor", 0), 2),
                    "actual_labor": round(t.get("actual_labor", 0), 2),
                    "efficiency": round(sum(t["efficiencies"]) / len(t["efficiencies"]), 2) if t.get("efficiencies") else 0,
                    "fail_details": {"models": {}, "steps": {}},
                    "model_stats": d_ms_list
                })
            data["trend"] = trend_list
            
            # Finalize hourly_trend
            def parse_slot_key(k):
                try:
                    parts = k.split(" ")
                    m, d = parse_date_key(f"{parts[0]} {parts[1]}")
                    slot_order = {"A-1": 1, "A": 2, "B": 3, "C": 4, "D": 5, "E": 6}
                    order = slot_order.get(parts[2].split("(")[0], 99)
                    return (m, d, order)
                except: return (99, 99, 99)

            ht_raw = data["hourly_trend"]
            ht_list = []
            for k in sorted(ht_raw.keys(), key=parse_slot_key):
                h = ht_raw[k]
                ht_list.append({
                    "time_slot": k,
                    "total": h["total"],
                    "pass": h["pass"],
                    "fail": h["fail"],
                    "pass_rate": 100.0,
                    "efficiency": round(sum(h["efficiencies"]) / len(h["efficiencies"]), 2) if h.get("efficiencies") else 0
                })
            data["hourly_trend"] = ht_list

            # Weekly Trend
            weekly_stats = {}
            for day_data in data["trend"]:
                day_int = int(day_data["date"].split(" ")[1].replace("일", ""))
                if day_int <= 7: w = "Week 1"
                elif day_int <= 14: w = "Week 2"
                elif day_int <= 21: w = "Week 3"
                else: w = "Week 4+"
                if w not in weekly_stats: weekly_stats[w] = {"total": 0, "pass": 0, "fail": 0, "days": []}
                weekly_stats[w]["total"] += day_data["total"]
                weekly_stats[w]["pass"] += day_data["pass"]
                weekly_stats[w]["days"].append(day_data)
            
            data["weekly_trend"] = []
            for w in sorted(weekly_stats.keys()):
                ws = weekly_stats[w]
                data["weekly_trend"].append({
                    "week": w, "total": ws["total"], "pass": ws["pass"], "fail": 0, "pass_rate": 100.0, "days": ws["days"]
                })

            final_lines.append(data)

        overall_total = sum([ls["total"] for ls in lines_summary.values()])
        overall_pass = sum([ls["pass"] for ls in lines_summary.values()])
        overall_fail = sum([ls["fail"] for ls in lines_summary.values()])
        
        # Calculate overall efficiency
        valid_efficiencies = [d["efficiency"] for d in daily_averages if d["efficiency"] > 0]
        overall_efficiency = round(sum(valid_efficiencies) / len(valid_efficiencies), 2) if valid_efficiencies else 0

        final_data = {
            "overall": {
                "total": overall_total,
                "pass": overall_pass,
                "fail": overall_fail,
                "pass_rate": 100.0,
                "efficiency": overall_efficiency,
                "summary_models": summary_models
            },
            "lines": final_lines,
            "daily_averages": daily_averages,
            "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }

        os.makedirs("data", exist_ok=True)
        with open("data/data.json", 'w', encoding='utf-8') as f:
            json.dump(final_data, f, ensure_ascii=False, indent=2)
            
        print("PBA Data extraction successful. JSON saved to: " + os.path.abspath("data/data.json"))
        
    except Exception as e:
        print(f"Error extracting PBA data: {e}")

if __name__ == "__main__":
    process_file()
