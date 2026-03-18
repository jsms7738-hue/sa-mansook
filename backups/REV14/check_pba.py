import json, os

# Check PBA data.json
with open('data/data.json', 'r', encoding='utf-8') as f:
    d = json.load(f)

print('=== PBA DATA CHECK ===')
print(f'Generated at: {d.get("generated_at")}')
print(f'Overall total: {d["overall"]["total"]}')
print(f'Overall efficiency: {d["overall"]["efficiency"]}%')
print(f'Lines: {[l["line_id"] for l in d["lines"]]}')
print(f'Daily averages count: {len(d["daily_averages"])}')
if d["daily_averages"]:
    print(f'Date range: {d["daily_averages"][0]["date"]} ~ {d["daily_averages"][-1]["date"]}')
print()

# Check each line
for line in d['lines']:
    lid = line['line_id']
    trend = line.get('trend', [])
    models = line.get('model_stats', [])
    ht = line.get('hourly_trend', [])
    print(f'  Line {lid}: total={line["total"]}, trend_days={len(trend)}, hourly_slots={len(ht)}, models={len(models)}')
    if trend:
        print(f'    Date range: {trend[0]["date"]} ~ {trend[-1]["date"]}')

print()
print('=== POTENTIAL ISSUES ===')
# Check for lines with 0 data
for line in d['lines']:
    if line['total'] == 0:
        print(f'WARNING: Line {line["line_id"]} has 0 total production!')
    if len(line.get('trend', [])) == 0:
        print(f'WARNING: Line {line["line_id"]} has no trend data!')
    if line.get('efficiency', 0) == 0:
        print(f'WARNING: Line {line["line_id"]} has 0% efficiency!')

# Check for model_stats bug (space in key name)
print()
print('=== KEY ISSUE CHECK in model_stats ===')
for line in d['lines']:
    for ms in line.get('model_stats', []):
        keys = list(ms.keys())
        if ' efficiencies' in keys:
            print(f'BUG FOUND: Line {line["line_id"]}, model {ms["name"]} has key " efficiencies" (leading space!)')
        if 'efficiency' not in keys:
            print(f'WARNING: Line {line["line_id"]}, model {ms["name"]} missing "efficiency" key. Keys: {keys}')
