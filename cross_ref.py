import json, os
from collections import defaultdict

src_base = "E:/claude/FMEA/HKMC_JG1_SBW_R44_4.2109.00_26210_Mood_R1 1/Static_Code/Source"

with open("E:/claude/FMEA/fmea_data.json", encoding='utf-8') as f:
    records = json.load(f)

# Collect all C/H source files
c_files = []
for root, dirs, files in os.walk(src_base):
    for fn in files:
        if fn.endswith(('.c', '.h')):
            c_files.append(os.path.join(root, fn))

print(f"Total source files: {len(c_files)}", flush=True)

# Read all source files
all_src_content = {}
for fp in c_files:
    try:
        with open(fp, encoding='utf-8', errors='ignore') as f:
            all_src_content[fp] = f.read()
    except:
        pass

print(f"Files read: {len(all_src_content)}", flush=True)

# Get unique variables per unit
vars_by_unit = defaultdict(set)
for r in records:
    unit = r['SW_Unit'] or 'Unknown'
    var = r['Variable']
    if var:
        vname = var.split('\n')[0].strip().split('(')[0].strip()
        base = vname.split('.')[0].split('[')[0].strip()
        if len(base) >= 4:
            vars_by_unit[unit].add(base)

# Cross-reference
results = {}
for unit, var_set in sorted(vars_by_unit.items()):
    found = []
    not_found = []
    for var in sorted(var_set):
        matched_files = []
        for fp, content in all_src_content.items():
            if var in content:
                rel = fp.replace(src_base, '').replace('\\', '/').lstrip('/')
                matched_files.append(rel)
        if matched_files:
            found.append((var, matched_files))
        else:
            not_found.append(var)
    results[unit] = {'found': found, 'not_found': not_found}

# Print summary
print("\n" + "="*70)
for unit, res in sorted(results.items()):
    total = len(res['found']) + len(res['not_found'])
    pct = int(100*len(res['found'])/total) if total else 0
    print(f"\n[{unit}] Found: {len(res['found'])}/{total} ({pct}%)")
    if res['not_found']:
        print(f"  NOT IN CODE: {', '.join(res['not_found'])}")

with open("E:/claude/FMEA/xref_results.json", 'w', encoding='utf-8') as f:
    json.dump(results, f, ensure_ascii=False, indent=2)
print("\nSaved xref_results.json")
