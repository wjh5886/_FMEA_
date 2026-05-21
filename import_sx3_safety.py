"""
SX3_SBW_Safety_Mechanism.xlsx → Supabase
SX3 포함 프로젝트 전체에 Safety Goal & Safety Mechanism 업서트
"""
import requests, urllib3, sys, openpyxl
urllib3.disable_warnings()
sys.stdout.reconfigure(encoding='utf-8')

URL = "https://itzgdbeiyvodhfhmvrfw.supabase.co"
KEY = ("eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Iml0"
       "emdkYmVpeXZvZGhmaG12cmZ3Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzY3NDE2MzcsImV4cCI"
       "6MjA5MjMxNzYzN30.iqzQr-3Lqf1O4UHFe9euTLyeIyBXreLPoSzUdtEaNP8")
H = {"apikey": KEY, "Authorization": f"Bearer {KEY}", "Content-Type": "application/json"}

SX3_PROJECTS = [
    ("a43d7d4e-b104-4e90-ab55-04c7b31aa3e7", "SX3"),
    ("6a4bf862-070d-4e24-af0e-4f36183be745", "SX3_ARXML"),
    ("1ba10b41-4717-4d1b-b0df-1b91fe2c4870", "SX3_ICE_TEST"),
    ("61e184f1-6f0b-469e-8241-1168b217a27c", "SX3_OPT_TEST"),
    ("3f961375-b1ff-412c-901f-1e34ac809f04", "SX3_TEST_AUTO"),
    ("e552c9b4-e392-48cc-b6b0-771caea6e0cd", "SX3_TEST_WEB"),
]

EXCEL = r"E:\claude\FMEA\SX3_SBW_Safety_Mechanism.xlsx"

# ── 엑셀 파싱 ─────────────────────────────────────────────────
wb = openpyxl.load_workbook(EXCEL, data_only=True)

# Safety Goals (SG01~SG05)
import re
sg_rows = []
for row in wb["Safety Goal & Safe State"].iter_rows(values_only=True):
    if row[1] and re.match(r"^SG\d+$", str(row[1]).strip()):
        sg_id    = str(row[1]).strip()
        desc     = str(row[2]).strip().replace("\xa0", "") if row[2] else ""
        asil_raw = str(row[3]).strip() if row[3] else ""
        # "ASIL A" → "A", "A" → "A", 그 외 → None
        m = re.search(r"\b([ABCDQM]+)\b", asil_raw)
        asil = m.group(1) if m and m.group(1) in ("QM","A","B","C","D") else None
        sg_rows.append({"sg_id": sg_id, "name": sg_id, "description": desc, "asil": asil})

# Safety Mechanisms (SM##)
sm_rows = []
for row in wb["Safety mechanism"].iter_rows(min_row=3, values_only=True):
    if not (row[0] and str(row[0]).strip().startswith("SM")):
        continue
    sm_id   = str(row[0]).strip()
    name    = str(row[1]).strip().replace("\n", " ") if row[1] else ""
    desc    = str(row[2]).strip() if row[2] else ""
    impl    = str(row[4]).strip() if row[4] else ""
    sm_rows.append({
        "sm_id": sm_id,
        "name": name,
        "description": desc,
        "type": "Both" if "HW" in impl and "SW" in impl else
                "Detection" if impl == "SW" else
                "Preventive" if impl == "HW" else None,
        "diagnostic_coverage": None,
        "related_sg_id": None,
    })

print(f"파싱 완료 — Safety Goal {len(sg_rows)}개, Safety Mechanism {len(sm_rows)}개")
print("Safety Goals:", [r["sg_id"] for r in sg_rows])
print("Safety Mechanisms:", [r["sm_id"] for r in sm_rows])

# ── Supabase 업서트 ────────────────────────────────────────────
def delete_existing(table, proj_id):
    requests.delete(f"{URL}/rest/v1/{table}?project_id=eq.{proj_id}", headers=H, verify=False)

def insert(table, rows):
    res = requests.post(f"{URL}/rest/v1/{table}", json=rows, headers=H, verify=False)
    return res.status_code, res.text[:200] if res.status_code >= 300 else f"{len(rows)}개 완료"

for proj_id, proj_name in SX3_PROJECTS:
    print(f"\n▶ {proj_name} ({proj_id[:8]}...)")

    delete_existing("safety_goals", proj_id)
    code, msg = insert("safety_goals", [{**r, "project_id": proj_id} for r in sg_rows])
    print(f"  SG: {code} {msg}")

    delete_existing("safety_mechanisms", proj_id)
    code, msg = insert("safety_mechanisms", [{**r, "project_id": proj_id} for r in sm_rows])
    print(f"  SM: {code} {msg}")

print("\n✓ 완료")
