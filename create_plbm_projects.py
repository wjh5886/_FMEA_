"""
PLBM_30 AY1/JW1/KU1 프로젝트 생성 스크립트
- ARXML에서 SW Unit + 인터페이스 추출
- 각 차종별 프로젝트 생성 후 FMEA 기본 틀 삽입
"""
import sys, os, lxml.etree as ET, requests, urllib3, json
from collections import defaultdict
urllib3.disable_warnings()
sys.stdout.reconfigure(encoding='utf-8')

URL = "https://itzgdbeiyvodhfhmvrfw.supabase.co"
KEY = ("eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Iml0"
       "emdkYmVpeXZvZGhmaG12cmZ3Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzY3NDE2MzcsImV4cCI"
       "6MjA5MjMxNzYzN30.iqzQr-3Lqf1O4UHFe9euTLyeIyBXreLPoSzUdtEaNP8")
H  = {"apikey": KEY, "Authorization": f"Bearer {KEY}", "Content-Type": "application/json"}

ARXML_DIR = r"E:\claude\FMEA\PLBM_SV_V01_BSWv36-4\Configuration\System\Swcd_App"
FAILURE_MODES = ["MORE", "LESS", "CORRUPT", "EARLY", "LATE", "STUCK", "ERRATIC"]
VARIANTS = [("PLBM_SV1", "SV1")]

# ── ARXML 파싱 ────────────────────────────────────────────────
def parse_swcd(path):
    tree = ET.parse(path)
    root = tree.getroot()
    ns = root.tag.split('}')[0].lstrip('{')
    results = []
    for swc in root.iter(f'{{{ns}}}APPLICATION-SW-COMPONENT-TYPE'):
        swc_name = swc.findtext(f'{{{ns}}}SHORT-NAME') or '?'
        for tag, cat, direction in [
            (f'{{{ns}}}R-PORT-PROTOTYPE', 'External', 'R'),
            (f'{{{ns}}}P-PORT-PROTOTYPE', 'Internal', 'P'),
        ]:
            for port in swc.iter(tag):
                pname = port.findtext(f'{{{ns}}}SHORT-NAME')
                if not pname:
                    continue
                # 인터페이스 타입 추출
                for ref_tag in ['PROVIDED-REQUIRED-INTERFACE-TREF',
                                'REQUIRED-INTERFACE-TREF',
                                'PROVIDED-INTERFACE-TREF']:
                    iref = port.find(f'.//{{{ns}}}{ref_tag}')
                    if iref is not None:
                        break
                itype = iref.text.split('/')[-1] if iref is not None else ''
                results.append({'swc': swc_name, 'category': cat,
                                'name': pname, 'type': itype})
    return results

all_ports = []
for f in sorted(os.listdir(ARXML_DIR)):
    if f.endswith('.arxml'):
        all_ports.extend(parse_swcd(os.path.join(ARXML_DIR, f)))

by_swc = defaultdict(list)
for p in all_ports:
    by_swc[p['swc']].append(p)

print(f"파싱 완료 — SW Unit {len(by_swc)}개, 인터페이스 {len(all_ports)}개")
for swc, ports in by_swc.items():
    print(f"  {swc}: {len(ports)}개")

# ── Supabase 헬퍼 ─────────────────────────────────────────────
def post(table, data):
    r = requests.post(f"{URL}/rest/v1/{table}", json=data, headers=H, verify=False)
    if r.status_code >= 300:
        raise RuntimeError(f"{table} 삽입 실패: {r.text[:200]}")
    return r.json() if r.text else []

def delete_project(pid):
    for table in ['fmea_items', 'sw_units', 'projects']:
        requests.delete(f"{URL}/rest/v1/{table}?project_id=eq.{pid}", headers=H, verify=False)
    requests.delete(f"{URL}/rest/v1/projects?id=eq.{pid}", headers=H, verify=False)

# 기존 PLBM 프로젝트 확인
r = requests.get(f"{URL}/rest/v1/projects?name=like.PLBM_*&select=id,name", headers=H, verify=False)
existing = {p['name']: p['id'] for p in r.json()}
if existing:
    print(f"\n기존 PLBM 프로젝트: {list(existing.keys())} → 삭제 후 재생성")
    for pid in existing.values():
        delete_project(pid)

# ── 프로젝트 생성 ─────────────────────────────────────────────
print()
for proj_name, vehicle in VARIANTS:
    print(f"▶ {proj_name} ({vehicle})")

    # 1. 프로젝트 생성
    r = requests.post(
        f"{URL}/rest/v1/projects",
        json={"name": proj_name, "vehicle_model": vehicle,
              "description": f"PLBM FBL SW FMEA — {vehicle}"},
        headers={**H, "Prefer": "return=representation"},
        verify=False,
    )
    proj_id = r.json()[0]['id']
    print(f"  프로젝트 ID: {proj_id}")

    # 2. SW Unit 생성
    unit_id_map = {}
    for swc_name in by_swc:
        r = requests.post(
            f"{URL}/rest/v1/sw_units",
            json={"project_id": proj_id, "name": swc_name},
            headers={**H, "Prefer": "return=representation"},
            verify=False,
        )
        unit_id_map[swc_name] = r.json()[0]['id']
    print(f"  SW Unit {len(unit_id_map)}개 생성")

    # 3. FMEA 항목 생성 (인터페이스 × Failure Mode)
    items = []
    item_no = 1
    for swc_name, ports in by_swc.items():
        unit_id = unit_id_map[swc_name]
        for port in ports:
            for fm in FAILURE_MODES:
                items.append({
                    "project_id": proj_id,
                    "sw_unit_id": unit_id,
                    "item_no": f"{item_no}.0",
                    "category": port['category'],
                    "variable_name": port['name'],
                    "variable_type": port['type'] or None,
                    "failure_mode": fm,
                    "status": "draft",
                    "ai_generated": False,
                })
            item_no += 1

    # 배치 삽입 (100개씩)
    for i in range(0, len(items), 100):
        batch = items[i:i+100]
        r = requests.post(f"{URL}/rest/v1/fmea_items", json=batch, headers=H, verify=False)
        if r.status_code >= 300:
            print(f"  ⚠ 항목 삽입 오류: {r.text[:100]}")
            break
    print(f"  FMEA 항목 {len(items)}개 생성 ({len(all_ports)}개 인터페이스 × {len(FAILURE_MODES)}개 FM)")

print("\n✓ 완료")
