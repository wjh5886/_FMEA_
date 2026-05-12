"""
ARXML 파싱 → SX3 FMEA 자동 생성
GenDoc 도구 대신 직접 ARXML을 파싱하여 인터페이스 추출
"""

import xml.etree.ElementTree as ET
import re, os, requests, urllib3, json
from pathlib import Path
from collections import defaultdict

urllib3.disable_warnings()

# ─── 설정 ───────────────────────────────────────────────────────────────────
SX3_DIR   = Path("E:/HKMC_SX3_SBW_ICE_R44-4.22.00_20251117_T037_Proto_2026-04-27 200109")
SWCD_DIR  = SX3_DIR / "Configuration/System/Swcd_App"
INTF_DIR  = SX3_DIR / "Configuration/System"
DBC_DIR   = SX3_DIR / "Configuration/System/DBImport"

SUPABASE_URL = "https://itzgdbeiyvodhfhmvrfw.supabase.co"
SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Iml0emdkYmVpeXZvZGhmaG12cmZ3Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzY3NDE2MzcsImV4cCI6MjA5MjMxNzYzN30.iqzQr-3Lqf1O4UHFe9euTLyeIyBXreLPoSzUdtEaNP8"
SX3_PROJECT_ID = "a43d7d4e-b104-4e90-ab55-04c7b31aa3e7"

SB_H = {
    "apikey": SUPABASE_KEY,
    "Authorization": f"Bearer {SUPABASE_KEY}",
    "Content-Type": "application/json",
    "Prefer": "return=representation",
}

NS = {"ar": "http://autosar.org/schema/r4.0"}

# FMEA 제외 컴포넌트 (진단/스냅샷/DID)
SKIP_COMP = {"Demo"}
SKIP_PREFIX = {"DID", "Dcm", "Snapot", "SnapShot", "Snapshot", "NoPub", "PubDtc", "DtcEn"}

# 외부 인터페이스 판별
EXTERNAL_SRC = {"BswIF_CtIoHwAb_IntfIn", "IoHwAb", "PCAN", "GCAN"}
EXTERNAL_COMP = {"CtIoHwAb_IntfIn"}

MODE_BY_TYPE = {
    "uint8":   ["MORE","LESS","CORRUPT","STUCK"],
    "uint16":  ["MORE","LESS","CORRUPT","STUCK","ERRATIC"],
    "uint32":  ["MORE","LESS","CORRUPT","STUCK"],
    "sint8":   ["MORE","LESS","CORRUPT","STUCK"],
    "sint16":  ["MORE","LESS","CORRUPT","STUCK"],
    "boolean": ["MORE","LESS","STUCK"],
    "default": ["MORE","LESS","CORRUPT","STUCK"],
}

# ─── ARXML 파싱 헬퍼 ────────────────────────────────────────────────────────
def find_text(elem, path):
    e = elem.find(path, NS)
    return e.text.strip() if e is not None and e.text else ""

def parse_arxml(fpath):
    try:
        tree = ET.parse(fpath)
        return tree.getroot()
    except Exception as e:
        print(f"  파싱 오류 {fpath.name}: {e}")
        return None

# ─── 1단계: 인터페이스 타입 맵 구축 ─────────────────────────────────────────
def build_interface_type_map():
    """모든 ARXML에서 인터페이스명 → 데이터타입 매핑 구축"""
    iface_type = {}   # interface_ref → data_type
    iface_de   = {}   # interface_ref → de_name (variable name)

    arxml_files = list(SWCD_DIR.glob("*.arxml")) + \
                  list((SX3_DIR / "Configuration/System").glob("*.arxml")) + \
                  list(DBC_DIR.glob("*.arxml"))

    for fpath in arxml_files:
        root = parse_arxml(fpath)
        if root is None:
            continue

        # SENDER-RECEIVER-INTERFACE 순회
        for sri in root.iter("{http://autosar.org/schema/r4.0}SENDER-RECEIVER-INTERFACE"):
            pkg_path = _get_package_path(root, sri)
            sname = find_text(sri, "ar:SHORT-NAME")
            full_ref = f"{pkg_path}/{sname}" if pkg_path else sname

            for vdp in sri.iter("{http://autosar.org/schema/r4.0}VARIABLE-DATA-PROTOTYPE"):
                de_name  = find_text(vdp, "ar:SHORT-NAME")
                type_ref = find_text(vdp, "ar:TYPE-TREF")
                dtype    = type_ref.split("/")[-1].lower() if type_ref else "uint8"
                iface_type[full_ref] = dtype
                iface_de[full_ref]   = de_name

    return iface_type, iface_de

def _get_package_path(root, target_elem):
    """ARXML 요소의 패키지 경로 반환"""
    paths = []
    for pkg in root.iter("{http://autosar.org/schema/r4.0}AR-PACKAGE"):
        sname = find_text(pkg, "ar:SHORT-NAME")
        if sname:
            paths.append(sname)
    return "/".join(paths[:2]) if paths else ""

# ─── 2단계: SW Component 인터페이스 추출 ─────────────────────────────────────
def extract_swc_interfaces(iface_type, iface_de):
    """Swcd_App ARXML에서 SW Component별 포트(인터페이스) 추출"""
    interfaces = []
    seen = set()

    for fpath in sorted(SWCD_DIR.glob("*.arxml")):
        root = parse_arxml(fpath)
        if root is None:
            continue

        # APPLICATION-SW-COMPONENT-TYPE, COMPOSITION-SW-COMPONENT-TYPE 모두 처리
        for comp_tag in ["APPLICATION-SW-COMPONENT-TYPE", "COMPOSITION-SW-COMPONENT-TYPE"]:
            for swc in root.iter(f"{{http://autosar.org/schema/r4.0}}{comp_tag}"):
                comp_name = find_text(swc, "ar:SHORT-NAME")

                # 제외 컴포넌트 필터
                if comp_name in SKIP_COMP:
                    continue
                if any(comp_name.startswith(p) for p in SKIP_PREFIX):
                    continue

                # R-PORT (수신 = 입력 인터페이스)
                for rport in swc.iter("{http://autosar.org/schema/r4.0}R-PORT-PROTOTYPE"):
                    port_name = find_text(rport, "ar:SHORT-NAME")
                    iface_ref = find_text(rport, "ar:REQUIRED-INTERFACE-TREF")

                    de_name, dtype, category = _resolve_interface(
                        port_name, iface_ref, iface_de, iface_type, is_rx=True
                    )
                    if not de_name:
                        continue

                    key = (comp_name, de_name, "R")
                    if key in seen:
                        continue
                    seen.add(key)

                    interfaces.append({
                        "sw_unit":   comp_name,
                        "variable":  de_name,
                        "type":      dtype,
                        "category":  category,
                        "port_name": port_name,
                        "direction": "R",
                    })

                # P-PORT (송신 = 출력 인터페이스) - SBWSigSet, TxCAN 같은 출력 컴포넌트
                for pport in swc.iter("{http://autosar.org/schema/r4.0}P-PORT-PROTOTYPE"):
                    port_name = find_text(pport, "ar:SHORT-NAME")
                    iface_ref = find_text(pport, "ar:PROVIDED-INTERFACE-TREF")

                    # 출력은 특정 컴포넌트만 포함
                    if not any(k in comp_name for k in ["TxMainCAN","TxSubCAN","SBWSig","TxCAN"]):
                        continue

                    de_name, dtype, _ = _resolve_interface(
                        port_name, iface_ref, iface_de, iface_type, is_rx=False
                    )
                    if not de_name:
                        continue

                    key = (comp_name, de_name, "P")
                    if key in seen:
                        continue
                    seen.add(key)

                    interfaces.append({
                        "sw_unit":   comp_name,
                        "variable":  de_name,
                        "type":      dtype,
                        "category":  "Internal",
                        "port_name": port_name,
                        "direction": "P",
                    })

    return interfaces

def _resolve_interface(port_name, iface_ref, iface_de, iface_type, is_rx):
    """인터페이스 레퍼런스에서 신호명, 타입, 카테고리 결정"""
    # 인터페이스 레퍼런스에서 데이터 타입 조회
    dtype = "uint8"
    de_name = ""

    for ref_key in iface_type:
        if iface_ref and iface_ref.split("/")[-1] in ref_key:
            dtype = iface_type[ref_key]
            de_name = iface_de.get(ref_key, "")
            break

    # de_name이 없으면 port_name에서 De 이후 추출
    if not de_name:
        m = re.search(r'_De([A-Za-z0-9_]+)$', port_name)
        de_name = m.group(1) if m else port_name.split("_")[-1]

    # De prefix 제거
    if de_name.startswith("De") and len(de_name) > 2:
        de_name = de_name[2:]

    # 카테고리 결정
    category = "Internal"
    if is_rx:
        for ext in EXTERNAL_SRC:
            if iface_ref and ext in iface_ref:
                category = "External"
                break
        for ext in EXTERNAL_SRC:
            if ext in port_name:
                category = "External"
                break

    return de_name, dtype, category

# ─── 3단계: CAN 신호 추가 ────────────────────────────────────────────────────
def extract_can_signals():
    """PCAN.arxml, GCAN.arxml에서 CAN 신호 직접 추출"""
    can_signals = []
    seen = set()

    for fpath in [DBC_DIR / "PCAN.arxml", DBC_DIR / "GCAN.arxml"]:
        if not fpath.exists():
            continue
        root = parse_arxml(fpath)
        if root is None:
            continue

        # I-SIGNAL-TO-I-PDU-MAPPING에서 신호 추출
        for sig in root.iter("{http://autosar.org/schema/r4.0}I-SIGNAL"):
            sname = find_text(sig, "ar:SHORT-NAME")
            # DLC/길이로 타입 추정
            dl = sig.find("ar:LENGTH", NS)
            bit_len = int(dl.text) if dl is not None and dl.text else 8
            dtype = "uint8" if bit_len <= 8 else ("uint16" if bit_len <= 16 else "uint32")

            if sname and sname not in seen:
                seen.add(sname)
                can_signals.append({
                    "sw_unit": "RxMainCAN",
                    "variable": sname,
                    "type": dtype,
                    "category": "External",
                    "port_name": sname,
                    "direction": "R",
                })

    return can_signals[:50]  # CAN 신호 최대 50개 (너무 많아지지 않도록)

# ─── 4단계: FMEA 행 생성 ─────────────────────────────────────────────────────
def build_fmea_rows(interfaces):
    rows = []
    item_counter = defaultdict(int)

    for sig in interfaces:
        dk = sig['type'] if sig['type'] in MODE_BY_TYPE else 'default'
        modes = list(MODE_BY_TYPE[dk])

        # CAN 관련 신호: LATE/EARLY 추가
        if "CAN" in sig['sw_unit'] or "CAN" in sig.get('port_name','') or "Msg" in sig['variable']:
            for m in ["LATE","EARLY"]:
                if m not in modes:
                    modes.append(m)

        # 같은 variable에 여러 SW Unit이 있는 경우 → item_no 공유
        item_no = str(len(rows) // 4 + 1)

        for mode in modes:
            rows.append({
                "item_no":   item_no,
                "sw_unit":   sig['sw_unit'],
                "variable":  sig['variable'],
                "type":      sig['type'],
                "category":  sig['category'],
                "failure_mode": mode,
            })

    return rows

# ─── 5단계: 기존 데이터 참조 매칭 ────────────────────────────────────────────
def load_reference_data():
    """기존 LQ2/TK1/GN7_FL 데이터에서 S/O/D 참조값 로드"""
    SOURCE_IDS = {
        "GN7_FL": "32cc5148-96b5-42c5-a3dc-2b8fa3b9f93e",
        "LQ2":    "89dc5818-2435-4d09-a1a9-36aea664d11d",
        "TK1":    "70f74c19-66b6-4b2e-a2b3-1ee04dd1b101",
    }

    all_refs = []
    for name, pid in SOURCE_IDS.items():
        offset, batch = 0, 1000
        while True:
            r = requests.get(f"{SUPABASE_URL}/rest/v1/fmea_items",
                headers={**SB_H, "Range": f"{offset}-{offset+batch-1}"},
                params={"project_id": f"eq.{pid}",
                        "select": "variable_name,failure_mode,severity,occurrence,detection,preventive_action,detection_action,effect_module,effect_system"},
                verify=False)
            data = r.json()
            if not data: break
            filled = [d for d in data if d.get('severity') and d.get('occurrence') and d.get('detection')]
            all_refs.extend(filled)
            if len(data) < batch: break
            offset += batch

    def norm(name):
        if not name: return ""
        n = re.sub(r'\(.*?\)', '', str(name))
        n = n.split('\n')[0].strip()
        return re.sub(r'[^a-zA-Z0-9]', '', n).lower()

    # (norm_name, mode) → item
    ref_exact  = {}
    ref_byname = {}
    for item in all_refs:
        k1 = (norm(item['variable_name']), item.get('failure_mode',''))
        if k1 not in ref_exact:
            ref_exact[k1] = item
        k2 = norm(item['variable_name'])
        if k2 and k2 not in ref_byname:
            ref_byname[k2] = item

    print(f"  참조 데이터: {len(all_refs)}개 (키: {len(ref_exact)}개)")
    return ref_exact, ref_byname, norm

# ─── 6단계: Supabase 업데이트 ────────────────────────────────────────────────
def delete_existing_items(project_id):
    r = requests.delete(f"{SUPABASE_URL}/rest/v1/fmea_items",
        headers=SB_H, params={"project_id": f"eq.{project_id}"}, verify=False)
    print(f"  기존 항목 삭제: {r.status_code}")

def ensure_sw_units(unit_names, project_id):
    existing = requests.get(f"{SUPABASE_URL}/rest/v1/sw_units",
        headers=SB_H, params={"project_id": f"eq.{project_id}", "select":"id,name"}, verify=False).json()
    unit_map = {u["name"]: u["id"] for u in existing}
    for name in unit_names:
        if name and name not in unit_map:
            r = requests.post(f"{SUPABASE_URL}/rest/v1/sw_units",
                headers=SB_H, json={"project_id": project_id, "name": name}, verify=False)
            d = r.json()
            uid = (d[0] if isinstance(d, list) else d)["id"]
            unit_map[name] = uid
    return unit_map

def insert_batch(rows, batch=200):
    inserted = 0
    for i in range(0, len(rows), batch):
        chunk = rows[i:i+batch]
        r = requests.post(f"{SUPABASE_URL}/rest/v1/fmea_items",
            headers=SB_H, json=chunk, verify=False)
        if r.status_code in (200, 201):
            inserted += len(chunk)
        else:
            print(f"  !! 오류 {r.status_code}: {r.text[:150]}")
    return inserted

# ─── 메인 ──────────────────────────────────────────────────────────────────
def main():
    print("=" * 60)
    print("SX3 ARXML 파싱 → FMEA 자동 생성")
    print("=" * 60)

    # 1. 인터페이스 타입 맵
    print("\n[1/5] 인터페이스 타입 맵 구축...")
    iface_type, iface_de = build_interface_type_map()
    print(f"  인터페이스 {len(iface_type)}개 타입 로드")

    # 2. SW Component 인터페이스 추출
    print("\n[2/5] SW Component 인터페이스 추출...")
    interfaces = extract_swc_interfaces(iface_type, iface_de)
    print(f"  {len(interfaces)}개 인터페이스 추출")

    # 컴포넌트별 통계
    from collections import Counter
    cnt = Counter(i['sw_unit'] for i in interfaces)
    for k, v in sorted(cnt.items()):
        ext = sum(1 for i in interfaces if i['sw_unit']==k and i['category']=='External')
        print(f"    {k:30s}: {v}개 (External:{ext})")

    # 3. FMEA 행 생성
    print("\n[3/5] FMEA 항목 생성...")
    fmea_rows = build_fmea_rows(interfaces)
    print(f"  총 {len(fmea_rows)}개 FMEA 항목")

    # 4. 기존 데이터 참조 매칭
    print("\n[4/5] 기존 데이터 매칭...")
    ref_exact, ref_byname, norm = load_reference_data()

    matched = 0
    for row in fmea_rows:
        vn = norm(row['variable'])
        fm = row.get('failure_mode','')

        ref = ref_exact.get((vn, fm)) or ref_byname.get(vn)
        if ref:
            row['severity']          = ref.get('severity')
            row['occurrence']        = ref.get('occurrence')
            row['detection']         = ref.get('detection')
            row['preventive_action'] = ref.get('preventive_action')
            row['detection_action']  = ref.get('detection_action')
            row['effect_module']     = ref.get('effect_module')
            row['effect_system']     = ref.get('effect_system')
            row['ai_generated']      = False
            matched += 1
        else:
            row['severity']    = None
            row['occurrence']  = None
            row['detection']   = None
            row['ai_generated'] = False

    print(f"  매칭 성공: {matched}/{len(fmea_rows)}개")

    # 5. Supabase 저장
    print("\n[5/5] Supabase 저장...")
    delete_existing_items(SX3_PROJECT_ID)

    unit_names = list({r['sw_unit'] for r in fmea_rows})
    unit_map = ensure_sw_units(unit_names, SX3_PROJECT_ID)

    items = []
    for r in fmea_rows:
        s, o, d = r.get('severity'), r.get('occurrence'), r.get('detection')
        items.append({
            "project_id":     SX3_PROJECT_ID,
            "sw_unit_id":     unit_map.get(r['sw_unit']),
            "item_no":        r['item_no'],
            "category":       r['category'],
            "variable_name":  r['variable'],
            "variable_type":  r['type'],
            "failure_mode":   r['failure_mode'],
            "effect_module":  r.get('effect_module'),
            "effect_system":  r.get('effect_system'),
            "severity":       s,
            "occurrence":     o,
            "detection":      d,
            "preventive_action": r.get('preventive_action'),
            "detection_action":  r.get('detection_action'),
            "status":         "in_review" if (s and o and d) else "draft",
            "ai_generated":   False,
        })

    inserted = insert_batch(items)

    filled = sum(1 for i in items if i['severity'])
    print(f"\n{'='*60}")
    print(f"완료!")
    print(f"  SW Component: {len(unit_names)}개")
    print(f"  인터페이스:   {len(interfaces)}개")
    print(f"  FMEA 항목:    {inserted}/{len(items)}개 저장")
    print(f"  S/O/D 입력:   {filled}개 ({filled*100//len(items) if items else 0}%)")
    print(f"  URL: https://fmea-web.vercel.app/projects/{SX3_PROJECT_ID}/fmea")

if __name__ == "__main__":
    main()
