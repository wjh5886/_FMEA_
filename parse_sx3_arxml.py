"""
SX3 ARXML → FMEA items 생성
- Configuration/System/Swcd_App/*.arxml  : SW 컴포넌트 + 포트 인터페이스 정의
- Configuration/System/Composition/RootComposition.arxml : 컴포넌트 간 연결
→ fmea_items 생성 후 Supabase 업로드 (기존 SX3 항목 교체)
"""

import xml.etree.ElementTree as ET
from pathlib import Path
from collections import defaultdict
import requests, urllib3, concurrent.futures

urllib3.disable_warnings()

# ── 경로 ──────────────────────────────────────────────────────────────
BASE = Path("E:/claude/FMEA/SBW_FMEA/SX3/"
            "HKMC_SX3_SBW_ICE_R44-4.22.00_20251117_T037_Proto_2026-04-27 200109")
SWCD_APP_DIR  = BASE / "Configuration/System/Swcd_App"
COMPOSITION   = BASE / "Configuration/System/Composition/RootComposition.arxml"
SOURCE_DIR    = BASE / "Static_Code/Source"

# ── Supabase ───────────────────────────────────────────────────────────
SX3_PROJECT_ID = "6a4bf862-070d-4e24-af0e-4f36183be745"  # SX3_ARXML (신규)
URL = "https://itzgdbeiyvodhfhmvrfw.supabase.co"
KEY = ("eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Iml0"
       "emdkYmVpeXZvZGhmaG12cmZ3Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzY3NDE2MzcsImV4cCI"
       "6MjA5MjMxNzYzN30.iqzQr-3Lqf1O4UHFe9euTLyeIyBXreLPoSzUdtEaNP8")
H = {"apikey": KEY, "Authorization": f"Bearer {KEY}",
     "Content-Type": "application/json", "Prefer": "return=representation"}

NS = "http://autosar.org/schema/r4.0"
NP = {"ar": NS}

# 실패모드 정의
MODES_SR = ["MORE", "LESS", "CORRUPT", "EARLY", "LATE", "STUCK", "ERRATIC"]
MODES_CS = ["CORRUPT", "EARLY", "LATE"]

DETAIL_TMPL = {
    "MORE":    "{var}이(가) 정상 범위를 초과하는 값을 출력/수신",
    "LESS":    "{var}이(가) 정상 범위 미만의 값을 출력/수신",
    "CORRUPT": "{var}이(가) 정상 범위 내이나 논리적으로 잘못된 값을 출력/수신",
    "EARLY":   "{var}이(가) 예상 시점보다 일찍 업데이트/발생됨",
    "LATE":    "{var}이(가) 예상 시점보다 늦게 업데이트됨 (타임아웃 가능성)",
    "STUCK":   "{var}이(가) 특정 값에 고착되어 변화하지 않음",
    "ERRATIC": "{var}이(가) 불규칙하게 값이 변동하거나 진동함",
}

# 외부(CAN) 수신 컴포넌트 — 이 컴포넌트의 P-PORT는 External 신호
CAN_RX_COMPONENTS = {"CtAp_RxMainCAN", "CtAp_RxSubCAN"}
# DCM/DEM/BSW는 FMEA 스코프 제외
SKIP_PREFIX = ("CtDcm_", "CtDem_", "CtWdgM", "CtNvM", "CtSbc",
               "CtSafetyLib", "CtAp_Demo", "CtIoHwAb", "CtCdd_")


# ── XML 헬퍼 ──────────────────────────────────────────────────────────
def sname(el):
    s = el.find("ar:SHORT-NAME", NP)
    return s.text.strip() if s is not None and s.text else ""

def ref_tail(el, tag):
    r = el.find(f"ar:{tag}", NP)
    return r.text.strip().split("/")[-1] if r is not None and r.text else ""

def ref_full(el, tag):
    r = el.find(f"ar:{tag}", NP)
    return r.text.strip() if r is not None and r.text else ""


# ── 1. 인터페이스 파싱 ────────────────────────────────────────────────
def parse_interfaces(arxml_files):
    """
    returns:
      if_map  : {short_name -> {type: 'SR'|'CS', signals: [sig_name, ...]}}
      full_map: {'/Pkg/PortInterfaces/IfName' -> short_name}  (경로→이름)
    """
    if_map  = {}
    full_map = {}

    for f in arxml_files:
        try:
            tree = ET.parse(f)
        except ET.ParseError as e:
            print(f"  XML parse error: {f.name}: {e}")
            continue
        root = tree.getroot()

        def walk(el, path_prefix):
            for child in el:
                local = child.tag.split("}")[-1] if "}" in child.tag else child.tag
                name = sname(child) if local not in ("AR-PACKAGES", "ELEMENTS") else ""
                cur_path = f"{path_prefix}/{name}" if name else path_prefix

                if local == "SENDER-RECEIVER-INTERFACE":
                    signals = []
                    for de in child.findall(".//ar:VARIABLE-DATA-PROTOTYPE", NP):
                        sn = sname(de)
                        if sn:
                            signals.append(sn)
                    if_map[name] = {"type": "SR", "signals": signals, "path": cur_path}
                    full_map[cur_path] = name

                elif local == "CLIENT-SERVER-INTERFACE":
                    ops = []
                    for op in child.findall(".//ar:CLIENT-SERVER-OPERATION", NP):
                        on = sname(op)
                        if on:
                            ops.append(on)
                    if_map[name] = {"type": "CS", "signals": ops, "path": cur_path}
                    full_map[cur_path] = name

                else:
                    walk(child, cur_path)

        walk(root, "")

    return if_map, full_map


# ── 2. SWC 포트 파싱 ──────────────────────────────────────────────────
def parse_swc_ports(arxml_files):
    """
    returns: {comp_name -> {port_name -> {kind:'P'|'R', if_ref: str}}}
    """
    comp_ports = {}

    for f in arxml_files:
        try:
            tree = ET.parse(f)
        except ET.ParseError:
            continue
        root = tree.getroot()

        for swc in root.findall(".//ar:APPLICATION-SW-COMPONENT-TYPE", NP):
            comp = sname(swc)
            if not comp:
                continue
            ports = {}
            for p in swc.findall(".//ar:P-PORT-PROTOTYPE", NP):
                pn = sname(p)
                ref = ref_full(p, "PROVIDED-INTERFACE-TREF")
                if pn and ref:
                    ports[pn] = {"kind": "P", "if_ref": ref.split("/")[-1],
                                 "if_path": ref}
            for r in swc.findall(".//ar:R-PORT-PROTOTYPE", NP):
                rn = sname(r)
                ref = ref_full(r, "REQUIRED-INTERFACE-TREF")
                if rn and ref:
                    ports[rn] = {"kind": "R", "if_ref": ref.split("/")[-1],
                                 "if_path": ref}
            comp_ports[comp] = ports

    return comp_ports


# ── 3. RootComposition 연결 파싱 ─────────────────────────────────────
def parse_connectors(composition_file):
    """
    returns: list of {provider_comp, provider_port, requester_comp, requester_port}
    """
    connectors = []
    tree = ET.parse(composition_file)
    root = tree.getroot()

    for conn in root.findall(".//ar:ASSEMBLY-SW-CONNECTOR", NP):
        prov_comp = ref_full(
            conn.find("ar:PROVIDER-IREF", NP) or ET.Element(""),
            "CONTEXT-COMPONENT-REF").split("/")[-1]
        prov_port = ref_full(
            conn.find("ar:PROVIDER-IREF", NP) or ET.Element(""),
            "TARGET-P-PORT-REF").split("/")[-1]
        req_comp  = ref_full(
            conn.find("ar:REQUESTER-IREF", NP) or ET.Element(""),
            "CONTEXT-COMPONENT-REF").split("/")[-1]
        req_port  = ref_full(
            conn.find("ar:REQUESTER-IREF", NP) or ET.Element(""),
            "TARGET-R-PORT-REF").split("/")[-1]
        if prov_comp and prov_port and req_comp:
            connectors.append({
                "provider_comp": prov_comp,
                "provider_port": prov_port,
                "requester_comp": req_comp,
                "requester_port": req_port,
            })

    return connectors


# ── 4. 소스코드 컨텍스트 수집 (선택) ─────────────────────────────────
def read_source_context(source_dir, comp_name):
    """컴포넌트 이름 기반으로 관련 .c 파일 내용 반환 (최대 3KB)"""
    parts = []
    for c_file in source_dir.rglob(f"{comp_name}*.c"):
        try:
            parts.append(c_file.read_text(encoding="utf-8", errors="ignore")[:1500])
        except Exception:
            pass
        if len(parts) >= 2:
            break
    return "\n\n".join(parts)[:3000] if parts else ""


# ── 5. FMEA 항목 생성 ────────────────────────────────────────────────
def generate_items(comp_ports, if_map, connectors, source_dir=None):
    # provider_comp+port → [requester_comp, ...]  수신 컴포넌트 목록
    receiver_map = defaultdict(list)
    for c in connectors:
        key = (c["provider_comp"], c["provider_port"])
        if c["requester_comp"] not in receiver_map[key]:
            receiver_map[key].append(c["requester_comp"])

    items = []
    item_no = 1
    seen = set()  # (comp, signal, mode) 중복 방지

    for comp, ports in sorted(comp_ports.items()):
        # DCM/DEM 등 스코프 제외
        if any(comp.startswith(p) for p in SKIP_PREFIX):
            continue

        # 이 컴포넌트가 External(CAN RX) 신호를 제공하는지 여부
        is_can_rx = comp in CAN_RX_COMPONENTS

        for port_name, port_info in ports.items():
            if port_info["kind"] != "P":
                continue  # Sender 포트만 처리

            if_name = port_info["if_ref"]
            iface   = if_map.get(if_name)
            if not iface:
                continue

            # 수신 컴포넌트 목록
            receivers = receiver_map.get((comp, port_name), [])
            effect_module = ", ".join(receivers) if receivers else None

            # 카테고리
            category = "External" if is_can_rx else "Internal"

            # 소스코드 컨텍스트
            src_ctx = ""
            if source_dir and not is_can_rx:
                src_ctx = read_source_context(source_dir, comp)

            # 신호별 FMEA
            modes = MODES_CS if iface["type"] == "CS" else MODES_SR
            for signal in iface["signals"]:
                for mode in modes:
                    key = (comp, signal, mode)
                    if key in seen:
                        continue
                    seen.add(key)

                    detail = DETAIL_TMPL[mode].format(var=signal)
                    if category == "External":
                        cause = f"CAN 통신 오류 / 외부 ECU 신호 이상 / 하드웨어 고장"
                    else:
                        cause = "SW 연산 오류 / 내부 상태 데이터 이상 / 메모리 손상"

                    items.append({
                        "comp":         comp,
                        "item_no":      str(item_no),
                        "category":     category,
                        "variable_name": signal,
                        "variable_type": "uint8" if iface["type"] == "CS" else iface["type"],
                        "failure_mode":  mode,
                        "failure_detail": detail,
                        "effect_module": effect_module,
                        "potential_cause": cause,
                        "port_name":     port_name,
                        "if_type":       iface["type"],
                        "source_ctx":    src_ctx[:500] if src_ctx else "",
                    })
                    item_no += 1

    return items


# ── Supabase 헬퍼 ─────────────────────────────────────────────────────
def fetch_sw_units():
    r = requests.get(f"{URL}/rest/v1/sw_units", headers=H,
                     params={"project_id": f"eq.{SX3_PROJECT_ID}",
                             "select": "id,name"}, verify=False)
    return {row["name"]: row["id"] for row in r.json()}

def ensure_sw_unit(name, unit_map):
    if name in unit_map:
        return unit_map[name]
    r = requests.post(f"{URL}/rest/v1/sw_units", headers=H,
                      json={"project_id": SX3_PROJECT_ID, "name": name},
                      verify=False)
    new_id = r.json()[0]["id"]
    unit_map[name] = new_id
    return new_id

def delete_all_items():
    r = requests.delete(f"{URL}/rest/v1/fmea_items", headers=H,
                        params={"project_id": f"eq.{SX3_PROJECT_ID}"},
                        verify=False)
    return r.status_code in (200, 204)

def batch_insert(rows, batch_size=200):
    total = 0
    for i in range(0, len(rows), batch_size):
        chunk = rows[i:i + batch_size]
        r = requests.post(f"{URL}/rest/v1/fmea_items", headers=H,
                          json=chunk, verify=False)
        if r.status_code not in (200, 201):
            print(f"  오류: {r.status_code} {r.text[:200]}")
        else:
            total += len(chunk)
        print(f"  삽입: {min(i+batch_size, len(rows))}/{len(rows)}", end="\r")
    print()
    return total


# ── main ──────────────────────────────────────────────────────────────
def main():
    print("=" * 60)
    print("SX3 ARXML → FMEA 생성")
    print("=" * 60)

    arxml_files = list(SWCD_APP_DIR.glob("*.arxml"))
    print(f"\n[1/5] ARXML 파일: {len(arxml_files)}개")
    for f in arxml_files:
        print(f"  {f.name}")

    print("\n[2/5] 인터페이스 파싱...")
    if_map, full_map = parse_interfaces(arxml_files)
    sr_count = sum(1 for v in if_map.values() if v["type"] == "SR")
    cs_count = sum(1 for v in if_map.values() if v["type"] == "CS")
    print(f"  SR 인터페이스: {sr_count}개, CS 인터페이스: {cs_count}개")

    print("\n[3/5] SW 컴포넌트 포트 파싱...")
    comp_ports = parse_swc_ports(arxml_files)
    print(f"  컴포넌트: {len(comp_ports)}개")
    for comp in sorted(comp_ports):
        p_cnt = sum(1 for v in comp_ports[comp].values() if v["kind"] == "P")
        r_cnt = sum(1 for v in comp_ports[comp].values() if v["kind"] == "R")
        print(f"  {comp:40s} P:{p_cnt:3d}  R:{r_cnt:3d}")

    print("\n[4/5] RootComposition 커넥터 파싱...")
    connectors = parse_connectors(COMPOSITION)
    print(f"  커넥터: {len(connectors)}개")

    print("\n[5/5] FMEA 항목 생성...")
    items = generate_items(comp_ports, if_map, connectors, SOURCE_DIR)
    print(f"  생성 항목: {len(items)}개")

    # 통계 출력
    from collections import Counter
    by_comp = Counter(it["comp"] for it in items)
    by_cat  = Counter(it["category"] for it in items)
    by_mode = Counter(it["failure_mode"] for it in items)
    print(f"\n  카테고리: {dict(by_cat)}")
    print(f"  실패모드: {dict(by_mode)}")
    print(f"\n  컴포넌트별 항목 수 (상위 15):")
    for comp, cnt in by_comp.most_common(15):
        print(f"    {comp:40s}: {cnt}")

    print("\n[샘플 10개]")
    for it in items[:10]:
        print(f"  {it['comp']:30s} | {it['variable_name']:30s} | {it['failure_mode']:8s}"
              f" | → {str(it['effect_module'])[:30]}")

    ans = input("\nDB에 업로드하시겠습니까? (y/N): ").strip().lower()
    if ans != "y":
        print("취소됨. 항목은 메모리에만 있습니다.")

        # JSON으로 저장
        import json
        out = Path("E:/claude/FMEA/sx3_arxml_fmea_preview.json")
        preview = [{k: v for k, v in it.items() if k != "source_ctx"}
                   for it in items[:200]]
        out.write_text(json.dumps(preview, ensure_ascii=False, indent=2),
                       encoding="utf-8")
        print(f"미리보기 200개 저장: {out}")
        return

    # Supabase 업로드
    print("\nSW Unit 조회...")
    unit_map = fetch_sw_units()

    print("기존 fmea_items 삭제...")
    ok = delete_all_items()
    print(f"  {'완료' if ok else '오류'}")

    print("삽입 중...")
    db_rows = []
    for it in items:
        uid = ensure_sw_unit(it["comp"], unit_map)
        db_rows.append({
            "project_id":      SX3_PROJECT_ID,
            "sw_unit_id":      uid,
            "item_no":         it["item_no"],
            "category":        it["category"],
            "variable_name":   it["variable_name"],
            "variable_type":   it["variable_type"],
            "failure_mode":    it["failure_mode"],
            "failure_detail":  it["failure_detail"],
            "effect_module":   it["effect_module"],
            "potential_cause": it["potential_cause"],
            "status":          "draft",
            "ai_generated":    False,
        })

    inserted = batch_insert(db_rows)
    print(f"\n완료! {inserted}개 삽입")


if __name__ == "__main__":
    main()
