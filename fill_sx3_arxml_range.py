"""
SX3_ARXML 프로젝트 signal_range 채우기
1. App_PortInterface.arxml → variable_name → data type
2. App_DataTypes.arxml      → custom type → base type (uint8/uint16)
3. SX3 DBC 3종             → 신호명 → VAL_(enum) 또는 물리 범위
4. DBC 매칭: ARXML variable_name에서 "De"/"Dg" 접두어 제거 후 매칭
"""

import re, requests, urllib3, xml.etree.ElementTree as ET
from pathlib import Path
from collections import defaultdict
import concurrent.futures

urllib3.disable_warnings()

# ── 설정 ──────────────────────────────────────────────────────────────────
BASE = Path("E:/claude/FMEA/SBW_FMEA/SX3/"
            "HKMC_SX3_SBW_ICE_R44-4.22.00_20251117_T037_Proto_2026-04-27 200109")
SYSTEM = BASE / "Configuration/System"

DBC_FILES = [
    SYSTEM / "DBImport/../../../References/DB/20241125_STD_DB_CAR_R2.0_2024_FD_P1_SBW_0.9_SX3_ICE_20250916_수정.dbc",
    Path("E:/claude/FMEA/SBW_FMEA/SX3/20250902_STD_DB_CAR_R2.0_2024_FD_P1_v25.08.01.dbc"),
    Path("E:/claude/FMEA/SBW_FMEA/SX3/20250902_STD_DB_CAR_R2.0_2024_FD_H_Local_v25.08.01.dbc"),
    Path("E:/claude/FMEA/SBW_FMEA/SX3/20250902_CANFD_마스터__SBW_SHFTR_FF_HS_송신_수신_25.08.01(CAN FD).dbc"),
]
# 실제 존재하는 것만
DBC_FILES = [p for p in DBC_FILES if p.exists()]

SX3_ARXML_PROJECT_ID = "6a4bf862-070d-4e24-af0e-4f36183be745"
URL = "https://itzgdbeiyvodhfhmvrfw.supabase.co"
KEY = ("eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Iml0"
       "emdkYmVpeXZvZGhmaG12cmZ3Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzY3NDE2MzcsImV4cCI"
       "6MjA5MjMxNzYzN30.iqzQr-3Lqf1O4UHFe9euTLyeIyBXreLPoSzUdtEaNP8")
H = {"apikey": KEY, "Authorization": f"Bearer {KEY}",
     "Content-Type": "application/json", "Prefer": "return=representation"}

NS = "http://autosar.org/schema/r4.0"
NP = {"ar": NS}

# 기본 타입 범위
BASE_RANGES = {
    "uint8":   "0 ~ 255",
    "uint8_t": "0 ~ 255",
    "uint16":  "0 ~ 65535",
    "uint32":  "0 ~ 4294967295",
    "sint8":   "-128 ~ 127",
    "sint16":  "-32768 ~ 32767",
    "sint32":  "-2147483648 ~ 2147483647",
    "boolean": "0(False) ~ 1(True)",
    "float":   "IEEE 754 float",
}

# External 신호 고정 범위
EXTERNAL_FIXED = {
    "CanMsg":          "CAN 메시지 수신 그룹 (유효/타임아웃 상태)",
    "TransformerError":"0(정상) ~ 1(E2E/Transformer 오류)",
    "u1_GetESigInfo":  "E-Signal 정보 (0=정상, 비0=오류 코드)",
    "u1_GetPSigInfo":  "P-Signal 정보 (0=정상, 비0=오류 코드)",
}


# ── 1. DBC 파싱 ────────────────────────────────────────────────────────────
def parse_dbc_all(dbc_files):
    """모든 DBC에서 신호 정보 + VAL_ 통합"""
    signals = {}  # sig_name -> {phy_min, phy_max}
    vals    = {}  # sig_name -> {raw: label}

    for dbc_path in dbc_files:
        print(f"  DBC: {dbc_path.name}")
        try:
            with open(dbc_path, encoding="utf-8", errors="replace") as f:
                lines = f.readlines()
        except Exception as e:
            print(f"    오류: {e}")
            continue

        for line in lines:
            line = line.strip()
            m = re.match(
                r'SG_\s+(\w+)\s*:\s*\d+\|(\d+)@\d+([+-])\s*'
                r'\(([^,]+),([^)]+)\)\s*\[([^|]*)\|([^\]]*)\]',
                line)
            if m:
                name, bit_len = m.group(1), int(m.group(2))
                is_signed = m.group(3) == "-"
                factor    = float(m.group(4))
                offset    = float(m.group(5))
                if is_signed:
                    raw_min, raw_max = -(2**(bit_len-1)), 2**(bit_len-1)-1
                else:
                    raw_min, raw_max = 0, 2**bit_len-1
                phy_min = factor*raw_min + offset
                phy_max = factor*raw_max + offset
                try:
                    mn_f, mx_f = float(m.group(6)), float(m.group(7))
                    if mn_f != 0.0 or mx_f != 0.0:
                        phy_min, phy_max = mn_f, mx_f
                except ValueError:
                    pass
                if name not in signals:
                    signals[name] = {"phy_min": phy_min, "phy_max": phy_max}

            m2 = re.match(r'VAL_\s+\d+\s+(\w+)\s+(.*?)\s*;', line)
            if m2:
                sname = m2.group(1)
                enum_map = {}
                for vm in re.finditer(r'(\d+)\s+"([^"]*)"', m2.group(2)):
                    enum_map[int(vm.group(1))] = vm.group(2)
                if enum_map and sname not in vals:
                    vals[sname] = enum_map

    return signals, vals


def make_range_str(sig_info, enum_map):
    if enum_map:
        items = sorted(enum_map.items())[:8]
        labels = ", ".join(f"{v}={l}" for v, l in items)
        if len(enum_map) > 8:
            labels += " ..."
        return f"[0 ~ {max(enum_map)}] {labels}"
    mn, mx = sig_info["phy_min"], sig_info["phy_max"]
    fmt = lambda x: str(int(x)) if x == int(x) else f"{x:.4g}"
    return f"{fmt(mn)} ~ {fmt(mx)}"


def build_dbc_lookup(signals):
    """신호명 → lower-case 역인덱스 (정확 + 접미어)"""
    lookup = {}
    for sname in signals:
        low = sname.lower()
        lookup[low] = sname
        parts = low.split("_")
        for i in range(1, len(parts)):
            suffix = "_".join(parts[i:])
            if suffix not in lookup:
                lookup[suffix] = sname
    return lookup


# ── 2. ARXML 타입 파싱 ────────────────────────────────────────────────────
def parse_arxml_types():
    """variable_name → range_str (ARXML 기반)"""
    # App_PortInterface.arxml: variable_name → type_name
    var_type = {}
    port_if = SYSTEM / "Swcd_App/App_PortInterface.arxml"
    if port_if.exists():
        for vdp in ET.parse(port_if).getroot().iter(f'{{{NS}}}VARIABLE-DATA-PROTOTYPE'):
            n = vdp.find(f'{{{NS}}}SHORT-NAME')
            t = vdp.find(f'{{{NS}}}TYPE-TREF')
            if n is not None and t is not None:
                var_type[n.text] = t.text.split('/')[-1]

    # App_DataTypes.arxml + AUTOSAR_DataTypes.arxml: custom type → base type
    custom_base = {}
    for dtpath in [
        SYSTEM / "Swcd_App/App_DataTypes.arxml",
        SYSTEM / "DBImport/DataTypes/AUTOSAR_DataTypes.arxml",
    ]:
        if not dtpath.exists():
            continue
        for idt in ET.parse(dtpath).getroot().iter(f'{{{NS}}}IMPLEMENTATION-DATA-TYPE'):
            n = idt.find(f'{{{NS}}}SHORT-NAME')
            if n is None:
                continue
            tname = n.text
            base_ref = idt.find(f'.//{{{NS}}}BASE-TYPE-REF')
            if base_ref is not None:
                custom_base[tname] = base_ref.text.split('/')[-1]
            # 명시 범위
            lo = idt.find(f'.//{{{NS}}}LOWER-LIMIT')
            hi = idt.find(f'.//{{{NS}}}UPPER-LIMIT')
            if lo is not None and hi is not None and tname not in custom_base:
                custom_base[tname] = f"{lo.text} ~ {hi.text}"

    # variable_name → range_str
    def resolve(type_name):
        if type_name in BASE_RANGES:
            return BASE_RANGES[type_name]
        base = custom_base.get(type_name, "")
        if base in BASE_RANGES:
            return BASE_RANGES[base]
        if "uint8" in base.lower():
            return BASE_RANGES["uint8"]
        if "uint16" in base.lower():
            return BASE_RANGES["uint16"]
        return None

    result = {}
    for vname, tname in var_type.items():
        r = resolve(tname)
        if r:
            result[vname] = r
    return result


# ── 3. DBC 매칭 ───────────────────────────────────────────────────────────
def match_dbc(variable_name, dbc_lookup, signals, vals):
    """ARXML variable_name → DBC 신호 매칭"""
    vn = variable_name

    # "De"/"Dg"/"Di"/"Dv" 접두어 제거 시도
    candidates = [vn]
    for prefix in ("De", "Dg", "Di", "Dv", "Dp", "Dm"):
        if vn.startswith(prefix) and len(vn) > len(prefix):
            candidates.append(vn[len(prefix):])

    for cand in candidates:
        low = cand.lower()
        # 정확 매칭
        if low in dbc_lookup:
            sname = dbc_lookup[low]
            return make_range_str(signals[sname], vals.get(sname, {}))
        # 포함 매칭
        for key, sname in dbc_lookup.items():
            if len(low) > 4 and (low in key or key.endswith(low)):
                return make_range_str(signals[sname], vals.get(sname, {}))
    return None


# ── 4. Supabase 조회/업데이트 ─────────────────────────────────────────────
def fetch_items():
    items = []
    offset = 0
    while True:
        r = requests.get(f"{URL}/rest/v1/fmea_items",
            headers={**H, "Range": f"{offset}-{offset+999}"},
            params={"project_id": f"eq.{SX3_ARXML_PROJECT_ID}",
                    "select": "id,variable_name,variable_type,category"},
            verify=False)
        batch = r.json()
        if not isinstance(batch, list) or not batch:
            break
        items.extend(batch)
        if len(batch) < 1000:
            break
        offset += 1000
    return items


def patch_range(vn, rng):
    r = requests.patch(f"{URL}/rest/v1/fmea_items",
        headers=H,
        params={"project_id": f"eq.{SX3_ARXML_PROJECT_ID}", "variable_name": f"eq.{vn}"},
        json={"signal_range": rng},
        verify=False)
    return r.status_code in (200, 204)


# ── 메인 ──────────────────────────────────────────────────────────────────
def main():
    print("=" * 65)
    print("SX3_ARXML signal_range 채우기")
    print("=" * 65)

    print(f"\n사용 DBC: {len(DBC_FILES)}개")

    # 1. DBC 파싱
    print("\n[1/4] DBC 파싱...")
    signals, vals = parse_dbc_all(DBC_FILES)
    print(f"  신호: {len(signals)}개, VAL_(enum): {len(vals)}개")
    dbc_lookup = build_dbc_lookup(signals)

    # 2. ARXML 타입 파싱
    print("\n[2/4] ARXML 타입 파싱...")
    arxml_ranges = parse_arxml_types()
    print(f"  variable_name → range: {len(arxml_ranges)}개")

    # 3. FMEA 항목 조회
    print("\n[3/4] SX3_ARXML 항목 조회...")
    items = fetch_items()
    print(f"  총 {len(items)}개")

    # 4. 범위 결정
    print("\n[4/4] 범위 결정 및 업데이트...")

    var_range = {}  # variable_name → range_str (중복 제거)
    stat = {"dbc": 0, "arxml": 0, "basetype": 0, "external": 0, "none": 0}

    for item in items:
        vn   = item["variable_name"]
        vtyp = item["variable_type"] or ""
        cat  = item["category"]

        if vn in var_range:
            continue

        # External 고정 범위
        if cat == "External" and vn in EXTERNAL_FIXED:
            var_range[vn] = EXTERNAL_FIXED[vn]
            stat["external"] += 1
            continue

        # DBC 매칭 시도
        rng = match_dbc(vn, dbc_lookup, signals, vals)
        if rng:
            var_range[vn] = rng
            stat["dbc"] += 1
            continue

        # ARXML 타입 기반
        if vn in arxml_ranges:
            var_range[vn] = arxml_ranges[vn]
            stat["arxml"] += 1
            continue

        # variable_type 직접 사용 (uint8 등)
        if vtyp.lower() in BASE_RANGES:
            var_range[vn] = BASE_RANGES[vtyp.lower()]
            stat["basetype"] += 1
            continue

        var_range[vn] = None
        stat["none"] += 1

    print(f"\n  결과:")
    print(f"    DBC 매칭:      {stat['dbc']}개 고유 변수")
    print(f"    ARXML 타입:    {stat['arxml']}개 고유 변수")
    print(f"    기본 타입:     {stat['basetype']}개 고유 변수")
    print(f"    External 고정: {stat['external']}개 고유 변수")
    print(f"    범위 없음:     {stat['none']}개 고유 변수")

    # DBC 매칭 샘플 출력
    print("\n  [DBC 매칭 샘플]")
    dbc_matched = [(vn, rng) for vn, rng in var_range.items()
                   if rng and rng not in BASE_RANGES.values() and rng not in EXTERNAL_FIXED.values()]
    for vn, rng in dbc_matched[:10]:
        print(f"    {vn:40s}: {rng}")

    # 업데이트
    to_update = {vn: rng for vn, rng in var_range.items() if rng}
    print(f"\n  업데이트 대상: {len(to_update)}개 고유 variable_name")
    print(f"  (아이템 기준 약 {len(items) * len(to_update) // max(len(var_range),1)}개 행)")

    updated = 0
    with concurrent.futures.ThreadPoolExecutor(max_workers=20) as ex:
        futures = {ex.submit(patch_range, vn, rng): vn for vn, rng in to_update.items()}
        for i, f in enumerate(concurrent.futures.as_completed(futures), 1):
            if f.result():
                updated += 1
            if i % 20 == 0:
                print(f"  진행: {i}/{len(futures)}", end="\r")
    print()
    print(f"\n완료! {updated}개 variable_name signal_range 저장")


if __name__ == "__main__":
    main()
