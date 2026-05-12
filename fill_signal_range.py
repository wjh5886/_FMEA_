"""
DBC 파일 → fmea_items.signal_range 채우기
- DBC VAL_ enum 레이블 우선 사용
- VAL_ 없으면 bit_len → 물리 범위 계산
- Internal 신호는 data_type → 기본 범위
"""

import re, requests, urllib3
from pathlib import Path

urllib3.disable_warnings()

DBC_PATH = Path(
    "E:/HKMC_SX3_SBW_ICE_R44-4.22.00_20251117_T037_Proto_2026-04-27 200109"
    "/References/DB/20241125_STD_DB_CAR_R2.0_2024_FD_P1_SBW_0.9_SX3_ICE_20250916_수정.dbc"
)

SUPABASE_URL = "https://itzgdbeiyvodhfhmvrfw.supabase.co"
SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Iml0emdkYmVpeXZvZGhmaG12cmZ3Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzY3NDE2MzcsImV4cCI6MjA5MjMxNzYzN30.iqzQr-3Lqf1O4UHFe9euTLyeIyBXreLPoSzUdtEaNP8"
SX3_PROJECT_ID = "70f74c19-66b6-4b2e-a2b3-1ee04dd1b101"  # TK1

SB_H = {
    "apikey": SUPABASE_KEY,
    "Authorization": f"Bearer {SUPABASE_KEY}",
    "Content-Type": "application/json",
    "Prefer": "return=representation",
}

# 기본 데이터 타입 범위
TYPE_RANGE = {
    "boolean": "0(False) ~ 1(True)",
    "uint8":   "0 ~ 255",
    "uint16":  "0 ~ 65535",
    "uint32":  "0 ~ 4294967295",
    "sint8":   "-128 ~ 127",
    "sint16":  "-32768 ~ 32767",
    "sint32":  "-2147483648 ~ 2147483647",
}


# ─── DBC 파싱 ──────────────────────────────────────────────────────────────
def parse_dbc(dbc_path):
    """DBC → {sig_name: {bit_len, factor, offset, is_signed, min_phy, max_phy}}
               vals     → {sig_name: {raw_val: label}}"""
    signals = {}
    vals    = {}

    with open(dbc_path, encoding="utf-8", errors="replace") as f:
        lines = f.readlines()

    for line in lines:
        line = line.strip()

        # SG_ SignalName : startBit|bitLen@byteOrder valueType(factor,offset) [min|max] "unit" receivers
        m = re.match(
            r'SG_\s+(\w+)\s*:\s*\d+\|(\d+)@\d+([+-])\s*'
            r'\(([^,]+),([^)]+)\)\s*\[([^|]*)\|([^\]]*)\]',
            line
        )
        if m:
            name      = m.group(1)
            bit_len   = int(m.group(2))
            is_signed = m.group(3) == "-"
            factor    = float(m.group(4))
            offset    = float(m.group(5))
            mn_str    = m.group(6).strip()
            mx_str    = m.group(7).strip()

            # 물리 범위 계산
            if is_signed:
                raw_min = -(2 ** (bit_len - 1))
                raw_max =  (2 ** (bit_len - 1)) - 1
            else:
                raw_min = 0
                raw_max = (2 ** bit_len) - 1

            phy_min = factor * raw_min + offset
            phy_max = factor * raw_max + offset

            # DBC에 명시된 범위가 있으면 우선 (단, 0~0은 무시)
            try:
                mn_f = float(mn_str)
                mx_f = float(mx_str)
                if mn_f != 0.0 or mx_f != 0.0:
                    phy_min, phy_max = mn_f, mx_f
            except ValueError:
                pass

            signals[name] = {
                "bit_len":   bit_len,
                "is_signed": is_signed,
                "factor":    factor,
                "offset":    offset,
                "phy_min":   phy_min,
                "phy_max":   phy_max,
            }

        # VAL_ msgId sigName val0 "label0" val1 "label1" ... ;
        m2 = re.match(r'VAL_\s+\d+\s+(\w+)\s+(.*?)\s*;', line)
        if m2:
            sig_name = m2.group(1)
            rest     = m2.group(2)
            enum_map = {}
            for vm in re.finditer(r'(\d+)\s+"([^"]*)"', rest):
                enum_map[int(vm.group(1))] = vm.group(2)
            if enum_map:
                vals[sig_name] = enum_map

    return signals, vals


def make_range_str(sig_name, sig_info, vals):
    """신호 정보 → 범위 문자열"""
    if sig_name in vals:
        enum = vals[sig_name]
        # 너무 많으면 처음 6개만
        items = sorted(enum.items())[:6]
        labels = ", ".join(f'{v}={l}' for v, l in items)
        max_v = max(enum.keys())
        if len(enum) < len(vals.get(sig_name, {})):
            labels += " ..."
        return f"[0 ~ {max_v}] {labels}"
    else:
        mn = sig_info["phy_min"]
        mx = sig_info["phy_max"]
        # 정수이면 int로
        mn_s = str(int(mn)) if mn == int(mn) else f"{mn:.4g}"
        mx_s = str(int(mx)) if mx == int(mx) else f"{mx:.4g}"
        return f"{mn_s} ~ {mx_s}"


# ─── 신호명 매칭 ────────────────────────────────────────────────────────────
def build_lookup(signals, vals):
    """DBC 신호명 → {lower_suffix: sig_name} 역인덱스 구축"""
    # 신호명을 소문자로 정규화
    # 예: BCM_GearPosPSta → 'bcm_gearpospta', 'gearpospta'
    lookup = {}
    for sname in signals:
        low = sname.lower()
        lookup[low] = sname
        # 마지막 '_' 이후 suffix도 등록
        parts = low.split("_")
        for i in range(1, len(parts)):
            suffix = "_".join(parts[i:])
            if suffix not in lookup:
                lookup[suffix] = sname
    return lookup


def match_signal(variable_name, lookup):
    """FMEA variable_name → DBC sig_name 매칭"""
    vn_low = variable_name.lower()
    # 1. 정확히 일치
    if vn_low in lookup:
        return lookup[vn_low]
    # 2. variable_name을 suffix로 포함하는 신호 (첫 번째 매칭)
    for key, sname in lookup.items():
        if key.endswith(vn_low) or vn_low in key:
            return sname
    return None


# ─── Supabase 조회/업데이트 ─────────────────────────────────────────────────
def fetch_all_items(project_id):
    items = []
    offset, batch = 0, 1000
    while True:
        r = requests.get(
            f"{SUPABASE_URL}/rest/v1/fmea_items",
            headers={**SB_H, "Range": f"{offset}-{offset+batch-1}"},
            params={"project_id": f"eq.{project_id}",
                    "select": "id,variable_name,variable_type,category"},
            verify=False,
        )
        data = r.json()
        if not data:
            break
        items.extend(data)
        if len(data) < batch:
            break
        offset += batch
    return items


def update_by_varname(var_range_map, project_id):
    """variable_name 기준으로 일괄 PATCH (요청 수 = 고유 변수명 수)"""
    import concurrent.futures

    def patch_one(vn, rng):
        r = requests.patch(
            f"{SUPABASE_URL}/rest/v1/fmea_items",
            headers=SB_H,
            params={"project_id": f"eq.{project_id}", "variable_name": f"eq.{vn}"},
            json={"signal_range": rng},
            verify=False,
        )
        return r.status_code in (200, 204)

    items_to_update = [(vn, rng) for vn, rng in var_range_map.items() if rng]
    updated = 0
    with concurrent.futures.ThreadPoolExecutor(max_workers=20) as ex:
        futures = {ex.submit(patch_one, vn, rng): vn for vn, rng in items_to_update}
        for i, f in enumerate(concurrent.futures.as_completed(futures), 1):
            if f.result():
                updated += 1
            if i % 50 == 0:
                print(f"  진행: {i}/{len(futures)}", end="\r")
    print()
    return updated


# ─── 메인 ──────────────────────────────────────────────────────────────────
def main():
    print("=" * 60)
    print("DBC → signal_range 채우기")
    print("=" * 60)

    # 1. DBC 파싱
    print("\n[1/3] DBC 파싱...")
    signals, vals = parse_dbc(DBC_PATH)
    print(f"  신호: {len(signals)}개, VAL_(enum): {len(vals)}개")

    # 역인덱스 구축
    lookup = build_lookup(signals, vals)

    # 2. FMEA 항목 조회
    print("\n[2/3] FMEA 항목 조회...")
    items = fetch_all_items(SX3_PROJECT_ID)
    print(f"  총 {len(items)}개")

    # 3. 매칭 + 범위 결정
    print("\n[3/3] 범위 매칭 및 업데이트...")

    # variable_name 기준으로 중복 제거해서 매칭 (id는 일괄 적용)
    var_range = {}  # variable_name → signal_range string

    for item in items:
        vn = item["variable_name"]
        if vn in var_range:
            continue

        cat  = item.get("category", "")
        vtyp = item.get("variable_type", "") or ""

        # CAN/External 신호 → DBC 매칭 시도
        matched_sig = match_signal(vn, lookup)

        if matched_sig and matched_sig in signals:
            rng = make_range_str(matched_sig, signals[matched_sig], vals)
        elif vtyp.lower() in TYPE_RANGE:
            # data type 기본 범위
            rng = TYPE_RANGE[vtyp.lower()]
        elif "boolean" in vtyp.lower():
            rng = TYPE_RANGE["boolean"]
        elif "uint8" in vtyp.lower():
            rng = TYPE_RANGE["uint8"]
        elif "uint16" in vtyp.lower():
            rng = TYPE_RANGE["uint16"]
        elif "sint" in vtyp.lower():
            bits = re.search(r'(\d+)', vtyp)
            rng = TYPE_RANGE.get(f"sint{bits.group(1)}", "N/A") if bits else "N/A"
        elif vtyp.startswith("msggr_"):
            # CAN 메시지 그룹 → 메시지 수신 상태
            rng = "0(Valid) ~ 1(Invalid) [Message Group]"
        else:
            rng = None

        var_range[vn] = rng

    # 통계
    matched_dbc  = sum(1 for v in var_range.values() if v and "Message Group" not in v and "~" in v and v not in TYPE_RANGE.values())
    type_default = sum(1 for v in var_range.values() if v in TYPE_RANGE.values())
    no_range     = sum(1 for v in var_range.values() if not v)
    print(f"  DBC 매칭:    {matched_dbc}개 변수")
    print(f"  타입 기본값: {type_default}개 변수")
    print(f"  범위 없음:   {no_range}개 변수")

    # 샘플 출력
    print("\n  [샘플]")
    count = 0
    for vn, rng in var_range.items():
        if rng and rng not in TYPE_RANGE.values():
            print(f"    {vn:35s}: {rng}")
            count += 1
            if count >= 10:
                break

    # 업데이트 준비
    updates = []
    for item in items:
        rng = var_range.get(item["variable_name"])
        if rng:
            updates.append({"id": item["id"], "signal_range": rng})

    print(f"\n  업데이트 대상 변수: {len([v for v in var_range.values() if v])}개 고유 변수명")
    updated = update_by_varname(var_range, SX3_PROJECT_ID)
    print(f"\n완료! {updated}개 변수명 signal_range 저장")


if __name__ == "__main__":
    main()
