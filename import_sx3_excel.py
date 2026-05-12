"""
SX3_Software_Composition_All_Interface_20251218.xlsx → SX3 fmea_items 재생성
개선:
- variable_name = Argument Name (Col10, 실제 AUTOSAR 변수명)
- effect_module = Destination (SENDER) 또는 Source (RECEIVER) - 영향받는 모듈
- failure_detail = 실패모드별 템플릿 자동 생성
- potential_cause = 카테고리별 템플릿
- CLIENT PORT: 번호 접두사 행 스킵
"""

import re, requests, urllib3
from pathlib import Path

urllib3.disable_warnings()

XLSX_PATH = Path("E:/claude/FMEA/SBW_FMEA/SX3/SX3_Software_Composition_All_Interface_20251218.xlsx")
SX3_PROJECT_ID = "a43d7d4e-b104-4e90-ab55-04c7b31aa3e7"

SUPABASE_URL = "https://itzgdbeiyvodhfhmvrfw.supabase.co"
SUPABASE_KEY = ("eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Iml0"
                "emdkYmVpeXZvZGhmaG12cmZ3Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzY3NDE2MzcsImV4cCI"
                "6MjA5MjMxNzYzN30.iqzQr-3Lqf1O4UHFe9euTLyeIyBXreLPoSzUdtEaNP8")
SB_H = {
    "apikey": SUPABASE_KEY,
    "Authorization": f"Bearer {SUPABASE_KEY}",
    "Content-Type": "application/json",
    "Prefer": "return=representation",
}

# 실패모드별 failure_detail 템플릿
DETAIL_TMPL = {
    "MORE":    "{var}이(가) 정상 범위{rng}를 초과하는 값을 출력/수신",
    "LESS":    "{var}이(가) 정상 범위{rng} 미만의 값을 출력/수신",
    "CORRUPT": "{var}이(가) 정상 범위 내이나 논리적으로 잘못된 값을 출력/수신",
    "EARLY":   "{var}이(가) 예상 시점보다 일찍 업데이트/발생됨",
    "LATE":    "{var}이(가) 예상 시점보다 늦게 업데이트됨 (타임아웃 가능성)",
    "STUCK":   "{var}이(가) 특정 값에 고착되어 변화하지 않음",
    "ERRATIC": "{var}이(가) 불규칙하게 값이 변동하거나 진동함",
}

# 카테고리별 potential_cause 템플릿
CAUSE_INTERNAL = "SW 연산 오류 / 내부 상태 데이터 이상 / 메모리 손상"
CAUSE_EXTERNAL = "CAN/통신 오류 / 외부 ECU 신호 이상 / 하드웨어 고장"


def clean_range(raw):
    if not raw:
        return None
    s = str(raw).strip()
    if not s:
        return None
    lines = [l.strip() for l in s.split("\n") if l.strip()]
    enum_parts = []
    for line in lines:
        m = re.match(r"(\d+)u?\s*:\s*(.+)", line)
        if m:
            enum_parts.append(f"{m.group(1)}={m.group(2).strip()}")
    if enum_parts:
        return ", ".join(enum_parts)
    s_clean = re.sub(r"u\b", "", s).strip()
    return s_clean


def clean_modules(raw):
    """'CstAp_CANMgt\nCstAp_DtcMgt\n' → 'CstAp_CANMgt, CstAp_DtcMgt'"""
    if not raw:
        return None
    parts = [p.strip() for p in str(raw).split("\n") if p.strip()]
    return ", ".join(parts) if parts else None


def is_numbered_param(s):
    """'1. CtNvM_Task_O_st_CalData' 처럼 번호로 시작하면 True"""
    return bool(re.match(r"^\d+\.", s.strip()))


def fetch_sw_units():
    r = requests.get(
        f"{SUPABASE_URL}/rest/v1/sw_units",
        headers=SB_H,
        params={"project_id": f"eq.{SX3_PROJECT_ID}", "select": "id,name"},
        verify=False,
    )
    return {row["name"]: row["id"] for row in r.json()}


def ensure_sw_unit(name, unit_map):
    if name in unit_map:
        return unit_map[name]
    r = requests.post(
        f"{SUPABASE_URL}/rest/v1/sw_units",
        headers=SB_H,
        json={"project_id": SX3_PROJECT_ID, "name": name},
        verify=False,
    )
    new_id = r.json()[0]["id"]
    unit_map[name] = new_id
    return new_id


def delete_all_items():
    r = requests.delete(
        f"{SUPABASE_URL}/rest/v1/fmea_items",
        headers=SB_H,
        params={"project_id": f"eq.{SX3_PROJECT_ID}"},
        verify=False,
    )
    return r.status_code in (200, 204)


def batch_insert(items, batch_size=200):
    total = 0
    for i in range(0, len(items), batch_size):
        chunk = items[i:i + batch_size]
        r = requests.post(
            f"{SUPABASE_URL}/rest/v1/fmea_items",
            headers=SB_H,
            json=chunk,
            verify=False,
        )
        if r.status_code not in (200, 201):
            print(f"  오류: {r.status_code} {r.text[:200]}")
        else:
            total += len(chunk)
        print(f"  삽입 진행: {min(i+batch_size, len(items))}/{len(items)}", end="\r")
    print()
    return total


def main():
    import openpyxl
    print("=" * 60)
    print("SX3 Interface Excel → FMEA 재생성 (개선)")
    print("=" * 60)

    print("\n[1/4] Excel 파싱...")
    wb = openpyxl.load_workbook(XLSX_PATH, read_only=True, data_only=True)
    ws = wb["Sheet"]
    rows = list(ws.iter_rows(values_only=True))
    data_rows = rows[2:]
    print(f"  데이터 행: {len(data_rows)}개")

    print("\n[2/4] SW Unit 조회...")
    unit_map = fetch_sw_units()
    print(f"  기존 SW Unit: {len(unit_map)}개")

    print("\n[3/4] 기존 fmea_items 삭제...")
    ok = delete_all_items()
    print(f"  {'완료' if ok else '오류'}")

    print("\n[4/4] FMEA 항목 생성 및 삽입...")
    items = []
    item_no = 1
    skip_reasons = {}

    for row in data_rows:
        # Col 인덱스
        # 0:No, 1:Composition, 2:Port Type, 3:IF Name, 4:IF Proto Name,
        # 5:Operation, 6:Param Type, 7:Param Name, 8:Param Dir,
        # 9:Arg Type, 10:Arg Name, 11:Value Range, 12:Source, 13:Dest
        composition = str(row[1]).strip() if row[1] else None
        port_type   = str(row[2]).strip() if row[2] else None
        if_name     = str(row[3]).strip() if row[3] else None
        param_type  = str(row[6]).strip() if row[6] else None
        param_name  = str(row[7]).strip() if row[7] else None
        arg_name    = str(row[10]).strip() if row[10] else None
        val_range   = row[11]
        src         = row[12]
        dst         = row[13]

        # 기본 필터
        if not composition or not port_type:
            skip_reasons["no_composition"] = skip_reasons.get("no_composition", 0) + 1
            continue

        # 번호 접두사 파라미터 스킵 (C/S 포트 numbered args)
        if param_name and is_numbered_param(param_name):
            skip_reasons["numbered_param"] = skip_reasons.get("numbered_param", 0) + 1
            continue

        # arg_name '-' 이거나 없으면 스킵
        if not arg_name or arg_name == "-":
            skip_reasons["no_arg_name"] = skip_reasons.get("no_arg_name", 0) + 1
            continue

        # 번호 접두사 arg_name도 스킵
        if is_numbered_param(arg_name):
            skip_reasons["numbered_arg"] = skip_reasons.get("numbered_arg", 0) + 1
            continue

        # variable_name = Argument Name (AUTOSAR 실제 변수명)
        variable_name = arg_name

        # SW Unit
        sw_unit_id = ensure_sw_unit(composition, unit_map)

        # category & effect_module
        if port_type in ("SENDER PORT", "SERVER PORT"):
            category = "Internal"
            effect_module = clean_modules(dst)   # 이 신호를 받는 모듈들이 영향받음
        else:  # RECEIVER PORT, CLIENT PORT
            category = "External"
            effect_module = composition          # 이 컴포넌트가 영향받음

        # 실패모드 선택
        if port_type in ("CLIENT PORT", "SERVER PORT"):
            modes = ["CORRUPT", "EARLY", "LATE"]
        else:
            modes = ["MORE", "LESS", "CORRUPT", "EARLY", "LATE", "STUCK", "ERRATIC"]

        sig_range = clean_range(val_range)
        rng_str = f"({sig_range})" if sig_range else ""

        # potential_cause
        cause = CAUSE_INTERNAL if category == "Internal" else CAUSE_EXTERNAL
        if category == "External" and src:
            src_clean = clean_modules(src)
            cause = f"{src_clean} 출력 이상 / CAN 통신 오류 / 하드웨어 고장"

        for mode in modes:
            detail = DETAIL_TMPL[mode].format(var=variable_name, rng=rng_str)
            items.append({
                "project_id":      SX3_PROJECT_ID,
                "sw_unit_id":      sw_unit_id,
                "item_no":         str(item_no),
                "category":        category,
                "variable_name":   variable_name,
                "variable_type":   param_type,
                "failure_mode":    mode,
                "failure_detail":  detail,
                "effect_module":   effect_module,
                "potential_cause": cause,
                "signal_range":    sig_range,
                "status":          "draft",
                "ai_generated":    False,
            })
            item_no += 1

    print(f"  생성 항목: {len(items)}개")
    print(f"  스킵 현황: {skip_reasons}")

    inserted = batch_insert(items)
    print(f"\n완료! {inserted}개 삽입")

    # 통계
    from collections import Counter
    by_unit = Counter(it["sw_unit_id"] for it in items)
    id_to_name = {v: k for k, v in unit_map.items()}
    print("\n[SW Unit별 항목 수]")
    for uid, cnt in sorted(by_unit.items(), key=lambda x: -x[1]):
        print(f"  {id_to_name.get(uid, uid):35s}: {cnt}개")

    print("\n[샘플 데이터]")
    for it in items[:3]:
        print(f"  {it['variable_name']}/{it['failure_mode']}: {it['failure_detail'][:60]}")
        print(f"    effect_module={it['effect_module']}, cause={it['potential_cause'][:50]}")


if __name__ == "__main__":
    main()
