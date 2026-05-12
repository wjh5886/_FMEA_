"""
bulk_fill_sod.py — 여러 프로젝트 S/O/D 빠른 일괄 채우기
1단계: SX3_ICE_TEST(완성) 교차 참조 복사
2단계: 나머지는 규칙 기반 자동 생성

사용법:
  python bulk_fill_sod.py                  # 모든 미완성 프로젝트
  python bulk_fill_sod.py "GN7_FL" "LQ2"  # 특정 프로젝트만
"""

import sys, re, requests, urllib3, concurrent.futures
from collections import defaultdict
urllib3.disable_warnings()

SB_URL = "https://itzgdbeiyvodhfhmvrfw.supabase.co"
SB_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Iml0emdkYmVpeXZvZGhmaG12cmZ3Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzY3NDE2MzcsImV4cCI6MjA5MjMxNzYzN30.iqzQr-3Lqf1O4UHFe9euTLyeIyBXreLPoSzUdtEaNP8"
SB_H = {
    "apikey": SB_KEY,
    "Authorization": f"Bearer {SB_KEY}",
    "Content-Type": "application/json",
    "Prefer": "return=minimal",
}

# 교차 참조 소스 (완성도 높은 프로젝트 우선)
SRC_IDS = [
    "1ba10b41-4717-4d1b-b0df-1b91fe2c4870",  # SX3_ICE_TEST (100%)
    "0715f883-d3a1-4ddd-8a3b-d3071da9ed3e",  # JG1
    "a43d7d4e-b104-4e90-ab55-04c7b31aa3e7",  # SX3
]

_NOISE = {"i", "p", "b", "u1", "u8", "u16", "u32", "ctap", "ctdcm", "raw"}
_DE_PFX = ("de", "dg", "di", "dv", "dp", "dm")

_SEV_KW = [
    (10, ("steering", "brake", "accel")),
    (9,  ("lvrpos", "gearpos", "shiftpos", "parkpos", "shiftact", "sbw")),
    (8,  ("ign", "ecumodeect", "drvrdyst", "lvr", "park", "busoff", "canbus")),
    (7,  ("hall", "sensor", "flt", "fault", "err", "idtflt", "idtsta")),
    (6,  ("bat", "vbat", "ldo", "power", "cal", "calib", "can", "rx", "tx")),
    (5,  ("sta", "status", "chk", "check", "mon", "info")),
]
_DET = {"MORE":3,"LESS":3,"CORRUPT":5,"STUCK":4,"EARLY":6,"LATE":6,"ERRATIC":7}
_OCC = {"MORE":3,"LESS":3,"CORRUPT":3,"STUCK":3,"EARLY":2,"LATE":3,"ERRATIC":2}


def normalize(vn):
    low = vn.lower()
    for pfx in _DE_PFX:
        if low.startswith(pfx) and len(low) > len(pfx) + 2:
            low = low[len(pfx):]
            break
    low = re.sub(r'^u\d+_', '', low)
    parts = low.split("_")
    meaningful = [p for p in parts if len(p) > 2 and p not in _NOISE]
    if len(meaningful) >= 2:
        return meaningful[-1]
    return meaningful[0] if meaningful else (parts[-1] if parts else low)


def rule_sod(vn, fm, cat):
    low = vn.lower()
    sev = 5
    for sv, kws in _SEV_KW:
        if any(k in low for k in kws):
            sev = sv
            break
    if cat == "External":
        sev = min(sev + 1, 10)
    occ = _OCC.get(fm, 3)
    if cat == "External":
        occ = min(occ + 1, 10)
    det = _DET.get(fm, 5)
    effects = {
        "MORE":    f"{vn} 과도한 값으로 인한 시스템 오동작",
        "LESS":    f"{vn} 부족한 값으로 인한 기능 미작동",
        "CORRUPT": f"{vn} 잘못된 값으로 인한 제어 오류",
        "STUCK":   f"{vn} 고착으로 인한 기능 상실",
        "EARLY":   f"{vn} 조기 발생으로 인한 순서 오류",
        "LATE":    f"{vn} 지연으로 인한 응답 지체",
        "ERRATIC": f"{vn} 불규칙 변동으로 인한 제어 불안정",
    }
    return {
        "severity": sev, "occurrence": occ, "detection": det,
        "rpn": sev * occ * det,
        "effect_system": effects.get(fm, f"{vn} 비정상 동작"),
        "preventive_action": "설계 리뷰 및 코드 인스펙션 수행",
    }


def sb_get(table, params):
    rows, offset = [], 0
    while True:
        r = requests.get(f"{SB_URL}/rest/v1/{table}",
                         headers={**SB_H, "Range": f"{offset}-{offset+999}"},
                         params=params, verify=False)
        data = r.json()
        if not isinstance(data, list) or not data:
            break
        rows.extend(data)
        if len(data) < 1000:
            break
        offset += 1000
    return rows


def sb_patch(item_id, patch):
    r = requests.patch(f"{SB_URL}/rest/v1/fmea_items",
                       headers=SB_H,
                       params={"id": f"eq.{item_id}"},
                       json=patch, verify=False)
    return r.status_code in (200, 204)


def build_lookup():
    """완성 프로젝트에서 S/O/D 교차 참조 테이블 구축"""
    print("  교차 참조 테이블 구축 중...")
    lookup = defaultdict(dict)
    for src_id in SRC_IDS:
        rows = sb_get("fmea_items", {
            "project_id": f"eq.{src_id}",
            "severity": "not.is.null",
            "select": "variable_name,failure_mode,severity,occurrence,detection,rpn,effect_system,preventive_action",
        })
        for row in rows:
            norm = normalize(row["variable_name"])
            fm = row.get("failure_mode") or "None"
            existing = lookup[norm].get(fm)
            if existing is None or (row.get("rpn") or 0) > (existing.get("rpn") or 0):
                lookup[norm][fm] = row
    print(f"  교차 참조: {len(lookup)}개 정규화 변수명")
    return dict(lookup)


def process_project(project_id, project_name, lookup):
    print(f"\n[{project_name}] S/O/D 미입력 항목 조회...")
    todo = sb_get("fmea_items", {
        "project_id": f"eq.{project_id}",
        "severity": "is.null",
        "select": "id,variable_name,failure_mode,category",
    })
    print(f"  미입력: {len(todo)}개")
    if not todo:
        print("  완료 상태.")
        return 0, 0

    cross_done = 0
    rule_done = 0

    def apply_one(item):
        nonlocal cross_done, rule_done
        norm = normalize(item["variable_name"])
        fm = item.get("failure_mode") or "None"
        cat = item.get("category") or "Internal"

        # 1단계: 교차 참조
        src = lookup.get(norm, {}).get(fm) or lookup.get(norm, {}).get("None")
        if src:
            patch = {
                "severity":          src["severity"],
                "occurrence":        src["occurrence"],
                "detection":         src["detection"],
                "effect_system":     src.get("effect_system"),
                "preventive_action": src.get("preventive_action"),
            }
            return sb_patch(item["id"], patch), "cross"

        # 2단계: 규칙 기반
        sod = rule_sod(item["variable_name"], fm if fm != "None" else "MORE", cat)
        sod.pop("rpn", None)
        return sb_patch(item["id"], sod), "rule"

    with concurrent.futures.ThreadPoolExecutor(max_workers=30) as ex:
        futures = {ex.submit(apply_one, item): item for item in todo}
        for i, f in enumerate(concurrent.futures.as_completed(futures), 1):
            ok, method = f.result()
            if ok:
                if method == "cross":
                    cross_done += 1
                else:
                    rule_done += 1
            if i % 100 == 0:
                print(f"  진행: {i}/{len(todo)} (교차:{cross_done} 규칙:{rule_done})", end="\r")

    print(f"  완료: 교차참조={cross_done}개, 규칙기반={rule_done}개")
    return cross_done, rule_done


def main():
    # 처리할 프로젝트 목록
    all_projects = [
        ("0715f883-d3a1-4ddd-8a3b-d3071da9ed3e", "JG1 SBW SW FMEA"),
        ("32cc5148-96b5-42c5-a3dc-2b8fa3b9f93e", "GN7_FL"),
        ("89dc5818-2435-4d09-a1a9-36aea664d11d", "LQ2"),
        ("70f74c19-66b6-4b2e-a2b3-1ee04dd1b101", "TK1"),
        ("a43d7d4e-b104-4e90-ab55-04c7b31aa3e7", "SX3"),
    ]

    # 인자로 특정 프로젝트만 지정 가능
    if len(sys.argv) > 1:
        names = [a.upper() for a in sys.argv[1:]]
        projects = [(pid, name) for pid, name in all_projects if name.upper() in names]
    else:
        projects = all_projects

    print("=" * 60)
    print("S/O/D 빠른 일괄 채우기")
    print(f"대상: {[n for _,n in projects]}")
    print("=" * 60)

    lookup = build_lookup()

    total_cross, total_rule = 0, 0
    for pid, name in projects:
        c, r = process_project(pid, name, lookup)
        total_cross += c
        total_rule += r

    print(f"\n전체 완료: 교차참조={total_cross}개, 규칙기반={total_rule}개, 합계={total_cross+total_rule}개")


if __name__ == "__main__":
    main()
