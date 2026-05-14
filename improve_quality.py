"""
improve_quality.py — FMEA 텍스트 품질 개선
형식적 preventive_action / effect_system을 JG1·SX3_ICE_TEST 양질 데이터로 교체

사용법:
  python improve_quality.py           # GN7_FL, LQ2, TK1, SX3 전체
  python improve_quality.py "LQ2"    # 특정 프로젝트만
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

# 품질 좋은 소스 프로젝트
SRC_PROJECTS = [
    ("0715f883-d3a1-4ddd-8a3b-d3071da9ed3e", "JG1"),
    ("1ba10b41-4717-4d1b-b0df-1b91fe2c4870", "SX3_ICE_TEST"),
]

# 개선 대상 프로젝트
TARGET_PROJECTS = [
    ("32cc5148-96b5-42c5-a3dc-2b8fa3b9f93e", "GN7_FL"),
    ("89dc5818-2435-4d09-a1a9-36aea664d11d", "LQ2"),
    ("70f74c19-66b6-4b2e-a2b3-1ee04dd1b101", "TK1"),
    ("a43d7d4e-b104-4e90-ab55-04c7b31aa3e7", "SX3"),
]

# 형식적(저품질) 텍스트 판별 패턴
GENERIC_PATTERNS = [
    "설계 리뷰 및 코드 인스펙션",
    "리뷰 및 코드 인스펙션",
    "코드 인스펙션 수행",
]

_NOISE = {"i", "p", "b", "u1", "u8", "u16", "u32", "ctap", "ctdcm", "raw"}
_DE_PFX = ("de", "dg", "di", "dv", "dp", "dm")


def normalize(vn: str) -> str:
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


def is_generic(text: str | None) -> bool:
    if not text:
        return True
    return any(pat in text for pat in GENERIC_PATTERNS)


def quality_score(row: dict) -> int:
    """텍스트 품질 점수 (높을수록 좋음)"""
    score = 0
    pa = row.get("preventive_action") or ""
    es = row.get("effect_system") or ""
    if not is_generic(pa):
        score += len(pa)
    if len(es) > 30:
        score += len(es)
    return score


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


def sb_patch(item_id: str, patch: dict) -> bool:
    r = requests.patch(f"{SB_URL}/rest/v1/fmea_items",
                       headers=SB_H,
                       params={"id": f"eq.{item_id}"},
                       json=patch, verify=False)
    return r.status_code in (200, 204)


def build_lookup() -> dict:
    """소스 프로젝트에서 품질 좋은 항목 lookup 구축"""
    print("  소스 데이터 로드 중...")
    lookup: dict[tuple, dict] = {}  # (norm_vn, failure_mode) → best_row

    for pid, name in SRC_PROJECTS:
        rows = sb_get("fmea_items", {
            "project_id": f"eq.{pid}",
            "severity": "not.is.null",
            "select": "variable_name,failure_mode,effect_system,preventive_action,severity,occurrence,detection",
        })
        good = [r for r in rows if not is_generic(r.get("preventive_action"))]
        print(f"    {name}: {len(rows)}개 중 양질 {len(good)}개")

        for row in good:
            norm = normalize(row["variable_name"])
            fm = row.get("failure_mode") or "ANY"
            key = (norm, fm)
            existing = lookup.get(key)
            if existing is None or quality_score(row) > quality_score(existing):
                lookup[key] = row

    # failure_mode=ANY 도 등록 (None인 항목 매칭용)
    print(f"  lookup: {len(lookup)}개 키")
    return lookup


def find_best(norm: str, fm: str | None, lookup: dict) -> dict | None:
    """정규화 변수명 + failure_mode로 최적 소스 찾기"""
    fm_key = fm or "ANY"

    # 1. 정확 매칭
    row = lookup.get((norm, fm_key))
    if row:
        return row

    # 2. failure_mode ANY 매칭 (소스에 failure_mode 없는 경우)
    row = lookup.get((norm, "ANY"))
    if row:
        return row

    # 3. 부분 문자열 매칭 (norm이 key에 포함되거나 key가 norm에 포함)
    for (k_norm, k_fm), row in lookup.items():
        if k_fm not in (fm_key, "ANY"):
            continue
        if len(norm) > 3 and (norm in k_norm or k_norm in norm):
            return row

    return None


def process_project(pid: str, name: str, lookup: dict) -> dict:
    print(f"\n[{name}] 형식적 텍스트 항목 조회...")

    # 형식적 preventive_action 항목 조회
    all_items = sb_get("fmea_items", {
        "project_id": f"eq.{pid}",
        "severity": "not.is.null",
        "select": "id,variable_name,failure_mode,preventive_action,effect_system",
    })

    # 형식적 텍스트인 항목만 추출
    todo = [i for i in all_items if is_generic(i.get("preventive_action"))]
    print(f"  전체 {len(all_items)}개 중 형식적 텍스트 {len(todo)}개")

    if not todo:
        print("  개선 필요 항목 없음.")
        return {"improved": 0, "not_found": 0}

    improved = 0
    not_found = 0

    def update_one(item):
        norm = normalize(item["variable_name"])
        fm = item.get("failure_mode")
        src = find_best(norm, fm, lookup)
        if not src:
            return "not_found"

        patch = {}
        if not is_generic(src.get("preventive_action")):
            patch["preventive_action"] = src["preventive_action"]
        if src.get("effect_system") and len(src["effect_system"]) > len(item.get("effect_system") or ""):
            patch["effect_system"] = src["effect_system"]

        if not patch:
            return "not_found"

        ok = sb_patch(item["id"], patch)
        return "ok" if ok else "fail"

    with concurrent.futures.ThreadPoolExecutor(max_workers=20) as ex:
        futures = {ex.submit(update_one, item): item for item in todo}
        for i, f in enumerate(concurrent.futures.as_completed(futures), 1):
            result = f.result()
            if result == "ok":
                improved += 1
            else:
                not_found += 1
            if i % 200 == 0:
                print(f"  진행: {i}/{len(todo)} (개선:{improved} 미매칭:{not_found})", end="\r")

    print(f"  완료: 개선={improved}개, 미매칭={not_found}개")
    return {"improved": improved, "not_found": not_found}


def main():
    filter_names = [a.upper() for a in sys.argv[1:]] if len(sys.argv) > 1 else []
    targets = [(pid, name) for pid, name in TARGET_PROJECTS
               if not filter_names or name.upper() in filter_names]

    print("=" * 60)
    print("FMEA 텍스트 품질 개선 (교차 참조 재강화)")
    print(f"대상: {[n for _, n in targets]}")
    print("=" * 60)

    lookup = build_lookup()

    total_improved = 0
    for pid, name in targets:
        result = process_project(pid, name, lookup)
        total_improved += result["improved"]

    print(f"\n전체 완료: {total_improved}개 항목 개선")


if __name__ == "__main__":
    main()
