"""
DB의 JG1 suffix 매칭으로 SX3 S/O/D 채우기
- JG1 variable_name이 SX3 variable_name 안에 포함되면 매칭
- 여러 JG1 후보 중 가장 긴 변수명 우선 (가장 구체적 매칭)
- S/O/D/RPN + effect_system, preventive_action, safety_mechanism_text 복사
"""

import requests, urllib3
from collections import defaultdict
import concurrent.futures

urllib3.disable_warnings()

SX3_PROJECT_ID = "a43d7d4e-b104-4e90-ab55-04c7b31aa3e7"
SRC_PROJECTS = [
    "0715f883-d3a1-4ddd-8a3b-d3071da9ed3e",  # JG1 SBW SW FMEA (우선)
    "32cc5148-96b5-42c5-a3dc-2b8fa3b9f93e",  # GN7_FL
    "89dc5818-2435-4d09-a1a9-36aea664d11d",  # LQ2
    "70f74c19-66b6-4b2e-a2b3-1ee04dd1b101",  # TK1
]

URL = "https://itzgdbeiyvodhfhmvrfw.supabase.co"
KEY = ("eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Iml0"
       "emdkYmVpeXZvZGhmaG12cmZ3Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzY3NDE2MzcsImV4cCI"
       "6MjA5MjMxNzYzN30.iqzQr-3Lqf1O4UHFe9euTLyeIyBXreLPoSzUdtEaNP8")
H = {"apikey": KEY, "Authorization": f"Bearer {KEY}", "Content-Type": "application/json"}

COPY_FIELDS = ["severity", "occurrence", "detection",
               "effect_system", "preventive_action", "safety_mechanism_text"]


def fetch_pages(project_id, select_fields, extra_params=None):
    """Range 헤더로 전체 페이지 조회"""
    all_rows, offset = [], 0
    params = {"project_id": f"eq.{project_id}", "select": select_fields}
    if extra_params:
        params.update(extra_params)
    while True:
        r = requests.get(f"{URL}/rest/v1/fmea_items",
            headers={**H, "Range": f"{offset}-{offset+999}"},
            params=params, verify=False)
        data = r.json()
        if not data or not isinstance(data, list): break
        all_rows.extend(data)
        if len(data) < 1000: break
        offset += 1000
    return all_rows


def patch_one(item_id, patch):
    r = requests.patch(f"{URL}/rest/v1/fmea_items",
        headers={**H, "Prefer": "return=minimal"},
        params={"id": f"eq.{item_id}"},
        json=patch, verify=False)
    return r.status_code in (200, 204)


def main():
    print("=" * 60)
    print("SX3 S/O/D 채우기 (DB suffix 매칭)")
    print("=" * 60)

    # ── 1. SX3 미입력 항목 ──────────────────────────────────────
    print("\n[1/3] SX3 항목 로드...")
    sx3_items = fetch_pages(SX3_PROJECT_ID, "id,variable_name,failure_mode,severity")
    sx3_todo  = [i for i in sx3_items if i.get("severity") is None]
    print(f"  전체: {len(sx3_items)}개 / 미입력: {len(sx3_todo)}개")

    # ── 2. 소스 DB 로드 (S있는 항목만 서버 필터) ──────────────────
    print("\n[2/3] 소스 프로젝트 로드...")
    lookup: dict[tuple, list] = defaultdict(list)

    for pid in SRC_PROJECTS:
        rows = fetch_pages(pid,
            ",".join(["variable_name", "failure_mode"] + COPY_FIELDS),
            {"severity": "not.is.null"})
        for row in rows:
            key = (row["variable_name"].lower(), row["failure_mode"])
            lookup[key].append(row)
        print(f"  {pid[:8]}... → {len(rows)}개")

    print(f"  lookup 고유 키: {len(lookup)}개")

    # ── 3. 매칭 + 업데이트 ────────────────────────────────────────
    print("\n[3/3] 매칭 및 업데이트...")

    updates = []  # (item_id, patch)

    for sx3 in sx3_todo:
        sv = sx3["variable_name"].lower()
        sm = sx3["failure_mode"]

        # 후보 수집: JG1 변수명이 SX3 변수명의 부분 문자열
        candidates = []
        for (jv, jm), rows in lookup.items():
            if jm != sm or len(jv) < 5:
                continue
            if jv in sv:
                # 같은 변수+모드에 여러 행 → RPN 최댓값 행 선택
                best = max(rows, key=lambda r: r.get("rpn") or 0)
                candidates.append((len(jv), best))

        if not candidates:
            continue

        # 가장 구체적 매칭 (가장 긴 JG1 변수명)
        candidates.sort(key=lambda x: -x[0])
        best = candidates[0][1]

        patch = {f: best[f] for f in COPY_FIELDS if best.get(f) is not None}
        if patch:
            updates.append((sx3["id"], patch))

    print(f"  매칭 성공: {len(updates)}개 / 미매칭: {len(sx3_todo)-len(updates)}개")
    print(f"  커버리지: {len(updates)}/{len(sx3_todo)} = {len(updates)/len(sx3_todo)*100:.1f}%")

    if not updates:
        print("  업데이트할 항목 없음")
        return

    # 병렬 PATCH
    print("  업데이트 중...")
    done = 0
    with concurrent.futures.ThreadPoolExecutor(max_workers=20) as ex:
        futures = {ex.submit(patch_one, iid, patch): iid for iid, patch in updates}
        for f in concurrent.futures.as_completed(futures):
            if f.result():
                done += 1

    print(f"\n완료! {done}/{len(updates)}개 업데이트 성공")

    # 샘플 출력
    print("\n[샘플 결과]")
    for sx3 in sx3_todo[:5]:
        sv = sx3["variable_name"].lower()
        sm = sx3["failure_mode"]
        for (jv, jm), rows in lookup.items():
            if jm == sm and len(jv) >= 5 and jv in sv:
                best = max(rows, key=lambda r: r.get("rpn") or 0)
                print(f"  {sx3['variable_name'][:35]} [{sm}]")
                print(f"    ← JG1:{best['variable_name']} S={best['severity']} O={best['occurrence']} D={best['detection']}")
                break


if __name__ == "__main__":
    main()
