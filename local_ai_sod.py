"""
local_ai_sod.py — Claude Code (MAX 플랜)으로 로컬에서 S/O/D 자동 채우기

사용법:
  python local_ai_sod.py SX3_ICE_TEST          # 프로젝트명
  python local_ai_sod.py 1ba10b41-4717-...     # 프로젝트 ID
"""

import sys, re, json, subprocess, time
import requests, urllib3
urllib3.disable_warnings()

SB_URL = "https://itzgdbeiyvodhfhmvrfw.supabase.co"
SB_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Iml0emdkYmVpeXZvZGhmaG12cmZ3Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzY3NDE2MzcsImV4cCI6MjA5MjMxNzYzN30.iqzQr-3Lqf1O4UHFe9euTLyeIyBXreLPoSzUdtEaNP8"
SB_H  = {
    "apikey": SB_KEY,
    "Authorization": f"Bearer {SB_KEY}",
    "Content-Type": "application/json",
    "Prefer": "return=representation",
}

BATCH_SIZE = 10  # 한 번에 Claude에게 보낼 항목 수


# ── Supabase 헬퍼 ──────────────────────────────────────────────────────────────

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


def sb_patch(table, params, data):
    r = requests.patch(f"{SB_URL}/rest/v1/{table}",
                       headers={**SB_H, "Prefer": "return=minimal"},
                       params=params, json=data, verify=False)
    return r.status_code in (200, 204)


def find_project(name_or_id: str) -> dict:
    # UUID 형식이면 ID로 검색
    if len(name_or_id) == 36 and name_or_id.count("-") == 4:
        rows = sb_get("projects", {"id": f"eq.{name_or_id}", "select": "id,name,vehicle_model"})
    else:
        rows = sb_get("projects", {"name": f"eq.{name_or_id}", "select": "id,name,vehicle_model"})
    if not rows:
        raise SystemExit(f"프로젝트를 찾을 수 없습니다: {name_or_id}")
    return rows[0]


# ── Claude Code 호출 ──────────────────────────────────────────────────────────

def analyze_batch(batch: list[dict], vehicle_model: str) -> list[dict]:
    items_text = "\n".join(
        f"{i+1}. SW={b.get('sw_unit_name','')}, var={b['variable_name']}, "
        f"mode={b['failure_mode']}, category={b.get('category','Internal')}, "
        f"range={b.get('signal_range') or '?'}"
        for i, b in enumerate(batch)
    )

    prompt = (
        f"당신은 {vehicle_model or 'SBW ECU'} SW FMEA 전문가입니다 (AIAG/VDA 기준).\n"
        f"아래 {len(batch)}개 항목의 S(Severity), O(Occurrence), D(Detection)(각 1~10)과\n"
        "effect_system(시스템 영향, 한국어), preventive_action(예방조치, 한국어)을 작성하세요.\n\n"
        f"{items_text}\n\n"
        "JSON 배열로만 응답 (마크다운 코드블록 없이, 순수 JSON만):\n"
        '[{"idx":1,"severity":7,"occurrence":3,"detection":4,'
        '"effect_system":"시스템 영향 설명","preventive_action":"예방조치 설명"},...]'
    )

    try:
        result = subprocess.run(
            ["claude", "-p", prompt],
            capture_output=True, text=True, timeout=180, encoding="utf-8"
        )
        text = result.stdout.strip()
        m = re.search(r'\[.*\]', text, re.DOTALL)
        if m:
            return json.loads(m.group())
        print(f"  [경고] JSON 파싱 실패. 응답:\n  {text[:200]}")
    except subprocess.TimeoutExpired:
        print("  [경고] Claude 응답 시간 초과 (180s)")
    except json.JSONDecodeError as e:
        print(f"  [경고] JSON 디코드 오류: {e}")
    except Exception as e:
        print(f"  [오류] {e}")
    return []


# ── 메인 ──────────────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) < 2:
        print("사용법: python local_ai_sod.py <프로젝트명 또는 ID>")
        sys.exit(1)

    project = find_project(sys.argv[1])
    project_id    = project["id"]
    vehicle_model = project.get("vehicle_model") or project["name"]
    print(f"프로젝트: {project['name']} ({project_id[:8]}...)")

    # S/O/D 미입력 항목 조회
    todo = sb_get("fmea_items", {
        "project_id": f"eq.{project_id}",
        "severity":   "is.null",
        "select":     "id,variable_name,failure_mode,category,signal_range,sw_unit_id",
    })

    # SW Unit 이름 매핑
    units = sb_get("sw_units", {"project_id": f"eq.{project_id}", "select": "id,name"})
    unit_map = {u["id"]: u["name"] for u in units}
    for row in todo:
        row["sw_unit_name"] = unit_map.get(row.get("sw_unit_id", ""), "")

    print(f"S/O/D 미입력: {len(todo)}개")
    if not todo:
        print("모두 입력됨. 종료.")
        return

    total_batches = (len(todo) + BATCH_SIZE - 1) // BATCH_SIZE
    done = 0

    for bi in range(total_batches):
        batch = todo[bi * BATCH_SIZE:(bi + 1) * BATCH_SIZE]
        print(f"\n배치 {bi+1}/{total_batches} ({len(batch)}개) 분석 중...", flush=True)

        results = analyze_batch(batch, vehicle_model)
        if not results:
            print("  응답 없음, 건너뜀")
            continue

        for res in results:
            idx = res.get("idx", 0) - 1
            if not (0 <= idx < len(batch)):
                continue
            item = batch[idx]
            s = res.get("severity")
            o = res.get("occurrence")
            d = res.get("detection")
            if not (s and o and d):
                continue
            patch = {
                "severity":          int(s),
                "occurrence":        int(o),
                "detection":         int(d),
                "effect_system":     res.get("effect_system") or None,
                "preventive_action": res.get("preventive_action") or None,
            }
            if sb_patch("fmea_items", {"id": f"eq.{item['id']}"}, patch):
                done += 1
                print(f"  OK {item['variable_name']} | {item['failure_mode']} -> S{s}/O{o}/D{d}")
            else:
                print(f"  FAIL DB 저장 실패: {item['variable_name']}")

        if bi < total_batches - 1:
            time.sleep(1)

    print(f"\n완료: {done}/{len(todo)}개 항목 업데이트")


if __name__ == "__main__":
    main()
