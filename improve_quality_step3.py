"""
improve_quality_step3.py — AI 기반 고위험 항목 품질 개선 (3단계)
RPN >= 200 항목 중 effect_system 이 짧거나 단순한 항목을 Claude AI 로 재작성

사용법:
  python improve_quality_step3.py
  python improve_quality_step3.py "LQ2"
  python improve_quality_step3.py --dry-run    # DB 업데이트 없이 결과만 출력
"""

import sys, os, re, json, time, requests, urllib3, concurrent.futures
urllib3.disable_warnings()

import anthropic

SB_URL = "https://itzgdbeiyvodhfhmvrfw.supabase.co"
SB_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Iml0emdkYmVpeXZvZGhmaG12cmZ3Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzY3NDE2MzcsImV4cCI6MjA5MjMxNzYzN30.iqzQr-3Lqf1O4UHFe9euTLyeIyBXreLPoSzUdtEaNP8"
SB_H = {
    "apikey": SB_KEY,
    "Authorization": f"Bearer {SB_KEY}",
    "Content-Type": "application/json",
    "Prefer": "return=minimal",
}

TARGET_PROJECTS = [
    ("32cc5148-96b5-42c5-a3dc-2b8fa3b9f93e", "GN7_FL"),
    ("89dc5818-2435-4d09-a1a9-36aea664d11d", "LQ2"),
    ("70f74c19-66b6-4b2e-a2b3-1ee04dd1b101", "TK1"),
    ("a43d7d4e-b104-4e90-ab55-04c7b31aa3e7", "SX3"),
]

RPN_THRESHOLD = 200
AI_WORKERS = 3  # concurrent Claude API calls

SYSTEM_PROMPT = """당신은 자동차 SBW(Shift By Wire) 시스템의 FMEA(Failure Mode and Effect Analysis) 전문가입니다.
HKMC(현대기아) SBW 제어 소프트웨어의 신호별 고장 모드 분석 문서를 작성합니다.

다음 규칙을 따르세요:
- 한국어로 작성
- effect_system: 해당 신호의 고장이 시스템/차량에 미치는 구체적 영향 서술 (2-3문장, 100자 이상)
  * 신호 이름을 그대로 반복하지 말고 기능 관점으로 서술
  * SBW 제어 로직, 변속 동작, 안전 기능에 대한 영향을 구체적으로 기술
- preventive_action: 해당 고장을 예방/감지하기 위한 구체적 SW 설계 대책 (2-3항목, 100자 이상)
  * 코드 리뷰/인스펙션 같은 형식적 문구 금지
  * 신호 범위 모니터링, 타임아웃 감지, 크로스체크, DTC 생성, 페일세이프 로직 등 구체적 기법 명시

출력 형식 (JSON only, 다른 텍스트 없음):
{"effect_system": "...", "preventive_action": "..."}"""


def needs_improvement(item: dict) -> bool:
    es = item.get("effect_system") or ""
    vn = item.get("variable_name") or ""
    # 짧은 경우
    if len(es) < 80:
        return True
    # 변수명을 그대로 포함한 단순 반복 패턴
    vn_short = vn.split("(")[0].strip()
    if vn_short and vn_short in es and len(es) < 200:
        return True
    return False


def clean_vn(vn: str) -> str:
    """변수명에서 범위 설명 부분만 남겨 읽기 쉽게"""
    return vn[:120]  # 너무 길면 잘라서 프롬프트 절약


def build_user_prompt(item: dict) -> str:
    vn = clean_vn(item.get("variable_name") or "")
    fm = item.get("failure_mode") or "N/A"
    cat = item.get("category") or "Internal"
    s = item.get("severity", "?")
    o = item.get("occurrence", "?")
    d = item.get("detection", "?")
    rpn = item.get("rpn", "?")
    cur_es = (item.get("effect_system") or "")[:300]
    cur_pa = (item.get("preventive_action") or "")[:300]

    return f"""신호명: {vn}
고장 모드: {fm}
분류: {cat}
S/O/D: {s}/{o}/{d} → RPN {rpn}
현재 effect_system: {cur_es or "(없음)"}
현재 preventive_action: {cur_pa or "(없음)"}

위 항목의 effect_system 과 preventive_action 을 개선하여 JSON으로 반환하세요."""


def ai_improve(client: anthropic.Anthropic, item: dict) -> dict | None:
    try:
        msg = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=600,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": build_user_prompt(item)}],
        )
        raw = msg.content[0].text.strip()
        # JSON 추출
        m = re.search(r'\{.*\}', raw, re.DOTALL)
        if not m:
            return None
        return json.loads(m.group())
    except Exception as e:
        print(f"  AI 오류 [{item.get('id','?')}]: {e}")
        return None


def sb_get_high_rpn(pid: str) -> list:
    rows, offset = [], 0
    while True:
        r = requests.get(f"{SB_URL}/rest/v1/fmea_items",
                         headers={**SB_H, "Range": f"{offset}-{offset+999}"},
                         params={
                             "project_id": f"eq.{pid}",
                             "rpn": f"gte.{RPN_THRESHOLD}",
                             "select": "id,variable_name,failure_mode,category,effect_system,preventive_action,severity,occurrence,detection,rpn",
                         }, verify=False)
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


def process_project(pid: str, name: str, client: anthropic.Anthropic, dry_run: bool) -> dict:
    print(f"\n[{name}] RPN>={RPN_THRESHOLD} 항목 조회...")
    all_items = sb_get_high_rpn(pid)
    todo = [i for i in all_items if needs_improvement(i)]
    skip = len(all_items) - len(todo)
    print(f"  전체 {len(all_items)}개 중 개선 필요 {len(todo)}개 (이미 양호 {skip}개 스킵)")

    if not todo:
        return {"improved": 0, "failed": 0}

    improved = failed = 0

    def do_one(item):
        result = ai_improve(client, item)
        if not result:
            return "fail"
        if dry_run:
            print(f"\n  [DRY] {item['variable_name'][:60]}")
            print(f"    ES: {result.get('effect_system','')[:100]}")
            print(f"    PA: {result.get('preventive_action','')[:100]}")
            return "ok"
        patch = {}
        if result.get("effect_system"):
            patch["effect_system"] = result["effect_system"]
        if result.get("preventive_action"):
            patch["preventive_action"] = result["preventive_action"]
        ok = sb_patch(item["id"], patch)
        return "ok" if ok else "fail"

    with concurrent.futures.ThreadPoolExecutor(max_workers=AI_WORKERS) as ex:
        futures = {ex.submit(do_one, item): item for item in todo}
        for i, f in enumerate(concurrent.futures.as_completed(futures), 1):
            res = f.result()
            if res == "ok":
                improved += 1
            else:
                failed += 1
            print(f"  진행: {i}/{len(todo)} (개선:{improved} 실패:{failed})", end="\r")

    print(f"  완료: 개선={improved}개, 실패={failed}개")
    return {"improved": improved, "failed": failed}


def main():
    args = sys.argv[1:]
    dry_run = "--dry-run" in args
    filter_names = [a.upper() for a in args if not a.startswith("--")]
    targets = [(pid, name) for pid, name in TARGET_PROJECTS
               if not filter_names or name.upper() in filter_names]

    print("=" * 60)
    print("FMEA 품질 개선 3단계 - Claude AI 고위험 항목 재작성")
    print(f"대상: {[n for _, n in targets]}  RPN>={RPN_THRESHOLD}")
    if dry_run:
        print("  [DRY-RUN 모드 - DB 업데이트 없음]")
    print("=" * 60)

    client = anthropic.Anthropic()

    total_improved = 0
    for pid, name in targets:
        result = process_project(pid, name, client, dry_run)
        total_improved += result["improved"]

    print(f"\n전체 완료: {total_improved}개 항목 AI 개선")


if __name__ == "__main__":
    main()
