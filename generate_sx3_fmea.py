"""
SX3 소스코드 → FMEA 자동 생성 스크립트
RTE 헤더에서 SW Component 인터페이스를 추출해 FMEA 항목을 생성하고
Claude AI로 S/O/D 값과 조치사항을 채워 Supabase에 저장
"""

import re
import os
import json
import time
import requests
import urllib3
from pathlib import Path
import anthropic

urllib3.disable_warnings()

# ─── 설정 ───────────────────────────────────────────────────────────────────
SX3_DIR = Path("E:/HKMC_SX3_SBW_ICE_R44-4.22.00_20251117_T037_Proto_2026-04-27 200109")
RTE_DIR  = SX3_DIR / "Generated/Bsw_Output/inc"

SUPABASE_URL = "https://itzgdbeiyvodhfhmvrfw.supabase.co"
SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Iml0emdkYmVpeXZvZGhmaG12cmZ3Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzY3NDE2MzcsImV4cCI6MjA5MjMxNzYzN30.iqzQr-3Lqf1O4UHFe9euTLyeIyBXreLPoSzUdtEaNP8"
ANTHROPIC_KEY = os.environ.get("ANTHROPIC_API_KEY", "")

PROJECT_NAME  = "SX3"
VEHICLE_MODEL = "SX3 ICE"

SB_HEADERS = {
    "apikey": SUPABASE_KEY,
    "Authorization": f"Bearer {SUPABASE_KEY}",
    "Content-Type": "application/json",
    "Prefer": "return=representation",
}

# FMEA에서 다룰 SW Component (진단/스냅샷 제외)
TARGET_COMPS = {
    "VBatStaChk",       # 배터리 전압 모니터링
    "IgnStaChk",        # IGN 스위치 상태
    "LvrPosChk",        # 레버 포지션 체크
    "LvrPosInfo",       # 레버 포지션 정보
    "ParkSWIn",         # 파킹 스위치 입력
    "HallDataSet",      # 홀센서 데이터
    "HallFltChk",       # 홀센서 고장
    "IdtCntl",          # IDT 제어
    "IdtCntlCdtChk",    # IDT 제어 조건
    "IdtFltChk",        # IDT 고장
    "CGWSigChk",        # CGW 신호 체크
    "SBWSigSet",        # SBW 신호 출력
    "DrvRdySigChk",     # 드라이브 준비 신호
    "ECUModeChk",       # ECU 모드
    "EcuModeCntl",      # ECU 모드 제어
    "ShiftActSigChk",   # 변속 액추에이터 신호
    "SysStaChk",        # 시스템 상태
    "LdoStaChk",        # LDO 상태
    "CalDataRead",      # 캘리브레이션 읽기
    "CalGapChk",        # 캘리브레이션 갭
    "RxMainCAN",        # 메인 CAN 수신
    "RxSubCAN",         # 서브 CAN 수신
    "TxMainCAN",        # 메인 CAN 송신
    "TxSubCAN",         # 서브 CAN 송신
    "CANBusOffChk",     # CAN BusOff 상태
    "PButtonSet",       # P버튼 입력
    "DimLvlSet",        # 조광 레벨
    "VehicleReset",     # 차량 리셋
    "VehicleReset_Manager",
}

# 외부 인터페이스 판별 키워드
EXTERNAL_KEYWORDS = ["IoHwAb", "PCAN_", "GCAN_", "ADC_", "Rx", "RxMain", "RxSub"]

VALID_MODES = ["MORE", "LESS", "CORRUPT", "EARLY", "LATE", "STUCK", "ERRATIC"]

# 데이터 타입별 기본 failure mode 집합
MODE_BY_TYPE = {
    "uint8":   ["MORE", "LESS", "CORRUPT", "STUCK"],
    "uint16":  ["MORE", "LESS", "CORRUPT", "STUCK", "ERRATIC"],
    "uint32":  ["MORE", "LESS", "CORRUPT", "STUCK"],
    "sint8":   ["MORE", "LESS", "CORRUPT", "STUCK"],
    "sint16":  ["MORE", "LESS", "CORRUPT", "STUCK"],
    "boolean": ["MORE", "LESS", "STUCK"],
    "default": ["MORE", "LESS", "CORRUPT", "STUCK"],
}

# ─── Supabase 헬퍼 ──────────────────────────────────────────────────────────
def sb_get(table, params=None):
    r = requests.get(f"{SUPABASE_URL}/rest/v1/{table}", headers=SB_HEADERS, params=params, verify=False)
    r.raise_for_status()
    return r.json()

def sb_post(table, data):
    r = requests.post(f"{SUPABASE_URL}/rest/v1/{table}", headers=SB_HEADERS, json=data, verify=False)
    r.raise_for_status()
    return r.json()

def sb_post_batch(table, rows, batch=200):
    inserted = 0
    for i in range(0, len(rows), batch):
        chunk = rows[i:i+batch]
        r = requests.post(f"{SUPABASE_URL}/rest/v1/{table}", headers=SB_HEADERS, json=chunk, verify=False)
        if r.status_code in (200, 201):
            inserted += len(chunk)
        else:
            print(f"  !! 배치 오류 {r.status_code}: {r.text[:200]}")
    return inserted

def get_or_create_project(name, vehicle_model):
    existing = sb_get("projects", {"name": f"eq.{name}", "select": "id,name"})
    if existing:
        print(f"  기존 프로젝트: {name} ({existing[0]['id']})")
        return existing[0]["id"]
    r = sb_post("projects", {"name": name, "vehicle_model": vehicle_model, "description": f"SBW FMEA - {name} (자동생성)"})
    pid = (r[0] if isinstance(r, list) else r)["id"]
    print(f"  새 프로젝트: {name} ({pid})")
    return pid

def ensure_sw_units(unit_names, project_id):
    existing = sb_get("sw_units", {"project_id": f"eq.{project_id}", "select": "id,name"})
    unit_map = {u["name"]: u["id"] for u in existing}
    for name in unit_names:
        if name and name not in unit_map:
            r = sb_post("sw_units", {"project_id": project_id, "name": name})
            uid = (r[0] if isinstance(r, list) else r)["id"]
            unit_map[name] = uid
    return unit_map

# ─── RTE 헤더 파싱 ─────────────────────────────────────────────────────────
def extract_de_name(port_and_element: str) -> str:
    """포트+요소 문자열에서 De 이후 실제 신호명 추출"""
    m = re.search(r'_De([A-Za-z0-9_]+)$', port_and_element)
    if m:
        return m.group(1)
    # De 없으면 마지막 언더스코어 이후
    parts = port_and_element.split('_')
    return parts[-1] if parts else port_and_element

def classify_category(port_name: str, comp_name: str) -> str:
    """External vs Internal 판별"""
    for kw in EXTERNAL_KEYWORDS:
        if kw in port_name:
            return "External"
    if comp_name in ("RxMainCAN", "RxSubCAN"):
        return "External"
    return "Internal"

def extract_data_type(func_decl: str) -> str:
    """함수 선언에서 데이터 타입 추출"""
    m = re.search(r'P2(?:VAR|CONST)\((\w+),', func_decl)
    if m:
        t = m.group(1).lower().replace("autosar_", "")
        return t
    m = re.search(r'CONST\((\w+),', func_decl)
    if m:
        return m.group(1).lower()
    return "uint8"

def parse_rte_headers() -> list[dict]:
    """모든 CtAp RTE 헤더에서 인터페이스 추출"""
    pattern = re.compile(
        r'extern FUNC\(Std_ReturnType, RTE_CODE\) '
        r'Rte_(Read|Write)_CtAp_(\w+?)_(.+?)\((.+?)\);'
    )
    signals = []
    seen = set()

    for hf in sorted(RTE_DIR.glob("Rte_CtAp_*.h")):
        if "_Type.h" in hf.name:
            continue
        comp = hf.stem.replace("Rte_CtAp_", "")
        if comp not in TARGET_COMPS:
            continue

        text = hf.read_text(encoding='utf-8', errors='ignore')
        for m in pattern.finditer(text):
            direction = m.group(1)   # Read or Write
            swcomp    = m.group(2)   # SW Component name
            port_elem = m.group(3)   # Port_Element
            func_sig  = m.group(4)   # function signature

            # Write는 출력 신호 - FMEA에서는 Read(입력) 중심
            # TxMainCAN Write는 External output으로 별도 처리
            if direction == "Write" and swcomp not in ("TxMainCAN", "TxSubCAN", "SBWSigSet"):
                continue

            de_name  = extract_de_name(port_elem)
            dtype    = extract_data_type(func_sig)
            category = classify_category(port_elem, swcomp)

            if direction == "Write":
                category = "Internal"  # output은 Internal 출력

            key = (swcomp, de_name)
            if key in seen:
                continue
            seen.add(key)

            signals.append({
                "sw_unit": swcomp,
                "variable": de_name,
                "type": dtype,
                "category": category,
                "port": port_elem,
                "direction": direction,
            })

    print(f"총 {len(signals)}개 인터페이스 추출")
    return signals

# ─── Claude AI 분석 ─────────────────────────────────────────────────────────
client = anthropic.Anthropic(api_key=ANTHROPIC_KEY)

def analyze_batch(signals_batch: list[dict]) -> list[dict]:
    """Claude API로 S/O/D + 조치사항 일괄 분석 (배치 처리)"""
    items_text = ""
    for i, sig in enumerate(signals_batch):
        items_text += (
            f"{i+1}. SW Unit: {sig['sw_unit']}, "
            f"Variable: {sig['variable']}, "
            f"Type: {sig['type']}, "
            f"Category: {sig['category']}, "
            f"Failure Mode: {sig['failure_mode']}\n"
        )

    prompt = f"""당신은 SBW(Shift-By-Wire) 자동차 ECU SW FMEA 전문가입니다.
아래 {len(signals_batch)}개의 FMEA 항목에 대해 각각의 S(Severity), O(Occurrence), D(Detection) 값(1~10)과
예방조치(preventive_action), 검출조치(detection_action)를 한국어로 작성하세요.

차량 모델: SX3 ICE (SBW 전자 변속기 제어기)
S/O/D 기준: AIAG/VDA FMEA 기준 (1=낮음, 10=매우 높음)

항목 목록:
{items_text}

반드시 아래 JSON 배열 형식으로만 응답하세요 (마크다운 없이):
[
  {{
    "idx": 1,
    "severity": 숫자,
    "occurrence": 숫자,
    "detection": 숫자,
    "effect_module": "SW Component 내부 영향",
    "effect_system": "시스템/차량 레벨 영향",
    "preventive_action": "예방 조치 (MISRA-C, SW 검증 등)",
    "detection_action": "검출 조치 (안전 메커니즘 등)"
  }},
  ...
]"""

    try:
        resp = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=4096,
            messages=[{"role": "user", "content": prompt}]
        )
        text = resp.content[0].text.strip()
        # JSON 추출
        m = re.search(r'\[.*\]', text, re.DOTALL)
        if m:
            return json.loads(m.group())
    except Exception as e:
        print(f"  AI 오류: {e}")
    return []

def build_fmea_rows(signals: list[dict]) -> list[dict]:
    """신호 목록 → FMEA 행 목록 (failure mode 분해)"""
    rows = []
    item_no = 1
    for sig in signals:
        dtype_key = sig['type'] if sig['type'] in MODE_BY_TYPE else 'default'
        modes = MODE_BY_TYPE[dtype_key]

        # CAN 타이밍 신호: LATE 추가
        if "CAN" in sig['sw_unit'] or "CAN" in sig['variable']:
            if "LATE" not in modes:
                modes = modes + ["LATE", "EARLY"]

        for mode in modes:
            rows.append({
                "item_no": str(item_no),
                "sw_unit": sig['sw_unit'],
                "variable": sig['variable'],
                "type": sig['type'],
                "category": sig['category'],
                "failure_mode": mode,
                # AI 분석 전 공백
                "effect_module": None,
                "effect_system": None,
                "severity": None,
                "occurrence": None,
                "detection": None,
                "preventive_action": None,
                "detection_action": None,
            })
        item_no += 1
    return rows

# ─── 메인 ──────────────────────────────────────────────────────────────────
def main():
    print("=" * 60)
    print("SX3 FMEA 자동 생성 시작")
    print("=" * 60)

    # 1. 인터페이스 추출
    print("\n[1/4] RTE 헤더 파싱...")
    signals = parse_rte_headers()

    # 2. FMEA 행 생성
    print("\n[2/4] FMEA 항목 생성...")
    fmea_rows = build_fmea_rows(signals)
    print(f"  총 {len(fmea_rows)}개 FMEA 항목 (failure mode 분해 후)")

    # 3. Claude AI로 S/O/D 분석 (신호 단위로 배치 처리, failure mode별 아님)
    print("\n[3/4] Claude AI 분석 (신호 단위 배치)...")
    # 신호별로 대표 1개만 AI 분석 (같은 신호의 다른 mode는 값 복사)
    unique_sigs = []
    seen_sigs = set()
    for row in fmea_rows:
        key = (row['sw_unit'], row['variable'])
        if key not in seen_sigs:
            seen_sigs.add(key)
            unique_sigs.append({**row, "failure_mode": row['failure_mode']})

    ai_results = {}
    batch_size = 10
    total_batches = (len(unique_sigs) + batch_size - 1) // batch_size

    for bi in range(total_batches):
        batch = unique_sigs[bi*batch_size:(bi+1)*batch_size]
        print(f"  배치 {bi+1}/{total_batches} ({len(batch)}개)...", end=" ", flush=True)
        results = analyze_batch(batch)
        for res in results:
            idx = res.get("idx", 0) - 1
            if 0 <= idx < len(batch):
                sig_key = (batch[idx]['sw_unit'], batch[idx]['variable'])
                ai_results[sig_key] = res
        print(f"완료 ({len(results)}개 응답)")
        time.sleep(1)  # API rate limit

    # AI 결과를 fmea_rows에 적용
    for row in fmea_rows:
        key = (row['sw_unit'], row['variable'])
        ai = ai_results.get(key, {})
        row['effect_module'] = ai.get('effect_module')
        row['effect_system'] = ai.get('effect_system')
        row['severity']      = ai.get('severity')
        row['occurrence']    = ai.get('occurrence')
        row['detection']     = ai.get('detection')
        row['preventive_action'] = ai.get('preventive_action')
        row['detection_action']  = ai.get('detection_action')

    # 4. Supabase 저장
    print("\n[4/4] Supabase 저장...")
    project_id = get_or_create_project(PROJECT_NAME, VEHICLE_MODEL)

    unit_names = list({r['sw_unit'] for r in fmea_rows})
    unit_map = ensure_sw_units(unit_names, project_id)
    print(f"  SW Units: {len(unit_map)}개")

    items = []
    for r in fmea_rows:
        s = r.get('severity')
        o = r.get('occurrence')
        d = r.get('detection')
        rpn = s * o * d if (s and o and d) else None
        items.append({
            "project_id":       project_id,
            "sw_unit_id":       unit_map.get(r['sw_unit']),
            "item_no":          r['item_no'],
            "category":         r['category'],
            "variable_name":    r['variable'],
            "variable_type":    r['type'],
            "failure_mode":     r['failure_mode'],
            "effect_module":    r['effect_module'],
            "effect_system":    r['effect_system'],
            "severity":         s,
            "occurrence":       o,
            "detection":        d,
            "preventive_action": r['preventive_action'],
            "detection_action":  r['detection_action'],
            "status":           "draft",
            "ai_generated":     bool(ai_results.get((r['sw_unit'], r['variable']))),
        })

    inserted = sb_post_batch("fmea_items", items)
    print(f"  삽입 완료: {inserted}/{len(items)}개")

    # 요약
    ai_filled = sum(1 for r in fmea_rows if ai_results.get((r['sw_unit'], r['variable'])))
    print("\n" + "=" * 60)
    print("완료!")
    print(f"  총 항목:     {len(items)}개")
    print(f"  AI 분석됨:  {ai_filled}개")
    print(f"  Supabase:  {inserted}개 저장")
    print(f"  프로젝트:  https://fmea-web.vercel.app/projects/{project_id}/fmea")

if __name__ == "__main__":
    main()
