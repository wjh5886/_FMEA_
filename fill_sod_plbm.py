"""
PLBM_SV1 FMEA S/O/D 자동 채우기 (Claude AI)
- failure_detail, effect_module, potential_cause, effect_system
- severity, occurrence, detection
- preventive_action, detection_action
"""
import sys, time, json, re, requests, urllib3
import anthropic

urllib3.disable_warnings()
sys.stdout.reconfigure(encoding='utf-8')

# ── 설정 ─────────────────────────────────────────────────────────
import os
API_KEY  = os.environ.get("ANTHROPIC_API_KEY", "")  # export ANTHROPIC_API_KEY=sk-ant-...
URL      = "https://itzgdbeiyvodhfhmvrfw.supabase.co"
SUPA_KEY = ("eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Iml0"
            "emdkYmVpeXZvZGhmaG12cmZ3Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzY3NDE2MzcsImV4cCI"
            "6MjA5MjMxNzYzN30.iqzQr-3Lqf1O4UHFe9euTLyeIyBXreLPoSzUdtEaNP8")
H        = {"apikey": SUPA_KEY, "Authorization": f"Bearer {SUPA_KEY}", "Content-Type": "application/json"}
PROJ_ID  = "fa631b49-e52a-4834-a019-3175009b2ddf"
BATCH    = 10   # 한 번에 분석할 항목 수

client = anthropic.Anthropic(api_key=API_KEY)

SYSTEM_PROMPT = """You are an automotive SW FMEA expert for PLBM (Platform Lithium Battery Module) systems.

PLBM is a Battery Management System (BMS) in HEV/EV vehicles.
Safety domain: ISO 26262, ASIL-B to ASIL-D depending on function.
Key safety risks: battery overcharge/overdischarge/over-temperature causing thermal runaway or fire.

SW architecture: AUTOSAR BSW-based with SW-Cs (SWC_DiagnosticService, SWC_AppMode, SWC_DiagnosticMonitor, SWC_SL_Interface, SWC_WdgMTest, SWC_AppFoD, App_Lin, App_Example)

Severity (1-10, AIAG-VDA):
  10 = thermal runaway / fire / loss of vehicle control without warning
  8-9 = battery unprotected overcharge/overdischarge, EV shutdown without warning
  6-7 = critical diagnostic loss, safety monitoring disabled, HV warning
  4-5 = degraded battery performance, reduced range, fault lamp
  2-3 = diagnostic/communication only, no direct safety impact
  1   = cosmetic / logging only

Occurrence (1-10):
  8-10 = frequent, known design weakness
  5-7  = moderate complexity, occasional
  2-4  = well-designed, unit-tested SW
  1    = extremely rare / proven design

Detection (1-10):
  1-3 = CRC check + range monitor + watchdog + EOL test
  4-6 = partial diagnostic coverage, periodic check
  7-9 = limited detection, manual review only
  10  = no detection mechanism

Domain guidance:
- NvM read/write errors → S=5-6 (calibration/config loss), D=3 (NvM job status monitored)
- DEM event callbacks → S=5-7 (monitoring), O=2, D=2 (always detected)
- DCM diagnostic services → S=3-4 (diagnostics only), O=2, D=3
- ComM/CanSM/LinSM mode errors → S=5-6 (comm loss), O=3, D=3
- EcuM state errors → S=6-7 (power mgmt), O=2, D=3
- WdgM errors → S=7-8 (watchdog = safety mechanism), O=2, D=2
- CAN group signal errors (MORE/LESS) → S=6-7, O=3, D=4
- IoHwAb ADC/Digital → S=6-7 (sensor input), O=2, D=3
- Internal callback errors (E_OK/E_NOT_OK) → S=3-5, O=2, D=3
- CORRUPT on safety signals → +1~2 to S
- STUCK/ERRATIC → harder to detect, D+1~2
- EARLY/LATE → timing, S-1, D+1"""

def build_prompt(items: list[dict]) -> str:
    lines = []
    for i, it in enumerate(items):
        lines.append(
            f"[{i}] Unit={it['sw_unit']} Cat={it['category']} "
            f"Var={it['variable_name']} Type={it['variable_type'] or '-'} "
            f"Mode={it['failure_mode']} Range={it['signal_range'] or 'N/A'}"
        )
    text = "\n".join(lines)

    return f"""Analyze these {len(items)} PLBM FMEA items and return ONLY a JSON array (no markdown):

{text}

Return exactly {len(items)} objects:
[{{"idx":0,
  "failure_detail":"<Korean: From X state To Y state OR specific fault description, 1 sentence>",
  "effect_module":"<Korean: effect within the SW module, 1 sentence>",
  "potential_cause":"<Korean: root cause, 1 sentence>",
  "effect_system":"<Korean: vehicle/battery system level effect, 1 sentence>",
  "severity":<1-10>,
  "occurrence":<1-10>,
  "detection":<1-10>,
  "preventive_action":"<Korean: design/SW control measure>",
  "detection_action":"<Korean: monitoring/diagnostic method>"
}},...]"""

# ── 1. FMEA 항목 로드 ────────────────────────────────────────────
print("FMEA 항목 로드 중...")
all_items = []
offset = 0
while True:
    r = requests.get(
        f"{URL}/rest/v1/fmea_items"
        f"?project_id=eq.{PROJ_ID}"
        f"&severity=is.null"
        f"&select=id,sw_unit_id,category,variable_name,variable_type,failure_mode,signal_range"
        f"&order=sw_unit_id,variable_name,failure_mode"
        f"&offset={offset}&limit=1000",
        headers=H, verify=False,
    )
    batch = r.json()
    if not batch: break
    all_items.extend(batch)
    offset += len(batch)
    if len(batch) < 1000: break

print(f"대상 항목: {len(all_items)}개")

# ── 2. SW Unit 이름 로드 ─────────────────────────────────────────
r = requests.get(
    f"{URL}/rest/v1/sw_units?project_id=eq.{PROJ_ID}&select=id,name",
    headers=H, verify=False,
)
unit_map = {u["id"]: u["name"] for u in r.json()}

for it in all_items:
    it["sw_unit"] = unit_map.get(it["sw_unit_id"], "Unknown")

print(f"SW Unit {len(unit_map)}개 로드 완료")
print(f"\n배치 크기: {BATCH}, 예상 배치 수: {(len(all_items)+BATCH-1)//BATCH}")
print(f"예상 소요 시간: ~{(len(all_items)+BATCH-1)//BATCH * 3 // 60}분\n")

# ── 3. 배치 처리 ─────────────────────────────────────────────────
ok = 0
err = 0
for batch_start in range(0, len(all_items), BATCH):
    batch = all_items[batch_start:batch_start + BATCH]
    batch_no = batch_start // BATCH + 1
    total_batches = (len(all_items) + BATCH - 1) // BATCH

    try:
        msg = client.messages.create(
            model="claude-haiku-4-5-20251001",  # 빠르고 저렴한 Haiku 사용
            max_tokens=2048,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": build_prompt(batch)}],
        )
        raw = msg.content[0].text
        # 마크다운 코드블록 제거
        raw = re.sub(r'^```[a-z]*\n?', '', raw, flags=re.MULTILINE)
        raw = re.sub(r'\n?```$', '', raw, flags=re.MULTILINE)
        results = json.loads(raw.strip())

        # Supabase 업데이트
        for res in results:
            idx = res.get("idx", 0)
            if idx >= len(batch): continue
            item = batch[idx]
            payload = {
                "failure_detail":    res.get("failure_detail"),
                "effect_module":     res.get("effect_module"),
                "potential_cause":   res.get("potential_cause"),
                "effect_system":     res.get("effect_system"),
                "severity":          res.get("severity"),
                "occurrence":        res.get("occurrence"),
                "detection":         res.get("detection"),
                "preventive_action": res.get("preventive_action"),
                "detection_action":  res.get("detection_action"),
                "ai_generated":      True,
                "status":            "in_review",
            }
            # None 값 제거
            payload = {k: v for k, v in payload.items() if v is not None}
            r2 = requests.patch(
                f"{URL}/rest/v1/fmea_items?id=eq.{item['id']}",
                json=payload, headers=H, verify=False,
            )
            if r2.status_code < 300:
                ok += 1
            else:
                print(f"  ⚠ DB 오류: {r2.text[:60]}")
                err += 1

        if batch_no % 10 == 0 or batch_no == total_batches:
            print(f"  진행: {min(batch_start+BATCH, len(all_items))}/{len(all_items)} "
                  f"({batch_no}/{total_batches} 배치) ✓{ok} ✗{err}")

        # 레이트 리밋 방지
        time.sleep(0.5)

    except Exception as e:
        print(f"  배치 {batch_no} 오류: {e}")
        err += len(batch)
        time.sleep(2)

print(f"\n✓ 완료: {ok}개 업데이트, {err}개 오류")
print(f"총 {len(all_items)}개 중 {ok}개 ({ok*100//len(all_items) if all_items else 0}%) 처리")
