"""
PLBM_SV1 나머지 112개 signal_range 채우기
- 13개 CAN 메시지: DBC raw bit 범위 사용 (물리 범위 [0|0] → raw 계산)
- modeNotificationPort_InitState / WakeupEvent: AUTOSAR EcuM 표준값
- FD_LOCAL_CAN_NM_CGW_CCU: NM 메시지 (범위 없음, NULL 유지)
"""
import re, sys, requests, urllib3
from collections import defaultdict
urllib3.disable_warnings()
sys.stdout.reconfigure(encoding='utf-8')

URL     = "https://itzgdbeiyvodhfhmvrfw.supabase.co"
KEY     = ("eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Iml0"
           "emdkYmVpeXZvZGhmaG12cmZ3Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzY3NDE2MzcsImV4cCI"
           "6MjA5MjMxNzYzN30.iqzQr-3Lqf1O4UHFe9euTLyeIyBXreLPoSzUdtEaNP8")
H       = {"apikey": KEY, "Authorization": f"Bearer {KEY}", "Content-Type": "application/json"}
PROJ_ID = "fa631b49-e52a-4834-a019-3175009b2ddf"

DBC_PDC = "E:/claude/FMEA/PLBM_SV_V01_BSWv36-4/References/DB/20250528_STD_LOCAL_PDC_2023_FD_Local_v25.05.01.dbc"
DBC_CAR = ("E:/claude/FMEA/HKMC_JG1_SBW_R44_4.2109.00_26210_Mood_R1 1/References/DB/"
           "20240530_STD_DB_CAR_R2.0_2024_FD_P1_v24.05.01_R05_Change_260313.dbc")

# ── 1. DBC 파싱: raw bit 범위 포함 ─────────────────────────────────
SG_RAW = re.compile(r'^\s*SG_\s+(\S+)\s*:\s*\d+\|(\d+)@\S+\s+\(([^,]+),([^)]+)\)\s*\[([^\|]+)\|([^\]]+)\]\s*"([^"]*)"')
BO_RE  = re.compile(r'^BO_\s+\d+\s+(\S+)\s*:')

def fmt(v):
    return str(int(v)) if v == int(v) else f"{v:.4g}"

def parse_dbc_full(path):
    """물리 범위가 있으면 물리 범위, 없으면 raw bit 범위 반환"""
    msg_signals = defaultdict(list)
    cur_msg = None
    for line in open(path, encoding='utf-8', errors='replace'):
        bm = BO_RE.match(line)
        if bm:
            cur_msg = bm.group(1).rstrip(':')
            continue
        if cur_msg:
            m = SG_RAW.match(line)
            if m:
                name, bits, factor, offset, rmin, rmax, unit = m.groups()
                try:
                    bits_n = int(bits)
                    f, o = float(factor), float(offset)
                    lo, hi = float(rmin), float(rmax)
                except:
                    continue
                if lo == 0 and hi == 0:
                    # raw bit 범위
                    raw_max = (1 << bits_n) - 1
                    rng = f"0 ~ {raw_max}"
                else:
                    # 물리 범위
                    phys_lo = lo * f + o
                    phys_hi = hi * f + o
                    rng = f"{fmt(phys_lo)} ~ {fmt(phys_hi)}"
                    if unit.strip():
                        rng += f" {unit.strip()}"
                msg_signals[cur_msg].append(f"{name}: {rng}")
    return msg_signals

print("DBC 파싱 중...")
sigs_pdc = parse_dbc_full(DBC_PDC)
sigs_car = parse_dbc_full(DBC_CAR)
print(f"  PDC_LOCAL: {len(sigs_pdc)}개 메시지")
print(f"  CAR:       {len(sigs_car)}개 메시지")

# ── 2. PORT_RANGE 매핑 구성 ─────────────────────────────────────────
PORT_RANGE = {}

# CAN 메시지: PDC_LOCAL DBC 우선, 없으면 CAR DBC
for db in [sigs_pdc, sigs_car]:
    for msg_name, parts in db.items():
        if not parts:
            continue
        for prefix in ["Gr_MsgGr_E2E_FD_LOCAL_CAN_", "Gr_MsgGr_FD_LOCAL_CAN_"]:
            vname = prefix + msg_name
            if vname not in PORT_RANGE:
                PORT_RANGE[vname] = " | ".join(parts)

# EcuM 모드 인터페이스
PORT_RANGE["modeNotificationPort_InitState"] = (
    "ECUM_STATE_STARTUP(0x04) / ECUM_STATE_APP_RUN(0x32) / "
    "ECUM_STATE_APP_POST_RUN(0x34) / ECUM_STATE_SHUTDOWN(0x44) / "
    "ECUM_STATE_SLEEP(0x52) / ECUM_STATE_OFF(0x80)"
)
PORT_RANGE["modeNotificationPort_WakeupEvent"] = (
    "0 ~ 4294967295 (uint32 bitmask, 각 bit = EcuM wakeup source)"
)

print(f"\nPORT_RANGE 총: {len(PORT_RANGE)}개")

# ── 3. Supabase에서 미채운 112개 로드 ───────────────────────────────
print("\n미채운 항목 로드 중...")
all_items = []
offset = 0
while True:
    r = requests.get(
        f"{URL}/rest/v1/fmea_items"
        f"?project_id=eq.{PROJ_ID}"
        f"&signal_range=is.null"
        f"&select=id,variable_name,variable_type"
        f"&offset={offset}&limit=1000",
        headers=H, verify=False,
    )
    batch = r.json()
    if not batch:
        break
    all_items.extend(batch)
    offset += len(batch)
    if len(batch) < 1000:
        break

print(f"미채운 항목: {len(all_items)}개")

# ── 4. 매칭 ────────────────────────────────────────────────────────
updates = []
no_match = defaultdict(int)

for item in all_items:
    vname = item.get("variable_name", "") or ""

    # NM 메시지 → NULL 유지
    if "NM_" in vname or vname.startswith("FD_LOCAL_CAN_NM"):
        no_match[vname] += 1
        continue

    rng = PORT_RANGE.get(vname)
    if rng:
        updates.append({"id": item["id"], "signal_range": rng})
    else:
        no_match[vname] += 1

print(f"\n업데이트 가능: {len(updates)}개")
print(f"매핑 없음:     {sum(no_match.values())}개")
for vn, cnt in sorted(no_match.items(), key=lambda x: -x[1])[:10]:
    print(f"  [{cnt:3d}개] {vn}")

if not updates:
    print("\n업데이트할 항목 없음.")
    import sys; sys.exit(0)

# ── 5. 샘플 출력 ────────────────────────────────────────────────────
print("\n[매핑 샘플]")
shown = set()
for u in updates:
    vn = next(i["variable_name"] for i in all_items if i["id"] == u["id"])
    if vn not in shown:
        shown.add(vn)
        print(f"  {vn:55s} → {u['signal_range'][:70]}")
    if len(shown) >= 10:
        break

# ── 6. 배치 업데이트 ─────────────────────────────────────────────────
print(f"\n업데이트 시작 ({len(updates)}개)...")
ok = 0
for i, u in enumerate(updates):
    r = requests.patch(
        f"{URL}/rest/v1/fmea_items?id=eq.{u['id']}",
        json={"signal_range": u["signal_range"]},
        headers=H, verify=False,
    )
    if r.status_code < 300:
        ok += 1
    else:
        print(f"  ⚠ {r.text[:80]}")
    if (i + 1) % 50 == 0:
        print(f"  진행: {i+1}/{len(updates)}")

print(f"\n✓ {ok}개 signal_range 업데이트 완료")
