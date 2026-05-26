"""
PLBM_SV1 signal_range 업데이트
  Gr_MsgGr_E2E_FD_LOCAL_CAN_<MSG>  → DBC 메시지 <MSG> 신호 범위
  FD_LOCAL_CAN_<IFACE>             → DBImport SR-인터페이스 개별 신호
  ICULIN1_*                        → LIN 프레임 신호 (LDF)
  R_IoHwAb_If_AnaInDir_*           → ADC 아날로그 입력
  R_IoHwAb_If_DigDir_*             → 디지털 I/O
  R_IoHwAb_If_Pwm_*               → PWM 출력
"""
import re, sys, requests, urllib3
import lxml.etree as ET
from collections import defaultdict
urllib3.disable_warnings()
sys.stdout.reconfigure(encoding='utf-8')

URL = "https://itzgdbeiyvodhfhmvrfw.supabase.co"
KEY = ("eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Iml0"
       "emdkYmVpeXZvZGhmaG12cmZ3Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzY3NDE2MzcsImV4cCI"
       "6MjA5MjMxNzYzN30.iqzQr-3Lqf1O4UHFe9euTLyeIyBXreLPoSzUdtEaNP8")
H = {"apikey": KEY, "Authorization": f"Bearer {KEY}", "Content-Type": "application/json"}

PROJ_ID  = "fa631b49-e52a-4834-a019-3175009b2ddf"
BASE     = "E:/claude/FMEA/PLBM_SV_V01_BSWv36-4"
DBC_FILE = f"{BASE}/References/DB/20250528_STD_LOCAL_PDC_2023_FD_Local_v25.05.01.dbc"
LDF_FILE = f"{BASE}/References/DB/P-LBM_LIN_DB_v02_220214_V02.ldf"

def fmt(v):
    return str(int(v)) if v == int(v) else f"{v:.4g}"

def range_str(lo, hi, unit=""):
    s = f"{fmt(lo)} ~ {fmt(hi)}"
    return f"{s} {unit}".strip() if unit else s

# ── 1. DBC 파싱: 메시지명 → 신호 범위 목록 ──────────────────────
SG_RE = re.compile(
    r'^\s*SG_\s+(\S+)\s*:.*?\(([^,]+),([^)]+)\)\s*\[([^\|]+)\|([^\]]+)\]\s*"([^"]*)"'
)
BO_RE = re.compile(r'^BO_\s+\d+\s+(\S+)\s*:')

msg_signals = defaultdict(list)  # msg_name → list of range strings
cur_msg = None
for line in open(DBC_FILE, encoding='utf-8', errors='replace'):
    bm = BO_RE.match(line)
    if bm:
        cur_msg = bm.group(1).rstrip(':')
        continue
    if cur_msg:
        m = SG_RE.match(line)
        if m:
            name, factor, offset, rmin, rmax, unit = m.groups()
            try:
                f, o, lo, hi = float(factor), float(offset), float(rmin), float(rmax)
            except ValueError:
                continue
            if not (lo == 0 and hi == 0):
                phys_lo = lo * f + o
                phys_hi = hi * f + o
                msg_signals[cur_msg].append(
                    f"{name}: {range_str(phys_lo, phys_hi, unit.strip())}"
                )

print(f"DBC 메시지 중 범위 있는 것: {len(msg_signals)}개")

# ── 2. LDF 파싱: 프레임명 → 신호 범위 목록 ──────────────────────
LDF_SIG_RE = re.compile(r'^\s*(\w+)\s*:\s*(\d+)\s*,\s*(\d+)\s*,')
LDF_FRAME_RE = re.compile(r'^\s*(\w+)\s*:\s*\d+\s*,\s*\S+\s*,\s*\d+')

lin_frame_signals = defaultdict(list)  # frame_name → signal names
lin_signal_bits   = {}                 # signal_name → bit_size

in_frames = in_signals = False
for line in open(LDF_FILE, encoding='utf-8', errors='replace'):
    if 'Signals' in line and '{' in line:
        in_signals = True; continue
    if 'Frames' in line and '{' in line:
        in_frames = True; continue
    if '}' in line:
        in_signals = in_frames = False; continue
    if in_signals:
        m = LDF_SIG_RE.match(line)
        if m:
            sig_name, bits, _ = m.groups()
            lin_signal_bits[sig_name] = int(bits)
    if in_frames:
        # 프레임 헤더
        mf = LDF_FRAME_RE.match(line)
        if mf:
            cur_lin_frame = mf.group(1)
            continue
        # 프레임 내 신호
        ms = re.match(r'^\s+(\w+)\s*,\s*\d+', line)
        if ms:
            lin_frame_signals[cur_lin_frame].append(ms.group(1))

# LIN 프레임 → 범위 목록
lin_msg_signals = {}
for frame, sigs in lin_frame_signals.items():
    parts = []
    for sig in sigs:
        if sig in lin_signal_bits:
            bits = lin_signal_bits[sig]
            parts.append(f"{sig}: 0 ~ {(1<<bits)-1}")
    if parts:
        lin_msg_signals[frame] = " | ".join(parts)

print(f"LDF 프레임 중 범위 있는 것: {len(lin_msg_signals)}개")

# ── 3. FMEA variable_name → range_str 매핑 테이블 ────────────────
PORT_RANGE = {}

#  Gr_MsgGr_E2E_FD_LOCAL_CAN_<MSG>
for port_prefix in ["Gr_MsgGr_E2E_FD_LOCAL_CAN_", "Gr_MsgGr_FD_LOCAL_CAN_"]:
    for msg_name, parts in msg_signals.items():
        vname = port_prefix + msg_name
        PORT_RANGE[vname] = " | ".join(parts)

#  FD_LOCAL_CAN_<IFACE> (ex: FD_LOCAL_CAN_NM_CGW_CCU)
#  → DBImport ARXML로부터 개별 신호 범위
tree = ET.parse(f"{BASE}/Configuration/System/DBImport/FD_LOCAL_CAN.arxml")
ns = tree.getroot().tag.split('}')[0].lstrip('{')
dbc_sig_ranges = {}
for msg, parts in msg_signals.items():
    for p in parts:
        sig_part = p.split(':')[0].strip()
        rng_part = p.split(':', 1)[1].strip()
        dbc_sig_ranges[sig_part] = rng_part

for iface in tree.getroot().iter(f'{{{ns}}}SENDER-RECEIVER-INTERFACE'):
    iname = iface.findtext(f'{{{ns}}}SHORT-NAME')
    if not iname:
        continue
    parts = []
    for vdp in iface.iter(f'{{{ns}}}VARIABLE-DATA-PROTOTYPE'):
        elem = vdp.findtext(f'{{{ns}}}SHORT-NAME')
        if elem and elem in dbc_sig_ranges:
            parts.append(f"{elem}: {dbc_sig_ranges[elem]}")
    if parts:
        PORT_RANGE[iname] = " | ".join(parts)

#  ICULIN1 LIN 프레임
for frame_name, rng in lin_msg_signals.items():
    for prefix in ["Gr_MsgGr_E2E_ICULIN1_", "ICULIN1_"]:
        PORT_RANGE[prefix + frame_name] = rng

#  R_IoHwAb_If 표준 범위
PORT_RANGE["R_IoHwAb_If_AnaInDir"] = "0 ~ 65535 (ADC raw)"
PORT_RANGE["R_IoHwAb_If_DigDir"]   = "0 / 1 (digital)"
PORT_RANGE["R_IoHwAb_If_Pwm"]      = "0 ~ 100 %"

print(f"\n포트 범위 매핑 테이블: {len(PORT_RANGE)}개")
print("\n[샘플 15개]")
for k, v in list(PORT_RANGE.items())[:15]:
    print(f"  {k:60s} → {v[:60]}")

# ── 4. Supabase FMEA 항목 로드 ──────────────────────────────────
print("\nFMEA 항목 조회 중...")
all_items = []
offset = 0
while True:
    r = requests.get(
        f"{URL}/rest/v1/fmea_items?project_id=eq.{PROJ_ID}"
        f"&select=id,variable_name&offset={offset}&limit=1000",
        headers=H, verify=False,
    )
    batch = r.json()
    if not batch: break
    all_items.extend(batch)
    offset += len(batch)
    if len(batch) < 1000: break

print(f"FMEA 항목 {len(all_items)}개 로드")

# ── 5. 매칭 ─────────────────────────────────────────────────────
updates = []
for item in all_items:
    vname = item.get("variable_name", "") or ""
    rng = PORT_RANGE.get(vname)
    # prefix 매칭 (R_IoHwAb_If_AnaInDir_*)
    if rng is None:
        for prefix, std_rng in PORT_RANGE.items():
            if vname.startswith(prefix + '_') or vname.startswith(prefix + ' '):
                rng = std_rng; break
    if rng:
        updates.append({"id": item["id"], "signal_range": rng})

print(f"\n업데이트 대상: {len(updates)}개 / 전체 {len(all_items)}개")

if not updates:
    print("업데이트할 항목이 없습니다.")
    sys.exit(0)

print("\n[매칭 샘플]")
shown = set()
for u in updates:
    vn = next(i["variable_name"] for i in all_items if i["id"] == u["id"])
    if vn not in shown:
        shown.add(vn)
        print(f"  {vn:55s} → {u['signal_range'][:60]}")
    if len(shown) >= 12: break

# ── 6. 배치 업데이트 ─────────────────────────────────────────────
ok = 0
for i in range(0, len(updates), 100):
    batch = updates[i:i+100]
    for u in batch:
        r = requests.patch(
            f"{URL}/rest/v1/fmea_items?id=eq.{u['id']}",
            json={"signal_range": u["signal_range"]},
            headers=H, verify=False,
        )
        if r.status_code < 300:
            ok += 1
        else:
            print(f"  ⚠ {r.text[:80]}")
    print(f"  진행: {min(i+100, len(updates))}/{len(updates)}")

print(f"\n✓ {ok}개 signal_range 업데이트 완료")

# ── 7. 0~0 잘못된 범위 정리 ─────────────────────────────────────
# NM 메시지 등 DBC에서 [0|0]으로 정의된 신호는 물리 범위 없음 → NULL로 정리
print("\n0~0 잘못된 범위 정리 중...")
import re as _re
# 로컬에서 0~0 항목 식별 (모든 서브범위가 0 ~ 0 인 경우)
zero_pat = _re.compile(r'^[\w\s]+:\s*0\s*~\s*0')
all_range_items = []
offset2 = 0
while True:
    r = requests.get(
        f"{URL}/rest/v1/fmea_items?project_id=eq.{PROJ_ID}"
        f"&signal_range=not.is.null&select=id,signal_range&offset={offset2}&limit=1000",
        headers=H, verify=False,
    )
    batch = r.json()
    if not batch: break
    all_range_items.extend(batch)
    offset2 += len(batch)
    if len(batch) < 1000: break

cleared = 0
for item in all_range_items:
    rng = item.get("signal_range") or ""
    # 세미콜론 또는 파이프로 분리된 모든 서브범위가 "xxx: 0 ~ 0" 인 경우
    parts = [p.strip() for p in _re.split(r'[;|]', rng) if p.strip()]
    if parts and all(_re.search(r':\s*0\s*~\s*0\s*$', p) for p in parts):
        r2 = requests.patch(
            f"{URL}/rest/v1/fmea_items?id=eq.{item['id']}",
            json={"signal_range": None},
            headers=H, verify=False,
        )
        if r2.status_code < 300:
            cleared += 1

print(f"✓ {cleared}개 0~0 잘못된 범위 NULL로 정리 완료")
