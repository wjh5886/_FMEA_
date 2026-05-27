"""
LQ2 SW FMEA 업데이트 스크립트
1. Safety Goal (SG01~SG03) → safety_goals 테이블
2. Safety Mechanism (SM01~SM23) → safety_mechanisms 테이블
3. SW_FMEA 시트 → fmea_items 업데이트 (S/O/D, 상세내용)
4. effect_safety_goal 기준 safety_goal_id 연결
"""
import sys, re, json, requests, urllib3, openpyxl
urllib3.disable_warnings()
sys.stdout.reconfigure(encoding='utf-8')

# ── 설정 ──────────────────────────────────────────────
URL      = "https://itzgdbeiyvodhfhmvrfw.supabase.co"
SUPA_KEY = ("eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Iml0"
            "emdkYmVpeXZvZGhmaG12cmZ3Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzY3NDE2MzcsImV4cCI"
            "6MjA5MjMxNzYzN30.iqzQr-3Lqf1O4UHFe9euTLyeIyBXreLPoSzUdtEaNP8")
H        = {"apikey": SUPA_KEY, "Authorization": f"Bearer {SUPA_KEY}", "Content-Type": "application/json"}
PROJ_ID  = "89dc5818-2435-4d09-a1a9-36aea664d11d"  # LQ2

SW_FMEA_FILE = r"E:\claude\FMEA\LQ2_SW_FMEA.xlsx"

# ── 헬퍼 ──────────────────────────────────────────────
def supa_get(path, params=""):
    r = requests.get(f"{URL}/rest/v1/{path}{params}", headers=H, verify=False)
    return r.json()

def supa_post(path, data):
    r = requests.post(f"{URL}/rest/v1/{path}", json=data, headers={**H, "Prefer": "return=representation"}, verify=False)
    return r

def supa_patch(path, data):
    r = requests.patch(f"{URL}/rest/v1/{path}", json=data, headers=H, verify=False)
    return r

def normalize_varname(name: str) -> str:
    """변수명 정규화: 첫 번째 토큰(공백/줄바꿈/괄호 전) 소문자"""
    if not name: return ""
    n = str(name).strip()
    # 줄바꿈 앞
    n = n.split('\n')[0]
    # 괄호/공백 앞
    for sep in ['(', ' ', '\t']:
        if sep in n:
            n = n.split(sep)[0]
    return n.lower().strip()

# ══════════════════════════════════════════════════════
# 1단계: Safety Goals 입력
# ══════════════════════════════════════════════════════
print("=" * 60)
print("1단계: Safety Goals 입력")
print("=" * 60)

# Excel에서 Safety Goal 읽기
wb = openpyxl.load_workbook(SW_FMEA_FILE, read_only=True, data_only=True)
ws_sg = wb['Safety Goal']

sg_rows = []
for row in ws_sg.iter_rows(min_row=4, values_only=True):
    sg_id = str(row[1]).strip() if row[1] else ""
    asil  = str(row[2]).strip() if row[2] else ""
    desc  = str(row[3]).strip() if row[3] else ""
    if sg_id.startswith("SG"):
        sg_rows.append({"sg_id": sg_id, "asil": asil, "description": desc})

print(f"Excel에서 {len(sg_rows)}개 SG 발견:")
for sg in sg_rows:
    print(f"  {sg['sg_id']} (ASIL-{sg['asil']}): {sg['description'][:50]}")

# 기존 SG 확인
existing_sgs = supa_get("safety_goals", f"?project_id=eq.{PROJ_ID}&select=sg_id")
existing_sg_ids = {r['sg_id'] for r in existing_sgs}
print(f"\n기존 SG: {existing_sg_ids}")

# 없는 것만 INSERT
sg_id_map = {}  # sg_id -> uuid
inserted_sgs = 0
for sg in sg_rows:
    if sg['sg_id'] not in existing_sg_ids:
        payload = {
            "project_id":  PROJ_ID,
            "sg_id":       sg['sg_id'],
            "name":        sg['sg_id'],
            "description": sg['description'],
            "asil":        sg['asil'],  # 'A', 'B', 'C', 'D' 형식
        }
        r = supa_post("safety_goals", payload)
        if r.status_code < 300:
            new_sg = r.json()
            if isinstance(new_sg, list): new_sg = new_sg[0]
            sg_id_map[sg['sg_id']] = new_sg['id']
            print(f"  ✓ {sg['sg_id']} 추가 → {new_sg['id']}")
            inserted_sgs += 1
        else:
            print(f"  ✗ {sg['sg_id']} 오류: {r.text[:80]}")

# 항상 최신 SG 목록 다시 로드 (신규 추가된 것 포함)
all_sgs = supa_get("safety_goals", f"?project_id=eq.{PROJ_ID}&select=id,sg_id")
for row in all_sgs:
    sg_id_map[row['sg_id']] = row['id']

print(f"\nSG 처리 완료: {inserted_sgs}개 신규 추가, 총 {len(sg_id_map)}개")

# ══════════════════════════════════════════════════════
# 2단계: Safety Mechanisms 입력
# ══════════════════════════════════════════════════════
print("\n" + "=" * 60)
print("2단계: Safety Mechanisms 입력")
print("=" * 60)

# SM 목록 (System FMEA에서 추출)
SM_LIST = [
    # (sm_id, name)  ← type/coverage는 System FMEA에 명시되지 않아 기본값 사용
    # type = 'Detection' (기본), diagnostic_coverage = 'N/A'
    ("SM01", "HW Watchdog"),
    ("SM02", "Power Supply#1 Voltage Monitoring"),
    ("SM03", "Battery Voltage Monitoring"),
    ("SM04", "Redundancy CAN"),
    ("SM05", "CAN Monitoring(Main CAN)"),
    ("SM06", "Redundancy CAN"),
    ("SM07", "MCU Safety Mechanism (Clock)"),
    ("SM08", "MCU Safety Mechanism (Memory)"),
    ("SM09", "IGN Signal Redundancy"),
    ("SM10", "Power Supply#2 Voltage Monitoring"),
    ("SM11", "MCU Safety Mechanism (Voltage)"),
    ("SM12", "CRC"),
    ("SM13", "Alive Counter"),
    ("SM14", "Alive Counter"),
    ("SM15", "Button/Dial Stuck Fault Monitoring"),
    ("SM16", "Button/Dial Stuck Fault Monitoring"),
    ("SM17", "VCU CRC Fault Check"),
    ("SM19", "Flow Control Monitoring"),
    ("SM20", "Flow Control Monitoring"),
    ("SM21", "Flow Control Monitoring"),
    ("SM23", "IGN Power Redundancy"),
]

# 기존 SM 확인
existing_sms = supa_get("safety_mechanisms", f"?project_id=eq.{PROJ_ID}&select=sm_id,id")
existing_sm_ids = {r['sm_id'] for r in existing_sms}
sm_id_map = {r['sm_id']: r['id'] for r in existing_sms}
print(f"기존 SM: {existing_sm_ids}")

inserted_sms = 0
for sm_id, name in SM_LIST:
    if sm_id not in existing_sm_ids:
        payload = {
            "project_id":          PROJ_ID,
            "sm_id":               sm_id,
            "name":                name,
            "type":                "Detection",    # 기본값 (Safety Mechanism은 기본 Detection)
            "diagnostic_coverage": "N/A",          # 미확인 → N/A
        }
        r = supa_post("safety_mechanisms", payload)
        if r.status_code < 300:
            new_sm = r.json()
            if isinstance(new_sm, list): new_sm = new_sm[0]
            sm_id_map[sm_id] = new_sm['id']
            print(f"  ✓ {sm_id} {name} 추가")
            inserted_sms += 1
        else:
            print(f"  ✗ {sm_id} 오류: {r.text[:80]}")

print(f"\nSM 처리 완료: {inserted_sms}개 신규 추가, 총 {len(sm_id_map)}개")

# ══════════════════════════════════════════════════════
# 3단계: SW FMEA 시트 → DB 업데이트
# ══════════════════════════════════════════════════════
print("\n" + "=" * 60)
print("3단계: SW_FMEA 시트 → DB 업데이트")
print("=" * 60)

# Excel SW_FMEA 시트 파싱
ws_sw = wb['SW_FMEA']

# 헤더 맵 (12행): col_idx → field_name
# [1]No [2]SW Unit [3]Cat [4]VarName [5]VarType [6]FailMode
# [7]Detail [8]EffMod [9]EffSys [10]EffSG [11]S [12]Prev [13]O [14]DetAct [16]D
# [17]RPN [18]CmReq [19]CM [20]S_after [21]O_after [22]D_after [23]RPN_after
# [24]TargetDate [25]Resp [26]RefResult [27]FinDate

excel_items = []   # list of dicts
current_unit = None
current_var  = None
current_type = None

for i, row in enumerate(ws_sw.iter_rows(min_row=14, max_row=2300, values_only=True), start=14):
    if row[1]: current_unit = str(row[1]).strip()
    if row[3]: current_var  = str(row[3]).strip()
    if row[4]: current_type = str(row[4]).strip()

    fail_mode = str(row[5]).strip() if row[5] else None
    if not fail_mode:
        continue

    def cv(v):
        if v is None: return None
        s = str(v).strip()
        return s if s else None

    def ci(v):
        if v is None: return None
        try: return int(float(v))
        except: return None

    sg_raw = cv(row[9])
    # SG 정규화: "SG1/SG2 / SG3" → ["SG01","SG02","SG03"]
    sg_normalized = None
    sg_ids_linked = []
    if sg_raw and sg_raw.upper() not in ('X', '-', 'N/A', 'NONE'):
        sg_normalized = sg_raw
        # SG01/SG02/SG03 스타일로 정규화
        sg_nums = re.findall(r'SG\s*0?(\d+)', sg_raw, re.I)
        sg_ids_linked = [f"SG{n.zfill(2)}" for n in sg_nums]

    excel_items.append({
        "sw_unit":        current_unit,
        "var_name":       current_var,
        "var_type":       current_type,
        "var_key":        normalize_varname(current_var),
        "unit_key":       current_unit.lower() if current_unit else "",
        "failure_mode":   fail_mode.upper().strip(),
        "failure_detail": cv(row[6]),
        "effect_module":  cv(row[7]),
        "effect_system":  cv(row[8]),
        "effect_sg_raw":  sg_raw,
        "sg_ids":         sg_ids_linked,
        "severity":       ci(row[10]),
        "preventive_action": cv(row[11]),
        "occurrence":     ci(row[12]),
        "detection_action": cv(row[13]),
        "detection":      ci(row[15]),
        "cm_required":    cv(row[17]),
        "countermeasure": cv(row[18]),
        "severity_after": ci(row[19]),
        "occurrence_after": ci(row[20]),
        "detection_after": ci(row[21]),
        "target_date":    cv(row[23]),
        "responsibility": cv(row[24]),
        "reference_result": cv(row[25]),
        "finish_date":    cv(row[26]),
    })

wb.close()
print(f"Excel SW_FMEA: {len(excel_items)}개 항목 파싱 완료")
sod_count = sum(1 for it in excel_items if it['severity'])
sg_count  = sum(1 for it in excel_items if it['sg_ids'])
print(f"  S/O/D 있음: {sod_count}개, SG 연결됨: {sg_count}개")

# DB에서 LQ2 항목 전체 로드
print("\nDB LQ2 항목 로드 중...")
db_items = []
offset = 0
while True:
    r = requests.get(
        f"{URL}/rest/v1/fmea_items"
        f"?project_id=eq.{PROJ_ID}"
        f"&select=id,sw_unit_id,variable_name,failure_mode,effect_safety_goal,severity,"
        f"failure_detail,effect_module,effect_system,preventive_action,detection_action,"
        f"safety_goal_id"
        f"&order=created_at"
        f"&offset={offset}&limit=1000",
        headers=H, verify=False
    )
    batch = r.json()
    if not batch: break
    db_items.extend(batch)
    offset += len(batch)
    if len(batch) < 1000: break

# SW Unit ID → Name 맵
unit_data = requests.get(
    f"{URL}/rest/v1/sw_units?project_id=eq.{PROJ_ID}&select=id,name",
    headers=H, verify=False
).json()
unit_id2name = {u['id']: u['name'] for u in unit_data}

print(f"DB 항목: {len(db_items)}개")

# DB 항목에 unit_name, var_key 추가
for it in db_items:
    it['unit_name'] = unit_id2name.get(it['sw_unit_id'], '')
    it['var_key']   = normalize_varname(it['variable_name'])
    it['unit_key']  = it['unit_name'].lower()

# Excel → DB 매핑 (unit_key + var_key + failure_mode)
# Excel 빌드 index: (unit_key, var_key, failure_mode) → item
excel_index = {}
for it in excel_items:
    k = (it['unit_key'], it['var_key'], it['failure_mode'])
    if k not in excel_index:
        excel_index[k] = it

print(f"Excel 인덱스: {len(excel_index)}개 키")

# 매핑 & 업데이트
ok = err = skip = sg_linked = 0
no_match = []

for db_it in db_items:
    k = (db_it['unit_key'], db_it['var_key'], (db_it['failure_mode'] or '').upper().strip())
    ex = excel_index.get(k)

    if not ex:
        skip += 1
        no_match.append(k)
        continue

    # 업데이트 페이로드 구성
    # 원칙: DB에 이미 값이 있으면 덮어쓰지 않음 (빈 필드만 채움)
    # 예외: S/O/D는 Excel에 수동 입력된 값이 있으면 더 신뢰할 수 있으므로 업데이트
    payload = {}

    # 텍스트 필드: DB가 비어있을 때만 채움
    def fill_if_empty(db_key, ex_key):
        db_val = db_it.get(db_key)
        ex_val = ex.get(ex_key)
        if not db_val and ex_val:
            payload[db_key] = ex_val

    fill_if_empty('failure_detail',    'failure_detail')
    fill_if_empty('effect_module',     'effect_module')
    fill_if_empty('effect_system',     'effect_system')
    fill_if_empty('preventive_action', 'preventive_action')
    fill_if_empty('detection_action',  'detection_action')
    fill_if_empty('countermeasure',    'countermeasure')
    fill_if_empty('target_date',       'target_date')
    fill_if_empty('responsibility',    'responsibility')
    fill_if_empty('reference_result',  'reference_result')
    fill_if_empty('finish_date',       'finish_date')

    # S/O/D: Excel에 수동 입력된 값이 있으면 덮어씀 (수동 값이 AI 생성값보다 신뢰)
    if ex['severity'] and ex['occurrence'] and ex['detection']:
        payload['severity']   = ex['severity']
        payload['occurrence'] = ex['occurrence']
        payload['detection']  = ex['detection']
    # rpn은 generated column → 직접 설정 불가

    # After-action S/O/D: DB가 비어있을 때만
    if not db_it.get('severity_after') and ex['severity_after']:
        payload['severity_after']   = ex['severity_after']
    if not db_it.get('occurrence_after') and ex['occurrence_after']:
        payload['occurrence_after'] = ex['occurrence_after']
    if not db_it.get('detection_after') and ex['detection_after']:
        payload['detection_after']  = ex['detection_after']

    # SG 연결: DB에 없을 때만
    if not db_it.get('effect_safety_goal') and ex['effect_sg_raw']:
        payload['effect_safety_goal'] = ex['effect_sg_raw']

    if not db_it.get('safety_goal_id') and ex['sg_ids']:
        first_sg = ex['sg_ids'][0]
        if first_sg in sg_id_map:
            payload['safety_goal_id'] = sg_id_map[first_sg]
            sg_linked += 1

    if not payload:
        skip += 1
        continue

    # DB 업데이트
    r2 = supa_patch(f"fmea_items?id=eq.{db_it['id']}", payload)
    if r2.status_code < 300:
        ok += 1
    else:
        print(f"  ✗ 오류: {r2.text[:80]}")
        err += 1

print(f"\n업데이트 결과: ✓{ok} 성공 / ✗{err} 오류 / -{skip} 매핑없음")
print(f"  SG 연결: {sg_linked}개")

# 매핑 안 된 것 중 샘플
if no_match:
    print(f"\n매핑 안 된 DB 항목 샘플 (상위 10개):")
    shown_units = set()
    for k in no_match[:50]:
        if k[0] not in shown_units:
            print(f"  unit={k[0]} var={k[1]} mode={k[2]}")
            shown_units.add(k[0])
            if len(shown_units) >= 10: break

# ══════════════════════════════════════════════════════
# 4단계: effect_safety_goal → safety_goal_id 일괄 연결
# ══════════════════════════════════════════════════════
print("\n" + "=" * 60)
print("4단계: 기존 effect_safety_goal → safety_goal_id 연결")
print("=" * 60)

# 이미 safety_goal_id가 NULL이고 effect_safety_goal에 SG가 있는 항목들
unlinked = requests.get(
    f"{URL}/rest/v1/fmea_items"
    f"?project_id=eq.{PROJ_ID}"
    f"&safety_goal_id=is.null"
    f"&effect_safety_goal=not.is.null"
    f"&select=id,effect_safety_goal"
    f"&limit=3000",
    headers=H, verify=False
).json()

sg_link_ok = sg_link_skip = 0
for it in unlinked:
    sg_text = it.get('effect_safety_goal', '') or ''
    sg_nums = re.findall(r'SG\s*0?(\d+)', sg_text, re.I)
    if not sg_nums:
        sg_link_skip += 1
        continue
    first_sg = f"SG{sg_nums[0].zfill(2)}"
    if first_sg not in sg_id_map:
        sg_link_skip += 1
        continue

    r2 = supa_patch(f"fmea_items?id=eq.{it['id']}", {"safety_goal_id": sg_id_map[first_sg]})
    if r2.status_code < 300:
        sg_link_ok += 1
    else:
        sg_link_skip += 1

print(f"SG 연결: {sg_link_ok}개 완료, {sg_link_skip}개 스킵")

# ══════════════════════════════════════════════════════
# 최종 통계
# ══════════════════════════════════════════════════════
print("\n" + "=" * 60)
print("최종 통계")
print("=" * 60)

stats = requests.get(
    f"{URL}/rest/v1/fmea_items"
    f"?project_id=eq.{PROJ_ID}"
    f"&select=id,severity,safety_goal_id,safety_mechanism_id"
    f"&limit=3000",
    headers=H, verify=False
).json()

total = len(stats)
has_sod = sum(1 for it in stats if it.get('severity'))
has_sg_link = sum(1 for it in stats if it.get('safety_goal_id'))
has_sm_link = sum(1 for it in stats if it.get('safety_mechanism_id'))

sg_count_final = len(supa_get("safety_goals", f"?project_id=eq.{PROJ_ID}&select=sg_id"))
sm_count_final = len(supa_get("safety_mechanisms", f"?project_id=eq.{PROJ_ID}&select=sm_id"))

print(f"  총 FMEA 항목: {total}개")
print(f"  S/O/D 있음: {has_sod}개 ({has_sod*100//total if total else 0}%)")
print(f"  SG 연결됨: {has_sg_link}개 ({has_sg_link*100//total if total else 0}%)")
print(f"  SM 연결됨: {has_sm_link}개")
print(f"  Safety Goals: {sg_count_final}개")
print(f"  Safety Mechanisms: {sm_count_final}개")
print("\n완료!")
