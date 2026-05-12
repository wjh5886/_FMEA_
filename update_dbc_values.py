"""
DBC VAL_ 을 참고해서 FMEA SW_FMEA 시트의
Interface(Variable) name 셀에 value description 을 업데이트한다.
"""
import re
import win32com.client

# ──────────────────────────────────────────────
# 1. DBC 파싱 – VAL_ 항목 수집
# ──────────────────────────────────────────────

def parse_val(dbc_file):
    val_map = {}
    with open(dbc_file, 'r', encoding='latin-1') as f:
        content = f.read()
    pattern = r'VAL_\s+\d+\s+(\w+)((?:\s+\d+\s+"[^"]*")+)\s*;'
    for m in re.finditer(pattern, content):
        sig = m.group(1)
        pairs = re.findall(r'(\d+)\s+"([^"]*)"', m.group(2))
        val_map[sig] = pairs
    return val_map

p1 = parse_val(r'E:\claude\FMEA\20260119_STD_DB_CAR_R2.0_2024_FD_P1_v25.11.01.dbc')
p2 = parse_val(r'E:\claude\FMEA\20260119_STD_DB_CAR_R2.0_2024_FD_P2_v25.11.01.dbc')
ALL_VALS = {**p1, **p2}

# ──────────────────────────────────────────────
# 2. FMEA 변수명 → DBC 신호명 수동 매핑
#    (prefix가 다르거나 이름이 바뀐 경우)
# ──────────────────────────────────────────────

MANUAL_MAP = {
    # FMEA 변수명(첫 줄, strip)       : DBC 신호명
    'TrmnlCtrlGrpStaBDCEV'            : 'SMK_TrmnlCtrlGrpStaBDCEV',
    'SMK_TrmnlCtrlGrpStaBDCEV'        : 'SMK_TrmnlCtrlGrpStaBDCEV',
    'SMK_TrmnlCtrlGrpStaBDC'          : 'SMK_TrmnlCtrlGrpStaBDC',
    'SMK_TrmnlCtrlStaBDC'             : 'SMK_TrmnlCtrlStaBDC',
    'SMK_TrmnlCtrlStaBDCEV'           : 'SMK_TrmnlCtrlStaBDCEV',
    'PwrOnModeSta'                     : 'SMK_PwrOnModeSta',

    # SSB
    'StrtStpBtnSw1Sta'                : 'SSB_StrtStpBtnSw1Sta',
    'SSB_StrtStpBtnSw1Sta'            : 'SSB_StrtStpBtnSw1Sta',
    'SSB_StrtStpBtnSw2Sta'            : 'SSB_StrtStpBtnSw2Sta',

    # RKE / DoorLock
    'RKESig'                           : 'RKE_BtnReq',
    'SMKSig'                           : 'DoorLock_PsvDrLkReq',
    'DrLockSta'                        : 'DoorLock_PsvDrLkReq',

    # USM
    'USM06Msg'                         : 'USM_SBW_RotateModeChangeReq',

    # HTCU / TCU / VCU
    'GetPSigSta'                       : 'HTCU_PrkReleaseReq',
    'GearPosSta'                       : 'VCU_GearPosSta',
    'ValGearSlctDis'                   : 'HTCU_GearSlctrDis',
    'gearPositionVcu'                  : 'VCU_GearPosSta',
    'SactSig'                          : 'SBW_GearSelSta',
    'SubSActSig'                       : 'SCU_FF_PosTarSta',
    'lvrPosInfo'                       : 'SBW_LvrPosSta',

    # ICC
    'FaceDetectStat'                   : 'ICC_FaceDetectStat',

    # MoodLamp
    'MoodLamp_MdLmpFadeSta'           : 'MoodLamp_MdLmpFadeSta',
    'MoodLamp_SlvBrgtnsVal'           : 'MoodLamp_SlvBrgtnsVal',
    'MoodLamp_SlvColor_X'             : 'MoodLamp_SlvColor_X',
    'MoodLamp_SlvColor_Y'             : 'MoodLamp_SlvColor_Y',
    'MoodLamp_SlvFadeInTimetVal'      : 'MoodLamp_SlvFadeInTimetVal',
    'MoodLamp_SlvFadeOutTimetVal'     : 'MoodLamp_SlvFadeOutTimetVal',
    'MdLmpFadeSta'                     : 'MoodLamp_MdLmpFadeSta',
    'SlvBrgtnsVal'                     : 'MoodLamp_SlvBrgtnsVal',
    'SlvXVal'                          : 'MoodLamp_SlvColor_X',
    'SlvYVal'                          : 'MoodLamp_SlvColor_Y',
    'SlvFadInTimetVal'                 : 'MoodLamp_SlvFadeInTimetVal',
    'SlvFadOutTimetVal'               : 'MoodLamp_SlvFadeOutTimetVal',

    # Lamp / BCM
    'InlTailLmpSta'                    : 'IntLamp_InlTailLmpSta',
    'IntTailLmpOnReqFlag'             : 'Lamp_IntTailLmpOnReq',
    'AutoLtSnrNightSta'               : 'Lamp_AutoLtSnsrNightSta',
    'AvTailLmpSta'                     : 'Lamp_AvTailLmpSta',

    # Door / Seat
    'DrvDrSwSta'                       : 'Warn_DrvDrSwSta',
    'AsstDrSwSta'                      : 'Warn_AsstDrSwSta',
    'RrLftDrSwSta'                     : 'Warn_RrLftDrSwSta',
    'RrRtDrSwSta'                      : 'Warn_RrRtDrSwSta',
    'DrvDrSwStaSBCM'                   : 'Warn_DrvDrSwSta_SBCM',
    'DrvDrSwSta_SBCM'                  : 'Warn_DrvDrSwSta_SBCM',
    'DrvStOccSta'                      : 'Warn_DrvStOccSta',

    # SBW 출력 신호
    'PosSta'                           : 'SBW_LvrPosSta',
    'UtilMode'                         : 'VCU_UtilModeActSta',
    'CLU_RhstaLvlSta'                 : 'CLU_RhstaLvlSta',
    'AutoBrightSta'                    : 'CLU_AutoBrightSta',
}

# ──────────────────────────────────────────────
# 3. 신호명 → DBC 검색 (자동 매칭 – prefix 제거)
# ──────────────────────────────────────────────

KNOWN_PREFIXES = [
    'SMK_', 'SSB_', 'RKE_', 'USM_', 'HTCU_', 'TCU_', 'VCU_',
    'ICC_', 'PDC_', 'BCM_', 'BDC_', 'CLU_', 'SBW_', 'SCU_FF_',
    'Lamp_', 'Warn_', 'WakeUp_', 'IntLamp_', 'MoodLamp_', 'PRAC_',
]

def find_dbc_signal(varname):
    """FMEA 변수명 → DBC VAL_ 항목 반환 (없으면 None)"""
    if varname in MANUAL_MAP:
        dbc_sig = MANUAL_MAP[varname]
        return ALL_VALS.get(dbc_sig)

    # 직접 매치
    if varname in ALL_VALS:
        return ALL_VALS[varname]

    # prefix 붙여서 검색
    for pfx in KNOWN_PREFIXES:
        candidate = pfx + varname
        if candidate in ALL_VALS:
            return ALL_VALS[candidate]

    # suffix 제거 후 검색 (점 포함 구조체 필드 무시)
    base = varname.split('.')[0]
    if base != varname:
        return find_dbc_signal(base)

    return None

# ──────────────────────────────────────────────
# 4. value description 텍스트 생성
# ──────────────────────────────────────────────

def make_val_text(signal_name, pairs):
    """
    예: TrmnlCtrlGrpStaBDCEV
        0 : Off
        1 : PowerOn
        ...
    중복 value(범위 표현 등)는 첫 번째만 사용.
    """
    seen = set()
    lines = [signal_name]
    for val, desc in pairs:
        if val in seen:
            continue
        seen.add(val)
        lines.append(f'{val} : {desc}')
    return '\n'.join(lines)

# ──────────────────────────────────────────────
# 5. Excel 업데이트
# ──────────────────────────────────────────────

SRC_FILE = r'E:\claude\FMEA\SBW_FMEA\JG1\JG1_SBW-Software FMEA.xlsx'
DST_FILE = r'E:\claude\FMEA\SBW_FMEA\JG1\JG1_SBW-Software FMEA_dbc_updated.xlsx'

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

wb = excel.Workbooks.Open(SRC_FILE)
ws = wb.Sheets('SW_FMEA')

header_row = 12   # No / SW Unit Name / Interface Category / Interface name / ...
data_start  = 14
max_row     = ws.UsedRange.Rows.Count + ws.UsedRange.Row - 1

COL_NAME = 4   # D: Interface(Variable) name
COL_TYPE = 5   # E: Interface(Variable) type

updated = 0
skipped = 0
no_val  = 0

for r in range(data_start, max_row + 1):
    cell_d = ws.Cells(r, COL_NAME)
    raw    = cell_d.Value
    if not raw:
        continue

    # 첫 줄이 신호/변수명
    first_line = str(raw).split('\n')[0].strip()
    # 괄호 앞 이름만 (예: "CLU_AutoBrightSta(0~ 200u)" 형태도 대응)
    varname = first_line.split('(')[0].strip()

    pairs = find_dbc_signal(varname)
    if pairs is None:
        no_val += 1
        # print(f'  [NO_VAL] row={r} sig={varname}')
        continue

    new_text = make_val_text(varname, pairs)

    if str(raw).strip() == new_text.strip():
        skipped += 1
        continue

    cell_d.Value = new_text
    updated += 1
    print(f'  [UPD] row={r:4d}  {varname}')

print(f'\n완료: updated={updated}, skipped(same)={skipped}, no_val={no_val}')

wb.SaveAs(DST_FILE)
wb.Close(False)
excel.Quit()
print(f'저장: {DST_FILE}')
