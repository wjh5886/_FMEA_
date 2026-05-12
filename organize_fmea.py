"""
JG1 SBW FMEA 정리 스크립트
- SW_FMEA 시트에 소스코드 대조 컬럼 추가 (AB, AC, AD)
- 코드대조_요약 시트 추가
"""
import win32com.client, json, os, sys
from collections import defaultdict

FMEA_IN  = r"E:\claude\FMEA\JG1_SBW-Software_FMEA_1_정리.xlsx"
FMEA_OUT = r"E:\claude\FMEA\JG1_SBW-Software_FMEA_1_정리.xlsx"

# -----------------------------------------------------------------
# 1. Load cross-reference data
# -----------------------------------------------------------------
with open(r"E:\claude\FMEA\xref_results.json", encoding='utf-8') as f:
    xref = json.load(f)

with open(r"E:\claude\FMEA\fmea_data.json", encoding='utf-8') as f:
    fmea_records = json.load(f)

# Build lookup: (unit, base_varname) -> (found: bool, files: list)
var_lookup = {}
for unit, res in xref.items():
    for var, files in res['found']:
        key = (unit, var)
        # Keep only short relative paths
        short_files = [f.split('Source/')[-1].split('Source\\')[-1] for f in files[:3]]
        var_lookup[key] = ('Y', ', '.join(short_files))
    for var in res['not_found']:
        var_lookup[(unit, var)] = ('N', '')

# Known naming corrections
NAME_CORRECTIONS = {
    'SactSig':   'SActSig (CtAp_ShiftActSigChk.c)',
    'SactSigTo': 'SActSigTo (CtAp_ShiftActSigChk.c)',
    'LeverWarning_Dial':    'LvrWrngMsg (CtAp_SBWSigSet.c)',
    'LeverWarning_Sphere':  'LvrWrngMsg (CtAp_SBWSigSet.c)',
    'RetryWarning_Dial':    'SBW_MotorWarn.RetryWarning (CtAp_SBWSigSet.c)',
    'RetryWarning_Sphere':  'SBW_MotorWarn.RetryWarning (CtAp_SBWSigSet.c)',
    'MotorFaultWarning':    'MotorFaultPinChek (CtIoHwAb_IntfOut.c)',
    'BDC02Timeout':         'BDC_02_Timeout (CAN_Management/)',
    'BDC05Timeout':         'BDC_05_Timeout (CAN_Management/)',
    'CLU01Timeout':         'CLU_01_Timeout (CAN_Management/)',
    'SMK03Timeout':         'SMK_03_Timeout (CAN_Management/)',
    'PositionSensorInfo_PButtonFltSta': 'PButtonFltSta (Position_Management/)',
    'PositionSensorInfo_PButtonSta':    'PButtonSta (Position_Management/)',
    'SnapShot0200': 'CtAp_SnapShot0xFD60.c (DTC FD60)',
    'SnapShot0202': 'CtAp_SnapShot0xFD51.c (DTC FD51)',
    'SnapShot0203': 'CtAp_SnapShot0xFD53.c (DTC FD53)',
    'SnapShot0204': 'CtAp_SnapShot0xFD54.c (DTC FD54)',
    'SnapShot0205': 'CtAp_SnapShot0xFD57.c (DTC FD57)',
    'SnapShot0206': 'CtAp_SnapShot0xFD58.c (DTC FD58)',
    'SnapShot0207': 'CtAp_SnapShot0xFD59.c (DTC FD59)',
    'SnapShot0208': 'CtAp_SnapShot0xFD5A.c (DTC FD5A)',
    'SnapShot0209': 'CtAp_SnapShot0xFD5B.c (DTC FD5B)',
    'SnapShot020A': 'CtAp_SnapShot0xFD5C.c (DTC FD5C)',
    'SnapShot020B': 'CtAp_SnapShot0xFD50.c (DTC FD50)',
    'SnapShot020D': 'CtAp_SnapShot0xFD53.c (DTC FD53)',
    'SnapShot020E': 'CtAp_SnapShot0xFD60.c (DTC FD60)',
    'HallSnrFltInfo': 'Generated/Bsw_Output (Dem)',
    'HallSnrFltVal':  'Generated/Bsw_Output (Dem)',
    'SlaveAddress':   'Generated RTE (I2C)',
    'TransmitLength': 'Generated RTE (I2C)',
}

# -----------------------------------------------------------------
# 2. Open Excel via COM
# -----------------------------------------------------------------
print("Excel 열기...", flush=True)
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(FMEA_IN)
    ws = wb.Sheets('SW_FMEA')
    total_rows = ws.UsedRange.Rows.Count
    print(f"SW_FMEA: {total_rows}행", flush=True)

    # -----------------------------------------------------------------
    # 3. Add header row for new columns (row 12, cols 28, 29, 30)
    # -----------------------------------------------------------------
    COL_FOUND  = 28   # AB
    COL_NOTE   = 29   # AC
    COL_FILES  = 30   # AD

    ws.Cells(12, COL_FOUND).Value  = "코드 발견"
    ws.Cells(12, COL_NOTE).Value   = "코드 실제 변수명 / 비고"
    ws.Cells(12, COL_FILES).Value  = "소스 파일"

    # Style headers
    for col in [COL_FOUND, COL_NOTE, COL_FILES]:
        cell = ws.Cells(12, col)
        cell.Font.Bold = True
        cell.Interior.Color = 0x2E75B6   # Blue (BGR: B=B6, G=75, R=2E → Excel uses BGR int)
        cell.Font.Color = 0xFFFFFF       # White
        cell.HorizontalAlignment = -4108  # xlCenter

    # -----------------------------------------------------------------
    # 4. Fill rows with cross-reference data
    # -----------------------------------------------------------------
    FAILURE_MODES = {'MORE','LESS','CORRUPT','OMISSION','COMMISSION',
                     'WRONG','EARLY','LATE','STUCK'}

    # Read bulk data
    data = ws.UsedRange.Value
    curr = {}
    filled = 0

    for r_idx in range(12, len(data)):
        row = data[r_idx]

        def clean(v):
            if v is None: return None
            return str(v).replace('\xa0', ' ').strip() or None

        no   = clean(row[0])
        unit = clean(row[1])
        var  = clean(row[3])
        fm   = clean(row[5])

        if no:   curr['no']   = no
        if unit: curr['unit'] = unit
        if var:  curr['var']  = var

        if fm not in FAILURE_MODES:
            continue

        cur_unit = curr.get('unit', '')
        cur_var  = curr.get('var', '')
        if not cur_var:
            continue

        # Extract base variable name
        base = cur_var.split('\n')[0].strip().split('(')[0].strip()
        base = base.split('.')[0].split('[')[0].strip()

        # Look up
        key = (cur_unit, base)
        found_val, files_val = var_lookup.get(key, ('?', ''))

        note_val = ''
        if found_val == 'N':
            note_val = NAME_CORRECTIONS.get(base, '미구현 또는 생성코드 확인 필요')
        elif found_val == 'Y':
            note_val = ''

        excel_row = r_idx + 1  # 1-based
        ws.Cells(excel_row, COL_FOUND).Value = found_val
        if note_val:
            ws.Cells(excel_row, COL_NOTE).Value = note_val
        if files_val:
            ws.Cells(excel_row, COL_FILES).Value = files_val

        # Color-code
        cell = ws.Cells(excel_row, COL_FOUND)
        if found_val == 'Y':
            cell.Interior.Color = 0xC6EFCE   # light green
            cell.Font.Color = 0x276221
        elif found_val == 'N':
            cell.Interior.Color = 0xFFC7CE   # light red
            cell.Font.Color = 0x9C0006

        filled += 1

    print(f"채운 행: {filled}", flush=True)

    # -----------------------------------------------------------------
    # 5. Add summary sheet "코드대조_요약"
    # -----------------------------------------------------------------
    SUMMARY_NAME = "코드대조_요약"
    # Remove if exists
    for i in range(wb.Sheets.Count, 0, -1):
        if wb.Sheets(i).Name == SUMMARY_NAME:
            wb.Sheets(i).Delete()
            break

    ws_sum = wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
    ws_sum.Name = SUMMARY_NAME

    # --- Header ---
    headers = ["SW Unit", "전체 변수", "코드 발견", "미발견", "커버리지(%)",
               "미발견 변수 목록", "비고"]
    for c, h in enumerate(headers, 1):
        cell = ws_sum.Cells(1, c)
        cell.Value = h
        cell.Font.Bold = True
        cell.Interior.Color = 0x2E75B6
        cell.Font.Color = 0xFFFFFF
        cell.HorizontalAlignment = -4108

    # --- Data rows ---
    row_idx = 2
    total_found = 0
    total_not = 0

    unit_order = sorted(xref.keys())
    for unit in unit_order:
        res = xref[unit]
        n_found = len(res['found'])
        n_not   = len(res['not_found'])
        n_total = n_found + n_not
        pct = int(100 * n_found / n_total) if n_total else 0

        not_found_str = ', '.join(sorted(res['not_found']))

        # Remarks
        remarks = []
        if n_not > 0:
            naming = [v for v in res['not_found'] if v in NAME_CORRECTIONS]
            unimpl = [v for v in res['not_found'] if v not in NAME_CORRECTIONS]
            if naming:
                remarks.append(f"명명불일치: {len(naming)}개")
            if unimpl:
                remarks.append(f"미구현/확인필요: {len(unimpl)}개")

        ws_sum.Cells(row_idx, 1).Value = unit.replace('\n', '')
        ws_sum.Cells(row_idx, 2).Value = n_total
        ws_sum.Cells(row_idx, 3).Value = n_found
        ws_sum.Cells(row_idx, 4).Value = n_not
        ws_sum.Cells(row_idx, 5).Value = pct
        ws_sum.Cells(row_idx, 6).Value = not_found_str
        ws_sum.Cells(row_idx, 7).Value = ' | '.join(remarks) if remarks else ''

        # Color by coverage
        pct_cell = ws_sum.Cells(row_idx, 5)
        if pct == 100:
            pct_cell.Interior.Color = 0xC6EFCE
            pct_cell.Font.Color = 0x276221
        elif pct >= 80:
            pct_cell.Interior.Color = 0xFFEB9C
            pct_cell.Font.Color = 0x9C6500
        else:
            pct_cell.Interior.Color = 0xFFC7CE
            pct_cell.Font.Color = 0x9C0006

        total_found += n_found
        total_not   += n_not
        row_idx += 1

    # Total row
    ws_sum.Cells(row_idx, 1).Value = "합계"
    ws_sum.Cells(row_idx, 2).Value = total_found + total_not
    ws_sum.Cells(row_idx, 3).Value = total_found
    ws_sum.Cells(row_idx, 4).Value = total_not
    ws_sum.Cells(row_idx, 5).Value = int(100 * total_found / (total_found + total_not))
    for c in range(1, 8):
        ws_sum.Cells(row_idx, c).Font.Bold = True

    # Auto-fit columns
    ws_sum.Columns.AutoFit()

    # -----------------------------------------------------------------
    # 6. Add FMEA 완성도 현황 섹션 to summary
    # -----------------------------------------------------------------
    ws_sum.Cells(row_idx + 2, 1).Value = "FMEA 완성도 현황"
    ws_sum.Cells(row_idx + 2, 1).Font.Bold = True
    ws_sum.Cells(row_idx + 2, 1).Font.Size = 12

    status_data = [
        ("항목", "입력 수", "전체", "완성도"),
        ("총 FMEA 항목", 1460, 1460, "100%"),
        ("S (Severity)", 0, 1460, "0% ⚠"),
        ("O (Occurrence)", 0, 1460, "0% ⚠"),
        ("D (Detection)", 0, 1460, "0% ⚠"),
        ("RPN", 0, 1460, "0% ⚠"),
        ("Preventive Action", 0, 1460, "0% ⚠"),
        ("Detection Action", 0, 1460, "0% ⚠"),
        ("Effect on SG", 193, 1460, "13%"),
    ]

    for i, row_data in enumerate(status_data):
        for j, val in enumerate(row_data):
            cell = ws_sum.Cells(row_idx + 3 + i, j + 1)
            cell.Value = val
            if i == 0:
                cell.Font.Bold = True
                cell.Interior.Color = 0x404040
                cell.Font.Color = 0xFFFFFF

    ws_sum.Columns.AutoFit()

    # -----------------------------------------------------------------
    # 7. Save
    # -----------------------------------------------------------------
    print("저장 중...", flush=True)
    wb.Save()
    wb.Close(False)
    excel.Quit()
    print("완료!", flush=True)

except Exception as e:
    print(f"오류: {e}", flush=True)
    import traceback
    traceback.print_exc()
    try:
        excel.Quit()
    except:
        pass
