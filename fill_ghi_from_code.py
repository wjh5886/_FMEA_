"""
JG1 SBW Software FMEA - G/H/I 컬럼 소스코드 기반 재작성
JG_FMEA 참조 형식:
  G: From : [현재 상태] / To : [이상 상태]
  H: [모듈 내 영향] - 오인식/오판단/오동작
  I: [시스템 영향] - DTC/모드/기능

규칙:
  - 기존에 값이 있는 셀은 유지
  - 빈 셀만 채움
  - S/O/D/RPN은 건드리지 않음
입력: JG1_SBW-Software_FMEA_4_필드완성.xlsx
출력: JG1_SBW-Software_FMEA_5_GHI완성.xlsx
"""
import win32com.client as win32, shutil

SRC  = r"E:\claude\FMEA\JG1_SBW-Software_FMEA_4_필드완성.xlsx"
DEST = r"E:\claude\FMEA\JG1_SBW-Software_FMEA_5_GHI완성.xlsx"

# ══════════════════════════════════════════════════════════════
# GHI DB: (unit_key, var_key, fm) → (G_detail, H_eff_module, I_eff_system)
# fm='ALL' 은 모든 실패모드 공통
# ══════════════════════════════════════════════════════════════
GHI = {}

NA_BOOL  = "- Boolean 타입이므로 해당없음"
NA_DIGIT = "- 디지털/정수 타입이므로 해당없음"
NA_ENUM  = "- Enum 타입이므로 해당없음"
NA_CAN   = "- CAN 수신 신호이므로 해당없음"

# ─────────────────────── CstAp_PwrMGT ───────────────────────
GHI[('PwrMGT','TrmnlCtrlGrpStaBDCEV','MORE')] = (
    "From : BDCEV_OFF\nTo : BDCEV_POWERON",
    "BDCEV_POWERON으로 오인식",
    "Normal 모드로 천이\n  - B/L On, DTC 저장 가능")
GHI[('PwrMGT','TrmnlCtrlGrpStaBDCEV','LESS')] = (
    "From : BDCEV_READY\nTo : BDCEV_CRANKING",
    "BDCEV_CRANKING으로 오인식",
    "DTC 저장 불가")
GHI[('PwrMGT','TrmnlCtrlGrpStaBDCEV','CORRUPT')] = (
    "From : BDCEV_OFF\nTo : Out of Range",
    "Power Off로 오인식",
    "BDCEV_READY, BDCEV_POWERON 진입 불가")
GHI[('PwrMGT','TrmnlCtrlGrpStaBDCEV','REVERSE')] = (
    NA_ENUM, NA_ENUM, NA_ENUM)
GHI[('PwrMGT','TrmnlCtrlGrpStaBDCEV','NO')] = (
    "From : BDCEV 신호 수신\nTo : 신호 없음 (CAN 타임아웃)",
    "TrmnlCtrlGrpStaBDCEV = 0 강제",
    "모터 활성화 조건 불충족 → 레버 동작 불가")
GHI[('PwrMGT','TrmnlCtrlGrpStaBDCEV','EARLY')] = (
    "From : 정상 수신 타이밍\nTo : 조기 수신",
    "이전 사이클 값 사용",
    "모드 전환 타이밍 오차 발생")
GHI[('PwrMGT','TrmnlCtrlGrpStaBDCEV','LATE')] = (
    "From : 정상 수신 타이밍\nTo : 지연 수신",
    "이전 사이클 값 유지",
    "모드 전환 지연 → 레버 응답성 저하")
GHI[('PwrMGT','TrmnlCtrlGrpStaBDCEV','AS WELL AS')] = (
    NA_CAN, NA_CAN, NA_CAN)

GHI[('PwrMGT','ECUSta','MORE')] = (
    "From : EXTER_ECU_WAKEUP\nTo : EXTER_ECU_DOOR_WAKEUP",
    "Door Wakeup으로 오인식",
    "Door Wakeup 상태 진입")
GHI[('PwrMGT','ECUSta','LESS')] = (
    "From : EXTER_ECU_SLEEP\nTo : EXTER_ECU_STANDBY",
    "Standby로 오인식",
    "Standby 상태 진입, Sleep 전환 불가")
GHI[('PwrMGT','ECUSta','CORRUPT')] = (
    "From : EXTER_ECU_WAKEUP\nTo : Out of Range",
    "유효 이외의 상태로 오인식",
    "Ldo2FltSta 판단 불가")
GHI[('PwrMGT','ECUSta','REVERSE')] = (NA_ENUM, NA_ENUM, NA_ENUM)
GHI[('PwrMGT','ECUSta','NO')] = (
    "From : ECU 상태 정상 수신\nTo : 신호 없음",
    "ECUSta = STANDBY(3) 기본값 적용",
    "LDO2/SBC 검사 활성화 조건 불충족")
GHI[('PwrMGT','ECUSta','EARLY')] = (
    "From : 정상 전이 타이밍\nTo : 조기 전이",
    "이전 상태 유지",
    "ECU 상태 전이 타이밍 오차")
GHI[('PwrMGT','ECUSta','LATE')] = (
    "From : 정상 전이 타이밍\nTo : 지연 전이",
    "이전 상태 유지",
    "ECU 상태 전이 지연 → 초기화 절차 지연")

GHI[('PwrMGT','BatVolt','MORE')] = (
    "From : Low Bat\nTo : Normal",
    "Normal Battery 상태로 오인식",
    "가용성 문제")
GHI[('PwrMGT','BatVolt','LESS')] = (
    "From : High Bat\nTo : Normal",
    "Normal Battery 상태로 오인식",
    "가용성 문제")
GHI[('PwrMGT','BatVolt','CORRUPT')] = (
    "From : 0~4095\nTo : Out of Range",
    "BatOverSta / BatUnderSta 오판단",
    "U3003_A2/A3 DTC 오설정 또는 미설정")
GHI[('PwrMGT','BatVolt','EARLY')] = (
    "From : 정상 전압 유지\nTo : 조기 이상 전압 판정",
    "200ms 디바운스 전 BatOverSta/UnderSta 조기 설정",
    "일시적 전압 변동에 의한 DTC 오설정")
GHI[('PwrMGT','BatVolt','LATE')] = (
    "From : 이상 전압 발생\nTo : 정상 전압으로 늦게 복귀",
    "BatOverSta/UnderSta 해제 지연",
    "DTC 해제 지연 → 가용성 저하")
GHI[('PwrMGT','BatVolt','REVERSE')] = (NA_DIGIT, NA_DIGIT, NA_DIGIT)
GHI[('PwrMGT','BatVolt','NO')] = (
    "From : ADC 배터리 전압 측정\nTo : 측정값 없음 (0)",
    "BatVolt = 0 → BatUnderSta = ON 오판단",
    "U3003_A2 DTC 오설정 → 저전압 경고")

GHI[('PwrMGT','IgnVolt','MORE')] = (
    "From : 4V Under\nTo : 7V Over",
    "IGN ON 상태로 오인식",
    "정상 시동 전 IGN ON 조건 충족 → DTC 미발생으로 Normal Mode 전이 가능")
GHI[('PwrMGT','IgnVolt','LESS')] = (
    "From : 7V Over\nTo : 4V Under",
    "IGN OFF 상태로 오인식",
    "정상 IGN ON 상태에서도 DTC 미발생 → N단 대기 Mode로 전이 가능")
GHI[('PwrMGT','IgnVolt','CORRUPT')] = (
    "From : 0~4095\nTo : Out of Range",
    "IGN 상태 판단 불가",
    "PowerOnSta 오판단 → 시스템 전원 상태 이상")
GHI[('PwrMGT','IgnVolt','REVERSE')] = (NA_DIGIT, NA_DIGIT, NA_DIGIT)
GHI[('PwrMGT','IgnVolt','NO')] = (
    "From : ADC IGN 전압 측정\nTo : 측정값 없음 (0)",
    "IgnVolt = 0 → HWIGN = OFF 오판단",
    "PowerOnSta = OFF → 시스템 전원 OFF 오전환")
GHI[('PwrMGT','IgnVolt','EARLY')] = (
    "From : 정상 IGN 타이밍\nTo : 조기 IGN ON 판정",
    "50ms 필터 전 HWIGN = ON 조기 설정",
    "시동 전 조기 전원 ON → 불필요한 초기화")
GHI[('PwrMGT','IgnVolt','LATE')] = (
    "From : 정상 IGN 타이밍\nTo : IGN OFF 판정 지연",
    "HWIGN = OFF 해제 지연",
    "시동 OFF 후 전원 유지 지연 → 배터리 소모")

GHI[('PwrMGT','Ldo2OnVolt','MORE')] = (
    "From : 정상 전압 (3.5V~3.6V)\nTo : 정상 범위 초과",
    "LDO2 과전압 조건 (별도 감지 없음)",
    "LDO2 과전압 미감지 → 회로 손상 가능성")
GHI[('PwrMGT','Ldo2OnVolt','LESS')] = (
    "From : 정상 전압 (>3.5V)\nTo : 저전압 (≤3.5V, ADC≤2866)",
    "Ldo2FltSta = ON 오판단",
    "시스템 전원 불안정 판단 → 인디케이터 비활성화")
GHI[('PwrMGT','Ldo2OnVolt','CORRUPT')] = (
    "From : 0~4095\nTo : Out of Range",
    "LDO2 전압 판단 오류",
    "Ldo2FltSta 오설정 → 전원 상태 오판단")
GHI[('PwrMGT','Ldo2OnVolt','REVERSE')] = (NA_DIGIT, NA_DIGIT, NA_DIGIT)
GHI[('PwrMGT','Ldo2OnVolt','NO')] = (
    "From : ADC LDO2 전압 측정\nTo : 측정값 없음 (0)",
    "Ldo2OnVolt = 0 → Ldo2FltSta = ON 오판단",
    "전원 불안정 판단 → 시스템 동작 제한")
GHI[('PwrMGT','Ldo2OnVolt','EARLY')] = (
    "From : 정상 LDO2 타이밍\nTo : 조기 저전압 판정",
    "100ms 필터 전 Ldo2FltSta 조기 설정",
    "일시적 전압 변동에 의한 오판단")
GHI[('PwrMGT','Ldo2OnVolt','LATE')] = (
    "From : LDO2 이상 발생\nTo : 정상 복귀 지연",
    "Ldo2FltSta 해제 지연",
    "전원 복구 후에도 기능 제한 지속")
GHI[('PwrMGT','Ldo2OnVolt','AS WELL AS')] = (NA_DIGIT, NA_DIGIT, NA_DIGIT)
GHI[('PwrMGT','Ldo2OnVolt','PART OF')] = (NA_DIGIT, NA_DIGIT, NA_DIGIT)

GHI[('PwrMGT','SbcFlt','MORE')] = (
    "From : Normal (1)\nTo : Fault (0) 반전 감지",
    "SbcFltSta = ON 오판단",
    "SysPwrSta = POWER_OFF 강제 → 시스템 전원 차단")
GHI[('PwrMGT','SbcFlt','LESS')] = (
    "From : Fault (0)\nTo : Normal (1) 반전 감지",
    "SbcFlt 정상으로 오인식",
    "SBC 결함 미감지 → 전원 이상 상태 지속")
GHI[('PwrMGT','SbcFlt','CORRUPT')] = (
    "From : 0 또는 1\nTo : Out of Range",
    "SbcFlt 상태 판단 불가",
    "SbcFltSta 오설정 → SysPwrSta 오판단")
GHI[('PwrMGT','SbcFlt','REVERSE')] = (
    "From : Fault (0)\nTo : Normal (1) 반전",
    "SBC 결함을 정상으로 오인식",
    "SBC 결함 미감지 → 시스템 전원 이상 지속")
GHI[('PwrMGT','SbcFlt','NO')] = (
    "From : SBC 결함 신호 수신\nTo : 신호 없음",
    "SbcFltSta 판단 불가",
    "SBC 결함 미감지 → 전원 이상 상태 지속")

# ─────────────────────── CstAp_ECUModeMgt ───────────────────
def _ecu_timeout(sig_name, effect):
    """ECUModeMgt 타임아웃 변수 GHI 생성"""
    return {
        'MORE':     (f"From : 정상 수신\nTo : {sig_name} 타임아웃", f"{sig_name} 타임아웃 → 관련 신호 무효", effect),
        'LESS':     (f"From : {sig_name} 타임아웃\nTo : 정상 복귀", f"{sig_name} 복귀 감지 지연", "이전 안전 기본값 유지"),
        'CORRUPT':  (f"From : 정상/타임아웃\nTo : 비정상 플래그 값", f"{sig_name} 타임아웃 플래그 오판단", effect),
        'ALL':      (f"From : 정상 수신\nTo : {sig_name} 타임아웃", f"{sig_name} 타임아웃 → 관련 신호 무효", effect),
    }

_timeout_effects = {
    'BDC02MsgTo':       "IgnSwStaFlag = OFF 강제 → 점화 신호 무효화",
    'BDC02Timeout':     "IgnSwStaFlag = OFF 강제 → 점화 신호 무효화",
    'BDC05MsgTo':       "테일램프/무드램프 신호 OFF 강제",
    'BDC05Timeout':     "IntTailLmpOnReqFlag = OFF 강제 → 인테리어 조명 차단",
    'CLUMsgTo':         "AutoBrightSta = OFF 강제 → 인디케이터 밝기 수동 모드",
    'CLU01Timeout':     "AutoBrightSta = OFF 강제",
    'PDC03MsgTo':       "DrvDrSwSta = OFF 강제 → 모터 활성화 조건 불충족",
    'PDC01Timeout':     "DrvStOccSta = OFF 강제 → 모터 활성화 조건 불충족",
    'PDC03Timeout':     "DrvDrSwSta = OFF 강제",
    'SMK03Timeout':     "TrmnlCtrlGrpStaBDCEV = 0 강제 → U1065_8C DTC → 모터 차단",
    'DeVCU01Timeout':   "P0A1D_8C DTC → 기어 변속 신호 무효화",
    'DeVCU04Timeout':   "P0A1D_8C DTC → 관련 신호 무효화",
    'ICC02Timeout':     "U01D0_8C DTC → ICU 신호 무효화",
    'SBCMDRV01Timeout': "DrvDrSwStaSBCM = OFF 강제 → 모터 활성화 조건 불충족",
    'MainShiftActMsgTo':"P0A1D_8C DTC → Sub CAN 폴백",
    'SubShiftActMsgTo': "P0A1D_8C DTC 조건 충족",
}

for vk, eff in _timeout_effects.items():
    d = _ecu_timeout(vk, eff)
    for fm, vals in d.items():
        GHI[('ECUModeMgt', vk, fm)] = vals

GHI[('ECUModeMgt','DriveSta','MORE')] = (
    "From : DriveSta = OFF\nTo : DriveSta = ON 오인식",
    "주행 준비 완료로 오인식",
    "모터 활성화 조건 충족 오판단 → 의도치 않은 레버 동작 가능")
GHI[('ECUModeMgt','DriveSta','LESS')] = (
    "From : DriveSta = ON\nTo : DriveSta = OFF 오인식",
    "주행 준비 미완료로 오인식",
    "모터 활성화 차단 → 기어 변속 불가")
GHI[('ECUModeMgt','DriveSta','CORRUPT')] = (
    "From : 정상 DriveSta\nTo : 비정상 값",
    "DriveSta 판단 불가",
    "모터 활성화 조건 판단 오류 → 레버 동작 이상")
GHI[('ECUModeMgt','DriveSta','NO')] = (
    "From : DriveSta 신호 수신\nTo : 신호 없음",
    "DriveSigAllTo = ON → 모터 활성화 차단",
    "레버 동작 불가")

GHI[('ECUModeMgt','DriveSigAllTo','ALL')] = (
    "From : DriveSigAllTo = OFF\nTo : DriveSigAllTo = ON",
    "Main/Sub CAN 양쪽 타임아웃 판정",
    "모터 활성화 조건 차단 → 레버 동작 불가")
GHI[('ECUModeMgt','DriveSigMainTo','ALL')] = (
    "From : Main CAN 정상\nTo : Main CAN 타임아웃",
    "Main CAN DrvRdySig 타임아웃",
    "Sub CAN 폴백; Sub도 실패 시 DriveSigAllTo = ON")
GHI[('ECUModeMgt','DriveSigSubTo','ALL')] = (
    "From : Sub CAN 정상\nTo : Sub CAN 타임아웃",
    "Sub CAN DrvRdySig 타임아웃",
    "Main CAN 유지; 양쪽 모두 실패 시 DriveSigAllTo = ON")
GHI[('ECUModeMgt','SysPwrSta','ALL')] = (
    "From : POWER_ON\nTo : POWER_OFF",
    "SysPwrSta ≠ POWER_ON 오판단",
    "모터/인디케이터 활성화 조건 불충족 → 기능 비활성화")
GHI[('ECUModeMgt','TrmnlCtrlGrpStaBDCEV','ALL')] = (
    "From : BDCEV_READY/POWERON\nTo : BDCEV_OFF",
    "TrmnlCtrlGrpStaBDCEV = 0 오인식",
    "모터 활성화 조건 불충족 → 레버 동작 불가 → U1065_8C DTC")
GHI[('ECUModeMgt','DiagSession','ALL')] = (
    "From : 진단 세션 제한 없음\nTo : 주행 중 쓰기 허용",
    "DriveSta 조건 검사 없이 세션 변경",
    "주행 중 ECU 파라미터 변경 → 안전 기능 손상 가능")
GHI[('ECUModeMgt','DrLockSta','ALL')] = (
    "From : 도어 잠금 해제\nTo : 도어 잠금 오인식",
    "도어 잠금 상태 오판단",
    "도어 잠금 연동 기능 오동작")
GHI[('ECUModeMgt','IntTailLmpOnReqFlag','ALL')] = (
    "From : 정상 조명 요청\nTo : 조명 요청 오신호",
    "테일램프 ON/OFF 오인식",
    "인테리어 테일램프 의도치 않은 ON/OFF")

# ─────────────────────── CstAp_CANMGT ───────────────────────
GHI[('CANMGT','SActSig','MORE')] = (
    "From : 정상 기어 위치 신호\nTo : CRC 오류 (E2E_P_ERROR)",
    "기어 위치 신호 CRC 오류 발생",
    "P0A1D_83 DTC → GearPosSta = NOT_DISPLAY → 인디케이터 표시 불가")
GHI[('CANMGT','SActSig','LESS')] = (
    "From : 정상 AlvCnt 순서\nTo : AlvCnt 반복 (E2E_P_REPEATED)",
    "기어 위치 신호 AlvCnt 오류",
    "P0A1D_82 DTC → 신호 폐기 → Sub CAN 폴백")
GHI[('CANMGT','SActSig','CORRUPT')] = (
    "From : 정상 E2E 검증\nTo : E2E 검증 실패",
    "기어 위치 신호 무효 처리",
    "P0A1D_82/83 DTC → GearPosSta = NOT_DISPLAY → 기어 표시 불가")
GHI[('CANMGT','SActSig','OMISSION')] = (
    "From : SActSig 정상 수신\nTo : 타임아웃 (신호 누락)",
    "MainSActMsgToFlag = ON",
    "P0A1D_8C DTC → Sub CAN 폴백 또는 레버 동작 차단")
GHI[('CANMGT','SActSig','REVERSE')] = (NA_ENUM, NA_ENUM, NA_ENUM)
GHI[('CANMGT','SActSig','EARLY')] = (
    "From : 정상 수신 주기 (10ms)\nTo : 조기 수신",
    "이전 사이클 값 이중 처리",
    "기어 위치 판단 타이밍 오차")
GHI[('CANMGT','SActSig','LATE')] = (
    "From : 정상 수신 주기 (10ms)\nTo : 지연 수신",
    "SActSigTo 카운터 증가",
    "타임아웃 카운터 누적 → P0A1D_8C DTC 조건 근접")

# SactSig (소문자 a) - SActSig와 동일
for fm in ['MORE','LESS','CORRUPT','OMISSION','REVERSE','EARLY','LATE']:
    k1 = ('CANMGT','SActSig',fm)
    k2 = ('CANMGT','SactSig',fm)
    if k1 in GHI:
        GHI[k2] = GHI[k1]

GHI[('CANMGT','SubSActSig','ALL')] = (
    "From : Sub CAN 정상 E2E\nTo : E2E 검증 실패",
    "Sub CAN 기어 신호 무효 처리",
    "P0A1D_82/83 DTC → Main CAN 폴백 유지")
GHI[('CANMGT','SactSigTo','ALL')] = (
    "From : 정상 수신\nTo : Main CAN SActSig 타임아웃",
    "MainSActMsgToFlag = ON",
    "P0A1D_8C DTC → Sub CAN 폴백")
GHI[('CANMGT','SubSActSigTo','ALL')] = (
    "From : 정상 수신\nTo : Sub CAN SActSig 타임아웃",
    "SubSActMsgToFlag = ON",
    "Main CAN 유지; 양쪽 타임아웃 시 P0A1D_8C DTC")

GHI[('CANMGT','MainCANBusOFF','MORE')] = (
    "From : CAN 버스 정상\nTo : Main CAN BusOff",
    "MainCANBusOFFSta = ON",
    "U0028_88 DTC → Main CAN 모든 신호 폐기 → Sub CAN 폴백")
GHI[('CANMGT','MainCANBusOFF','LESS')] = (
    "From : Main CAN BusOff\nTo : 정상 복귀",
    "MainCANBusOFFSta 해제",
    "Main CAN 신호 재수신 시작")
GHI[('CANMGT','MainCANBusOFF','CORRUPT')] = (
    "From : 정상/BusOff\nTo : BusOff 플래그 오설정",
    "MainCANBusOFFSta 오판단",
    "U0028_88 DTC 오설정 또는 미설정")
GHI[('CANMGT','SubCANBusOFF','MORE')] = (
    "From : CAN 버스 정상\nTo : Sub CAN BusOff",
    "SubCANBusOFFSta = ON",
    "U0028_88 DTC → Sub CAN 모든 신호 폐기 → Main CAN 유지")
GHI[('CANMGT','SubCANBusOFF','LESS')] = (
    "From : Sub CAN BusOff\nTo : 정상 복귀",
    "SubCANBusOFFSta 해제",
    "Sub CAN 신호 재수신 시작")
GHI[('CANMGT','SubCANBusOFF','CORRUPT')] = (
    "From : 정상/BusOff\nTo : BusOff 플래그 오설정",
    "SubCANBusOFFSta 오판단",
    "U0028_88 DTC 오설정 또는 미설정")

def _can_timeout_ghi(sig, dtc_effect):
    return {
        'MORE':    (f"From : 정상 수신\nTo : {sig} 타임아웃", f"{sig} 타임아웃 플래그 = ON", dtc_effect),
        'LESS':    (f"From : {sig} 타임아웃\nTo : 정상 복귀", f"{sig} 타임아웃 해제", "정상 신호 처리 재개"),
        'CORRUPT': (f"From : 정상 수신\nTo : 신호 오염", f"{sig} 신호 유효성 실패", dtc_effect),
        'ALL':     (f"From : 정상 수신\nTo : {sig} 타임아웃", f"{sig} 타임아웃 플래그 = ON", dtc_effect),
    }

_can_timeouts = {
    'BDC_02_Timeout':     "BDC02MsgTo = ON → IgnSwStaFlag = OFF 강제",
    'BDC_03_Timeout':     "무드램프 관련 신호 OFF 강제",
    'BDC_04_Timeout':     "관련 신호 무효화 → 안전 기본값 유지",
    'BDC_05_Timeout':     "BDC05MsgTo = ON → 테일램프/무드램프 OFF 강제",
    'BDC_06_Timeout':     "관련 신호 무효화",
    'CGWCLU_01_20ms_Timeout': "CLUMsgTo = ON → AutoBrightSta = OFF 강제",
    'SMK_02_Timeout':     "SMK_02 관련 신호 무효화",
    'SMK_03_Timeout':     "Main+Sub 양쪽 타임아웃 → TrmnlCtrlGrpStaBDCEV = 0 → U1065_8C DTC",
    'PDC_02_Timeout':     "관련 신호 무효화",
    'PDC_03_Timeout':     "PDC03MsgTo = ON → DrvDrSwSta = OFF 강제",
    'VCU04_Timeout':      "P0A1D_8C DTC → 기어 변속 신호 무효화",
    'ICC_02_50ms':        "U01D0_8C DTC → ICU 신호 무효화",
    'SBCM_DRV_01_Timeout':"DrvDrSwStaSBCM = OFF 강제",
}
for vk, eff in _can_timeouts.items():
    for fm, vals in _can_timeout_ghi(vk, eff).items():
        GHI[('CANMGT', vk, fm)] = vals

def _can_signal_ghi(sig, flag_effect, dtc_effect):
    return {
        'MORE':    (f"From : 정상 범위\nTo : 정상 범위 초과", f"{sig} 비정상 고값 수신", dtc_effect),
        'LESS':    (f"From : 정상 범위\nTo : 정상 범위 미달", f"{sig} 비정상 저값 수신", dtc_effect),
        'CORRUPT': (f"From : 유효 신호\nTo : Out of Range 또는 오염 신호", f"{sig} 유효성 검사 실패", dtc_effect),
        'NO':      (f"From : {sig} 정상 수신\nTo : 신호 없음", flag_effect, "안전 기본값 적용"),
        'REVERSE': (NA_CAN, NA_CAN, NA_CAN),
        'AS WELL AS': (NA_CAN, NA_CAN, NA_CAN),
        'ALL':     (f"From : 유효 신호\nTo : 비정상 신호", f"{sig} 신호 이상", dtc_effect),
    }

for vk in ['DrvDrSwSta','DrvDrSwSta_SBCM','DrvStOccSta']:
    for fm, vals in _can_signal_ghi(vk, f"{vk} = OFF 기본값 적용", "모터 활성화 조건 불충족 → 레버 동작 불가").items():
        GHI[('CANMGT', vk, fm)] = vals

GHI[('CANMGT','SMK_TrmnlCtrlGrpStaBDCEV','ALL')] = (
    "From : 정상 SMK 신호\nTo : Dual-CAN 양쪽 타임아웃",
    "TrmnlCtrlGrpStaBDCEV = 0 강제",
    "모터 활성화 차단 → U1065_8C DTC")
GHI[('CANMGT','SMK_TrmnlCtrlStaBDC','ALL')] = (
    "From : 정상 SMK_03 수신\nTo : 타임아웃",
    "SMK_TrmnlCtrlStaBDC 신호 무효화",
    "TrmnlCtrl 관련 기능 비활성화")
GHI[('CANMGT','SMK_TrmnlCtrlGrpStaBDC','ALL')] = GHI[('CANMGT','SMK_TrmnlCtrlStaBDC','ALL')]
GHI[('CANMGT','SMK_PwrOnModeSta','ALL')] = (
    "From : SMK_03 정상 수신\nTo : 타임아웃 또는 오염",
    "PwrOnModeSta = OFF 강제",
    "인디케이터/무드램프 제어 비활성화")
GHI[('CANMGT','SubDrvRdySig','ALL')] = (
    "From : Sub CAN DrvRdySig 정상\nTo : 타임아웃",
    "DriveSigSubTo = ON",
    "Main CAN 유지; 양쪽 타임아웃 시 DriveSigAllTo = ON → 모터 차단")
GHI[('CANMGT','SubDrvRdySigTo','ALL')] = (
    "From : Sub CAN 정상\nTo : Sub DrvRdySig 타임아웃",
    "DriveSigSubTo = ON",
    "Main CAN 유지; 양쪽 모두 실패 시 DriveSigAllTo = ON")
GHI[('CANMGT','RotateState','ALL')] = (
    "From : 유효값 (0 또는 1)\nTo : 유효 범위 외",
    "RotateState 유효성 실패",
    "상태 머신 진입 불가 → 모터 비활성화 유지")
GHI[('CANMGT','PosSta','ALL')] = (
    "From : 정상 위치 상태\nTo : 위치 결함 상태 (LVR_Flt)",
    "LvrPosSta = LVR_Flt(0x0F) 오인식",
    "P2E00_01 DTC → CAN으로 FAULT 상태 전송")
GHI[('CANMGT','PosStuck','ALL')] = (
    "From : 위치 정상 변화\nTo : 위치 고착 (500ms 이상 ±2 이내)",
    "PosStuck = ON",
    "LvrWrngMsg = LEVER_STUCK(0x12) → 경고 3회 후 모터 정지")
GHI[('CANMGT','PButtonSta','ALL')] = (
    "From : P버튼 정상\nTo : P버튼 결함 또는 고착",
    "LvrPSta = P_FAULT(0x03)",
    "C1181_96 DTC → P위치 전환 불가")
GHI[('CANMGT','PButtonFault','ALL')] = (
    "From : PSw1/PSw2 일치\nTo : PSw1/PSw2 불일치",
    "PButtonFault = P_BUTTON_FAIL",
    "C1181_96 DTC → LvrPSta = P_FAULT → P위치 전환 불가")
GHI[('CANMGT','PButtonStuck','ALL')] = (
    "From : P버튼 정상 해제\nTo : 18000ms 이상 눌림 (Stuck)",
    "PButtonStuck = ON",
    "C1181_96 DTC → LvrPSta = P_FAULT")
GHI[('CANMGT','IdtFltSta','ALL')] = (
    "From : 인디케이터 정상\nTo : 회로 결함",
    "LvrIdtSta = INDI_FAULT",
    "C1182_96 DTC 설정")
GHI[('CANMGT','IdtSta','ALL')] = (
    "From : 인디케이터 정상 상태\nTo : 비정상 상태",
    "LvrIdtSta 이상값",
    "C1182_96 DTC 조건 충족")
GHI[('CANMGT','PosSnrFlt','ALL')] = (
    "From : 위치 센서 정상\nTo : 결함 (범위 외)",
    "LvrPosSta = LVR_Flt",
    "P2E00_01 DTC → 모터 정지")
GHI[('CANMGT','MotorActivation','ALL')] = (
    "From : 모터 활성화 정상\nTo : 활성화 조건 불충족",
    "MotorActivation = OFF 전환",
    "Step_EN = 1 (하드웨어 비활성화) → 모터 정지")
GHI[('CANMGT','MotorStopSig','ALL')] = (
    "From : MotorStopSig = 0 (정상)\nTo : MotorStopSig = 1 (정지 명령)",
    "MotorControlStop() 호출",
    "Motor_State = 0, Step_EN = 1, tmp_Speed = 0 → 영구 정지")
GHI[('CANMGT','MotorFaultWarning','ALL')] = (
    "From : 모터 정상\nTo : 모터 고장 핀 감지",
    "MotorFaultPinChek 오인식",
    "모터 정지 명령 → 경고 CAN 전송")
GHI[('CANMGT','LeverWarning_Dial','ALL')] = (
    "From : 정상 레버 상태\nTo : 레버 경고 발생",
    "LvrWrngMsg 경고 코드 설정",
    "CAN 경고 코드 전송 → 3회 반복 후 모터 정지")
GHI[('CANMGT','LeverWarning_Sphere','ALL')] = GHI[('CANMGT','LeverWarning_Dial','ALL')]
GHI[('CANMGT','RetryWarning_Dial','ALL')] = (
    "From : 정상 동작\nTo : 재시도 경고 발생",
    "SBW_MotorWarn.RetryWarning = ON",
    "경고 CAN 전송 → 3회 재시도 실패 시 MotorStopSig = ON")
GHI[('CANMGT','RetryWarning_Sphere','ALL')] = GHI[('CANMGT','RetryWarning_Dial','ALL')]
GHI[('CANMGT','tmpPos1','ALL')] = (
    "From : 유효 범위 (100~900)\nTo : 범위 외 또는 오염",
    "PosSnrFlt = ON",
    "P2E00_01 DTC → 모터 정지")
GHI[('CANMGT','Main_Vcu_E2E_Error_Sta','ALL')] = (
    "From : E2E 정상 (0x00)\nTo : E2E 오류 (0x01~0x05)",
    "E2E 오류 코드별 플래그 설정",
    "P0A1D_82/83 DTC → 비정상 기어 신호 폐기")
GHI[('CANMGT','Main_Vcu_E2E_Error_Return','ALL')] = (
    "From : E2E_P_OK\nTo : E2E_P_ERROR",
    "CrcErrFlag = ON",
    "P0A1D_83 DTC → 해당 데이터 폐기")

for vk in ['SysPwrSta','ECUSta']:
    GHI[('CANMGT', vk, 'ALL')] = (
        f"From : 정상 상태\nTo : {vk} 비정상",
        f"{vk} 오판단 → 관련 기능 이상",
        "모터/인디케이터 활성화 조건 불충족")

for vk in ['AsstDrSwSta','RrLftDrSwSta','RrRtDrSwSta']:
    GHI[('CANMGT', vk, 'ALL')] = (
        f"From : 도어 정상 상태\nTo : {vk} 오신호",
        f"{vk} 오인식",
        "도어 상태 연동 기능 오동작")

for vk in ['AvTailLmpSta','InlTailLmpSta','IntTailLmpOnReq']:
    GHI[('CANMGT', vk, 'ALL')] = (
        f"From : 정상 조명 신호\nTo : {vk} 오신호",
        f"{vk} = OFF 오인식 또는 비정상 ON",
        "테일램프/인테리어 조명 오동작")

for vk in ['MoodLamp_MdLmpFadeSta','MoodLamp_SlvBrgtnsVal',
           'MoodLamp_SlvColor_X','MoodLamp_SlvColor_Y',
           'MoodLamp_SlvFadeInTimetVal','MoodLamp_SlvFadeOutTimetVal']:
    GHI[('CANMGT', vk, 'ALL')] = (
        f"From : 정상 무드램프 신호\nTo : {vk} 오신호",
        f"{vk} 오인식",
        "무드램프 색상/밝기/페이드 오동작 (안전에 직접 영향 없음)")

for vk in ['AutoLtSnsrNightSta','BtnIllumiAlwaysOnSta','USM_IllAlwaysOnwithPSTNSta']:
    GHI[('CANMGT', vk, 'ALL')] = (
        f"From : 정상 USM/CLU 신호\nTo : {vk} 오신호",
        f"{vk} 오인식",
        "실내 조명 자동 제어 오동작")

for vk in ['RKESig','SMKSig','GetPSigSta','GetPSigTo',
           'StrtStpBtnSw1Sta','SSB_StrtStpBtnSw2Sta']:
    GHI[('CANMGT', vk, 'ALL')] = (
        f"From : 정상 P위치 요청 신호\nTo : {vk} 오신호",
        f"{vk} 오인식",
        "P위치 전환 오동작 또는 차단")

for vk in ['Naccept','OverRideWarning','SpecOption','ComCtrlMode',
           'FaceDetectStat','UtilModeActStaSig','USM06Msg',
           'MainCanSigSet_LvrMsg','SubCanSigSet_LvrMsg',
           'Ign1InSta','Trunk_IBU_OpnDiagSetSta','Trunk_TrnkTlgtReleaseRly']:
    GHI[('CANMGT', vk, 'ALL')] = (
        f"From : 정상 신호\nTo : {vk} 오신호",
        f"{vk} 오인식",
        "관련 기능 오동작 또는 비활성화")

# ─────────────────────── CstAp_MotorControlMgt ──────────────
GHI[('MotorControlMgt','tmp_Position','MORE')] = (
    "From : 유효 범위 (100~900)\nTo : 900 초과",
    "Moving_Position 미갱신 → 위치 업데이트 중단",
    "500ms Stuck 판정 → P2E00_01 DTC → 모터 정지")
GHI[('MotorControlMgt','tmp_Position','LESS')] = (
    "From : 유효 범위 (100~900)\nTo : 100 미달",
    "Moving_Position 미갱신",
    "Stuck 감지 → P2E00_01 DTC → 모터 정지")
GHI[('MotorControlMgt','tmp_Position','CORRUPT')] = (
    "From : 유효 위치 값\nTo : Out of Range",
    "LvrModePosInfo = 3 (오류 상태)",
    "CAN으로 Not Available 전송 → P2E00_01 DTC")
GHI[('MotorControlMgt','tmp_Position','STUCK')] = (
    "From : 위치 정상 변화 (>±2)\nTo : 500ms 이상 ±2 이내 고착",
    "TurnDialError / TurnSphereError 카운터 증가",
    "3회 실패 → MotorStopSig = 1 → 모터 영구 정지 → P2E00_01 DTC")
GHI[('MotorControlMgt','tmp_Position','REVERSE')] = (
    "From : 정방향 이동\nTo : 역방향 오인식",
    "모터 방향 판단 오류",
    "목표 위치 반대 방향 이동 → Stuck 감지")
GHI[('MotorControlMgt','tmp_Position','NO')] = (
    "From : 위치 센서 정상 수신\nTo : 신호 없음 (0)",
    "tmp_Position = 0 → 범위 미달 판정",
    "Stuck 감지 → P2E00_01 DTC")

GHI[('MotorControlMgt','RotateState','ALL')] = (
    "From : 유효값 (0=Sphere, 1=Dial)\nTo : 유효 범위 외",
    "RotateState 유효성 실패 → 상태 머신 진입 불가",
    "모터 비활성화 유지 (안전 상태)")
GHI[('MotorControlMgt','SysPwrSta','ALL')] = (
    "From : POWER_ON\nTo : POWER_OFF 또는 이상",
    "SysPwrSta ≠ POWER_ON → 모터 활성화 조건 불충족",
    "경고 클리어 → 모터 안전 상태 유지")
GHI[('MotorControlMgt','TrmnlCtrlGrpStaBDCEV','ALL')] = (
    "From : BDCEV_READY(3) 또는 POWERON(1)\nTo : BDCEV_OFF(0)",
    "모터 활성화 진입 조건 불충족",
    "상태 머신 진입 불가 → 모터 비활성화")
GHI[('MotorControlMgt','DriveSta','ALL')] = (
    "From : DriveSta = ON\nTo : DriveSta = OFF 오인식",
    "DriveSigAllTo = ON → 모터 활성화 차단",
    "모터 정지 상태 유지 → 레버 동작 불가")
GHI[('MotorControlMgt','DriveSigAllTo','ALL')] = (
    "From : DriveSigAllTo = OFF\nTo : DriveSigAllTo = ON",
    "모터 활성화 조건 차단",
    "모터 정지 유지 → 레버 동작 불가")
GHI[('MotorControlMgt','DriveSigMainTo','ALL')] = (
    "From : Main CAN 정상\nTo : Main CAN 타임아웃",
    "DriveSigMainTo = ON",
    "Sub CAN 폴백; 양쪽 실패 시 DriveSigAllTo = ON")
GHI[('MotorControlMgt','DriveSigSubTo','ALL')] = (
    "From : Sub CAN 정상\nTo : Sub CAN 타임아웃",
    "DriveSigSubTo = ON",
    "Main CAN 유지; 양쪽 실패 시 DriveSigAllTo = ON")
GHI[('MotorControlMgt','MotorActivation','ALL')] = (
    "From : 모터 활성화 조건 유지\nTo : 100ms 이내 재활성화 실패",
    "MotorActivation = OFF 자동 전환",
    "Step_EN = 1 (하드웨어 비활성화) → 모터 정지")
GHI[('MotorControlMgt','MotorStopSig','ALL')] = (
    "From : MotorStopSig = 0\nTo : MotorStopSig = 1 (정지 명령)",
    "MotorControlStop() 호출",
    "Motor_State = 0, Step_EN = 1, tmp_Speed = 0 → 영구 정지")
GHI[('MotorControlMgt','DrvDrSwSta','ALL')] = (
    "From : 도어 닫힘\nTo : 도어 열림 또는 오인식",
    "모터 활성화 조건 불충족",
    "모드 변경 비활성화 → 의도치 않은 레버 동작 방지")
GHI[('MotorControlMgt','DrvDrSwStaSBCM','ALL')] = (
    "From : SBCM 도어 닫힘\nTo : 도어 열림 오인식",
    "SBCM 채널 도어 상태 오인식",
    "PDC 채널 우선 사용 → 양쪽 불일치 시 안전 기본값 적용")
GHI[('MotorControlMgt','DrvStOccSta','ALL')] = (
    "From : 운전자 착석\nTo : 미착석 오인식",
    "모터 활성화 조건 불충족",
    "모드 변경 비활성화 → 10사이클 지연 후 SystemState = 2")
GHI[('MotorControlMgt','FaceDetectStat','ALL')] = (
    "From : 정상 얼굴 인식 신호\nTo : 오신호",
    "FaceDetectStat 오인식",
    "모터 활성화에 직접 영향 없음 (참고 신호)")
GHI[('MotorControlMgt','PwrOnModeSta','ALL')] = (
    "From : PwrOnModeSta = READY(2)\nTo : PwrOnModeSta ≠ READY",
    "모터 활성화 조건 불충족",
    "모터 비활성화 유지 → 안전 상태")
for vk in ['Channel','Motor_Speed','STEPMOTOR_SEQ_Write',
           'Step_DIR','Step_EN','SrcDataBufferPtr','DesDataBufferPtr','Length']:
    GHI[('MotorControlMgt', vk, 'ALL')] = (
        f"From : 유효값\nTo : {vk} 비정상",
        f"{vk} 오류 → SPI/모터 제어 이상",
        "모터 명령 전달 실패 → 안전 상태 유지")

# ─────────────────────── CstAp_PosMgt ───────────────────────
GHI[('PosMgt','PositionSensor_PosSnrRaw1','MORE')] = (
    "From : 유효 범위 (100~900)\nTo : 900 초과",
    "PosSnrFlt = ON",
    "P2E00_01 DTC → LvrPosSta = LVR_Flt → 모터 정지")
GHI[('PosMgt','PositionSensor_PosSnrRaw1','LESS')] = (
    "From : 유효 범위 (100~900)\nTo : 100 미달",
    "PosSnrFlt = ON",
    "P2E00_01 DTC → LvrPosSta = LVR_Flt")
GHI[('PosMgt','PositionSensor_PosSnrRaw1','CORRUPT')] = (
    "From : 유효 위치 값\nTo : Out of Range",
    "PosSnrFlt = ON",
    "P2E00_01 DTC → LvrPosSta = LVR_Flt")
GHI[('PosMgt','PositionSensor_PosSnrRaw1','STUCK')] = (
    "From : 위치 정상 변화\nTo : 위치 고착",
    "Delta 변화량 ≤ ±2 지속 → PosStuck = ON",
    "LvrWrngMsg = LEVER_STUCK → 경고 3회 후 모터 정지")
for fm in ['MORE','LESS','CORRUPT','STUCK']:
    if ('PosMgt','PositionSensor_PosSnrRaw1',fm) in GHI:
        GHI[('PosMgt','PositionSensor_PosSnrRaw2',fm)] = GHI[('PosMgt','PositionSensor_PosSnrRaw1',fm)]

GHI[('PosMgt','PositionSensorInfo_PButtonFltSta','ALL')] = (
    "From : PSw1/PSw2 일치\nTo : PSw1/PSw2 불일치",
    "PButtonFaultSta = ON",
    "C1181_96 DTC → LvrPSta = P_FAULT → P위치 전환 불가")
GHI[('PosMgt','PositionSensorInfo_PButtonSta','ALL')] = (
    "From : P버튼 정상\nTo : P버튼 결함 또는 비정상",
    "PButtonFault 감지",
    "C1181_96 DTC → P위치 진입 불가")

# ─────────────────────── CstAp_MovingMgt ────────────────────
for _vbase in ['MovSnr1','MovSnr2','PosSnr1','PosSnr2']:
    for _sub in ['ActiveTime','PeriodTime']:
        _vk = f"{_vbase}.{_sub}"
        GHI[('MovingMgt', _vk, 'MORE')] = (
            f"From : 정상 {_sub} 범위\nTo : 정상 범위 초과",
            f"{_vk} 비정상 고값 → 위치 계산 오류",
            "Moving_Position 오산출 → Stuck 판정 → P2E00_01 DTC")
        GHI[('MovingMgt', _vk, 'LESS')] = (
            f"From : 정상 {_sub} 범위\nTo : 정상 범위 미달",
            f"{_vk} 비정상 저값 → 위치 계산 오류",
            "Moving_Position 오산출 → 모터 제어 정확도 저하")
        GHI[('MovingMgt', _vk, 'CORRUPT')] = (
            f"From : 유효 {_sub} 값\nTo : Out of Range",
            f"{_vk} 유효성 실패",
            "위치 계산 불가 → Stuck 감지 → P2E00_01 DTC")
        GHI[('MovingMgt', _vk, 'STUCK')] = (
            f"From : 정상 변화\nTo : {_vk} 고착",
            "위치 변화 없음 → Stuck 판정",
            "P2E00_01 DTC → 모터 정지")
        GHI[('MovingMgt', _vk, 'NO')] = (
            f"From : {_vk} 정상 측정\nTo : 신호 없음 (0)",
            "위치 측정 불가",
            "Stuck 판정 → P2E00_01 DTC")
        GHI[('MovingMgt', _vk, 'ALL')] = GHI[('MovingMgt', _vk, 'MORE')]

# ─────────────────────── CstAp_IdtMgt ───────────────────────
for _g in ['DIdtCntl','NIdtCntl','PIdtCntl','RIdtCntl']:
    GHI[('IdtMgt', _g, 'MORE')] = (
        f"From : 정상 PWM 범위\nTo : 정상 범위 초과",
        f"{_g} 과전류/과밝기",
        "MntrVolt 임계값 초과 → IdtFltSta = ON → C1182_96 DTC")
    GHI[('IdtMgt', _g, 'LESS')] = (
        f"From : 정상 PWM 범위\nTo : 정상 범위 미달",
        f"{_g} 저밝기 또는 미출력",
        "인디케이터 표시 불량 → C1182_96 DTC")
    GHI[('IdtMgt', _g, 'CORRUPT')] = (
        f"From : 유효 PWM 값\nTo : Out of Range",
        f"{_g} PWM 오류",
        "인디케이터 표시 오류 → C1182_96 DTC")
    GHI[('IdtMgt', _g, 'NO')] = (
        f"From : {_g} 정상 출력\nTo : 신호 없음",
        f"{_g} = OFF (인디케이터 소등)",
        "기어 위치 표시 불가 → 운전자 정보 손실")
    GHI[('IdtMgt', _g, 'ALL')] = GHI[('IdtMgt', _g, 'MORE')]

for _v in ['DMntrVolt','NMntrVolt','PMntrVolt','RMntrVolt',
           'DMntrVolt2','NMntrVolt12','NMntrVolt2','PMntrVolt12','PMntrVolt2','RMntrVolt2']:
    GHI[('IdtMgt', _v, 'MORE')] = (
        f"From : 정상 모니터링 전압\nTo : 임계값 초과",
        f"{_v} 과전압 감지",
        "IdtFltSta = ON → C1182_96 DTC → PWM 출력 OFF")
    GHI[('IdtMgt', _v, 'LESS')] = (
        f"From : 정상 모니터링 전압\nTo : 임계값 미달",
        f"{_v} 저전압 감지 (회로 단선/단락)",
        "IdtFltSta = ON → C1182_96 DTC")
    GHI[('IdtMgt', _v, 'CORRUPT')] = (
        f"From : 유효 전압 값\nTo : Out of Range",
        f"{_v} 측정 오류",
        "IdtFltSta 오설정 가능 → C1182_96 DTC 오설정")
    GHI[('IdtMgt', _v, 'ALL')] = GHI[('IdtMgt', _v, 'MORE')]

GHI[('IdtMgt','GearPosSta','ALL')] = (
    "From : 유효 기어 위치\nTo : NOT_DISPLAY 또는 오류값",
    "기어 위치 표시 불가",
    "인디케이터 표시 오류 → 운전자 기어 위치 확인 불가")
GHI[('IdtMgt','AutoBrightSta','ALL')] = (
    "From : 자동 밝기 ON\nTo : 자동 밝기 OFF 오인식",
    "AutoBrightSta = OFF 강제",
    "수동 밝기 모드 전환 → 야간 인디케이터 밝기 고정")
GHI[('IdtMgt','BltDimLvl','ALL')] = (
    "From : 정상 BLT 밝기\nTo : 비정상 밝기 레벨",
    "인디케이터 BLT 밝기 오설정",
    "인디케이터 과밝기 또는 저밝기 표시")
GHI[('IdtMgt','HltDimLvl','ALL')] = (
    "From : 정상 HLT 밝기\nTo : 비정상 밝기 레벨",
    "인디케이터 HLT 밝기 오설정",
    "인디케이터 하이라이트 밝기 오류")
GHI[('IdtMgt','Indi_tmp_Pos','ALL')] = (
    "From : 유효 위치 값\nTo : 범위 외 위치값",
    "인디케이터 표시 위치 오류",
    "기어 위치 표시 오류 → 운전자 오인")
for vk in ['ECUSta','Ldo2FltSta','MainCANBusOffSta','MotorStopSig',
           'SysPwrSta','BatStbSta','CanBatNormalSta.DebounceBatNorSta',
           'CanBatNormalSta.ImmediateBatNorSta','TotalAlvCntFlt','TotalCrcFlt',
           'PwrOnModeSta','CLU01MsgTo','BDC05MsgTo','PowerOn12','PowerOn2']:
    GHI[('IdtMgt', vk, 'ALL')] = (
        f"From : 정상 {vk}\nTo : {vk} 이상",
        f"{vk} 오인식 → 인디케이터 제어 조건 오판단",
        "인디케이터 비활성화 또는 오표시")

# ─────────────────────── CstAp_ButtonMgt ────────────────────
GHI[('ButtonMgt','P_SW1_Raw','MORE')] = (
    "From : 정상 ADC 범위\nTo : 정상 범위 초과 (>2604 또는 >889)",
    "CONTACT_FAULT(2) 판정",
    "MechaErrCnt 증가 → PButtonFault = ON → C1181_96 DTC")
GHI[('ButtonMgt','P_SW1_Raw','LESS')] = (
    "From : 정상 ADC 범위\nTo : 정상 범위 미달 (<2356 또는 <804)",
    "CONTACT_FAULT(2) 판정",
    "MechaErrCnt 증가 → PButtonFault = ON → C1181_96 DTC")
GHI[('ButtonMgt','P_SW1_Raw','CORRUPT')] = (
    "From : 유효 ADC 값\nTo : Out of Range",
    "PSw1 상태 판단 불가",
    "PButtonFault = ON → C1181_96 DTC")
GHI[('ButtonMgt','P_SW1_Raw','STUCK')] = (
    "From : 정상 버튼 해제\nTo : 18000ms 이상 눌림",
    "PSw1stuckflag = ON",
    "PButtonStuck = ON → C1181_96 DTC")
GHI[('ButtonMgt','P_SW1_Raw','REVERSE')] = (NA_DIGIT, NA_DIGIT, NA_DIGIT)
GHI[('ButtonMgt','P_SW1_Raw','NO')] = (
    "From : P_SW1_Raw 정상 측정\nTo : 신호 없음 (0)",
    "PSw1value = 0 → CONTACT_FAULT 판정",
    "PButtonFault = ON → C1181_96 DTC")
GHI[('ButtonMgt','P_SW1_Raw','EARLY')] = (
    "From : 정상 버튼 타이밍\nTo : 조기 버튼 인식",
    "3 사이클 디바운스 전 조기 상태 전환",
    "P위치 전환 타이밍 오차")
GHI[('ButtonMgt','P_SW1_Raw','LATE')] = (
    "From : 정상 버튼 타이밍\nTo : 버튼 인식 지연",
    "디바운스 카운터 초과",
    "P위치 전환 지연")

for fm in ['MORE','LESS','CORRUPT','STUCK','REVERSE','NO','EARLY','LATE']:
    if ('ButtonMgt','P_SW1_Raw',fm) in GHI:
        g,h,i = GHI[('ButtonMgt','P_SW1_Raw',fm)]
        g2 = g.replace('P_SW1_Raw','P_SW2_Raw').replace('2604','3488').replace('889','1778').replace('2356','3156').replace('804','1608')
        GHI[('ButtonMgt','P_SW2_Raw',fm)] = (g2, h.replace('PSw1','PSw2'), i)

# ─────────────────────── CstAp_HapticControlMgt ─────────────
GHI[('HapticControlMgt','lvrPosInfo','MORE')] = (
    "From : 유효 기어 위치 정보\nTo : 유효 범위 초과",
    "햅틱 패턴 매핑 실패",
    "기본 패턴(OFF) 적용 → 햅틱 피드백 없음")
GHI[('HapticControlMgt','lvrPosInfo','LESS')] = (
    "From : 유효 기어 위치 정보\nTo : 유효 범위 미달",
    "햅틱 패턴 매핑 실패",
    "기본 패턴(OFF) 적용 → 햅틱 피드백 없음")
GHI[('HapticControlMgt','lvrPosInfo','CORRUPT')] = (
    "From : 유효 lvrPosInfo\nTo : Out of Range",
    "햅틱 패턴 인덱스 오류",
    "기본 패턴 적용 → 햅틱 피드백 오동작")
GHI[('HapticControlMgt','lvrPosInfo','ALL')] = GHI[('HapticControlMgt','lvrPosInfo','MORE')]

GHI[('HapticControlMgt','gearPositionVcu','MORE')] = (
    "From : 유효 기어 위치\nTo : 유효 범위 초과",
    "gearPositionVcu 매핑 실패",
    "기본 햅틱 패턴(OFF) 적용")
GHI[('HapticControlMgt','gearPositionVcu','LESS')] = (
    "From : 유효 기어 위치\nTo : 유효 범위 미달",
    "gearPositionVcu 매핑 실패",
    "기본 햅틱 패턴(OFF) 적용")
GHI[('HapticControlMgt','gearPositionVcu','CORRUPT')] = (
    "From : 유효 기어 위치\nTo : E2E 오류 또는 Out of Range",
    "gearPositionVcu 유효성 실패",
    "햅틱 피드백 오동작")
GHI[('HapticControlMgt','gearPositionVcu','ALL')] = GHI[('HapticControlMgt','gearPositionVcu','MORE')]

GHI[('HapticControlMgt','sysPwrSta','ALL')] = (
    "From : POWER_ON\nTo : POWER_OFF 또는 이상",
    "sysPwrSta ≠ POWER_ON",
    "햅틱 I2C 통신 비활성화 → 햅틱 피드백 OFF")
GHI[('HapticControlMgt','DataBufferPtr','ALL')] = (
    "From : 유효 버퍼 포인터\nTo : NULL 또는 잘못된 주소",
    "I2C 전송 데이터 참조 오류",
    "I2C 통신 실패 → 햅틱 피드백 전달 불가")
GHI[('HapticControlMgt','SlaveAddress','ALL')] = (
    "From : 유효 I2C 슬레이브 주소\nTo : 잘못된 주소",
    "I2C ACK 없음 → 통신 실패",
    "햅틱 액추에이터 응답 없음 → 피드백 동작 불가")
GHI[('HapticControlMgt','TransmitLength','ALL')] = (
    "From : 유효 전송 길이 (고정값)\nTo : 잘못된 길이",
    "I2C 데이터 오염 또는 통신 실패",
    "햅틱 명령 전달 불가 → 피드백 없음")

# ─────────────────────── CstAp_MoodControlMgt ───────────────
GHI[('MoodControlMgt','BDC05MsgTimeout','ALL')] = (
    "From : BDC_05 정상 수신\nTo : BDC_05 타임아웃",
    "BDC05MsgTimeout = ON",
    "무드램프 PWM 출력 OFF 강제 → 무드램프 소등")
GHI[('MoodControlMgt','MdLmpFadeSta','MORE')] = (
    "From : 정상 페이드 상태 (1/2)\nTo : 유효 범위 초과",
    "FadeInOutRdy 설정 오류",
    "무드램프 페이드 오동작")
GHI[('MoodControlMgt','MdLmpFadeSta','LESS')] = (
    "From : FadeInOut ON (2)\nTo : FadeInOut OFF (1) 오인식",
    "페이드 기능 비활성화 오인식",
    "무드램프 즉시 ON/OFF 전환 (페이드 없음)")
GHI[('MoodControlMgt','MdLmpFadeSta','CORRUPT')] = (
    "From : 유효 FadeSta (1 또는 2)\nTo : Out of Range",
    "FadeInOutRdy 설정 불가",
    "무드램프 페이드 오동작 또는 비활성화")
GHI[('MoodControlMgt','MdLmpFadeSta','ALL')] = GHI[('MoodControlMgt','MdLmpFadeSta','CORRUPT')]

for _rgb, _color in [('MoodLed_RPWM','빨간색'),('MoodLed_GPWM','초록색'),('MoodLed_BPWM','파란색')]:
    GHI[('MoodControlMgt', _rgb, 'MORE')] = (
        f"From : 정상 PWM 범위 (0~255)\nTo : 정상 범위 초과",
        f"{_color} LED 과밝기",
        f"무드램프 {_color} 색상 오표현 (안전에 직접 영향 없음)")
    GHI[('MoodControlMgt', _rgb, 'LESS')] = (
        f"From : 정상 PWM 범위 (0~255)\nTo : 0 (OFF) 오인식",
        f"{_color} LED 소등",
        f"무드램프 {_color} 색상 표현 불가")
    GHI[('MoodControlMgt', _rgb, 'CORRUPT')] = (
        f"From : 유효 PWM 값\nTo : Out of Range",
        f"{_color} LED PWM 오류",
        f"무드램프 색상 오표현")
    GHI[('MoodControlMgt', _rgb, 'ALL')] = GHI[('MoodControlMgt', _rgb, 'MORE')]

GHI[('MoodControlMgt','PwrOnModeSta','ALL')] = (
    "From : PwrOnModeSta = READY(2)\nTo : PwrOnModeSta ≠ READY",
    "무드램프 활성화 조건 불충족",
    "무드램프 PWM 출력 OFF 강제 → 무드램프 소등")
GHI[('MoodControlMgt','SlvBrgtnsVal','MORE')] = (
    "From : 정상 밝기 (1~250)\nTo : 정상 범위 초과 (≥251)",
    "SlvBrgtnsVal = 0 처리 (최소 밝기)",
    "무드램프 밝기 OFF → 표시 불량")
GHI[('MoodControlMgt','SlvBrgtnsVal','LESS')] = (
    "From : 정상 밝기 (1~250)\nTo : 0 (최소값)",
    "SlvBrgtnsVal = 0 → BrgtStep = 0",
    "무드램프 최소 밝기 고정")
GHI[('MoodControlMgt','SlvBrgtnsVal','CORRUPT')] = (
    "From : 유효 밝기 값\nTo : Out of Range",
    "SlvBrgtnsVal 오인식",
    "무드램프 밝기 오표현")
GHI[('MoodControlMgt','SlvBrgtnsVal','ALL')] = GHI[('MoodControlMgt','SlvBrgtnsVal','MORE')]

for vk in ['SlvFadInTimetVal','SlvFadOutTimetVal','SlvXVal','SlvYVal']:
    GHI[('MoodControlMgt', vk, 'MORE')] = (
        f"From : 정상 범위\nTo : 정상 범위 초과",
        f"{vk} 비정상 고값",
        "무드램프 페이드/색상 오표현 (안전에 직접 영향 없음)")
    GHI[('MoodControlMgt', vk, 'LESS')] = (
        f"From : 정상 범위\nTo : 0 (최소값)",
        f"{vk} 최소값 클램프 적용",
        "무드램프 즉시 전환 (페이드 시간 최소)")
    GHI[('MoodControlMgt', vk, 'CORRUPT')] = (
        f"From : 유효 값\nTo : Out of Range",
        f"{vk} 오인식",
        "무드램프 오표현")
    GHI[('MoodControlMgt', vk, 'ALL')] = GHI[('MoodControlMgt', vk, 'MORE')]

for vk in ['UtilMode','ValGearSlctDis']:
    GHI[('MoodControlMgt', vk, 'ALL')] = (
        f"From : 정상 {vk}\nTo : {vk} 오신호",
        f"{vk} 오인식",
        "무드램프 모드 오동작 (안전에 직접 영향 없음)")

# ─────────────────────── CstAp_DtcMgt ───────────────────────
_SS_G = "From : DTC 미발생 상태\nTo : DTC 발생 → 스냅샷 캡처"
_SS_H = "DTC 발생 시점 시스템 상태 스냅샷 기록 오류"
_SS_I = "진단 스냅샷 데이터 손실 → 사후 분석 불가"
for _ss in ['SnapShot0200','SnapShot0202','SnapShot0203','SnapShot0204','SnapShot0205',
            'SnapShot0206','SnapShot0207','SnapShot0208','SnapShot0209','SnapShot020A',
            'SnapShot020B','SnapShot020D','SnapShot020E']:
    GHI[('DtcMgt', _ss, 'ALL')] = (_SS_G, _SS_H, _SS_I)

# DtcMgt 상태 변수들
_dtc_vars = {
    'BatVolt':              ("From : 정상 배터리 전압\nTo : 이상 전압", "BatVolt 이상값 DTC 조건 충족", "U3003_A2/A3 DTC 설정"),
    'BatOverSta':           ("From : BatOverSta = OFF\nTo : BatOverSta = ON", "과전압 판정", "U3003_A3 DTC 설정"),
    'BatUnderSta':          ("From : BatUnderSta = OFF\nTo : BatUnderSta = ON", "저전압 판정", "U3003_A2 DTC 설정"),
    'BatStbSta':            ("From : BatStbSta = ON (안정)\nTo : BatStbSta = OFF", "배터리 불안정", "LDO2/SBC 검사 비활성화"),
    'IgnVolt':              ("From : 정상 IGN 전압\nTo : 이상 전압", "IgnVolt 이상값", "PowerOnSta 오판단 → 시스템 모드 오전환"),
    'Ldo2FltSta':           ("From : Ldo2FltSta = OFF\nTo : Ldo2FltSta = ON", "LDO2 저전압 결함", "시스템 동작 제한"),
    'Ldo2OnVolt':           ("From : 정상 LDO2 전압 (>3.5V)\nTo : 저전압 (≤3.5V)", "Ldo2FltSta = ON", "전원 불안정 → 기능 제한"),
    'SbcFlt':               ("From : SBC 정상 (1)\nTo : SBC 결함 (0)", "SbcFltSta = ON", "SysPwrSta = POWER_OFF 강제"),
    'SysPwrSta':            ("From : POWER_ON\nTo : POWER_OFF", "DTC 인에이블 조건 불충족", "일부 DTC 설정 불가"),
    'ECUSta':               ("From : WAKEUP 상태\nTo : 비정상 상태", "DTC 인에이블 조건 불충족", "ECU 초기화 불완전"),
    'DriveSta':             ("From : DriveSta = ON\nTo : DriveSta = OFF", "주행 관련 DTC 조건 불충족", "주행 관련 DTC 미설정"),
    'DriveSigAllTo':        ("From : DriveSigAllTo = OFF\nTo : DriveSigAllTo = ON", "Main/Sub 양쪽 타임아웃", "P0A1D_8C DTC 설정"),
    'DriveSigMainTo':       ("From : Main CAN 정상\nTo : Main CAN 타임아웃", "DriveSigMainTo = ON", "Sub CAN 폴백; 양쪽 실패 시 DriveSigAllTo = ON"),
    'DriveSigSubTo':        ("From : Sub CAN 정상\nTo : Sub CAN 타임아웃", "DriveSigSubTo = ON", "Main CAN 유지"),
    'MainCANBusOffSta':     ("From : CAN 정상\nTo : Main BusOff", "MainCANBusOffSta = ON", "U0028_88 DTC 설정"),
    'SubCanBusOffSta':      ("From : CAN 정상\nTo : Sub BusOff", "SubCanBusOffSta = ON", "U0028_88 DTC 설정 (Sub)"),
    'MainCanRxSta':         ("From : Main CAN Rx 정상\nTo : Rx 오류", "MainCanRxSta 오류", "CAN 수신 오류 DTC 조건"),
    'MainCanTxSta':         ("From : Main CAN Tx 정상\nTo : Tx 오류", "MainCanTxSta 오류", "CAN 송신 오류 DTC 조건"),
    'SubCanRxSta':          ("From : Sub CAN Rx 정상\nTo : Rx 오류", "SubCanRxSta 오류", "Sub CAN 수신 오류 DTC 조건"),
    'SubCanTxSta':          ("From : Sub CAN Tx 정상\nTo : Tx 오류", "SubCanTxSta 오류", "Sub CAN 송신 오류 DTC 조건"),
    'MainShiftActMsgTo':    ("From : Main SActSig 정상\nTo : 타임아웃", "MainShiftActMsgTo = ON", "P0A1D_8C DTC 설정"),
    'SubShiftActMsgTo':     ("From : Sub SActSig 정상\nTo : 타임아웃", "SubShiftActMsgTo = ON", "P0A1D_8C DTC 조건 충족"),
    'BDC02MsgTo':           ("From : BDC_02 정상\nTo : 타임아웃", "BDC02MsgTo = ON", "U0840_8C DTC 설정"),
    'BDC05MsgTo':           ("From : BDC_05 정상\nTo : 타임아웃", "BDC05MsgTo = ON", "U0840_8C DTC 설정"),
    'CLUMsgTo':             ("From : CLU_01 정상\nTo : 타임아웃", "CLUMsgTo = ON", "U0855_8C DTC 설정"),
    'CLU_01_20ms_Timeout':  ("From : CLU_01 정상\nTo : 타임아웃 원시 플래그", "CLU_01_20ms_Timeout = ON", "CLUMsgTo = ON → U0855_8C DTC"),
    'SMK03MsgTo':           ("From : SMK_03 정상 (Main+Sub)\nTo : 양쪽 타임아웃", "SMK03MsgTo = ON", "U1065_8C DTC 설정"),
    'PDC03MsgTo':           ("From : PDC_03 정상\nTo : 타임아웃", "PDC03MsgTo = ON", "DTC 조건 충족; DrvDrSwSta = OFF"),
    'SactSig':              ("From : 정상 E2E 기어 신호\nTo : E2E 오류", "SActSig E2E 오류", "P0A1D_82/83 DTC 설정"),
    'SactSigTo':            ("From : SActSig 정상\nTo : 타임아웃", "SactSigTo = ON", "P0A1D_8C DTC 설정"),
    'SubSActSig':           ("From : Sub CAN 정상 E2E\nTo : E2E 오류", "Sub SActSig E2E 오류", "P0A1D_82/83 DTC"),
    'SubSActSigTo':         ("From : Sub SActSig 정상\nTo : 타임아웃", "SubSActSigTo = ON", "양쪽 타임아웃 시 P0A1D_8C DTC"),
    'PosSnrFlt':            ("From : 위치 센서 정상\nTo : 결함", "PosSnrFlt = ON", "P2E00_01 DTC → 모터 정지"),
    'PButtonStuck':         ("From : P버튼 정상\nTo : Stuck 판정 (18000ms)", "PButtonStuck = ON", "C1181_96 DTC 설정"),
    'IdtFltSta':            ("From : 인디케이터 정상\nTo : 회로 결함", "IdtFltSta = ON", "C1182_96 DTC 설정"),
    'IdtSta':               ("From : 인디케이터 정상\nTo : 이상", "IdtSta 이상", "C1182_96 DTC 조건 판단"),
    'DFlReadErrSta':        ("From : NVM 읽기 정상\nTo : 읽기 오류", "DFlReadErrSta = ON", "NVM 데이터 신뢰성 저하"),
    'DFlWriteErrSta':       ("From : NVM 쓰기 정상\nTo : 쓰기 오류", "DFlWriteErrSta = ON", "NVM 저장 실패"),
    'AlvCntFlt':            ("From : E2E AlvCnt 정상\nTo : 오류 누적", "AlvCntFlt = ON", "P0A1D_82 DTC 설정"),
    'AlvCntRep':            ("From : AlvCnt 정상\nTo : 반복 오류", "AlvCntRep = ON", "P0A1D_82 DTC 조건"),
    'AlvCntDiff':           ("From : AlvCnt 정상\nTo : 불일치", "AlvCntDiff = ON", "P0A1D_82 DTC 조건"),
    'SubAlvCntFlt':         ("From : Sub CAN AlvCnt 정상\nTo : 오류", "SubAlvCntFlt = ON", "Sub CAN 신호 무효화"),
    'SubAlvCntRep':         ("From : Sub AlvCnt 정상\nTo : 반복 오류", "SubAlvCntRep = ON", "Sub CAN 신호 무효화"),
    'SubAlvCntDiff':        ("From : Sub AlvCnt 정상\nTo : 불일치", "SubAlvCntDiff = ON", "Sub CAN 신호 무효화"),
    'CrcFltInfo.CrcFit':    ("From : CRC 정상\nTo : CRC 오류", "CrcFlt = ON", "P0A1D_83 DTC 설정"),
    'CrcFltInfo.SubCrcFit': ("From : Sub CRC 정상\nTo : Sub CRC 오류", "SubCrcFit = ON", "Sub CAN 신호 무효화"),
    'HallSnrFltInfo.Alpha': ("From : Hall Alpha 정상\nTo : Alpha 결함", "HallSnrFltInfo.TotalFlt 기여", "P2E00_01 DTC 조건"),
    'HallSnrFltInfo.Beta':  ("From : Hall Beta 정상\nTo : Beta 결함", "HallSnrFltInfo.TotalFlt 기여", "P2E00_01 DTC 조건"),
    'HallSnrFltInfo.TotalFlt':("From : Hall 정상\nTo : Hall 전체 결함", "HallSnrFltInfo.TotalFlt = ON", "P2E00_01 DTC → 모터 정지"),
    'HallSnrFltInfo.VG':    ("From : Hall VG 정상\nTo : VG 결함", "HallSnrFltInfo.TotalFlt 기여", "P2E00_01 DTC 조건"),
    'HallSnrFltVal.AlphaVal':("From : Alpha 정상 값\nTo : 이상 값", "Alpha 채널 결함 판정", "P2E00_01 DTC"),
    'HallSnrFltVal.BetaVal': ("From : Beta 정상 값\nTo : 이상 값", "Beta 채널 결함 판정", "P2E00_01 DTC"),
    'AutoBrightSta':        ("From : 자동 밝기 정상\nTo : 이상", "AutoBrightSta 이상", "C1182_96 DTC 조건 판단"),
    'GetPSigSta':           ("From : P위치 신호 정상\nTo : 이상", "GetPSigSta 이상", "P위치 전환 실패 DTC"),
    'GetPSigTo':            ("From : P위치 신호 정상\nTo : 타임아웃", "GetPSigTo = ON", "P위치 신호 타임아웃 DTC"),
    'RKESig':               ("From : RKE 정상 신호\nTo : 이상", "RKESig 이상", "원격 P위치 전환 차단"),
    'SMKSig':               ("From : SMK 정상 신호\nTo : 이상", "SMKSig 이상", "모터 활성화 조건 불충족"),
    'SleepModeFlag':        ("From : 정상 슬립 조건\nTo : 조건 불충족", "슬립 모드 진입 불가", "전력 소모 증가"),
    'CanBatNormalSta.DebounceBatNorSta': ("From : 배터리 정상\nTo : 배터리 비정상", "배터리 비정상 DTC 조건", "DTC 인에이블 조건 불충족"),
    'CanBatNormalSta.ImmediateBatNorSta':("From : 배터리 즉시 정상\nTo : 즉시 비정상", "배터리 즉시 이상", "DTC 즉시 설정"),
}
for vk, (g,h,i) in _dtc_vars.items():
    GHI[('DtcMgt', vk, 'ALL')] = (g, h, i)

for _idt in ['DIdtCntl','NIdtCntl','PIdtCntl','RIdtCntl']:
    for _sub in ['BltPwm','HltPwm','IdtOnsta']:
        vk = f"{_idt}.{_sub}"
        GHI[('DtcMgt', vk, 'ALL')] = (
            f"From : 정상 {_sub} 값\nTo : DTC 발생 시점 값",
            f"DTC 스냅샷에 {vk} 기록",
            "진단 도구로 DTC 발생 시 인디케이터 상태 확인 가능")
for _v in ['DMntrVolt','NMntrVolt','PMntrVolt','RMntrVolt']:
    GHI[('DtcMgt', _v, 'ALL')] = (
        f"From : 정상 모니터링 전압\nTo : 임계값 이상",
        f"{_v} 전압 이상",
        "IdtFltSta = ON → C1182_96 DTC")

# ══════════════════════════════════════════════════════════════
# 유틸 함수
# ══════════════════════════════════════════════════════════════
import re

def var_key(cell_val):
    if not cell_val: return ''
    return re.sub(r'\s*\(.*','', str(cell_val).split('\n')[0]).strip()

def unit_key(cell_val):
    if not cell_val: return ''
    s = str(cell_val)
    for k in ['PwrMGT','ECUModeMgt','CANMGT','MotorControlMgt','PosMgt',
              'MovingMgt','IdtMgt','ButtonMgt','HapticControlMgt','MoodControlMgt','DtcMgt']:
        if k in s: return k
    return s

def lookup_ghi(ukey, vkey, fm):
    r = GHI.get((ukey, vkey, fm))
    if r: return r
    r = GHI.get((ukey, vkey, 'ALL'))
    if r: return r
    short = vkey.split('.')[0]
    r = GHI.get((ukey, short, fm))
    if r: return r
    r = GHI.get((ukey, short, 'ALL'))
    if r: return r
    # 부분 매칭
    for (uk, vk2, fmk), val in GHI.items():
        if uk != ukey: continue
        if vk2 and (vk2 in vkey or vkey in vk2):
            if fmk == fm or fmk == 'ALL':
                return val
    return None

def make_g_fallback(vkey, fm, vtype):
    """DB 미매칭 시 From/To 패턴 자동생성"""
    fm_map = {
        'MORE':     ('정상 범위', '정상 범위 초과'),
        'LESS':     ('정상 범위', '정상 범위 미달'),
        'CORRUPT':  ('유효 신호', 'Out of Range 또는 오염'),
        'REVERSE':  ('정상 값', '반전 값'),
        'NO':       (f'{vkey} 정상 수신', '신호 없음'),
        'AS WELL AS': ('단일 신호', '추가 불필요 신호 발생'),
        'PART OF':  ('완전한 신호', '일부만 발생'),
        'EARLY':    ('정상 타이밍', '조기 발생'),
        'LATE':     ('정상 타이밍', '지연 발생'),
        'STUCK':    ('정상 변화', '고착 (Stuck)'),
        'OMISSION': (f'{vkey} 정상 수신', '신호 누락 (타임아웃)'),
        'WRONG':    ('올바른 신호', '잘못된 신호'),
        'COMMISSION':('신호 없음', '불필요한 신호 발생'),
    }
    fr, to = fm_map.get(fm, ('정상', '이상'))
    return f"From : {fr}\nTo   : {to}"

# ══════════════════════════════════════════════════════════════
# 메인 처리
# ══════════════════════════════════════════════════════════════
print("파일 복사 중...")
shutil.copy2(SRC, DEST)

print("Excel 시작...")
excel = win32.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

wb = excel.Workbooks.Open(DEST)
ws = wb.Worksheets('SW_FMEA')
max_row = ws.UsedRange.Rows.Count
print(f"총 {max_row}행 처리 시작")

filled_g = filled_h = filled_i = 0
skipped = 0
cur_unit = ''

for r in range(14, max_row + 1):
    u = ws.Cells(r, 2).Value
    if u: cur_unit = str(u)

    fm_raw = ws.Cells(r, 6).Value
    if not fm_raw: continue
    fm = str(fm_raw).strip().upper()

    ukey = unit_key(cur_unit)
    vkey = var_key(ws.Cells(r, 4).Value)
    vtype = str(ws.Cells(r, 5).Value or '')

    g_cell = ws.Cells(r, 7)
    h_cell = ws.Cells(r, 8)
    i_cell = ws.Cells(r, 9)

    g_val = g_cell.Value
    h_val = h_cell.Value
    i_val = i_cell.Value

    result = lookup_ghi(ukey, vkey, fm)

    # G: 비어있을 때만
    if not g_val:
        new_g = result[0] if result else make_g_fallback(vkey, fm, vtype)
        g_cell.Value = new_g
        filled_g += 1

    # H: 비어있을 때만
    if not h_val:
        new_h = result[1] if result else f"{vkey} {fm} → 모듈 내 처리 오류"
        h_cell.Value = new_h
        filled_h += 1

    # I: 비어있을 때만
    if not i_val:
        new_i = result[2] if result else f"{ukey} 기능 이상 → 시스템 영향 분석 필요"
        i_cell.Value = new_i
        filled_i += 1

    if not result:
        skipped += 1

    if r % 300 == 0:
        print(f"  {r}/{max_row} (G={filled_g}, H={filled_h}, I={filled_i})")

print(f"\n저장 중...")
wb.Save()
wb.Close(False)
excel.Quit()

print(f"\n완료!")
print(f"  G 채움: {filled_g}")
print(f"  H 채움: {filled_h}")
print(f"  I 채움: {filled_i}")
print(f"  DB 미매칭 (fallback): {skipped}")
print(f"  출력: {DEST}")
