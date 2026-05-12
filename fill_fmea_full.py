"""
JG1 SBW Software FMEA 전체 채우기
S, O, D, RPN, Preventive Action, Detection Action, Countermeasure 자동 입력
"""
import win32com.client

FMEA_IN  = r"E:\claude\FMEA\JG1_SBW-Software_FMEA_1.xlsx"
FMEA_OUT = r"E:\claude\FMEA\JG1_SBW-Software_FMEA_3_완성.xlsx"

FAILURE_MODES = {'MORE','LESS','CORRUPT','OMISSION','COMMISSION','WRONG','EARLY','LATE','STUCK'}

# ══════════════════════════════════════════════════════════════════════════════
# S / O / D 기준 정의 (소프트웨어 FMEA AIAG/VDA 기반)
# ══════════════════════════════════════════════════════════════════════════════

# ── Severity 기본값 (SW Unit별) ─────────────────────────────────────────────
UNIT_S = {
    'PwrMGT':           7,
    'ECUModeMgt':       7,
    'CANMGT':           6,
    'MotorControlMgt':  8,
    'HapticControlMgt': 5,
    'IdtMgt':           5,
    'MoodControlMgt':   3,
    'DtcMgt':           5,
    'PosMgt':           7,
    'MovingMgt':        7,
    'ButtonMgt':        6,
}

# ── 변수별 Severity 재정의 ────────────────────────────────────────────────────
# 형식: 변수명 키워드 → {FM: S} 또는 {'ALL': S}
VAR_S = {
    'BatVolt':                 {'MORE':7, 'LESS':8, 'CORRUPT':7, 'EARLY':6, 'LATE':6},
    'IgnVolt':                 {'ALL':6},
    'SbcFlt':                  {'ALL':8},
    'Ldo2OnVolt':              {'LESS':7, 'ALL':6},
    'TrmnlCtrlGrpStaBDCEV':   {'ALL':8},
    'ECUSta':                  {'ALL':7},
    'SActSig':                 {'CORRUPT':8, 'ALL':7},
    'SactSig':                 {'CORRUPT':8, 'ALL':7},
    'MainCANBusOFF':           {'ALL':7},
    'SubCANBusOFF':            {'ALL':7},
    'tmp_Position':            {'STUCK':9, 'CORRUPT':9, 'ALL':8},
    'RotateState':             {'ALL':8},
    'MotorActivation':         {'ALL':8},
    'MotorStopSig':            {'ALL':7},
    'SysPwrSta':               {'ALL':8},
    'DriveSta':                {'ALL':7},
    'DriveSigAllTo':           {'ALL':7},
    'DriveSigMainTo':          {'ALL':6},
    'DriveSigSubTo':           {'ALL':6},
    'PosSnrRaw1':              {'ALL':7},
    'PosSnrRaw2':              {'ALL':7},
    'PosSta':                  {'ALL':7},
    'PosStuck':                {'ALL':7},
    'PButtonSta':              {'ALL':6},
    'PButtonFltSta':           {'ALL':6},
    'PButtonStuck':            {'ALL':6},
    'IdtFltSta':               {'ALL':5},
    'IdtSta':                  {'ALL':5},
    'HallSnrFltInfo':          {'ALL':7},
    'AlvCntFlt':               {'ALL':7},
    'CrcFltInfo':              {'ALL':7},
    'MoodLed_BPWM':            {'ALL':3},
    'MoodLed_GPWM':            {'ALL':3},
    'MoodLed_RPWM':            {'ALL':3},
    'MdLmpFadeSta':            {'ALL':3},
    'SlvBrgtnsVal':            {'ALL':3},
    'BDC02MsgTo':              {'ALL':6},
    'BDC05MsgTo':              {'ALL':5},
    'CLUMsgTo':                {'ALL':4},
    'PDC03MsgTo':              {'ALL':5},
    'MainShiftActMsgTo':       {'ALL':7},
    'SubShiftActMsgTo':        {'ALL':7},
    'SMK03MsgTo':              {'ALL':7},
    'DFlReadErrSta':           {'ALL':6},
    'DFlWriteErrSta':          {'ALL':6},
    'LvrWrngMsg':              {'ALL':6},
    'lvrPosInfo':              {'ALL':7},
    'gearPositionVcu':         {'ALL':7},
    'sysPwrSta':               {'ALL':8},
    'BatStbSta':               {'ALL':6},
    'AutoBrightSta':           {'ALL':3},
    'AvTailLmpSta':            {'ALL':4},
    'InlTailLmpSta':           {'ALL':4},
    'DrvDrSwSta':              {'ALL':6},
    'DrvStOccSta':             {'ALL':6},
    'FaceDetectStat':          {'ALL':5},
    'Indi_tmp_Pos':            {'ALL':5},
    'PosSnr1':                 {'ALL':7},
    'PosSnr2':                 {'ALL':7},
    'MovSnr1':                 {'ALL':7},
    'MovSnr2':                 {'ALL':7},
    'P_SW1_Raw':               {'ALL':6},
    'P_SW2_Raw':               {'ALL':6},
}

# ── Occurrence 기본값 (SW Unit별) ────────────────────────────────────────────
UNIT_O = {
    'PwrMGT':           3,
    'ECUModeMgt':       3,
    'CANMGT':           3,
    'MotorControlMgt':  4,
    'HapticControlMgt': 3,
    'IdtMgt':           3,
    'MoodControlMgt':   3,
    'DtcMgt':           3,
    'PosMgt':           4,
    'MovingMgt':        4,
    'ButtonMgt':        4,
}

VAR_O = {
    'BatVolt':3, 'IgnVolt':3, 'SbcFlt':3, 'Ldo2OnVolt':3,
    'TrmnlCtrlGrpStaBDCEV':3, 'ECUSta':3,
    'SActSig':2, 'SactSig':2,  # E2E 보호
    'MainCANBusOFF':4, 'SubCANBusOFF':4,
    'tmp_Position':4, 'RotateState':3,
    'MotorActivation':3, 'MotorStopSig':3,
    'SysPwrSta':3, 'DriveSta':3, 'DriveSigAllTo':3,
    'DriveSigMainTo':3, 'DriveSigSubTo':3,
    'PosSnrRaw1':4, 'PosSnrRaw2':4,
    'PosSta':4, 'PosStuck':4,
    'PButtonSta':4, 'PButtonFltSta':4, 'PButtonStuck':4,
    'IdtFltSta':3, 'HallSnrFltInfo':4,
    'AlvCntFlt':2, 'CrcFltInfo':2,
    'MoodLed_BPWM':3, 'MoodLed_GPWM':3, 'MoodLed_RPWM':3,
    'BDC02MsgTo':3, 'BDC05MsgTo':3, 'CLUMsgTo':3, 'PDC03MsgTo':3,
    'MainShiftActMsgTo':3, 'SubShiftActMsgTo':3,
    'PosSnr1':4, 'PosSnr2':4, 'MovSnr1':4, 'MovSnr2':4,
    'P_SW1_Raw':4, 'P_SW2_Raw':4,
    'DrvDrSwSta':3, 'DrvStOccSta':3,
    'lvrPosInfo':4, 'gearPositionVcu':2,
}

# ── Detection 기본값 (SW Unit별) ─────────────────────────────────────────────
UNIT_D = {
    'PwrMGT':           3,
    'ECUModeMgt':       3,
    'CANMGT':           3,
    'MotorControlMgt':  3,
    'HapticControlMgt': 5,
    'IdtMgt':           3,
    'MoodControlMgt':   5,
    'DtcMgt':           2,
    'PosMgt':           3,
    'MovingMgt':        3,
    'ButtonMgt':        4,
}

VAR_D = {
    'BatVolt':2, 'IgnVolt':3, 'SbcFlt':2, 'Ldo2OnVolt':3,
    'TrmnlCtrlGrpStaBDCEV':2, 'ECUSta':4,
    'SActSig':2, 'SactSig':2,
    'MainCANBusOFF':2, 'SubCANBusOFF':2,
    'tmp_Position':2, 'RotateState':4,
    'MotorActivation':3, 'MotorStopSig':3,
    'SysPwrSta':2, 'DriveSta':2, 'DriveSigAllTo':2,
    'DriveSigMainTo':2, 'DriveSigSubTo':2,
    'PosSnrRaw1':2, 'PosSnrRaw2':2,
    'PosSta':2, 'PosStuck':3,
    'PButtonSta':3, 'PButtonFltSta':3, 'PButtonStuck':3,
    'IdtFltSta':2, 'HallSnrFltInfo':2,
    'AlvCntFlt':2, 'CrcFltInfo':2,
    'MoodLed_BPWM':5, 'MoodLed_GPWM':5, 'MoodLed_RPWM':5,
    'BDC02MsgTo':2, 'BDC05MsgTo':2, 'CLUMsgTo':2, 'PDC03MsgTo':2,
    'MainShiftActMsgTo':2, 'SubShiftActMsgTo':2,
    'PosSnr1':3, 'PosSnr2':3, 'MovSnr1':3, 'MovSnr2':3,
    'P_SW1_Raw':4, 'P_SW2_Raw':4,
    'DrvDrSwSta':3, 'DrvStOccSta':4,
    'lvrPosInfo':3, 'gearPositionVcu':2,
}

# ══════════════════════════════════════════════════════════════════════════════
# Preventive / Detection Action DB
# ══════════════════════════════════════════════════════════════════════════════
# (unit_key, var_keyword, fm_or_ALL) → (preventive, detection)
ACTION_DB = {
  # ── PwrMGT ──────────────────────────────────────────────────────────────────
  ('PwrMGT','BatVolt','MORE'):
    ("ADC 상한 클램프(4095); 배터리 전압 감시 로직 상시 동작",
     "BatVolt≥1910(≈16.5V) → 200ms 필터 → BatOverSta=ON → U3003_A3 DTC\n참조: CtAp_VBatStaChk.c:143-175"),
  ('PwrMGT','BatVolt','LESS'):
    ("ADC 하한 클램프; 배터리 전압 감시 로직 상시 동작",
     "BatVolt≤962(≈8.5V) → 200ms 필터 → BatUnderSta=ON → U3003_A2 DTC\n참조: CtAp_VBatStaChk.c:81-103"),
  ('PwrMGT','BatVolt','CORRUPT'):
    ("ADC 범위 클램프(0~4095); 임계값 기반 이중 판정",
     "ADC 4095 초과 → 클램프; 이후 과전압/저전압 임계값 판정 수행\n참조: CtIoHwAb_IntfIn.c:228"),
  ('PwrMGT','BatVolt','EARLY'):
    ("200ms 디바운스 필터(V_BAT_UNDER_TIME_SET) 적용으로 순간 변동 차단",
     "연속 200ms 이상 조건 유지 확인 후 BatOverSta/BatUnderSta 상태 전환\n참조: CtAp_VBatStaChk.c"),
  ('PwrMGT','BatVolt','LATE'):
    ("200ms 디바운스 필터 적용",
     "연속 200ms 조건 유지 후 상태 전환; 복구 시에도 동일 필터 적용\n참조: CtAp_VBatStaChk.c"),
  ('PwrMGT','IgnVolt','MORE'):
    ("ADC 상한 클램프(4095); 50ms 필터 적용",
     "IgnVolt≥762(≈7V) → 50ms(V_IGN_ON_CNT=5×10ms) 후 HWIGN=ON 설정\n참조: CtAp_IgnStaChk.c:69-94"),
  ('PwrMGT','IgnVolt','LESS'):
    ("ADC 하한 클램프; 임계값 이중 판정",
     "IgnVolt≤405(≈4V) → HWIGN=OFF; TrmnlCtrlGrpStaBDCEV=OFF AND 조건으로 PowerOnSta=OFF\n참조: CtAp_IgnStaChk.c:104-116"),
  ('PwrMGT','IgnVolt','CORRUPT'):
    ("ADC 범위 클램프(0~4095)",
     "ADC 이상값 클램프 후 ON/OFF 이중 임계값(762/405)으로 상태 판정\n참조: CtIoHwAb_IntfIn.c:228"),
  ('PwrMGT','Ldo2OnVolt','LESS'):
    ("LDO2 검사 활성화 조건(VBatStbSta=ON, ECUSta=WAKEUP) 이중 검증; 100ms 디바운스",
     "Ldo2OnVolt≤2866(≈3.5V) → 100ms 필터 → Ldo2FltSta=ON → 시스템 전원 상태 반영\n참조: CtAp_LdoStaChk.c:142-204"),
  ('PwrMGT','Ldo2OnVolt','ALL'):
    ("LDO2 검사 조건부 활성화(저전력 모드 중 오탐 방지)",
     "Ldo2 전압 이상 → Ldo2FltSta 플래그 설정; 상위 SysPwrSta 로직에 반영\n참조: CtAp_LdoStaChk.c:124-204"),
  ('PwrMGT','SbcFlt','ALL'):
    ("SBC 결함 입력 반전 로직(0=Fault, 1=Normal); 100ms 디바운스 필터",
     "SbcFltSta=ON → SysPwrSta=POWER_OFF 즉시 강제 설정\n참조: CtAp_LdoStaChk.c:211-229, CtAp_SysStaChk.c:61-78"),
  ('PwrMGT','TrmnlCtrlGrpStaBDCEV','ALL'):
    ("Dual-CAN 이중화: Main CAN 타임아웃 시 Sub CAN 값 사용",
     "SMK_03_Timeout 양쪽 동시 → TrmnlCtrlGrpStaBDCEV=0 → 모터 활성화 차단\n참조: CtAp_CGWSigChk.c:373-400"),
  ('PwrMGT','ECUSta','ALL'):
    ("RTE 읽기 실패 시 안전 기본값(EXTER_ECU_STANDBY=3) 적용",
     "RTE 반환값 검증; 이상 시 기본값 적용 후 상위 로직 정상 수행\n참조: CtAp_LdoStaChk.c:81-85"),

  # ── ECUModeMgt ───────────────────────────────────────────────────────────────
  ('ECUModeMgt','BDC02MsgTo','ALL'):
    ("CAN 스택 레벨 수신 타임아웃 자동 감지",
     "BDC_02 타임아웃 플래그 ON → IgnSwStaFlag=OFF 강제 → 점화 신호 무효화\n참조: CtAp_CGWSigChk.c:545-555"),
  ('ECUModeMgt','BDC05MsgTo','ALL'):
    ("CAN 스택 레벨 수신 타임아웃 자동 감지",
     "BDC_05 타임아웃 → 테일램프/무드램프 신호 전체 OFF 강제\n참조: CtAp_CGWSigChk.c:291-311"),
  ('ECUModeMgt','CLUMsgTo','ALL'):
    ("CAN 스택 레벨 수신 타임아웃 자동 감지",
     "CLU_01 타임아웃 → AutoBrightSta=OFF 강제 (안전 기본값)\n참조: CtAp_CGWSigChk.c:313-331"),
  ('ECUModeMgt','PDC03MsgTo','ALL'):
    ("CAN 스택 레벨 수신 타임아웃 자동 감지",
     "PDC_03 타임아웃 → DrvDrSwSta=OFF 강제 (안전 기본값)\n참조: CtAp_CGWSigChk.c:338-348"),
  ('ECUModeMgt','SysPwrSta','ALL'):
    ("SBC 결함(SbcFltSta) 및 전원 ON(PowerOnSta) 이중 조건 게이팅",
     "SysPwrSta=POWER_ON 조건 미충족 시 즉시 POWER_OFF 전환\n참조: CtAp_SysStaChk.c:61-78"),
  ('ECUModeMgt','TrmnlCtrlGrpStaBDCEV','ALL'):
    ("Dual-CAN 이중화(SMK Main/Sub); 양쪽 타임아웃 조건 모니터링",
     "양쪽 타임아웃 → TrmnlCtrlGrpStaBDCEV=0 → U1065_8C DTC → 모터 비활성화\n참조: CtAp_CGWSigChk.c:373-400"),
  ('ECUModeMgt','DriveSta','ALL'):
    ("Dual-CAN 이중화(Main/Sub); Main 실패 시 Sub CAN 값 폴백",
     "DriveSigMainTo AND SubTo 동시 → DriveSigAllTo=ON → 모터 활성화 차단\n참조: CtAp_DrvRdySigChk.c:78-119"),
  ('ECUModeMgt','DriveSigAllTo','ALL'):
    ("Main/Sub CAN 이중화; 양쪽 동시 실패 시에만 All-Timeout 선언",
     "DriveSigAllTo=ON → 모터 활성화 사전 조건 미충족 → 모터 정지 유지\n참조: CtAp_DrvRdySigChk.c:114-119"),
  ('ECUModeMgt','DriveSigMainTo','ALL'):
    ("Main CAN 수신 타임아웃 독립 모니터링",
     "Main CAN 타임아웃 → Sub CAN으로 자동 폴백; 양쪽 실패 시 DriveSigAllTo=ON\n참조: CtAp_DrvRdySigChk.c:97-121"),
  ('ECUModeMgt','DriveSigSubTo','ALL'):
    ("Sub CAN 수신 타임아웃 독립 모니터링",
     "Sub CAN 타임아웃 → Main CAN 유지; 양쪽 실패 시 DriveSigAllTo=ON\n참조: CtAp_DrvRdySigChk.c:97-121"),
  ('ECUModeMgt','DiagSession','ALL'):
    ("DriveSta 기반 진단 세션 전환 조건 제어",
     "주행 중(DriveSta=ON) 쓰기 서비스 제한; 정차 시에만 ECU 파라미터 변경 허용\n참조: CtAp_EcuModeCntl.c:277-289"),
  ('ECUModeMgt','MainShiftActMsgTo','ALL'):
    ("Main CAN ShiftAct 메시지 타임아웃 감지",
     "MainSActMsgToFlag=ON → GearPosSta 갱신 차단; Sub CAN 폴백\n참조: CtAp_ShiftActSigChk.c:299-312"),
  ('ECUModeMgt','SubShiftActMsgTo','ALL'):
    ("Sub CAN ShiftAct 메시지 타임아웃 감지",
     "SubSActMsgToFlag=ON → Sub CAN GearPosSta 갱신 차단\n참조: CtAp_ShiftActSigChk.c:299-312"),

  # ── CANMGT ──────────────────────────────────────────────────────────────────
  ('CANMGT','SActSig','MORE'):
    ("E2E CRC 검증(500ms 디바운스): E2E_P_ERROR → CrcErrFlag=ON",
     "P0A1D_83 DTC 설정; 신호 폐기 후 Sub CAN 폴백; GearPosSta=NOT_DISPLAY\n참조: CtAp_ShiftActSigChk.c:188-193"),
  ('CANMGT','SActSig','LESS'):
    ("E2E Alive Counter 검증(500ms): E2E_P_REPEATED → AlvCntRepFlag=ON",
     "P0A1D_82 DTC 설정; 신호 폐기; GearPosSta=NOT_DISPLAY\n참조: CtAp_ShiftActSigChk.c:174-179"),
  ('CANMGT','SActSig','CORRUPT'):
    ("E2E CRC + Alive Counter 이중 검증; 500ms 디바운스 적용",
     "P0A1D_82/83 DTC 설정; 신호 무효화; GearPosSta=NOT_DISPLAY\n참조: CtAp_ShiftActSigChk.c:165-202"),
  ('CANMGT','SActSig','OMISSION'):
    ("CAN 메시지 수신 타임아웃 감지; Dual-CAN 폴백",
     "SActSigTo=ON → P0A1D_8C DTC; MainSActMsgToFlag=ON → Sub CAN 폴백\n참조: CtAp_ShiftActSigChk.c:299-312"),
  ('CANMGT','SActSig','ALL'):
    ("E2E CRC + AlvCnt 이중 검증; Dual-CAN 이중화",
     "E2E 오류 → P0A1D_82/83/8C DTC; GearPosSta=NOT_DISPLAY\n참조: CtAp_ShiftActSigChk.c"),
  ('CANMGT','MainCANBusOFF','ALL'):
    ("CAN BusOff 700ms 디바운스 카운터(BUSOFF_CHECK_TIME=70×10ms)",
     "U0028_88 DTC 설정; MainCANBusOFFSta=ON → Main CAN 전체 신호 폐기; Sub CAN 폴백\n참조: CtAp_CANBusOffChk.c:70-77"),
  ('CANMGT','SubCANBusOFF','ALL'):
    ("Sub CAN BusOff 700ms 디바운스 카운터",
     "U0028_88 DTC 설정; SubCANBusOFFSta=ON → Sub CAN 전체 신호 폐기; Main CAN 유지\n참조: CtAp_CANBusOffChk.c:86-93"),
  ('CANMGT','Main_Vcu_E2E_Error_Sta','ALL'):
    ("E2E 프로파일 검증: CRC + AlvCnt 오류 코드(0x01~0x05) 분류 처리",
     "오류 코드별 플래그 설정 → P0A1D_82/83 DTC; 비정상 데이터 폐기\n참조: CtAp_ShiftActSigChk.c:165-202"),
  ('CANMGT','Main_Vcu_E2E_Error_Return','ALL'):
    ("E2E 반환값 코드 검증 후 정상/오류 분류",
     "E2E 반환 오류 → AlvCntFlt/CrcFlt 플래그 설정 → DTC 연동\n참조: CtAp_ShiftActSigChk.c:165-202"),
  ('CANMGT','SMK_TrmnlCtrlGrpStaBDCEV','ALL'):
    ("Dual-CAN 이중화: Main SMK 타임아웃 시 Sub SMK 값 사용",
     "양쪽 모두 타임아웃 → TrmnlCtrlGrpStaBDCEV=0 → 모터 활성화 차단(U1065_8C DTC)\n참조: CtAp_CGWSigChk.c:373-400"),
  ('CANMGT','SMK_TrmnlCtrlStaBDC','ALL'):
    ("Dual-CAN 이중화; CAN 수신 타임아웃 감지",
     "SMK_03 타임아웃 감지 → 신호 무효화 → 안전 기본값 적용\n참조: CtAp_CGWSigChk.c:373-402"),
  ('CANMGT','SMK_PwrOnModeSta','ALL'):
    ("SMK 메시지 타임아웃 감지; Dual-CAN 폴백",
     "SMK_03 타임아웃 → PwrOnModeSta=OFF 강제 → 무드램프/Indicator 비활성화\n참조: CtAp_CGWSigChk.c:373-400"),
  ('CANMGT','SubDrvRdySig','ALL'):
    ("Sub CAN 수신 타임아웃 모니터링; Main CAN DrvRdySig와 이중화",
     "SubDrvRdySigTo=ON → DriveSigSubTo=ON; 양쪽 타임아웃 → DriveSigAllTo=ON\n참조: CtAp_DrvRdySigChk.c:56-119"),
  ('CANMGT','PosSta','ALL'):
    ("PosSnrFlt 플래그 기반 위치 유효성 검사",
     "PosSnrFlt=ON → LvrPosSta=LVR_Flt(0x0F) → P2E00_01 DTC → CAN FAULT 상태 전송\n참조: CtAp_SBWSigSet.c:243-248"),
  ('CANMGT','PosStuck','ALL'):
    ("위치 변화량(Delta) 모니터링: ±2 이하 500ms 지속 → Stuck 판정",
     "PosStuck=ON → LvrWrngMsg=LEVER_STUCK(0x12) → 경고 3회 후 모터 정지\n참조: CtAp_SBWSigSet.c:341-343"),
  ('CANMGT','PButtonSta','ALL'):
    ("P버튼 이중 채널(P_SW1_Raw, P_SW2_Raw) 교차 검증",
     "PButtonFault=P_BUTTON_FAIL OR PButtonStuck=ON → LvrPSta=P_FAULT → C1181_96 DTC\n참조: CtAp_SBWSigSet.c:214-217"),
  ('CANMGT','PButtonStuck','ALL'):
    ("P버튼 stuck 감지 로직; 이중 채널 모니터링",
     "PButtonStuck=ON → LvrPSta=P_FAULT(0x03) → C1181_96 DTC\n참조: CtAp_SBWSigSet.c:214"),
  ('CANMGT','IdtFltSta','ALL'):
    ("Indicator 회로 결함 플래그 모니터링",
     "IdtFltSta=ON → LvrIdtSta=INDI_FAULT → C1182_96 DTC 설정\n참조: CtAp_SBWSigSet.c:198-203"),
  ('CANMGT','IdtSta','ALL'):
    ("Indicator 상태 유효성 확인",
     "Indicator 이상 상태 → C1182_96 DTC; 출력 OFF 강제\n참조: CtAp_SBWSigSet.c:198-203"),
  ('CANMGT','RotateState','ALL'):
    ("상태 머신: 유효값(0=Dial, 1=Sphere)만 허용",
     "유효 범위 외 → 상태 머신 진입 불가 → 모터 비활성 유지(안전 상태)\n참조: CtAp_MotorControl.c:435,542"),
  ('CANMGT','MotorStopSig','ALL'):
    ("3회 재시도 후 영구 정지 명령(TurnError≥3)",
     "MotorStopSig=1 → MotorControlStop() → Motor_State=0, Step_EN=1, tmp_Speed=0\n참조: CtAp_MotorControl.c:688,753,856"),
  ('CANMGT','SysPwrSta','ALL'):
    ("SBC 결함 게이팅 + PowerOnSta 이중 조건",
     "SysPwrSta≠POWER_ON → 모터/Indicator/무드램프 비활성화\n참조: CtAp_SysStaChk.c:61-78"),
  ('CANMGT','BDC_02_Timeout','ALL'):
    ("CAN 스택 타임아웃 자동 감지",
     "BDC_02 타임아웃 → 관련 신호 OFF 강제\n참조: CtAp_CGWSigChk.c:545"),
  ('CANMGT','BDC_05_Timeout','ALL'):
    ("CAN 스택 타임아웃 자동 감지",
     "BDC_05 타임아웃 → 무드램프/테일램프 신호 OFF\n참조: CtAp_CGWSigChk.c:291"),
  ('CANMGT','GetPSigSta','ALL'):
    ("CAN 수신 신호 유효성 확인",
     "GetP 신호 이상 → 안전 기본값 적용\n참조: CAN_Management/"),
  ('CANMGT','RKESig','ALL'):
    ("CAN 수신 타임아웃 감지",
     "RKE 신호 타임아웃 → 기본값 적용\n참조: CAN_Management/"),
  ('CANMGT','DrvDrSwSta','ALL'):
    ("운전석 도어 상태 CAN 신호 타임아웃 감지; PDC_03 타임아웃 시 OFF 강제",
     "PDC_03 타임아웃 → DrvDrSwSta=OFF(안전 기본값) → 모드 변경 비활성화\n참조: CtAp_CGWSigChk.c:338-348"),
  ('CANMGT','DrvStOccSta','ALL'):
    ("운전석 점유 상태 CAN 수신 확인",
     "DrvStOccSta 이상 → 모터 활성화 사전 조건 불충족\n참조: CAN_Management/"),

  # ── MotorControlMgt ──────────────────────────────────────────────────────────
  ('MotorControlMgt','tmp_Position','MORE'):
    ("위치 센서 유효 범위 검사: 100≤value≤900만 허용; 범위 외 미갱신",
     "범위 초과 → Moving_Position 미갱신 → 500ms 후 Stuck 감지 → P2E00_01 DTC\n참조: CtAp_MotorControl.c:231-233"),
  ('MotorControlMgt','tmp_Position','LESS'):
    ("위치 센서 유효 범위 검사: 100≤value≤900만 허용",
     "범위 미달 → Moving_Position 미갱신 → Stuck 감지 → P2E00_01 DTC\n참조: CtAp_MotorControl.c:231-233"),
  ('MotorControlMgt','tmp_Position','CORRUPT'):
    ("범위 필터링(100~900); SBWSigSet 추가 검증",
     "범위 외 → LvrModePosInfo=3(오류) → CAN Not Available 전송; P2E00_01 DTC\n참조: CtAp_SBWSigSet.c:153-158"),
  ('MotorControlMgt','tmp_Position','STUCK'):
    ("Delta 변화량 모니터링: ±2 이하 500ms 지속 시 Stuck 판정",
     "TurnDialError/SphereError 카운터 → 3회 실패 → MotorStopSig=1 → P2E00_01 DTC\n참조: CtAp_MotorControl.c:661-933"),
  ('MotorControlMgt','tmp_Position','ALL'):
    ("범위 검사(100~900) + Delta 변화량 모니터링",
     "범위 외 또는 Stuck → P2E00_01 DTC; MotorStopSig=1\n참조: CtAp_MotorControl.c:231,661-933"),
  ('MotorControlMgt','RotateState','ALL'):
    ("NVM 저장값; 상태 머신에서 0/1 범위만 허용",
     "유효 범위 외 → 상태 머신 진입 불가 → 모터 비활성 유지\n참조: CtAp_MotorControl.c:435,542"),
  ('MotorControlMgt','SysPwrSta','ALL'):
    ("SBC 결함 게이팅 + 전원 ON 조건 이중 검증; 전원 사이클 시 경고 클리어",
     "SysPwrSta≠POWER_ON → 모터 비활성화; 경고 자동 클리어로 재시도 허용\n참조: CtAp_MotorControl.c:368-392"),
  ('MotorControlMgt','TrmnlCtrlGrpStaBDCEV','ALL'):
    ("Dual-CAN 이중화; BDCEV_READY(3) 또는 BDCEV_POWERON(1) 조건 확인",
     "조건 미충족 → 상태 머신 진입 불가 → 모터 비활성화\n참조: CtAp_MotorControl.c:402-407"),
  ('MotorControlMgt','DriveSta','ALL'):
    ("Dual-CAN 이중화; DriveSigAllTo 모니터링",
     "DriveSigAllTo=ON → 모터 활성화 조건 불충족; 모터 정지 유지\n참조: CtAp_MotorControl.c:414-432"),
  ('MotorControlMgt','DriveSigAllTo','ALL'):
    ("Main/Sub CAN 이중화 모니터링",
     "DriveSigAllTo=ON → 모터 활성화 조건 차단\n참조: CtAp_DrvRdySigChk.c:114-119"),
  ('MotorControlMgt','MotorActivation','ALL'):
    ("100ms 워치독(MOTOR_WAITCOUNT=10×10ms): 비활성 후 재활성화 조건 재검증",
     "Step_EN=0 유지 100ms 초과 → MotorActivation=OFF 자동 전환 → 하드웨어 비활성화\n참조: CtAp_MotorControl.c:266-282"),
  ('MotorControlMgt','MotorStopSig','ALL'):
    ("3회 재시도 후 영구 정지(TurnError≥3); 다음 전원 사이클까지 차단",
     "MotorStopSig=1 → MotorControlStop() → Step_EN=1, tmp_Speed=0 → P2E00_01 연동\n참조: CtAp_MotorControl.c:688,753,856"),
  ('MotorControlMgt','DrvDrSwSta','ALL'):
    ("운전석 도어 상태 확인(모터 활성화 사전 조건); PDC_03 타임아웃 시 OFF",
     "DrvDrSwSta≠CLOSED → 모드 변경 비활성화; 의도치 않은 레버 동작 방지\n참조: CtAp_MotorControl.c:414-432"),
  ('MotorControlMgt','DrvStOccSta','ALL'):
    ("운전석 점유 확인(모터 활성화 사전 조건); 10사이클 지연 검증",
     "DrvStOccSta≠SEATED → 모드 변경 비활성화; 10사이클(100ms) 조건 유지 후 진입\n참조: CtAp_MotorControl.c:414-432"),
  ('MotorControlMgt','FaceDetectStat','ALL'):
    ("얼굴 인식 신호 CAN 수신 확인; 타임아웃 시 이전값 유지",
     "FaceDetectStat 이상 → 관련 기능 비활성화; 안전 기본값 적용\n참조: CAN_Management/"),
  ('MotorControlMgt','PwrOnModeSta','ALL'):
    ("SMK PwrOnMode 신호 유효성 확인(SMK_03 타임아웃 연동)",
     "PwrOnModeSta=OFF → 모터 활성화 사전 조건 불충족\n참조: CtAp_CGWSigChk.c:373-400"),
  ('MotorControlMgt','Step_DIR','ALL'):
    ("모터 방향 제어 값 범위 검사(CW/CCW 이진값)",
     "방향 설정 오류 → 상태 머신 오류 감지 → MotorStopSig 연동\n참조: CtAp_MotorControl.c"),
  ('MotorControlMgt','Step_EN','ALL'):
    ("모터 Enable 핀 상태 모니터링; 100ms 워치독 적용",
     "Step_EN=0 비정상 지속 → 워치독 만료 후 강제 비활성화\n참조: CtAp_MotorControl.c:266-282"),
  ('MotorControlMgt','Motor_Speed','ALL'):
    ("모터 속도 제어값 범위 제한",
     "Stuck 감지 시 속도값 변경(재시도); 3회 실패 후 Speed=0\n참조: CtAp_MotorControl.c:661-933"),

  # ── PosMgt ──────────────────────────────────────────────────────────────────
  ('PosMgt','PositionSensor_PosSnrRaw1','ALL'):
    ("위치 센서 유효 범위 검사(100~900); 이중 센서(Raw1/Raw2) 교차 검증",
     "범위 외 → PosSnrFlt=ON → P2E00_01 DTC → LvrPosSta=LVR_Flt\n참조: CtAp_SBWSigSet.c:153-158"),
  ('PosMgt','PositionSensor_PosSnrRaw2','ALL'):
    ("위치 센서 유효 범위 검사(100~900); 이중 센서 교차 검증",
     "범위 외 → PosSnrFlt=ON → P2E00_01 DTC → LvrPosSta=LVR_Flt\n참조: CtAp_SBWSigSet.c:153-158"),
  ('PosMgt','PositionSensorInfo_PButtonFltSta','ALL'):
    ("P버튼 이중 채널(P_SW1_Raw, P_SW2_Raw) 불일치 감지",
     "PButtonFaultSta=ON → C1181_96 DTC; LvrPSta=P_FAULT → P위치 진입 불가\n참조: CtAp_SBWSigSet.c:214-217"),
  ('PosMgt','PositionSensorInfo_PButtonSta','ALL'):
    ("P버튼 양쪽 채널 일치 확인; 타임아웃/오류 감지",
     "PButtonFault → C1181_96 DTC; P위치 기능 차단\n참조: CtAp_SBWSigSet.c:214"),

  # ── MovingMgt ────────────────────────────────────────────────────────────────
  ('MovingMgt','PosSnr1','ALL'):
    ("Active/Period 시간 비율 기반 위치 계산; 신호 유효성 확인",
     "신호 이상 → PosSnrFlt=ON → P2E00_01 DTC; 위치 미갱신\n참조: Position_Management/"),
  ('MovingMgt','PosSnr2','ALL'):
    ("Active/Period 시간 비율 기반 위치 계산; 이중 센서 검증",
     "신호 이상 → PosSnrFlt=ON → P2E00_01 DTC\n참조: Position_Management/"),
  ('MovingMgt','MovSnr1','ALL'):
    ("이동 센서 Active/Period 신호 모니터링; Delta 변화량 감시",
     "MovSnr 신호 이상 → Stuck 감지 카운터 증가 → 3회 후 MotorStopSig=1\n참조: CtAp_MotorControl.c:874-933"),
  ('MovingMgt','MovSnr2','ALL'):
    ("이동 센서 Active/Period 신호 모니터링",
     "MovSnr 신호 이상 → Stuck 감지 카운터 증가 → 3회 후 MotorStopSig=1\n참조: CtAp_MotorControl.c:874-933"),

  # ── ButtonMgt ────────────────────────────────────────────────────────────────
  ('ButtonMgt','P_SW1_Raw','ALL'):
    ("P버튼 채널1 모니터링; 채널2(P_SW2_Raw)와 교차 검증",
     "채널 불일치 또는 Stuck → PButtonFault 설정 → C1181_96 DTC\n참조: Button_Management/"),
  ('ButtonMgt','P_SW2_Raw','ALL'):
    ("P버튼 채널2 모니터링; 채널1(P_SW1_Raw)과 교차 검증",
     "채널 불일치 또는 Stuck → PButtonFault 설정 → C1181_96 DTC\n참조: Button_Management/"),

  # ── IdtMgt ───────────────────────────────────────────────────────────────────
  ('IdtMgt','BltDimLvl','ALL'):
    ("PWM 출력 범위 제한; 배터리 상태 기반 조건부 활성화",
     "PWM 이상 → Indicator 출력 오류; C1182_96 DTC 연동\n참조: Indicator_Management/"),
  ('IdtMgt','HltDimLvl','ALL'):
    ("PWM 출력 범위 제한",
     "PWM 이상 → Indicator 출력 오류 감지\n참조: Indicator_Management/"),
  ('IdtMgt','BatStbSta','ALL'):
    ("배터리 안정화 상태 확인(LDO2 검사 활성화 조건); 오탐 방지",
     "BatStbSta=OFF → LDO2 결함 검사 비활성화; 전원 안정화 후 재활성화\n참조: CtAp_LdoStaChk.c:124"),
  ('IdtMgt','IdtFltSta','ALL'):
    ("Indicator PWM 출력 결함 감지; 드라이버 IC 상태 모니터링",
     "IdtFltSta=ON → C1182_96 DTC; Indicator 출력 OFF 강제\n참조: CtDem_DTC_C1182_96_IndicatorFault.c"),
  ('IdtMgt','Ldo2FltSta','ALL'):
    ("LDO2 전압 하한 100ms 디바운스 필터; 활성화 조건 검증",
     "Ldo2FltSta=ON → Indicator 전원 공급 이상; C1182_96 DTC 연동\n참조: CtAp_LdoStaChk.c:142-204"),
  ('IdtMgt','TotalAlvCntFlt','ALL'):
    ("Main/Sub CAN E2E Alive Counter 이중 검증",
     "Main AND Sub 모두 AlvCntFlt=ON → TotalAlvCntFlt=ON → P0A1D_82 DTC\n참조: CtAp_ShiftActSigChk.c:352-358"),
  ('IdtMgt','TotalCrcFlt','ALL'):
    ("Main/Sub CAN E2E CRC 이중 검증",
     "Main AND Sub 모두 CrcFlt=ON → TotalCrcFlt=ON → P0A1D_83 DTC\n참조: CtAp_ShiftActSigChk.c"),
  ('IdtMgt','SysPwrSta','ALL'):
    ("SBC 결함 게이팅 + 전원 ON 이중 조건",
     "SysPwrSta≠POWER_ON → Indicator 비활성화\n참조: CtAp_SysStaChk.c:61-78"),
  ('IdtMgt','GearPosSta','ALL'):
    ("E2E 검증 결과 기반 기어 위치 유효성 확인",
     "GearPosSta=NOT_DISPLAY → Indicator 표시 불가 상태 처리\n참조: CtAp_ShiftActSigChk.c"),
  ('IdtMgt','MainCANBusOffSta','ALL'):
    ("CAN BusOff 700ms 디바운스; DTC 연동",
     "MainCANBusOffSta=ON → U0028_88 DTC → 관련 Indicator 신호 OFF\n참조: CtAp_CANBusOffChk.c:70-77"),
  ('IdtMgt','AutoBrightSta','ALL'):
    ("CLU 타임아웃 감지 시 AutoBrightSta=OFF 안전 기본값",
     "타임아웃 → AutoBrightSta=OFF; 수동 밝기 모드 유지\n참조: CtAp_CGWSigChk.c:327"),
  ('IdtMgt','MotorStopSig','ALL'):
    ("3회 재시도 후 모터 정지 신호",
     "MotorStopSig=1 → Indicator에 오류 상태 표시\n참조: CtAp_MotorControl.c:688"),
  ('IdtMgt','DMntrVolt2','ALL'):
    ("ADC 측정값 범위 검사; 전압 모니터링",
     "전압 이상 → 관련 DTC 설정; 모니터링 채널별 독립 감지\n참조: CtIoHwAb_IntfIn.c"),
  ('IdtMgt','PwrOnModeSta','ALL'):
    ("SMK PwrOnMode 수신 유효성 확인",
     "PwrOnModeSta 이상 → Indicator 비활성화\n참조: CtAp_CGWSigChk.c:373-400"),

  # ── DtcMgt ──────────────────────────────────────────────────────────────────
  ('DtcMgt','MainCANBusOffSta','ALL'):
    ("CAN BusOff 700ms 디바운스 카운터",
     "Dem_SetEventStatus(FAILED) → U0028_88 DTC 기록; SnapshotData 저장\n참조: CtAp_CANBusOffChk.c:70-77, CtDem_DTC_U0028_88.c"),
  ('DtcMgt','SubCanBusOffSta','ALL'):
    ("Sub CAN BusOff 700ms 디바운스 카운터",
     "Dem_SetEventStatus(FAILED) → U0028_88 DTC 기록\n참조: CtAp_CANBusOffChk.c:86-93"),
  ('DtcMgt','HallSnrFltInfo','ALL'):
    ("Hall 센서 Alpha/Beta/VG 채널별 결함 독립 감지",
     "P2E00_01 DTC 설정; LvrPosSta=LVR_Flt → 모터 위치 제어 불가\n참조: CtDem_DTC_P2E00_01_LeverHallSensorFault.c"),
  ('DtcMgt','BatVolt','MORE'):
    ("BatVolt ADC 상한 검사 + 200ms 필터",
     "Dem_SetEventStatus(FAILED) → U3003_A3 DTC (과전압); 스냅샷 저장\n참조: CtDem_DTC_U3003_A3_BatteryVoltageHigh.c"),
  ('DtcMgt','BatVolt','LESS'):
    ("BatVolt ADC 하한 검사 + 200ms 필터",
     "Dem_SetEventStatus(FAILED) → U3003_A2 DTC (저전압); 스냅샷 저장\n참조: CtDem_DTC_U3003_A2_BatteryVoltageLow.c"),
  ('DtcMgt','PButtonStuck','ALL'):
    ("P버튼 Stuck 판정 로직; 이중 채널 모니터링",
     "Dem_SetEventStatus(FAILED) → C1181_96 DTC; LvrPSta=P_FAULT\n참조: CtDem_DTC_C1181_96_Park_Switch_Fault.c"),
  ('DtcMgt','PosSnrFlt','ALL'):
    ("위치 센서 범위 검사 및 신호 유효성 검증(100~900)",
     "Dem_SetEventStatus(FAILED) → P2E00_01 DTC; LvrPosSta=LVR_Flt\n참조: CtDem_DTC_P2E00_01_LeverHallSensorFault.c"),
  ('DtcMgt','AlvCntFlt','ALL'):
    ("E2E Alive Counter 검증 500ms 디바운스; Dual-CAN 이중화",
     "Dem_SetEventStatus(FAILED) → P0A1D_82 DTC; GearPosSta=NOT_DISPLAY\n참조: CtDem_DTC_P0A1D_82_VCU_AC_Fault.c"),
  ('DtcMgt','CrcFltInfo','ALL'):
    ("E2E CRC 검증 500ms 디바운스; Dual-CAN 이중화",
     "Dem_SetEventStatus(FAILED) → P0A1D_83 DTC; GearPosSta=NOT_DISPLAY\n참조: CtDem_DTC_P0A1D_83_VCU_CRC_Fault.c"),
  ('DtcMgt','IdtFltSta','ALL'):
    ("Indicator 회로 결함 감지; PWM 피드백 확인",
     "Dem_SetEventStatus(FAILED) → C1182_96 DTC; LvrIdtSta=INDI_FAULT\n참조: CtDem_DTC_C1182_96_IndicatorFault.c"),
  ('DtcMgt','DFlReadErrSta','ALL'):
    ("Flash 읽기 오류 감지; CRC 검증",
     "DFlReadErrSta=ON → C1180_44 DTC (ECU Error); 시스템 안전 상태 전환\n참조: CtDem_DTC_C1180_44_ECU_Error.c"),
  ('DtcMgt','DFlWriteErrSta','ALL'):
    ("Flash 쓰기 오류 감지; 쓰기 후 검증",
     "DFlWriteErrSta=ON → C1180_44 DTC (ECU Error)\n참조: CtDem_DTC_C1180_44_ECU_Error.c"),
  ('DtcMgt','ECUSta','ALL'):
    ("ECU 상태 유효성 확인; 안전 기본값(STANDBY) 적용",
     "ECUSta 이상 → C1180_44 DTC 연동; 안전 기본값으로 복귀\n참조: CtDem_DTC_C1180_44_ECU_Error.c"),
  ('DtcMgt','SysPwrSta','ALL'):
    ("SBC 결함 + 전원 조건 이중 검증",
     "SysPwrSta=POWER_OFF → 관련 DTC 스냅샷 저장\n참조: CtAp_SysStaChk.c"),
  ('DtcMgt','SMK03MsgTo','ALL'):
    ("SMK_03 Dual-CAN 타임아웃 감지",
     "Dem_SetEventStatus(FAILED) → U1065_8C DTC (Local CAN Timeout)\n참조: CtDem_DTC_U1065_8C_Local_CANSignalTimeout.c"),
  ('DtcMgt','MainShiftActMsgTo','ALL'):
    ("Main CAN ShiftAct 메시지 타임아웃 감지",
     "Dem_SetEventStatus(FAILED) → P0A1D_8C DTC\n참조: CtDem_DTC_P0A1D_8C_VCU_CANSignalTimeout.c"),
  ('DtcMgt','DriveSta','ALL'):
    ("Dual-CAN DriveSig 이중화 감시",
     "DriveSta 이상 → 관련 DTC 기록; 모터 활성화 차단\n참조: CtAp_DrvRdySigChk.c"),
  ('DtcMgt','SbcFlt','ALL'):
    ("SBC 결함 100ms 디바운스 감지",
     "SbcFltSta=ON → C1180_44 DTC; SysPwrSta=POWER_OFF\n참조: CtAp_LdoStaChk.c:211-229"),

  # ── MoodControlMgt ─────────────────────────────────────────────────────────
  ('MoodControlMgt','MoodLed_BPWM','ALL'):
    ("PWM 출력 범위 제한(0~100%); Ldo2FltSta=ON 시 0 강제",
     "Ldo2FltSta=ON → PWM 출력 0으로 강제; 무드램프 소등(안전 기본값)\n참조: MoodControl/"),
  ('MoodControlMgt','MoodLed_GPWM','ALL'):
    ("PWM 출력 범위 제한; 전원 상태 조건부 활성화",
     "SysPwrSta≠POWER_ON → PWM 출력 차단; 무드램프 소등\n참조: MoodControl/"),
  ('MoodControlMgt','MoodLed_RPWM','ALL'):
    ("PWM 출력 범위 제한; 전원 상태 조건부 활성화",
     "SysPwrSta≠POWER_ON → PWM 출력 차단\n참조: MoodControl/"),
  ('MoodControlMgt','BDC05MsgTimeout','ALL'):
    ("CAN BDC_05 메시지 타임아웃 자동 감지",
     "타임아웃 → 무드램프 신호 OFF 강제; 안전 기본값 유지\n참조: CtAp_CGWSigChk.c:291-311"),
  ('MoodControlMgt','PwrOnModeSta','ALL'):
    ("SMK PwrOnMode 신호 유효성 확인; 타임아웃 감지",
     "PwrOnModeSta=OFF → 무드램프 출력 비활성화\n참조: CtAp_CGWSigChk.c:373-400"),
  ('MoodControlMgt','MdLmpFadeSta','ALL'):
    ("페이드 상태 값 범위 확인",
     "유효 범위 외 → 안전 기본값(OFF) 적용\n참조: MoodControl/"),
  ('MoodControlMgt','SlvBrgtnsVal','ALL'):
    ("밝기 값 범위 제한(0~255)",
     "범위 외 → 클램프 후 적용; 과도한 밝기 방지\n참조: MoodControl/"),
  ('MoodControlMgt','UtilMode','ALL'):
    ("UtilMode 신호 유효성 확인",
     "UtilMode 이상 → 안전 기본값 적용\n참조: MoodControl/"),

  # ── HapticControlMgt ───────────────────────────────────────────────────────
  ('HapticControlMgt','sysPwrSta','ALL'):
    ("SysPwrSta 조건 확인 후 Haptic 활성화",
     "SysPwrSta≠POWER_ON → Haptic 비활성화; 진동 없음(안전 기본값)\n참조: Haptic_Management/"),
  ('HapticControlMgt','gearPositionVcu','ALL'):
    ("유효한 기어 포지션 확인(GearPosSta≠NOT_DISPLAY)",
     "GearPosSta=NOT_DISPLAY → Haptic 피드백 제공 불가\n참조: Haptic_Management/"),
  ('HapticControlMgt','lvrPosInfo','ALL'):
    ("레버 위치 정보 유효성 확인(LvrModePosInfo≠오류)",
     "LvrModePosInfo=오류 → Haptic 피드백 비활성화\n참조: Haptic_Management/"),
  ('HapticControlMgt','SlaveAddress','ALL'):
    ("I2C 슬레이브 주소 유효성 확인; 통신 오류 감지",
     "I2C 통신 오류 → Haptic 명령 미전달; 진동 없음\n참조: CtCdd_IoHwAb_I2c.c"),
  ('HapticControlMgt','TransmitLength','ALL'):
    ("I2C 전송 길이 범위 검사",
     "길이 오류 → I2C 전송 실패; Haptic 비활성화\n참조: CtCdd_IoHwAb_I2c.c"),
}

# ── Countermeasure 데이터베이스 ───────────────────────────────────────────────
# RPN ≥ 100 또는 S ≥ 8 인 경우 권장 조치
CM_DB = {
    'SbcFlt':      "SBC 결함 시 ECU 리셋 또는 안전 모드 전환 절차 명확화; 하드웨어 SBC 이중화 검토",
    'TrmnlCtrlGrpStaBDCEV': "Dual-CAN 이중화 유지; 두 채널 동시 실패 시 경고등 점등 및 현재 위치 고정",
    'SActSig':     "E2E 파라미터 최적화(AlvCnt 윈도우 확대); 통신 오류 시 GearPosSta=마지막 유효값 유지 검토",
    'tmp_Position':"센서 이중화 구성(PosSnr1/PosSnr2 교차 검증); Hall 센서 교정 주기 점검",
    'RotateState': "NVM 무결성 검사(CRC) 추가; NVM 읽기 실패 시 기본값(Dial=0) 사용",
    'MotorActivation': "모터 활성화 조건 다중 게이팅 유지; 하드웨어 인터록 회로 추가 검토",
    'SysPwrSta':   "SBC 결함 발생 시 경고등 점등 및 서비스 요청 메시지 출력",
    'tmp_Position_STUCK': "모터 재시도 횟수 검토(현재 3회); Stuck 반복 시 DTC 우선순위 상향",
    'MotorStopSig': "모터 정지 후 운전자에게 경고 표시; 강제 정지 DTC 스냅샷 저장 항목 확인",
    'MainCANBusOFF': "CAN 버스 오프 복구 절차 점검; 복구 후 신호 재초기화 로직 확인",
    'HallSnrFltInfo': "Hall 센서 3채널 독립 감지 유지; 단일 채널 오류 시 제한 동작 모드 정의",
    'PosSnrRaw1':  "이중 센서(Raw1/Raw2) 교차 검증 로직 강화; 불일치 임계값 검토",
    'DriveSta':    "DriveSig 양쪽 모두 실패 시 모터 즉시 정지 확인; 서브 채널 복구 절차 정의",
}

def get_s(unit_key, var_base, fm):
    # Variable-specific lookup
    for key, val in VAR_S.items():
        if key.lower() in var_base.lower() or var_base.lower() in key.lower():
            if fm in val: return val[fm]
            if 'ALL' in val: return val['ALL']
    return UNIT_S.get(unit_key, 6)

def get_o(unit_key, var_base):
    for key, val in VAR_O.items():
        if key.lower() in var_base.lower() or var_base.lower() in key.lower():
            return val
    return UNIT_O.get(unit_key, 3)

def get_d(unit_key, var_base):
    for key, val in VAR_D.items():
        if key.lower() in var_base.lower() or var_base.lower() in key.lower():
            return val
    return UNIT_D.get(unit_key, 4)

def get_action(unit_raw, var_raw, fm):
    if not unit_raw or not var_raw:
        return '', ''
    unit = unit_raw.replace('\n','').strip()
    for pfx in ['CstAp_']:
        if unit.startswith(pfx):
            unit = unit[len(pfx):]
    var_base = var_raw.split('\n')[0].strip().split('(')[0].strip()
    var_base = var_base.split('.')[0].split('[')[0].strip()

    # Exact match
    for (uk, vk, fk), (p, d) in ACTION_DB.items():
        if uk == unit and vk == var_base and (fk == fm or fk == 'ALL'):
            return p, d
    # Partial var match
    for (uk, vk, fk), (p, d) in ACTION_DB.items():
        if uk == unit and (vk in var_base or var_base in vk) and (fk == fm or fk == 'ALL'):
            return p, d
    return '', ''

def get_cm(var_base, s_val, fm):
    for key, cm in CM_DB.items():
        if key.lower() in var_base.lower() or var_base.lower() in key.lower():
            return cm
    if s_val >= 8:
        return "안전 영향도 높음(S≥8): 해당 고장 모드에 대한 별도 안전 분석 및 추가 조치 검토 필요"
    return ""

# ══════════════════════════════════════════════════════════════════════════════
# Excel 작업
# ══════════════════════════════════════════════════════════════════════════════
print("Excel 열기...", flush=True)
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(FMEA_IN)
    ws = wb.Sheets('SW_FMEA')
    total_rows = ws.UsedRange.Rows.Count
    print(f"총 {total_rows}행", flush=True)

    # Bulk read
    data = ws.UsedRange.Value

    # 열 번호 (1-based)
    C_S    = 11   # K: Severity
    C_PREV = 12   # L: Preventive Action
    C_O    = 13   # M: Occurrence
    C_DET  = 14   # N: Detection Action
    C_D    = 16   # P: Detection rating
    C_RPN  = 17   # Q: RPN
    C_CM_Y = 18   # R: Is countermeasure required
    C_CM   = 19   # S: Countermeasure
    C_S2   = 20   # T: S (after)
    C_O2   = 21   # U: O (after)
    C_D2   = 22   # V: D (after)
    C_RPN2 = 23   # W: RPN (after)

    # Colors
    CLR_S     = 0xFCE4D6   # 연주황
    CLR_O     = 0xEBF3FB   # 연파랑
    CLR_D     = 0xEDF7EE   # 연녹색
    CLR_RPN_H = 0xFFC7CE   # 빨강(RPN높음)
    CLR_RPN_M = 0xFFEB9C   # 노랑(중간)
    CLR_RPN_L = 0xC6EFCE   # 녹색(낮음)
    CLR_PREV  = 0xD9EAD3
    CLR_DACT  = 0xCFE2F3
    CLR_CM    = 0xFFF2CC

    curr = {}
    stats = {'total':0,'prev_filled':0,'det_filled':0,'sod_filled':0,'cm_filled':0}

    print("행 처리 중...", flush=True)

    for r_idx in range(12, len(data)):
        row = data[r_idx]

        def clean(v):
            if v is None: return None
            s = str(v).replace('\xa0',' ').strip()
            return s if s else None

        no   = clean(row[0])
        unit = clean(row[1])
        var  = clean(row[3])
        fm   = clean(row[5])

        if no:   curr['no']   = no
        if unit: curr['unit'] = unit
        if var:  curr['var']  = var

        if fm not in FAILURE_MODES:
            continue

        stats['total'] += 1
        cur_unit = curr.get('unit','')
        cur_var  = curr.get('var','')
        if not cur_var:
            continue

        # Extract unit key
        unit_key = cur_unit.replace('\n','').strip()
        for pfx in ['CstAp_']:
            if unit_key.startswith(pfx):
                unit_key = unit_key[len(pfx):]

        # Extract base var
        var_base = cur_var.split('\n')[0].strip().split('(')[0].strip()
        var_base = var_base.split('.')[0].split('[')[0].strip()

        # S, O, D
        s_val = get_s(unit_key, var_base, fm)
        o_val = get_o(unit_key, var_base)
        d_val = get_d(unit_key, var_base)
        rpn   = s_val * o_val * d_val

        # Countermeasure needed?
        cm_needed = 'Y' if (rpn >= 100 or s_val >= 8) else 'N'
        cm_text   = get_cm(var_base, s_val, fm) if cm_needed == 'Y' else ''

        # After-countermeasure values (조치 후 O,D 감소)
        if cm_needed == 'Y':
            s2 = s_val
            o2 = max(1, o_val - 1)
            d2 = max(1, d_val - 1)
            rpn2 = s2 * o2 * d2
        else:
            s2, o2, d2, rpn2 = s_val, o_val, d_val, rpn

        # Preventive / Detection actions
        prev_act, det_act = get_action(cur_unit, cur_var, fm)

        # Generic fallback by unit
        if not prev_act:
            generic_prev = {
                'PwrMGT':           "전원 모니터링 로직 상시 동작; 임계값 기반 상태 제어",
                'ECUModeMgt':       "ECU 모드 상태 머신 기반 안전 상태 관리; CAN 타임아웃 감지",
                'CANMGT':           "CAN 수신 타임아웃 감지; Dual-CAN 이중화 또는 안전 기본값 적용",
                'MotorControlMgt':  "모터 활성화 사전 조건 다중 검증; 3회 재시도 후 정지 로직",
                'HapticControlMgt': "전원 상태 및 유효 신호 확인 후 Haptic 활성화",
                'IdtMgt':           "Indicator 전원 및 PWM 상태 모니터링; DTC 기록",
                'MoodControlMgt':   "PWM 출력 범위 제한; 전원 상태 기반 조건부 활성화",
                'DtcMgt':           "Dem_SetEventStatus 기반 DTC 기록; 스냅샷 데이터 저장",
                'PosMgt':           "위치 센서 범위 검사(100~900); 이중 센서 교차 검증",
                'MovingMgt':        "이동 센서 Active/Period 신호 모니터링; Delta 변화량 감시",
                'ButtonMgt':        "P버튼 이중 채널 교차 검증; Stuck 감지 로직",
            }
            prev_act = generic_prev.get(unit_key, "소프트웨어 내부 유효성 검사 로직 적용")

        if not det_act:
            generic_det = {
                'PwrMGT':           "전압/전류 임계값 초과 시 상태 플래그 설정; DTC 기록",
                'ECUModeMgt':       "CAN 타임아웃 플래그 감지 → 신호 OFF 강제 또는 DTC 설정",
                'CANMGT':           "CAN 타임아웃/BusOff 플래그 감지 → DTC 설정; 신호 무효화",
                'MotorControlMgt':  "위치 변화량(Delta) 모니터링; P2E00_01 DTC; 재시도 카운터",
                'HapticControlMgt': "I2C 통신 상태 확인; 전송 실패 감지",
                'IdtMgt':           "Indicator PWM 피드백 확인; C1182_96 DTC 기록",
                'MoodControlMgt':   "PWM 출력 상태 소프트웨어 모니터링",
                'DtcMgt':           "Dem_SetEventStatus(FAILED) → DTC 기록 및 스냅샷 저장",
                'PosMgt':           "센서 범위 외 → PosSnrFlt=ON → P2E00_01 DTC",
                'MovingMgt':        "센서 신호 이상 → PosSnrFlt=ON → P2E00_01 DTC",
                'ButtonMgt':        "채널 불일치/Stuck → C1181_96 DTC",
            }
            det_act = generic_det.get(unit_key, "소프트웨어 런타임 감지 로직; DTC 기록")

        # Write to Excel
        er = r_idx + 1  # 1-based row

        # S
        c = ws.Cells(er, C_S)
        c.Value = s_val
        c.Interior.Color = CLR_S
        c.HorizontalAlignment = -4108
        c.NumberFormat = "0"

        # Preventive
        c = ws.Cells(er, C_PREV)
        c.Value = prev_act
        c.Interior.Color = CLR_PREV
        c.WrapText = True

        # O
        c = ws.Cells(er, C_O)
        c.Value = o_val
        c.Interior.Color = CLR_O
        c.HorizontalAlignment = -4108
        c.NumberFormat = "0"

        # Detection Action
        c = ws.Cells(er, C_DET)
        c.Value = det_act
        c.Interior.Color = CLR_DACT
        c.WrapText = True

        # D
        c = ws.Cells(er, C_D)
        c.Value = d_val
        c.Interior.Color = CLR_D
        c.HorizontalAlignment = -4108
        c.NumberFormat = "0"

        # RPN
        c = ws.Cells(er, C_RPN)
        c.Value = rpn
        c.HorizontalAlignment = -4108
        c.Font.Bold = True
        c.NumberFormat = "0"
        if rpn >= 100:
            c.Interior.Color = CLR_RPN_H
            c.Font.Color = 0x9C0006
        elif rpn >= 50:
            c.Interior.Color = CLR_RPN_M
            c.Font.Color = 0x9C6500
        else:
            c.Interior.Color = CLR_RPN_L
            c.Font.Color = 0x276221

        # Countermeasure Required
        c = ws.Cells(er, C_CM_Y)
        c.Value = cm_needed
        c.HorizontalAlignment = -4108
        if cm_needed == 'Y':
            c.Interior.Color = 0xFFC7CE
            c.Font.Color = 0x9C0006
            c.Font.Bold = True

        # Countermeasure text
        if cm_text:
            c = ws.Cells(er, C_CM)
            c.Value = cm_text
            c.Interior.Color = CLR_CM
            c.WrapText = True

        # After-countermeasure
        ws.Cells(er, C_S2).Value = s2
        ws.Cells(er, C_O2).Value = o2
        ws.Cells(er, C_D2).Value = d2
        c2 = ws.Cells(er, C_RPN2)
        c2.Value = rpn2
        c2.Font.Bold = True
        if rpn2 >= 100:
            c2.Interior.Color = CLR_RPN_H
        elif rpn2 >= 50:
            c2.Interior.Color = CLR_RPN_M
        else:
            c2.Interior.Color = CLR_RPN_L

        stats['sod_filled'] += 1
        if prev_act: stats['prev_filled'] += 1
        if det_act:  stats['det_filled']  += 1
        if cm_text:  stats['cm_filled']   += 1

        if stats['total'] % 200 == 0:
            print(f"  처리 중: {stats['total']}행...", flush=True)

    print(f"\n완료 통계:", flush=True)
    for k, v in stats.items():
        print(f"  {k}: {v}", flush=True)

    # ── 저장 ──────────────────────────────────────────────────────────────────
    print(f"\n저장: {FMEA_OUT}", flush=True)
    wb.SaveAs(FMEA_OUT, FileFormat=51)
    wb.Close(False)
    excel.Quit()
    print("저장 완료!", flush=True)

except Exception as e:
    print(f"오류: {e}", flush=True)
    import traceback; traceback.print_exc()
    try: excel.Quit()
    except: pass
