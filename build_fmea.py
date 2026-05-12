"""
JG1 SBW Software FMEA 자동 생성
- 소스코드 분석 결과를 기반으로 Preventive/Detection Action 자동 기입
- 기존 FMEA 데이터 + 소스코드 대조 결과 통합
"""
import win32com.client, json
from collections import defaultdict

FMEA_IN  = r"E:\claude\FMEA\JG1_SBW-Software_FMEA_1.xlsx"
FMEA_OUT = r"E:\claude\FMEA\JG1_SBW-Software_FMEA_2_소스대조완료.xlsx"

# ══════════════════════════════════════════════════════════════════════════════
# 소스코드 분석 기반 Preventive / Detection Action 데이터베이스
# Key: (SW_Unit_prefix, variable_base, failure_mode)  → (preventive, detection, file_ref)
# failure_mode: 'ALL' = 모든 failure mode 공통
# ══════════════════════════════════════════════════════════════════════════════
ACTION_DB = {

    # ─────────────────────────────────────────────────────────────────────────
    # CstAp_PwrMGT
    # ─────────────────────────────────────────────────────────────────────────
    ('PwrMGT','BatVolt','MORE'): (
        "ADC 입력값 상한 클램프 (ADC_RANGE_LIMIT_VALUE=4095)",
        "배터리 전압 과전압 검출: BatVolt≥1910(ADC) → 200ms 후 BatOverSta=ON → U3003_A3 DTC 설정",
        "CtAp_VBatStaChk.c:143-175"),
    ('PwrMGT','BatVolt','LESS'): (
        "ADC 입력값 하한 클램프",
        "배터리 전압 저전압 검출: BatVolt≤962(ADC) → 200ms 후 BatUnderSta=ON → U3003_A2 DTC 설정",
        "CtAp_VBatStaChk.c:81-103"),
    ('PwrMGT','BatVolt','CORRUPT'): (
        "ADC 값 4095 초과 시 클램프 처리 (CtIoHwAb_IntfIn.c:228-240)",
        "ADC 범위 초과 감지: 4095 초과 → 클램프 후 상위 로직에서 과전압/저전압 판정",
        "CtIoHwAb_IntfIn.c:228"),
    ('PwrMGT','BatVolt','EARLY'): (
        "200ms 디바운스 필터 적용 (V_BAT_UNDER_TIME_SET)",
        "일시적 전압 변동 필터링: 연속 200ms 이상 조건 충족 시만 BatOverSta/BatUnderSta 설정",
        "CtAp_VBatStaChk.c:81-175"),
    ('PwrMGT','BatVolt','LATE'): (
        "200ms 디바운스 필터 적용",
        "연속 200ms 조건 유지 확인 후 상태 전환",
        "CtAp_VBatStaChk.c:81-175"),

    ('PwrMGT','IgnVolt','MORE'): (
        "ADC 값 4095 초과 시 클램프 처리",
        "IgnVolt≥762(ADC/7V) → 50ms 필터 후 HWIGN=ON 설정",
        "CtAp_IgnStaChk.c:69-94"),
    ('PwrMGT','IgnVolt','LESS'): (
        "ADC 클램프 처리",
        "IgnVolt≤405(ADC/4V) → HWIGN=OFF; TrmnlCtrlGrpStaBDCEV=OFF와 AND 조건으로 PowerOnSta=OFF",
        "CtAp_IgnStaChk.c:69-94"),
    ('PwrMGT','IgnVolt','CORRUPT'): (
        "ADC 범위 클램프 (0~4095)",
        "ADC 이상값 감지 → 클램프 후 임계값 기반 ON/OFF 판정",
        "CtIoHwAb_IntfIn.c:228"),

    ('PwrMGT','Ldo2OnVolt','MORE'): (
        "LDO2 결함 검사는 VBatStbSta=ON 및 ECUSta=WAKEUP 조건에서만 활성화",
        "Ldo2OnVolt 과전압 조건 없음 (저전압만 감지); 3.3V 정상 범위 이상은 별도 처리 없음",
        "CtAp_LdoStaChk.c:124"),
    ('PwrMGT','Ldo2OnVolt','LESS'): (
        "LDO2 검사 활성화 조건(VBatStbSta=ON, ECUSta=WAKEUP) 검증; 100ms 디바운스 필터",
        "Ldo2OnVolt≤2866(ADC/3.5V) → 100ms 후 Ldo2FltSta=ON → 시스템 전원 상태에 반영",
        "CtAp_LdoStaChk.c:142-204"),
    ('PwrMGT','Ldo2OnVolt','CORRUPT'): (
        "ADC 클램프 (0~4095)",
        "ADC 이상값 클램프 후 저전압 임계값 판정",
        "CtIoHwAb_IntfIn.c:228"),

    ('PwrMGT','SbcFlt','MORE'): (
        "SbcFlt 입력 반전 로직 (0=Fault, 1=Normal → 논리 반전 후 처리)",
        "100ms 디바운스 후 SbcFltSta=ON → SysPwrSta=POWER_OFF 강제 설정",
        "CtAp_LdoStaChk.c:211-229, CtAp_SysStaChk.c:61-78"),
    ('PwrMGT','SbcFlt','CORRUPT'): (
        "SbcFlt 반전 로직으로 이상값 처리",
        "SbcFlt 비정상 → SbcFltSta=ON → 시스템 전원 OFF 강제",
        "CtAp_LdoStaChk.c:211"),
    ('PwrMGT','TrmnlCtrlGrpStaBDCEV','CORRUPT'): (
        "Dual-CAN 이중화: Main CAN 수신 실패 시 Sub CAN 값 사용",
        "SMK_03_Timeout 감지: 양쪽(Main/Sub) 타임아웃 시 TrmnlCtrlGrpStaBDCEV=0 → 모터 비활성화",
        "CtAp_CGWSigChk.c:373-400"),
    ('PwrMGT','ECUSta','CORRUPT'): (
        "RTE 읽기 실패 시 ECUSta=EXTER_ECU_STANDBY(3) 기본값 사용",
        "RTE 반환값 검증 후 이상 시 안전 기본값 적용",
        "CtAp_LdoStaChk.c:81-85"),

    # ─────────────────────────────────────────────────────────────────────────
    # CstAp_ECUModeMgt
    # ─────────────────────────────────────────────────────────────────────────
    ('ECUModeMgt','BDC02MsgTo','ALL'): (
        "CAN 스택 레벨 수신 타임아웃 감지",
        "BDC_02 타임아웃 플래그 감지 → IgnSwStaFlag=OFF 강제 → 점화 신호 무효화",
        "CtAp_CGWSigChk.c:545-555"),
    ('ECUModeMgt','BDC05MsgTo','ALL'): (
        "CAN 스택 레벨 수신 타임아웃 감지",
        "BDC_05 타임아웃 → 테일램프/무드램프 신호 모두 OFF 강제",
        "CtAp_CGWSigChk.c:291-311"),
    ('ECUModeMgt','CLUMsgTo','ALL'): (
        "CAN 스택 레벨 수신 타임아웃 감지",
        "CLU_01 타임아웃 → AutoBrightSta=OFF 강제",
        "CtAp_CGWSigChk.c:313-331"),
    ('ECUModeMgt','PDC03MsgTo','ALL'): (
        "CAN 스택 레벨 수신 타임아웃 감지",
        "PDC_03 타임아웃 → DrvDrSwSta=OFF 강제 (안전 기본값)",
        "CtAp_CGWSigChk.c:338-348"),
    ('ECUModeMgt','DriveSta','ALL'): (
        "Dual-CAN 이중화: Main 실패 시 Sub CAN 값 폴백",
        "DriveSigMainTo/SubTo 플래그 감지; 양쪽 모두 타임아웃 시 DriveSigAllTo=ON → 모터 활성화 차단",
        "CtAp_DrvRdySigChk.c:78-119"),
    ('ECUModeMgt','DriveSigAllTo','ALL'): (
        "Main/Sub CAN 이중화 모니터링",
        "Main AND Sub 타임아웃 동시 발생 시 DriveSigAllTo=ON → 모터 활성화 조건 불충족",
        "CtAp_DrvRdySigChk.c:114-119"),
    ('ECUModeMgt','DriveSigMainTo','ALL'): (
        "Main CAN 수신 타임아웃 모니터링",
        "Main CAN 타임아웃 → Sub CAN으로 폴백; Sub도 실패 시 DriveSigAllTo=ON",
        "CtAp_DrvRdySigChk.c:97-121"),
    ('ECUModeMgt','DriveSigSubTo','ALL'): (
        "Sub CAN 수신 타임아웃 모니터링",
        "Sub CAN 타임아웃 → Main CAN 유지; 양쪽 모두 실패 시 DriveSigAllTo=ON",
        "CtAp_DrvRdySigChk.c:97-121"),
    ('ECUModeMgt','SysPwrSta','ALL'): (
        "SBC 결함 게이팅: SbcFltSta=ON 시 즉시 POWER_OFF 강제",
        "SysPwrSta = POWER_ON 조건: SbcFltSta=OFF AND PowerOnSta=ON; 조건 미충족 시 POWER_OFF",
        "CtAp_SysStaChk.c:61-78"),
    ('ECUModeMgt','TrmnlCtrlGrpStaBDCEV','ALL'): (
        "Dual-CAN 이중화 (SMK Main/Sub); SMK_03_Timeout 양쪽 감지",
        "양쪽 타임아웃 시 TrmnlCtrlGrpStaBDCEV=0 → U1065_8C DTC → 모터 활성화 차단",
        "CtAp_CGWSigChk.c:373-400"),
    ('ECUModeMgt','DiagSession','ALL'): (
        "DriveSta 조건 기반 진단 세션 전환 제어",
        "DriveSta=ON(주행 중) 시 쓰기 서비스 제한; DriveSta=OFF 시만 ECU 파라미터 변경 허용",
        "CtAp_EcuModeCntl.c:277-289"),

    # ─────────────────────────────────────────────────────────────────────────
    # CstAp_CANMGT
    # ─────────────────────────────────────────────────────────────────────────
    ('CANMGT','SActSig','MORE'): (
        "E2E CRC 검증 (500ms 디바운스): E2E_P_ERROR → CrcErrFlag",
        "P0A1D_83 DTC 설정; GearPosSta=NOT_DISPLAY; Sub CAN 폴백",
        "CtAp_ShiftActSigChk.c:188-193"),
    ('CANMGT','SActSig','LESS'): (
        "E2E Alive Counter 검증 (500ms 디바운스): E2E_P_REPEATED → AlvCntRepFlag",
        "P0A1D_82 DTC 설정; 신호 폐기; Sub CAN 폴백",
        "CtAp_ShiftActSigChk.c:174-179"),
    ('CANMGT','SActSig','CORRUPT'): (
        "E2E CRC + Alive Counter 이중 검증",
        "E2E 오류 → P0A1D_82/83 DTC; GearPosSta=NOT_DISPLAY; 기어 위치 표시 불가",
        "CtAp_ShiftActSigChk.c:165-202"),
    ('CANMGT','SActSig','OMISSION'): (
        "CAN 메시지 수신 타임아웃 감지",
        "SActSigTo=ON → P0A1D_8C DTC; MainSActMsgToFlag=ON → Sub CAN 폴백",
        "CtAp_ShiftActSigChk.c:299-312"),

    ('CANMGT','MainCANBusOFF','MORE'): (
        "CAN 버스 오프 카운터 (700ms 디바운스: BUSOFF_CHECK_TIME=70@10ms)",
        "U0028_88 DTC 설정; MainCANBusOFFSta=ON → Main CAN 모든 신호 폐기; Sub CAN 폴백",
        "CtAp_CANBusOffChk.c:70-77"),
    ('CANMGT','SubCANBusOFF','MORE'): (
        "CAN 버스 오프 카운터 (700ms 디바운스)",
        "U0028_88 DTC 설정; SubCANBusOFFSta=ON → Sub CAN 모든 신호 폐기; Main CAN 유지",
        "CtAp_CANBusOffChk.c:86-93"),

    ('CANMGT','Main_Vcu_E2E_Error_Sta','ALL'): (
        "E2E 프로파일 검증 (CRC + AlvCnt): 오류 코드 0x01~0x05 분류",
        "오류 코드별 플래그 설정 → P0A1D_82/83 DTC; 비정상 데이터 폐기",
        "CtAp_ShiftActSigChk.c:165-202"),

    ('CANMGT','SMK_TrmnlCtrlGrpStaBDCEV','ALL'): (
        "Dual-CAN 이중화: Main SMK 타임아웃 시 Sub SMK 값 사용",
        "양쪽 모두 타임아웃 시 TrmnlCtrlGrpStaBDCEV=0 → 모터 활성화 차단 (U1065_8C DTC)",
        "CtAp_CGWSigChk.c:373-400"),
    ('CANMGT','SMK_TrmnlCtrlStaBDC','ALL'): (
        "CAN 수신 타임아웃 감지; Dual-CAN 폴백",
        "SMK_03 타임아웃 감지 → 신호 무효화",
        "CtAp_CGWSigChk.c:373-402"),

    ('CANMGT','SubDrvRdySig','ALL'): (
        "Sub CAN 수신 타임아웃 모니터링; Main CAN DrvRdySig와 이중화",
        "SubDrvRdySigTo=ON → Sub 신호 폐기; DriveSigSubTo=ON; 양쪽 타임아웃 시 DriveSigAllTo=ON",
        "CtAp_DrvRdySigChk.c:56-119"),

    ('CANMGT','RotateState','ALL'): (
        "상태 머신 기반 유효값 (0 또는 1만 허용)",
        "유효 범위 외 값 → 상태 머신 진입 불가 → 모터 비활성화 유지",
        "CtAp_MotorControl.c:435,542"),
    ('CANMGT','PosSta','ALL'): (
        "PosSnrFlt 플래그 기반 위치 유효성 검사",
        "PosSnrFlt=ON → LvrPosSta=LVR_Flt(0x0F) → P2E00_01 DTC → CAN으로 FAULT 상태 전송",
        "CtAp_SBWSigSet.c:243-248"),
    ('CANMGT','PosStuck','ALL'): (
        "위치 센서 변화량(Delta) 모니터링 (500ms 동안 ±2 이내 시 stuck 판정)",
        "PosStuck=ON → LvrWrngMsg=LEVER_STUCK(0x12) → 경고 3회 후 모터 정지",
        "CtAp_SBWSigSet.c:341-343"),
    ('CANMGT','PButtonSta','ALL'): (
        "P버튼 이중 채널 모니터링 (P_SW1_Raw, P_SW2_Raw)",
        "PButtonFault=P_BUTTON_FAIL OR PButtonStuck=ON → LvrPSta=P_FAULT(0x03) → C1181_96 DTC",
        "CtAp_SBWSigSet.c:214-217"),
    ('CANMGT','IdtFltSta','ALL'): (
        "Indicator 회로 결함 플래그 모니터링",
        "IdtFltSta=ON → LvrIdtSta=INDI_FAULT → C1182_96 DTC 설정",
        "CtAp_SBWSigSet.c:198-203"),

    # ─────────────────────────────────────────────────────────────────────────
    # CstAp_MotorControlMgt
    # ─────────────────────────────────────────────────────────────────────────
    ('MotorControlMgt','tmp_Position','MORE'): (
        "위치 센서 유효 범위 검사: 100≤value≤900만 허용",
        "범위 초과 → Moving_Position 미갱신 → 500ms 후 Stuck 감지 → P2E00_01 DTC",
        "CtAp_MotorControl.c:231-233"),
    ('MotorControlMgt','tmp_Position','LESS'): (
        "위치 센서 유효 범위 검사: 100≤value≤900만 허용",
        "범위 미달 → Moving_Position 미갱신 → Stuck 감지 → P2E00_01 DTC",
        "CtAp_MotorControl.c:231-233"),
    ('MotorControlMgt','tmp_Position','CORRUPT'): (
        "범위 필터링 (100~900); SBWSigSet에서 추가 검증",
        "범위 외 → LvrModePosInfo=3(오류) → CAN으로 Not Available 전송; P2E00_01 DTC",
        "CtAp_SBWSigSet.c:153-158"),
    ('MotorControlMgt','tmp_Position','STUCK'): (
        "Delta 변화량 모니터링: ±2 이하 변화 500ms 지속 시 Stuck 판정",
        "TurnDialError/TurnSphereError 카운터; 3회 실패 → MotorStopSig=1 → P2E00_01 DTC",
        "CtAp_MotorControl.c:661-933"),

    ('MotorControlMgt','RotateState','ALL'): (
        "NVM 저장값; 상태 머신에서 0/1 범위만 사용",
        "유효 범위 외 → 상태 머신 진입 불가 → 모터 비활성 유지 (안전 상태)",
        "CtAp_MotorControl.c:435,542"),
    ('MotorControlMgt','SysPwrSta','ALL'): (
        "SBC 결함 및 전원 ON 조건 이중 게이팅",
        "SysPwrSta≠POWER_ON → 모터 활성화 조건 불충족 → 경고 클리어 및 안전 상태 유지",
        "CtAp_MotorControl.c:368-392"),
    ('MotorControlMgt','TrmnlCtrlGrpStaBDCEV','ALL'): (
        "Dual-CAN 이중화; BDCEV_READY(3) 또는 BDCEV_POWERON(1) 조건 확인",
        "TrmnlCtrlGrpStaBDCEV≠READY/POWERON → 상태 머신 진입 불가 → 모터 비활성화",
        "CtAp_MotorControl.c:402-407"),
    ('MotorControlMgt','DriveSta','ALL'): (
        "Dual-CAN 이중화; DriveSigAllTo=ON 시 모터 활성화 조건 불충족",
        "DriveSigAllTo=ON → 모터 활성화 조건 불충족; 모터 정지 상태 유지",
        "CtAp_MotorControl.c:414-432"),
    ('MotorControlMgt','DriveSigAllTo','ALL'): (
        "Main/Sub CAN 이중화 모니터링",
        "DriveSigAllTo=ON → 모터 활성화 조건 차단",
        "CtAp_DrvRdySigChk.c:114-119"),
    ('MotorControlMgt','MotorActivation','ALL'): (
        "100ms 워치독: Step_EN=1(비활성) 후 10 사이클(100ms) 이내 재활성화 조건 충족해야 함",
        "모터 활성화 타임아웃 → MotorActivation=OFF 자동 전환; Step_EN=1(하드웨어 비활성화)",
        "CtAp_MotorControl.c:266-282"),
    ('MotorControlMgt','MotorStopSig','ALL'): (
        "3회 재시도 후 영구 정지 명령 (TurnDialError/SphereError≥3)",
        "MotorStopSig=1 → MotorControlStop() 호출 → Motor_State=0, Step_EN=1, tmp_Speed=0",
        "CtAp_MotorControl.c:688,753,856"),
    ('MotorControlMgt','DrvDrSwSta','ALL'): (
        "운전석 도어 상태 확인 (모터 활성화 사전 조건)",
        "DrvDrSwSta≠CLOSED → 모드 변경 비활성화; 의도치 않은 레버 동작 방지",
        "CtAp_MotorControl.c:414-432"),
    ('MotorControlMgt','DrvStOccSta','ALL'): (
        "운전석 점유 상태 확인 (모터 활성화 사전 조건)",
        "DrvStOccSta≠SEATED → 모드 변경 비활성화; 10사이클 지연 후 SystemState=2 진입",
        "CtAp_MotorControl.c:414-432"),

    # ─────────────────────────────────────────────────────────────────────────
    # CstAp_PosMgt
    # ─────────────────────────────────────────────────────────────────────────
    ('PosMgt','PositionSensor_PosSnrRaw1','CORRUPT'): (
        "위치 센서 유효 범위 검사 (100~900)",
        "범위 외 → PosSnrFlt=ON → P2E00_01 DTC → LvrPosSta=LVR_Flt",
        "CtAp_SBWSigSet.c:153-158"),
    ('PosMgt','PositionSensor_PosSnrRaw2','CORRUPT'): (
        "위치 센서 유효 범위 검사 (100~900)",
        "범위 외 → PosSnrFlt=ON → P2E00_01 DTC → LvrPosSta=LVR_Flt",
        "CtAp_SBWSigSet.c:153-158"),
    ('PosMgt','PositionSensorInfo_PButtonFltSta','ALL'): (
        "P버튼 이중 채널 (P_SW1_Raw, P_SW2_Raw) 교차 검증",
        "불일치 → PButtonFaultSta=ON → C1181_96 DTC; LvrPSta=P_FAULT",
        "CtAp_SBWSigSet.c:214-217"),
    ('PosMgt','PositionSensorInfo_PButtonSta','ALL'): (
        "P버튼 양쪽 채널 일치 확인",
        "PButtonFault 감지 → C1181_96 DTC; P위치 진입 불가",
        "CtAp_SBWSigSet.c:214"),

    # ─────────────────────────────────────────────────────────────────────────
    # CstAp_MovingMgt
    # ─────────────────────────────────────────────────────────────────────────
    ('MovingMgt','PosSnr1','ALL'): (
        "Active/Period 시간 비율 기반 위치 계산; 센서 신호 유효성 확인",
        "센서 신호 이상 → PosSnrFlt=ON → P2E00_01 DTC; 위치 미갱신",
        "Position_Management/"),
    ('MovingMgt','PosSnr2','ALL'): (
        "Active/Period 시간 비율 기반 위치 계산",
        "센서 신호 이상 → PosSnrFlt=ON → P2E00_01 DTC",
        "Position_Management/"),
    ('MovingMgt','MovSnr1','ALL'): (
        "이동 센서 Active/Period 신호 모니터링",
        "MovSnr 신호 이상 → Stuck 감지 카운터 증가 → 3회 후 모터 정지",
        "CtAp_MotorControl.c:874-933"),
    ('MovingMgt','MovSnr2','ALL'): (
        "이동 센서 Active/Period 신호 모니터링",
        "MovSnr 신호 이상 → Stuck 감지 카운터 증가",
        "CtAp_MotorControl.c:874-933"),

    # ─────────────────────────────────────────────────────────────────────────
    # CstAp_ButtonMgt
    # ─────────────────────────────────────────────────────────────────────────
    ('ButtonMgt','P_SW1_Raw','ALL'): (
        "P버튼 이중 채널 중 채널1 모니터링; 채널2와 교차 검증",
        "P_SW1_Raw 이상 → PButtonFault 설정 → C1181_96 DTC; LvrPSta=P_FAULT",
        "Button_Management/"),
    ('ButtonMgt','P_SW2_Raw','ALL'): (
        "P버튼 이중 채널 중 채널2 모니터링; 채널1과 교차 검증",
        "P_SW2_Raw 이상 → PButtonFault 설정 → C1181_96 DTC; LvrPSta=P_FAULT",
        "Button_Management/"),

    # ─────────────────────────────────────────────────────────────────────────
    # CstAp_DtcMgt
    # ─────────────────────────────────────────────────────────────────────────
    ('DtcMgt','MainCANBusOffSta','ALL'): (
        "CAN BusOff 700ms 디바운스 카운터",
        "U0028_88 DTC 설정; Dem_SetEventStatus(FAILED) 호출",
        "CtAp_CANBusOffChk.c:70-77, CtDem_DTC_U0028_88.c"),
    ('DtcMgt','SubCanBusOffSta','ALL'): (
        "Sub CAN BusOff 700ms 디바운스 카운터",
        "U0028_88 DTC 설정",
        "CtAp_CANBusOffChk.c:86-93"),
    ('DtcMgt','HallSnrFltInfo','ALL'): (
        "Hall 센서 Alpha/Beta/VG 채널별 결함 분리 감지",
        "P2E00_01 DTC 설정; LvrPosSta=LVR_Flt → 모터 위치 제어 불가",
        "CtDem_DTC_P2E00_01_LeverHallSensorFault.c"),
    ('DtcMgt','BatVolt','MORE'): (
        "BatVolt ADC 상한 검사 + 200ms 필터",
        "U3003_A3 DTC 설정 (과전압); Dem_SetEventStatus(FAILED)",
        "CtDem_DTC_U3003_A3_BatteryVoltageHigh.c"),
    ('DtcMgt','BatVolt','LESS'): (
        "BatVolt ADC 하한 검사 + 200ms 필터",
        "U3003_A2 DTC 설정 (저전압); Dem_SetEventStatus(FAILED)",
        "CtDem_DTC_U3003_A2_BatteryVoltageLow.c"),
    ('DtcMgt','PButtonStuck','ALL'): (
        "P버튼 stuck 판정 로직 (일정 시간 이상 눌림 유지)",
        "C1181_96 DTC 설정; LvrPSta=P_FAULT",
        "CtDem_DTC_C1181_96_Park_Switch_Fault.c"),
    ('DtcMgt','PosSnrFlt','ALL'): (
        "위치 센서 범위 검사 및 신호 유효성 검증",
        "P2E00_01 DTC 설정; LvrPosSta=LVR_Flt",
        "CtDem_DTC_P2E00_01_LeverHallSensorFault.c"),
    ('DtcMgt','AlvCntFlt','ALL'): (
        "E2E Alive Counter 검증 500ms 디바운스",
        "P0A1D_82 DTC 설정; GearPosSta=NOT_DISPLAY",
        "CtDem_DTC_P0A1D_82_VCU_AC_Fault.c"),
    ('DtcMgt','CrcFltInfo','ALL'): (
        "E2E CRC 검증 500ms 디바운스",
        "P0A1D_83 DTC 설정; GearPosSta=NOT_DISPLAY",
        "CtDem_DTC_P0A1D_83_VCU_CRC_Fault.c"),
    ('DtcMgt','IdtFltSta','ALL'): (
        "Indicator 회로 결함 감지 (PWM 출력 확인)",
        "C1182_96 DTC 설정; LvrIdtSta=INDI_FAULT",
        "CtDem_DTC_C1182_96_IndicatorFault.c"),

    # ─────────────────────────────────────────────────────────────────────────
    # CstAp_IdtMgt
    # ─────────────────────────────────────────────────────────────────────────
    ('IdtMgt','BatStbSta','ALL'): (
        "배터리 안정화 상태 확인 (LDO2 검사 활성화 조건)",
        "BatStbSta=OFF 시 LDO2 결함 검사 비활성화 (오탐 방지)",
        "CtAp_LdoStaChk.c:124"),
    ('IdtMgt','Ldo2FltSta','ALL'): (
        "LDO2 전압 하한 100ms 디바운스 필터",
        "Ldo2FltSta=ON → 표시기(Indicator) 전원 공급 이상 → C1182_96 연동",
        "CtAp_LdoStaChk.c:142-204"),
    ('IdtMgt','IdtFltSta','ALL'): (
        "Indicator PWM 출력 결함 감지",
        "IdtFltSta=ON → C1182_96 DTC; Indicator 출력 OFF 강제",
        "CtDem_DTC_C1182_96_IndicatorFault.c"),
    ('IdtMgt','TotalAlvCntFlt','ALL'): (
        "Main/Sub CAN E2E AlvCnt 이중 검증",
        "Main AND Sub 모두 AlvCntFlt=ON → TotalAlvCntFlt=ON → P0A1D_82 DTC",
        "CtAp_ShiftActSigChk.c:352-358"),
    ('IdtMgt','TotalCrcFlt','ALL'): (
        "Main/Sub CAN E2E CRC 이중 검증",
        "Main AND Sub 모두 CrcFlt=ON → TotalCrcFlt=ON → P0A1D_83 DTC",
        "CtAp_ShiftActSigChk.c"),

    # ─────────────────────────────────────────────────────────────────────────
    # CstAp_MoodControlMgt
    # ─────────────────────────────────────────────────────────────────────────
    ('MoodControlMgt','BDC05MsgTimeout','ALL'): (
        "CAN BDC_05 메시지 타임아웃 감지",
        "타임아웃 → 무드램프 신호 OFF 강제; 안전 상태 유지",
        "CtAp_CGWSigChk.c:291-311"),
    ('MoodControlMgt','PwrOnModeSta','ALL'): (
        "SysPwrSta 기반 전원 상태 확인",
        "PwrOnModeSta=OFF → 무드램프 출력 비활성화",
        "MoodControl/"),
    ('MoodControlMgt','MoodLed_BPWM','ALL'): (
        "PWM 출력 범위 제한 (0~100%)",
        "Ldo2FltSta=ON 시 PWM 출력 0으로 강제 → 무드램프 소등",
        "MoodControl/"),

    # ─────────────────────────────────────────────────────────────────────────
    # CstAp_HapticControlMgt
    # ─────────────────────────────────────────────────────────────────────────
    ('HapticControlMgt','sysPwrSta','ALL'): (
        "SysPwrSta 조건 확인 후 Haptic 활성화",
        "SysPwrSta≠POWER_ON → Haptic 비활성화",
        "Haptic_Management/"),
    ('HapticControlMgt','gearPositionVcu','ALL'): (
        "유효한 기어 포지션 확인 (GearPosSta≠NOT_DISPLAY)",
        "GearPosSta=NOT_DISPLAY → Haptic 피드백 제공 불가 (안전 기본값: 진동 없음)",
        "Haptic_Management/"),
    ('HapticControlMgt','lvrPosInfo','ALL'): (
        "레버 위치 정보 유효성 확인",
        "LvrModePosInfo=오류 → Haptic 피드백 비활성화",
        "Haptic_Management/"),
}

# ──────────────────────────────────────────────────────────────────────────────
# Helper: lookup action
# ──────────────────────────────────────────────────────────────────────────────
def get_action(unit_raw, var_raw, fm):
    if not unit_raw or not var_raw:
        return '', '', ''

    # Extract unit suffix
    unit = unit_raw.replace('\n','').strip()
    unit_key = ''
    for prefix in ['CstAp_', 'Cst_']:
        if unit.startswith(prefix):
            unit_key = unit[len(prefix):]
            break

    # Extract base var
    var_base = var_raw.split('\n')[0].strip().split('(')[0].strip()
    var_base = var_base.split('.')[0].split('[')[0].strip()

    # Try specific FM first, then 'ALL'
    for key in [
        (unit_key, var_base, fm),
        (unit_key, var_base, 'ALL'),
    ]:
        if key in ACTION_DB:
            prev, det, fref = ACTION_DB[key]
            return prev, det, fref

    # Partial match on unit key
    for (uk, vk, fk), val in ACTION_DB.items():
        if uk == unit_key and (vk in var_base or var_base in vk) and (fk == fm or fk == 'ALL'):
            return val

    return '', '', ''


# ══════════════════════════════════════════════════════════════════════════════
# Excel 처리
# ══════════════════════════════════════════════════════════════════════════════
print("Excel 열기...", flush=True)
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

# Color constants (Excel uses BGR integer)
CLR_HEADER_BLUE  = 0xC55A11   # dark orange → actually use Blue for new cols
CLR_HEADER_GREEN = 0x375623
CLR_FILL_PREV    = 0xD9EAD3   # light green
CLR_FILL_DET     = 0xCFE2F3   # light blue
CLR_FILL_REF     = 0xFFF2CC   # light yellow
CLR_WHITE        = 0xFFFFFF
CLR_DARK_BLUE    = 0x2E75B6
CLR_DARK_GREEN   = 0x375623

try:
    wb = excel.Workbooks.Open(FMEA_IN)
    ws = wb.Sheets('SW_FMEA')
    total_rows = ws.UsedRange.Rows.Count
    print(f"SW_FMEA: {total_rows}행", flush=True)

    # Bulk read
    data = ws.UsedRange.Value

    # New column positions (after existing 27 cols)
    COL_PREV   = 28   # AB: Preventive Action (소스코드 근거)
    COL_DET    = 29   # AC: Detection Action (소스코드 근거)
    COL_REF    = 30   # AD: 소스 파일 참조
    COL_FOUND  = 31   # AE: 코드 발견 여부

    # Header row (row 12)
    hdr_data = [
        (COL_PREV,  "Preventive Action\n(소스코드 근거)", CLR_DARK_GREEN),
        (COL_DET,   "Detection Action\n(소스코드 근거)",  CLR_DARK_BLUE),
        (COL_REF,   "소스 파일 참조",                     0x7F6000),
        (COL_FOUND, "코드 발견",                           0x4472C4),
    ]
    for col, txt, clr in hdr_data:
        cell = ws.Cells(12, col)
        cell.Value = txt
        cell.Font.Bold = True
        cell.Interior.Color = clr
        cell.Font.Color = CLR_WHITE
        cell.HorizontalAlignment = -4108
        cell.VerticalAlignment = -4108
        cell.WrapText = True
    ws.Rows(12).RowHeight = 42

    # Load xref data
    with open(r"E:\claude\FMEA\xref_results.json", encoding='utf-8') as f:
        xref = json.load(f)

    not_found_vars = set()
    for unit, res in xref.items():
        for v in res['not_found']:
            not_found_vars.add(v)

    FAILURE_MODES = {'MORE','LESS','CORRUPT','OMISSION','COMMISSION','WRONG','EARLY','LATE','STUCK'}

    filled = 0
    curr = {}

    for r_idx in range(12, len(data)):
        row = data[r_idx]

        def clean(v):
            if v is None: return None
            s = str(v).replace('\xa0', ' ').strip()
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

        cur_unit = curr.get('unit', '')
        cur_var  = curr.get('var', '')

        prev_act, det_act, file_ref = get_action(cur_unit, cur_var, fm)

        excel_row = r_idx + 1

        if prev_act:
            c = ws.Cells(excel_row, COL_PREV)
            c.Value = prev_act
            c.Interior.Color = CLR_FILL_PREV
            c.WrapText = True

        if det_act:
            c = ws.Cells(excel_row, COL_DET)
            c.Value = det_act
            c.Interior.Color = CLR_FILL_DET
            c.WrapText = True

        if file_ref:
            c = ws.Cells(excel_row, COL_REF)
            c.Value = file_ref
            c.Interior.Color = CLR_FILL_REF
            c.WrapText = True

        # Code found flag
        var_base = cur_var.split('\n')[0].strip().split('(')[0].strip()
        var_base = var_base.split('.')[0].split('[')[0].strip()
        found = 'N' if var_base in not_found_vars else 'Y'
        fc = ws.Cells(excel_row, COL_FOUND)
        fc.Value = found
        if found == 'Y':
            fc.Interior.Color = 0xC6EFCE
            fc.Font.Color = 0x276221
        else:
            fc.Interior.Color = 0xFFC7CE
            fc.Font.Color = 0x9C0006

        if prev_act or det_act:
            filled += 1

    print(f"Preventive/Detection 채운 행: {filled}", flush=True)

    # Set column widths
    ws.Columns(COL_PREV).ColumnWidth = 50
    ws.Columns(COL_DET).ColumnWidth = 55
    ws.Columns(COL_REF).ColumnWidth = 35
    ws.Columns(COL_FOUND).ColumnWidth = 10

    # ── Summary sheet ──────────────────────────────────────────────────────────
    SUM_NAME = "코드대조_요약"
    for i in range(wb.Sheets.Count, 0, -1):
        if wb.Sheets(i).Name == SUM_NAME:
            wb.Sheets(i).Delete()

    ws_s = wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
    ws_s.Name = SUM_NAME

    # Header
    hdrs = ["SW Unit", "변수수", "코드발견", "미발견", "커버리지(%)",
            "Preventive입력", "Detection입력", "미발견 변수", "조치 사항"]
    for c, h in enumerate(hdrs, 1):
        cell = ws_s.Cells(1, c)
        cell.Value = h
        cell.Font.Bold = True
        cell.Interior.Color = 0x2E75B6
        cell.Font.Color = 0xFFFFFF

    with open(r"E:\claude\FMEA\fmea_data.json", encoding='utf-8') as f:
        records = json.load(f)

    # Count prev/det filled per unit
    from collections import Counter
    prev_count = Counter()
    det_count  = Counter()
    for r in records:
        unit = r['SW_Unit'] or ''
        var  = r['Variable'] or ''
        fm   = r['Failure_Mode'] or ''
        p, d, _ = get_action(unit, var, fm)
        if p: prev_count[unit.replace('\n','')] += 1
        if d: det_count[unit.replace('\n','')]  += 1

    row_idx = 2
    for unit, res in sorted(xref.items()):
        n_found  = len(res['found'])
        n_not    = len(res['not_found'])
        n_total  = n_found + n_not
        pct = int(100 * n_found / n_total) if n_total else 0
        unit_clean = unit.replace('\n','')

        not_str = ', '.join(sorted(res['not_found']))
        remarks = []
        for v in res['not_found']:
            if v in ['SactSig','SactSigTo','BDC02Timeout','BDC05Timeout',
                     'CLU01Timeout','SMK03Timeout','PositionSensorInfo_PButtonFltSta',
                     'PositionSensorInfo_PButtonSta','LeverWarning_Dial','LeverWarning_Sphere',
                     'RetryWarning_Dial','RetryWarning_Sphere','MotorFaultWarning']:
                remarks.append(f"{v}: 명명불일치")
            elif 'SnapShot' in v:
                remarks.append(f"{v}: DTC ID 불일치(0x020x→0xFD5x)")
            elif v in ['HallSnrFltInfo','HallSnrFltVal','SlaveAddress','TransmitLength']:
                remarks.append(f"{v}: Generated 코드")
            else:
                remarks.append(f"{v}: 미구현 확인 필요")

        pn = prev_count.get(unit_clean, 0)
        dn = det_count.get(unit_clean, 0)

        ws_s.Cells(row_idx, 1).Value = unit_clean
        ws_s.Cells(row_idx, 2).Value = n_total
        ws_s.Cells(row_idx, 3).Value = n_found
        ws_s.Cells(row_idx, 4).Value = n_not
        ws_s.Cells(row_idx, 5).Value = pct
        ws_s.Cells(row_idx, 6).Value = pn
        ws_s.Cells(row_idx, 7).Value = dn
        ws_s.Cells(row_idx, 8).Value = not_str
        ws_s.Cells(row_idx, 9).Value = '\n'.join(remarks[:5])

        # Color pct
        pc = ws_s.Cells(row_idx, 5)
        if pct == 100:
            pc.Interior.Color = 0xC6EFCE; pc.Font.Color = 0x276221
        elif pct >= 80:
            pc.Interior.Color = 0xFFEB9C; pc.Font.Color = 0x9C6500
        else:
            pc.Interior.Color = 0xFFC7CE; pc.Font.Color = 0x9C0006
        row_idx += 1

    # Total
    total_found = sum(len(r['found']) for r in xref.values())
    total_not   = sum(len(r['not_found']) for r in xref.values())
    ws_s.Cells(row_idx, 1).Value = "합계"
    ws_s.Cells(row_idx, 2).Value = total_found + total_not
    ws_s.Cells(row_idx, 3).Value = total_found
    ws_s.Cells(row_idx, 4).Value = total_not
    ws_s.Cells(row_idx, 5).Value = int(100*total_found/(total_found+total_not))
    ws_s.Cells(row_idx, 6).Value = sum(prev_count.values())
    ws_s.Cells(row_idx, 7).Value = sum(det_count.values())
    for c in range(1,8):
        ws_s.Cells(row_idx, c).Font.Bold = True
    ws_s.Columns.AutoFit()

    # FMEA 완성도 현황
    row_idx += 2
    ws_s.Cells(row_idx, 1).Value = "FMEA 완성도 현황"
    ws_s.Cells(row_idx, 1).Font.Bold = True
    status = [
        ("항목", "기존 입력", "이번 자동 입력", "전체", "완성도"),
        ("총 FMEA 항목",        1460, "-",  1460, "100%"),
        ("Preventive Action",     0, sum(prev_count.values()), 1460, f"{int(100*sum(prev_count.values())/1460)}%"),
        ("Detection Action",      0, sum(det_count.values()),  1460, f"{int(100*sum(det_count.values())/1460)}%"),
        ("코드 발견 여부",          0, 1460, 1460, "100%"),
        ("S (Severity)",          0, "수동 입력 필요", 1460, "0%"),
        ("O (Occurrence)",        0, "수동 입력 필요", 1460, "0%"),
        ("D (Detection)",         0, "수동 입력 필요", 1460, "0%"),
        ("RPN",                   0, "수동 입력 필요", 1460, "0%"),
    ]
    for i, row_data in enumerate(status):
        for j, val in enumerate(row_data):
            cell = ws_s.Cells(row_idx + 1 + i, j + 1)
            cell.Value = val
            if i == 0:
                cell.Font.Bold = True
                cell.Interior.Color = 0x404040
                cell.Font.Color = 0xFFFFFF
    ws_s.Columns.AutoFit()

    # Save As
    print(f"저장: {FMEA_OUT}", flush=True)
    wb.SaveAs(FMEA_OUT, FileFormat=51)  # xlOpenXMLWorkbook
    wb.Close(False)
    excel.Quit()
    print("완료!", flush=True)

except Exception as e:
    print(f"오류: {e}", flush=True)
    import traceback; traceback.print_exc()
    try: excel.Quit()
    except: pass
