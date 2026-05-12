"""
JG1 SBW Software FMEA - 전체 필드 채우기 (S/O/D/RPN 제외)
- _3_완성.xlsx 읽어서 빈 Preventive/Detection/Effect 필드 채움
- S, O, D, RPN은 모두 비움 (사람이 직접 입력)
- 출력: JG1_SBW-Software_FMEA_4_필드완성.xlsx
"""
import win32com.client as win32
import re

SRC  = r"E:\claude\FMEA\JG1_SBW-Software_FMEA_3_완성.xlsx"
DEST = r"E:\claude\FMEA\JG1_SBW-Software_FMEA_4_필드완성.xlsx"

# ══════════════════════════════════════════════════════════════════
# ACTION DB: (unit_keyword, var_keyword, fm) → (preventive, detection)
# fm='ALL' → 모든 실패모드 공통
# ══════════════════════════════════════════════════════════════════
A = {}

# ─────────────────────────────── CstAp_PwrMGT ─────────────────────
A[('PwrMGT','BatVolt','MORE')] = (
    "ADC 입력값 상한 클램프 (ADC_RANGE_LIMIT_VALUE=4095)",
    "배터리 전압 과전압 검출: BatVolt≥1910(ADC) → 200ms 후 BatOverSta=ON → U3003_A3 DTC 설정")
A[('PwrMGT','BatVolt','LESS')] = (
    "ADC 입력값 하한 클램프",
    "배터리 전압 저전압 검출: BatVolt≤962(ADC) → 200ms 후 BatUnderSta=ON → U3003_A2 DTC 설정")
A[('PwrMGT','BatVolt','CORRUPT')] = (
    "ADC 값 4095 초과 시 클램프 처리 (CtIoHwAb_IntfIn.c:228-240)",
    "ADC 범위 초과 감지: 4095 초과 → 클램프 후 과전압/저전압 판정")
A[('PwrMGT','BatVolt','EARLY')] = (
    "200ms 디바운스 필터 적용 (V_BAT_UNDER_TIME_SET)",
    "일시적 전압 변동 필터링: 연속 200ms 이상 조건 충족 시만 BatOverSta/BatUnderSta 설정")
A[('PwrMGT','BatVolt','LATE')] = (
    "200ms 디바운스 필터 적용",
    "연속 200ms 조건 유지 확인 후 상태 전환 (CtAp_VBatStaChk.c:81-175)")
A[('PwrMGT','IgnVolt','MORE')] = (
    "ADC 값 4095 초과 시 클램프 처리",
    "IgnVolt≥762(ADC/7V) → 50ms 필터 후 HWIGN=ON 설정 (CtAp_IgnStaChk.c:69-94)")
A[('PwrMGT','IgnVolt','LESS')] = (
    "ADC 클램프 처리",
    "IgnVolt≤405(ADC/4V) → HWIGN=OFF → TrmnlCtrlGrpStaBDCEV=OFF AND 조건으로 PowerOnSta=OFF")
A[('PwrMGT','IgnVolt','CORRUPT')] = (
    "ADC 범위 클램프 (0~4095)",
    "ADC 이상값 감지 → 클램프 후 임계값 기반 ON/OFF 판정 (CtIoHwAb_IntfIn.c:228)")
A[('PwrMGT','Ldo2OnVolt','MORE')] = (
    "LDO2 결함 검사는 VBatStbSta=ON 및 ECUSta=WAKEUP 조건에서만 활성화",
    "Ldo2OnVolt 과전압 조건 없음 (저전압만 감지); 정상 범위 이상은 별도 처리 없음")
A[('PwrMGT','Ldo2OnVolt','LESS')] = (
    "LDO2 검사 활성화 조건(VBatStbSta=ON, ECUSta=WAKEUP) 검증; 100ms 디바운스 필터",
    "Ldo2OnVolt≤2866(ADC/3.5V) → 100ms 후 Ldo2FltSta=ON → 시스템 전원 상태 반영 (CtAp_LdoStaChk.c:142-204)")
A[('PwrMGT','Ldo2OnVolt','CORRUPT')] = (
    "ADC 클램프 (0~4095)",
    "ADC 이상값 클램프 후 저전압 임계값 판정 (CtIoHwAb_IntfIn.c:228)")
A[('PwrMGT','SbcFlt','MORE')] = (
    "SbcFlt 입력 반전 로직 (0=Fault, 1=Normal → 논리 반전 후 처리)",
    "100ms 디바운스 후 SbcFltSta=ON → SysPwrSta=POWER_OFF 강제 (CtAp_LdoStaChk.c:211-229)")
A[('PwrMGT','SbcFlt','CORRUPT')] = (
    "SbcFlt 반전 로직으로 이상값 처리",
    "SbcFlt 비정상 → SbcFltSta=ON → 시스템 전원 OFF 강제 (CtAp_SysStaChk.c:61-78)")
A[('PwrMGT','TrmnlCtrlGrpStaBDCEV','ALL')] = (
    "Dual-CAN 이중화: Main CAN 타임아웃 시 Sub CAN 값 사용",
    "SMK_03_Timeout 감지: 양쪽 타임아웃 시 TrmnlCtrlGrpStaBDCEV=0 → 모터 활성화 차단 (CtAp_CGWSigChk.c:373-400)")
A[('PwrMGT','ECUSta','CORRUPT')] = (
    "RTE 읽기 실패 시 ECUSta=EXTER_ECU_STANDBY(3) 기본값 사용",
    "RTE 반환값 검증 후 이상 시 안전 기본값 적용 (CtAp_LdoStaChk.c:81-85)")

# ─────────────────────────────── CstAp_ECUModeMgt ─────────────────
A[('ECUModeMgt','BDC02MsgTo','ALL')] = (
    "CAN 스택 레벨 수신 타임아웃 감지",
    "BDC_02 타임아웃 플래그 감지 → IgnSwStaFlag=OFF 강제 → 점화 신호 무효화 (CtAp_CGWSigChk.c:545-555)")
A[('ECUModeMgt','BDC02Timeout','ALL')] = (
    "CAN 스택 레벨 BDC_02 메시지 타임아웃 플래그 입력",
    "BDC_02 타임아웃 → IgnSwStaFlag=OFF 강제 → ECU 모드 판단에서 점화 신호 무효 처리")
A[('ECUModeMgt','BDC05MsgTo','ALL')] = (
    "CAN 스택 레벨 수신 타임아웃 감지",
    "BDC_05 타임아웃 → 테일램프/무드램프 신호 모두 OFF 강제 (CtAp_CGWSigChk.c:291-311)")
A[('ECUModeMgt','BDC05Timeout','ALL')] = (
    "CAN 스택 레벨 BDC_05 메시지 타임아웃 플래그 입력",
    "BDC_05 타임아웃 → IntTailLmpOnReqFlag=OFF 강제; 무드램프 OFF 전환")
A[('ECUModeMgt','CLUMsgTo','ALL')] = (
    "CAN 스택 레벨 수신 타임아웃 감지",
    "CLU_01 타임아웃 → AutoBrightSta=OFF 강제 (CtAp_CGWSigChk.c:313-331)")
A[('ECUModeMgt','CLU01Timeout','ALL')] = (
    "CAN 스택 레벨 CLU_01 메시지 타임아웃 플래그 입력",
    "CLU_01 타임아웃 → AutoBrightSta=OFF → 인디케이터 밝기 자동조절 비활성화")
A[('ECUModeMgt','PDC03MsgTo','ALL')] = (
    "CAN 스택 레벨 수신 타임아웃 감지",
    "PDC_03 타임아웃 → DrvDrSwSta=OFF 강제 (안전 기본값) (CtAp_CGWSigChk.c:338-348)")
A[('ECUModeMgt','PDC01Timeout','ALL')] = (
    "CAN 스택 레벨 PDC_01 메시지 타임아웃 플래그 입력",
    "PDC_01 타임아웃 → DrvStOccSta=OFF 강제 → 모터 활성화 조건 불충족")
A[('ECUModeMgt','PDC03Timeout','ALL')] = (
    "CAN 스택 레벨 PDC_03 메시지 타임아웃 플래그 입력",
    "PDC_03 타임아웃 → DrvDrSwSta=OFF 강제 (안전 기본값)")
A[('ECUModeMgt','SMK03Timeout','ALL')] = (
    "Dual-CAN 이중화: Main SMK_03 타임아웃 시 Sub CAN 폴백",
    "Main AND Sub 양쪽 타임아웃 → TrmnlCtrlGrpStaBDCEV=0 강제 → U1065_8C DTC (CtAp_CGWSigChk.c:373-400)")
A[('ECUModeMgt','DeVCU01Timeout','ALL')] = (
    "CAN 스택 레벨 VCU_01 메시지 타임아웃 플래그 입력",
    "VCU_01 타임아웃 → P0A1D_8C DTC → 기어 변속 신호 무효화")
A[('ECUModeMgt','DeVCU04Timeout','ALL')] = (
    "CAN 스택 레벨 VCU_04 메시지 타임아웃 플래그 입력",
    "VCU_04 타임아웃 → P0A1D_8C DTC → 관련 신호 무효화")
A[('ECUModeMgt','ICC02Timeout','ALL')] = (
    "CAN 스택 레벨 ICC_02 메시지 타임아웃 플래그 입력",
    "ICC_02 타임아웃 → U01D0_8C DTC → 관련 ICU 신호 무효화")
A[('ECUModeMgt','SBCMDRV01Timeout','ALL')] = (
    "CAN 스택 레벨 SBCM_DRV_01 메시지 타임아웃 플래그 입력",
    "SBCM_DRV_01 타임아웃 → DrvDrSwStaSBCM=OFF 강제 → 모터 활성화 조건 불충족")
A[('ECUModeMgt','DriveSta','ALL')] = (
    "Dual-CAN 이중화: Main 실패 시 Sub CAN 값 폴백",
    "DriveSigMainTo/SubTo 플래그 감지; 양쪽 모두 타임아웃 시 DriveSigAllTo=ON → 모터 활성화 차단")
A[('ECUModeMgt','DriveSigAllTo','ALL')] = (
    "Main/Sub CAN 이중화 모니터링",
    "Main AND Sub 타임아웃 동시 발생 시 DriveSigAllTo=ON → 모터 활성화 조건 불충족")
A[('ECUModeMgt','DriveSigMainTo','ALL')] = (
    "Main CAN 수신 타임아웃 모니터링",
    "Main CAN 타임아웃 → Sub CAN으로 폴백; Sub도 실패 시 DriveSigAllTo=ON (CtAp_DrvRdySigChk.c:97-121)")
A[('ECUModeMgt','DriveSigSubTo','ALL')] = (
    "Sub CAN 수신 타임아웃 모니터링",
    "Sub CAN 타임아웃 → Main CAN 유지; 양쪽 모두 실패 시 DriveSigAllTo=ON")
A[('ECUModeMgt','SysPwrSta','ALL')] = (
    "SBC 결함 게이팅: SbcFltSta=ON 시 즉시 POWER_OFF 강제",
    "SysPwrSta = POWER_ON 조건: SbcFltSta=OFF AND PowerOnSta=ON; 조건 미충족 시 POWER_OFF (CtAp_SysStaChk.c:61-78)")
A[('ECUModeMgt','TrmnlCtrlGrpStaBDCEV','ALL')] = (
    "Dual-CAN 이중화 (SMK Main/Sub); SMK_03_Timeout 양쪽 감지",
    "양쪽 타임아웃 시 TrmnlCtrlGrpStaBDCEV=0 → U1065_8C DTC → 모터 활성화 차단")
A[('ECUModeMgt','DiagSession','ALL')] = (
    "DriveSta 조건 기반 진단 세션 전환 제어",
    "DriveSta=ON(주행 중) 시 쓰기 서비스 제한; DriveSta=OFF 시만 ECU 파라미터 변경 허용 (CtAp_EcuModeCntl.c:277-289)")
A[('ECUModeMgt','DrLockSta','ALL')] = (
    "도어 잠금 상태 CAN 신호 타임아웃 감지; Dual-CAN 폴백",
    "DrLockSta 수신 실패 → 기본값(잠금해제) 유지; 모터 활성화에 영향 없음")
A[('ECUModeMgt','IntTailLmpOnReqFlag','ALL')] = (
    "BDC_05 메시지 정상 수신 조건 게이팅",
    "BDC_05 타임아웃 시 IntTailLmpOnReqFlag=OFF 강제 → 인테리어 테일램프 요청 차단")
A[('ECUModeMgt','MainShiftActMsgTo','ALL')] = (
    "Main CAN SActSig 메시지 타임아웃 감지",
    "MainShiftActMsgTo=ON → Sub CAN 폴백 또는 P0A1D_8C DTC → 기어 위치 신호 무효화")
A[('ECUModeMgt','SubShiftActMsgTo','ALL')] = (
    "Sub CAN SActSig 메시지 타임아웃 감지",
    "SubShiftActMsgTo=ON → Main CAN 유지; 양쪽 모두 타임아웃 시 DriveSigAllTo 연동")

# ─────────────────────────────── CstAp_CANMGT ─────────────────────
A[('CANMGT','SActSig','MORE')] = (
    "E2E CRC 검증 (500ms 디바운스): E2E_P_ERROR → CrcErrFlag",
    "P0A1D_83 DTC 설정; GearPosSta=NOT_DISPLAY; Sub CAN 폴백 (CtAp_ShiftActSigChk.c:188-193)")
A[('CANMGT','SActSig','LESS')] = (
    "E2E Alive Counter 검증 (500ms 디바운스): E2E_P_REPEATED → AlvCntRepFlag",
    "P0A1D_82 DTC 설정; 신호 폐기; Sub CAN 폴백 (CtAp_ShiftActSigChk.c:174-179)")
A[('CANMGT','SActSig','CORRUPT')] = (
    "E2E CRC + Alive Counter 이중 검증",
    "E2E 오류 → P0A1D_82/83 DTC; GearPosSta=NOT_DISPLAY; 기어 위치 표시 불가")
A[('CANMGT','SActSig','OMISSION')] = (
    "CAN 메시지 수신 타임아웃 감지",
    "SActSigTo=ON → P0A1D_8C DTC; MainSActMsgToFlag=ON → Sub CAN 폴백")
A[('CANMGT','SactSig','MORE')] = A[('CANMGT','SActSig','MORE')]
A[('CANMGT','SactSig','LESS')] = A[('CANMGT','SActSig','LESS')]
A[('CANMGT','SactSig','CORRUPT')] = A[('CANMGT','SActSig','CORRUPT')]
A[('CANMGT','SactSig','OMISSION')] = A[('CANMGT','SActSig','OMISSION')]
A[('CANMGT','SubSActSig','ALL')] = (
    "Sub CAN E2E CRC + AlvCnt 이중 검증; Main CAN 폴백과 이중화",
    "Sub SActSig E2E 오류 → P0A1D_82/83 DTC; Main CAN 폴백 유지")
A[('CANMGT','SactSigTo','ALL')] = (
    "Main CAN SActSig 메시지 타임아웃 감지",
    "SactSigTo=ON → P0A1D_8C DTC; Sub CAN 폴백 또는 기어 신호 무효화")
A[('CANMGT','SubSActSigTo','ALL')] = (
    "Sub CAN SActSig 메시지 타임아웃 감지",
    "SubSActSigTo=ON → Main CAN 유지; 양쪽 모두 타임아웃 시 P0A1D_8C DTC")
A[('CANMGT','MainCANBusOFF','ALL')] = (
    "CAN 버스 오프 카운터 (700ms 디바운스: BUSOFF_CHECK_TIME=70@10ms)",
    "U0028_88 DTC 설정; MainCANBusOFFSta=ON → Main CAN 모든 신호 폐기; Sub CAN 폴백 (CtAp_CANBusOffChk.c:70-77)")
A[('CANMGT','SubCANBusOFF','ALL')] = (
    "CAN 버스 오프 카운터 (700ms 디바운스)",
    "U0028_88 DTC 설정; SubCANBusOFFSta=ON → Sub CAN 모든 신호 폐기; Main CAN 유지 (CtAp_CANBusOffChk.c:86-93)")
A[('CANMGT','Main_Vcu_E2E_Error_Sta','ALL')] = (
    "E2E 프로파일 검증 (CRC + AlvCnt): 오류 코드 0x01~0x05 분류",
    "오류 코드별 플래그 설정 → P0A1D_82/83 DTC; 비정상 데이터 폐기 (CtAp_ShiftActSigChk.c:165-202)")
A[('CANMGT','Main_Vcu_E2E_Error_Return','ALL')] = (
    "E2E 프로파일 반환값 유효성 검사",
    "E2E_P_ERROR 반환 시 CrcErrFlag=ON → P0A1D_83 DTC; 해당 데이터 폐기")
A[('CANMGT','SMK_TrmnlCtrlGrpStaBDCEV','ALL')] = (
    "Dual-CAN 이중화: Main SMK 타임아웃 시 Sub SMK 값 사용",
    "양쪽 모두 타임아웃 시 TrmnlCtrlGrpStaBDCEV=0 → 모터 활성화 차단 (U1065_8C DTC)")
A[('CANMGT','SMK_TrmnlCtrlStaBDC','ALL')] = (
    "CAN 수신 타임아웃 감지; Dual-CAN 폴백",
    "SMK_03 타임아웃 감지 → 신호 무효화 (CtAp_CGWSigChk.c:373-402)")
A[('CANMGT','SMK_TrmnlCtrlGrpStaBDC','ALL')] = (
    "CAN 수신 타임아웃 감지; Dual-CAN 폴백",
    "SMK_03 타임아웃 감지 → 신호 무효화")
A[('CANMGT','SMK_PwrOnModeSta','ALL')] = (
    "SMK_03 메시지 수신 조건 게이팅; Dual-CAN 이중화",
    "SMK_03 타임아웃 시 PwrOnModeSta=OFF 강제 → 인디케이터/무드램프 제어에 반영 (CtAp_CGWSigChk.c:378)")
A[('CANMGT','SubDrvRdySig','ALL')] = (
    "Sub CAN 수신 타임아웃 모니터링; Main CAN DrvRdySig와 이중화",
    "SubDrvRdySigTo=ON → Sub 신호 폐기; DriveSigSubTo=ON; 양쪽 타임아웃 시 DriveSigAllTo=ON (CtAp_DrvRdySigChk.c:56-119)")
A[('CANMGT','SubDrvRdySigTo','ALL')] = (
    "Sub CAN DrvRdySig 타임아웃 플래그",
    "SubDrvRdySigTo=ON → DriveSigSubTo=ON; Main CAN 유지; 양쪽 타임아웃 시 DriveSigAllTo=ON")
A[('CANMGT','RotateState','ALL')] = (
    "상태 머신 기반 유효값 (0 또는 1만 허용)",
    "유효 범위 외 값 → 상태 머신 진입 불가 → 모터 비활성화 유지 (CtAp_MotorControl.c:435,542)")
A[('CANMGT','PosSta','ALL')] = (
    "PosSnrFlt 플래그 기반 위치 유효성 검사",
    "PosSnrFlt=ON → LvrPosSta=LVR_Flt(0x0F) → P2E00_01 DTC → CAN으로 FAULT 상태 전송 (CtAp_SBWSigSet.c:243-248)")
A[('CANMGT','PosStuck','ALL')] = (
    "위치 센서 변화량(Delta) 모니터링 (500ms 동안 ±2 이내 시 stuck 판정)",
    "PosStuck=ON → LvrWrngMsg=LEVER_STUCK(0x12) → 경고 3회 후 모터 정지 (CtAp_SBWSigSet.c:341-343)")
A[('CANMGT','PButtonSta','ALL')] = (
    "P버튼 이중 채널 모니터링 (P_SW1_Raw, P_SW2_Raw)",
    "PButtonFault=P_BUTTON_FAIL OR PButtonStuck=ON → LvrPSta=P_FAULT(0x03) → C1181_96 DTC (CtAp_SBWSigSet.c:214-217)")
A[('CANMGT','PButtonFault','ALL')] = (
    "P버튼 이중 채널 교차 검증 (PSw1/PSw2 불일치 감지)",
    "불일치 → PButtonFault=P_BUTTON_FAIL → C1181_96 DTC; LvrPSta=P_FAULT → CAN 전송")
A[('CANMGT','PButtonStuck','ALL')] = (
    "P버튼 stuck 감지: 18000ms(3분) 이상 눌림 지속 시 stuck 판정 (CtApPButtonSet_PSwSutckTime=18000)",
    "PButtonStuck=ON → LvrPSta=P_FAULT → C1181_96 DTC (CtAp_SBWSigSet.c:214)")
A[('CANMGT','IdtFltSta','ALL')] = (
    "Indicator 회로 결함 플래그 모니터링",
    "IdtFltSta=ON → LvrIdtSta=INDI_FAULT → C1182_96 DTC 설정 (CtAp_SBWSigSet.c:198-203)")
A[('CANMGT','IdtSta','ALL')] = (
    "Indicator 상태 플래그 모니터링",
    "IdtSta 이상 → LvrIdtSta 반영 → C1182_96 DTC 설정 가능")
A[('CANMGT','BDC_02_Timeout','ALL')] = (
    "CAN 스택 레벨 BDC_02 메시지 타임아웃 감지 (CAN Management 레벨)",
    "BDC_02 타임아웃 → BDC02MsgTo=ON → IgnSwStaFlag=OFF 강제 (CtAp_CGWSigChk.c)")
A[('CANMGT','BDC_03_Timeout','ALL')] = (
    "CAN 스택 레벨 BDC_03 메시지 타임아웃 감지",
    "BDC_03 타임아웃 → 무드램프 관련 신호 OFF 강제")
A[('CANMGT','BDC_04_Timeout','ALL')] = (
    "CAN 스택 레벨 BDC_04 메시지 타임아웃 감지",
    "BDC_04 타임아웃 → 관련 신호 무효화; 안전 기본값 유지")
A[('CANMGT','BDC_05_Timeout','ALL')] = (
    "CAN 스택 레벨 BDC_05 메시지 타임아웃 감지",
    "BDC_05 타임아웃 → BDC05MsgTo=ON → 테일램프/무드램프 OFF 강제 (CtAp_CGWSigChk.c:291-311)")
A[('CANMGT','BDC_06_Timeout','ALL')] = (
    "CAN 스택 레벨 BDC_06 메시지 타임아웃 감지",
    "BDC_06 타임아웃 → 관련 신호 무효화; 안전 기본값 유지")
A[('CANMGT','CGWCLU_01_20ms_Timeout','ALL')] = (
    "CAN 스택 레벨 CLU_01 메시지 타임아웃 감지",
    "CLU_01 타임아웃 → CLUMsgTo=ON → AutoBrightSta=OFF 강제 (CtAp_CGWSigChk.c:313-331)")
A[('CANMGT','SMK_02_Timeout','ALL')] = (
    "CAN 스택 레벨 SMK_02 메시지 타임아웃 감지",
    "SMK_02 타임아웃 → 관련 신호 무효화")
A[('CANMGT','SMK_03_Timeout','ALL')] = (
    "Dual-CAN 이중화: Main SMK_03 타임아웃 시 Sub CAN 폴백",
    "Main AND Sub 양쪽 타임아웃 → TrmnlCtrlGrpStaBDCEV=0 → U1065_8C DTC (CtAp_CGWSigChk.c:373-400)")
A[('CANMGT','PDC_02_Timeout','ALL')] = (
    "CAN 스택 레벨 PDC_02 메시지 타임아웃 감지",
    "PDC_02 타임아웃 → 관련 신호 무효화")
A[('CANMGT','PDC_03_Timeout','ALL')] = (
    "CAN 스택 레벨 PDC_03 메시지 타임아웃 감지",
    "PDC_03 타임아웃 → PDC03MsgTo=ON → DrvDrSwSta=OFF 강제 (CtAp_CGWSigChk.c:338-348)")
A[('CANMGT','VCU04_Timeout','ALL')] = (
    "CAN 스택 레벨 VCU_04 메시지 타임아웃 감지",
    "VCU_04 타임아웃 → P0A1D_8C DTC → 기어 변속 신호 무효화")
A[('CANMGT','ICC_02_50ms','ALL')] = (
    "CAN 스택 레벨 ICC_02 메시지 타임아웃 감지",
    "ICC_02 타임아웃 → U01D0_8C DTC → 관련 ICU 신호 무효화")
A[('CANMGT','SBCM_DRV_01_Timeout','ALL')] = (
    "CAN 스택 레벨 SBCM_DRV_01 메시지 타임아웃 감지",
    "SBCM_DRV_01 타임아웃 → DrvDrSwStaSBCM=OFF 강제 → 모터 활성화 조건 불충족")
A[('CANMGT','DrvDrSwSta','ALL')] = (
    "PDC_03 메시지 정상 수신 조건 게이팅; Dual-CAN 폴백",
    "PDC_03 타임아웃 시 DrvDrSwSta=OFF 강제 → 모터 활성화 조건 불충족 (CtAp_CGWSigChk.c:338-348)")
A[('CANMGT','DrvDrSwSta_SBCM','ALL')] = (
    "SBCM_DRV_01 메시지 정상 수신 조건 게이팅",
    "SBCM_DRV_01 타임아웃 시 DrvDrSwStaSBCM=OFF 강제 → 모터 활성화 조건 불충족")
A[('CANMGT','DrvStOccSta','ALL')] = (
    "PDC_01 메시지 정상 수신 조건 게이팅",
    "PDC_01 타임아웃 시 DrvStOccSta=OFF 강제 → 모터 활성화 조건 불충족 (CtAp_CGWSigChk.c:404-415)")
A[('CANMGT','SysPwrSta','ALL')] = (
    "SBC 결함 및 전원 ON 조건 이중 게이팅",
    "SysPwrSta≠POWER_ON → 모터 활성화 조건 불충족 → 경고 클리어 및 안전 상태 유지")
A[('CANMGT','ECUSta','ALL')] = (
    "RTE 읽기 실패 시 안전 기본값 사용",
    "ECUSta 이상 → 안전 기본값(STANDBY) 유지 → 관련 기능 비활성화")
A[('CANMGT','MotorActivation','ALL')] = (
    "100ms 워치독: 비활성화 후 10 사이클 이내 재활성화 조건 충족 여부 확인",
    "조건 미충족 → MotorActivation=OFF 자동 전환; Step_EN=1(하드웨어 비활성화)")
A[('CANMGT','MotorStopSig','ALL')] = (
    "3회 재시도 후 영구 정지 명령 (TurnDialError/SphereError≥3)",
    "MotorStopSig=1 → MotorControlStop() 호출 → Motor_State=0, Step_EN=1, tmp_Speed=0")
A[('CANMGT','MotorFaultWarning','ALL')] = (
    "모터 고장 핀 상태 모니터링 (CtIoHwAb_IntfOut.c)",
    "MotorFaultPinChek 감지 → 모터 정지 명령 → 경고 메시지 CAN 전송")
A[('CANMGT','PosSnrFlt','ALL')] = (
    "위치 센서 유효 범위 검사 (100~900); 이중 채널 교차 검증",
    "PosSnrFlt=ON → LvrPosSta=LVR_Flt → P2E00_01 DTC → 모터 정지 (CtAp_SBWSigSet.c:153-158)")
A[('CANMGT','LeverWarning_Dial','ALL')] = (
    "모터 경고 상태 모니터링; LvrWrngMsg enum 값 기반 분류",
    "LvrWrngMsg 설정 → CAN으로 경고 코드 전송; 3회 경고 후 모터 정지 (CtAp_SBWSigSet.c)")
A[('CANMGT','LeverWarning_Sphere','ALL')] = A[('CANMGT','LeverWarning_Dial','ALL')]
A[('CANMGT','RetryWarning_Dial','ALL')] = (
    "모터 재시도 경고 모니터링; SBW_MotorWarn.RetryWarning 기반",
    "RetryWarning=ON → 경고 CAN 전송; 3회 재시도 실패 시 MotorStopSig=ON (CtAp_SBWSigSet.c)")
A[('CANMGT','RetryWarning_Sphere','ALL')] = A[('CANMGT','RetryWarning_Dial','ALL')]
A[('CANMGT','OverRideWarning','ALL')] = (
    "오버라이드 조건 감지 및 경고 상태 모니터링",
    "OverRideWarning=ON → CAN으로 경고 전송; 모터 동작 제한")
A[('CANMGT','Naccept','ALL')] = (
    "N단 접수 신호 유효성 검증",
    "Naccept 신호 이상 → N단 전환 거부; 안전 상태 유지")
A[('CANMGT','GetPSigSta','ALL')] = (
    "P위치 신호 상태 게이팅 (CAN 수신 정상 조건)",
    "GetPSigSta 이상 → P위치 전환 차단; 현재 기어 상태 유지")
A[('CANMGT','GetPSigTo','ALL')] = (
    "P위치 신호 타임아웃 감지",
    "GetPSigTo=ON → P위치 전환 차단; 타임아웃 DTC 설정")
A[('CANMGT','RKESig','ALL')] = (
    "RKE(원격 키 입력) 신호 유효성 검증",
    "RKESig 이상 → 원격 P위치 전환 차단; 안전 상태 유지")
A[('CANMGT','SMKSig','ALL')] = (
    "SMK 신호 유효성 검증; Dual-CAN 이중화",
    "SMKSig 이상 → TrmnlCtrlGrpStaBDCEV 무효화 → 모터 활성화 조건 불충족")
A[('CANMGT','SpecOption','ALL')] = (
    "차량 사양 옵션 코드 유효성 검증",
    "SpecOption 이상 → 기본 사양으로 동작; 사양별 기능 비활성화")
A[('CANMGT','ComCtrlMode','ALL')] = (
    "통신 제어 모드 상태 모니터링",
    "ComCtrlMode 이상 → 기본 통신 모드 유지; 진단 기능 제한")
A[('CANMGT','FaceDetectStat','ALL')] = (
    "얼굴 인식 상태 CAN 신호 유효성 검증",
    "FaceDetectStat 이상 → 기본값 유지; 관련 기능 비활성화")
A[('CANMGT','AvTailLmpSta','ALL')] = (
    "BDC 메시지 수신 조건 게이팅",
    "BDC 타임아웃 시 AvTailLmpSta=OFF 강제 → 테일램프 제어에 반영")
A[('CANMGT','InlTailLmpSta','ALL')] = (
    "BDC 메시지 수신 조건 게이팅",
    "BDC 타임아웃 시 InlTailLmpSta=OFF 강제 → 인테리어 테일램프 제어에 반영")
A[('CANMGT','IntTailLmpOnReq','ALL')] = (
    "BDC_05 메시지 정상 수신 조건 게이팅",
    "BDC_05 타임아웃 시 IntTailLmpOnReq=OFF 강제")
A[('CANMGT','AutoLtSnsrNightSta','ALL')] = (
    "CLU 메시지 수신 조건 게이팅",
    "CLU 타임아웃 시 AutoLtSnsrNightSta=OFF 강제 → 야간 모드 판단에 반영")
A[('CANMGT','BtnIllumiAlwaysOnSta','ALL')] = (
    "USM 메시지 수신 조건 게이팅",
    "USM 타임아웃 시 BtnIllumiAlwaysOnSta=OFF 강제 → 버튼 조명 항상 ON 기능 비활성화")
A[('CANMGT','USM_IllAlwaysOnwithPSTNSta','ALL')] = (
    "USM 메시지 수신 조건 게이팅",
    "USM 타임아웃 시 기본값 유지 → P위치 연동 조명 기능 비활성화")
A[('CANMGT','UtilModeActStaSig','ALL')] = (
    "Utility Mode 활성화 신호 유효성 검증",
    "UtilModeActStaSig 이상 → Utility Mode 비활성화 유지 (안전 기본값)")
A[('CANMGT','USM06Msg','ALL')] = (
    "USM_06 메시지 수신 조건 게이팅",
    "USM_06 타임아웃 시 관련 신호 기본값 유지")
A[('CANMGT','StrtStpBtnSw1Sta','ALL')] = (
    "SSB(Start/Stop Button) 신호 유효성 검증",
    "StrtStpBtnSw1Sta 이상 → P위치 전환 조건 불충족 → 안전 상태 유지")
A[('CANMGT','SSB_StrtStpBtnSw2Sta','ALL')] = (
    "SSB(Start/Stop Button) 이중 채널 검증",
    "SSB2 신호 이상 → P위치 전환 조건 불충족 → 안전 상태 유지")
A[('CANMGT','Ign1InSta','ALL')] = (
    "점화 신호 상태 CAN 게이팅; BDC_02 타임아웃 감지",
    "BDC_02 타임아웃 시 IgnSwStaFlag=OFF 강제 → 점화 기반 기능 비활성화")
A[('CANMGT','AsstDrSwSta','ALL')] = (
    "조수석 도어 상태 CAN 신호 게이팅",
    "CAN 타임아웃 시 AsstDrSwSta=OFF 기본값 유지")
A[('CANMGT','RrLftDrSwSta','ALL')] = (
    "뒷좌석 왼쪽 도어 상태 CAN 신호 게이팅",
    "CAN 타임아웃 시 RrLftDrSwSta=OFF 기본값 유지")
A[('CANMGT','RrRtDrSwSta','ALL')] = (
    "뒷좌석 오른쪽 도어 상태 CAN 신호 게이팅",
    "CAN 타임아웃 시 RrRtDrSwSta=OFF 기본값 유지")
A[('CANMGT','Trunk_IBU_OpnDiagSetSta','ALL')] = (
    "트렁크 진단 설정 상태 CAN 신호 게이팅",
    "CAN 타임아웃 시 트렁크 진단 기능 비활성화")
A[('CANMGT','Trunk_TrnkTlgtReleaseRly','ALL')] = (
    "트렁크 릴레이 제어 신호 유효성 검증",
    "CAN 타임아웃 시 트렁크 릴레이 명령 차단")
A[('CANMGT','MoodLamp_MdLmpFadeSta','ALL')] = (
    "BDC 메시지 수신 조건 게이팅",
    "BDC 타임아웃 시 MdLmpFadeSta=OFF 강제 → 무드램프 페이드 기능 비활성화")
A[('CANMGT','MoodLamp_SlvBrgtnsVal','ALL')] = (
    "BDC 메시지 수신 조건 게이팅; 범위 클램프 (0~250)",
    "BDC 타임아웃 시 SlvBrgtnsVal=0 강제 → 무드램프 밝기 OFF")
A[('CANMGT','MoodLamp_SlvColor_X','ALL')] = (
    "BDC 메시지 수신 조건 게이팅",
    "BDC 타임아웃 시 무드램프 색상 기본값 유지")
A[('CANMGT','MoodLamp_SlvColor_Y','ALL')] = A[('CANMGT','MoodLamp_SlvColor_X','ALL')]
A[('CANMGT','MoodLamp_SlvFadeInTimetVal','ALL')] = (
    "BDC 메시지 수신 조건 게이팅; 0 입력 시 1로 클램프",
    "BDC 타임아웃 시 FadeInTimeVal=1 강제 (최소값) → 즉시 페이드인 (CtAp_MoodControl.c:195-200)")
A[('CANMGT','MoodLamp_SlvFadeOutTimetVal','ALL')] = (
    "BDC 메시지 수신 조건 게이팅; 0 입력 시 1로 클램프",
    "BDC 타임아웃 시 FadeOutTimeVal=1 강제 (최소값) (CtAp_MoodControl.c:204-209)")
A[('CANMGT','MainCanSigSet_LvrMsg','ALL')] = (
    "Main CAN 레버 메시지 전송 전 유효성 검증",
    "LvrMsg 오류 시 NOT_DISPLAY 상태로 CAN 전송 → 클러스터 표시 불가")
A[('CANMGT','SubCanSigSet_LvrMsg','ALL')] = (
    "Sub CAN 레버 메시지 전송 전 유효성 검증",
    "SubLvrMsg 오류 시 NOT_DISPLAY 상태로 CAN 전송")
A[('CANMGT','tmpPos1','ALL')] = (
    "위치 센서 유효 범위 검사 (100~900); 이중 채널 검증",
    "범위 외 → PosSnrFlt=ON → P2E00_01 DTC → 모터 정지 (CtAp_SBWSigSet.c:121-158)")

# ─────────────────────────────── CstAp_MotorControlMgt ────────────
A[('MotorControlMgt','tmp_Position','MORE')] = (
    "위치 센서 유효 범위 검사: 100≤value≤900만 허용",
    "범위 초과 → Moving_Position 미갱신 → 500ms 후 Stuck 감지 → P2E00_01 DTC (CtAp_MotorControl.c:231-233)")
A[('MotorControlMgt','tmp_Position','LESS')] = (
    "위치 센서 유효 범위 검사: 100≤value≤900만 허용",
    "범위 미달 → Moving_Position 미갱신 → Stuck 감지 → P2E00_01 DTC")
A[('MotorControlMgt','tmp_Position','CORRUPT')] = (
    "범위 필터링 (100~900); SBWSigSet에서 추가 검증",
    "범위 외 → LvrModePosInfo=3(오류) → CAN으로 Not Available 전송; P2E00_01 DTC (CtAp_SBWSigSet.c:153-158)")
A[('MotorControlMgt','tmp_Position','STUCK')] = (
    "Delta 변화량 모니터링: ±2 이하 변화 500ms 지속 시 Stuck 판정",
    "TurnDialError/TurnSphereError 카운터; 3회 실패 → MotorStopSig=1 → P2E00_01 DTC (CtAp_MotorControl.c:661-933)")
A[('MotorControlMgt','RotateState','ALL')] = (
    "NVM 저장값; 상태 머신에서 0/1 범위만 사용",
    "유효 범위 외 → 상태 머신 진입 불가 → 모터 비활성 유지 (안전 상태)")
A[('MotorControlMgt','SysPwrSta','ALL')] = (
    "SBC 결함 및 전원 ON 조건 이중 게이팅",
    "SysPwrSta≠POWER_ON → 모터 활성화 조건 불충족 → 경고 클리어 및 안전 상태 유지")
A[('MotorControlMgt','TrmnlCtrlGrpStaBDCEV','ALL')] = (
    "Dual-CAN 이중화; BDCEV_READY(3) 또는 BDCEV_POWERON(1) 조건 확인",
    "TrmnlCtrlGrpStaBDCEV≠READY/POWERON → 상태 머신 진입 불가 → 모터 비활성화 (CtAp_MotorControl.c:402-407)")
A[('MotorControlMgt','DriveSta','ALL')] = (
    "Dual-CAN 이중화; DriveSigAllTo=ON 시 모터 활성화 조건 불충족",
    "DriveSigAllTo=ON → 모터 활성화 조건 불충족; 모터 정지 상태 유지")
A[('MotorControlMgt','DriveSigAllTo','ALL')] = (
    "Main/Sub CAN 이중화 모니터링",
    "DriveSigAllTo=ON → 모터 활성화 조건 차단 (CtAp_DrvRdySigChk.c:114-119)")
A[('MotorControlMgt','DriveSigMainTo','ALL')] = (
    "Main CAN 수신 타임아웃 모니터링",
    "Main CAN 타임아웃 → Sub CAN으로 폴백; Sub도 실패 시 DriveSigAllTo=ON")
A[('MotorControlMgt','DriveSigSubTo','ALL')] = (
    "Sub CAN 수신 타임아웃 모니터링",
    "Sub CAN 타임아웃 → Main CAN 유지; 양쪽 모두 실패 시 DriveSigAllTo=ON")
A[('MotorControlMgt','MotorActivation','ALL')] = (
    "100ms 워치독: Step_EN=1(비활성) 후 10 사이클(100ms) 이내 재활성화 조건 충족해야 함",
    "모터 활성화 타임아웃 → MotorActivation=OFF 자동 전환; Step_EN=1(하드웨어 비활성화)")
A[('MotorControlMgt','MotorStopSig','ALL')] = (
    "3회 재시도 후 영구 정지 명령 (TurnDialError/SphereError≥3)",
    "MotorStopSig=1 → MotorControlStop() 호출 → Motor_State=0, Step_EN=1, tmp_Speed=0 (CtAp_MotorControl.c:688,753,856)")
A[('MotorControlMgt','DrvDrSwSta','ALL')] = (
    "운전석 도어 상태 확인 (모터 활성화 사전 조건); PDC/SBCM 이중화",
    "DrvDrSwSta≠CLOSED → 모드 변경 비활성화; 의도치 않은 레버 동작 방지 (CtAp_MotorControl.c:414-432)")
A[('MotorControlMgt','DrvDrSwStaSBCM','ALL')] = (
    "SBCM_DRV_01 기반 운전석 도어 상태 이중화 채널",
    "DrvDrSwStaSBCM 이상 → PDC 채널 우선 사용; 양쪽 불일치 시 안전 기본값(CLOSED) 적용")
A[('MotorControlMgt','DrvStOccSta','ALL')] = (
    "운전석 점유 상태 확인 (모터 활성화 사전 조건)",
    "DrvStOccSta≠SEATED → 모드 변경 비활성화; 10사이클 지연 후 SystemState=2 진입 (CtAp_MotorControl.c:414-432)")
A[('MotorControlMgt','FaceDetectStat','ALL')] = (
    "얼굴 인식 상태 CAN 신호 유효성 검증",
    "FaceDetectStat 이상 → 기본값 유지; 모터 활성화에 직접 영향 없음")
A[('MotorControlMgt','PwrOnModeSta','ALL')] = (
    "SMK PwrOnModeSta 정상 수신 조건 게이팅",
    "PwrOnModeSta≠READY(2) → 모터 활성화 조건 불충족 → 안전 상태 유지 (CtAp_MotorControl.c:368-392)")
A[('MotorControlMgt','Channel','ALL')] = (
    "SPI 채널 번호 유효성 검증 (0 또는 유효 채널만 허용)",
    "채널 번호 이상 → SPI 통신 실패 → 모터 제어 명령 전달 불가 → 안전 상태")
A[('MotorControlMgt','Motor_Speed','ALL')] = (
    "모터 속도 파라미터 유효 범위 검사; NVM 저장값 검증",
    "Motor_Speed 이상 → 기본 속도값 사용 → 모터 정상 속도로 동작")
A[('MotorControlMgt','STEPMOTOR_SEQ_Write','ALL')] = (
    "스텝 모터 시퀀스 출력 유효성 검증",
    "시퀀스 오류 시 모터 스텝 오동작 → Stuck 감지 → P2E00_01 DTC")
A[('MotorControlMgt','Step_DIR','ALL')] = (
    "방향 신호 유효성 (0 또는 1만 허용)",
    "Step_DIR 이상 → 모터 방향 오동작 → 목표 위치 도달 불가 → Stuck 감지")
A[('MotorControlMgt','Step_EN','ALL')] = (
    "활성화 신호 이중 게이팅 (SysPwrSta + DriveSta 조건)",
    "Step_EN=1(비활성) 유지 → 모터 하드웨어 비활성화 → 안전 상태")
A[('MotorControlMgt','SrcDataBufferPtr','ALL')] = (
    "SPI 데이터 버퍼 포인터 유효성 검증",
    "포인터 이상 → SPI 통신 실패 → 모터 제어 명령 전달 불가")
A[('MotorControlMgt','DesDataBufferPtr','ALL')] = (
    "SPI 데이터 버퍼 포인터 유효성 검증",
    "포인터 이상 → SPI 수신 데이터 처리 불가 → 모터 상태 확인 불가")
A[('MotorControlMgt','Length','ALL')] = (
    "SPI 전송 길이 유효성 검증 (고정값)",
    "Length 이상 → SPI 통신 실패 → 모터 제어 명령 전달 불가")

# ─────────────────────────────── CstAp_PosMgt ─────────────────────
A[('PosMgt','PositionSensor_PosSnrRaw1','ALL')] = (
    "위치 센서 유효 범위 검사 (100~900); Snr1/Snr2 이중 채널 교차 검증",
    "범위 외 → PosSnrFlt=ON → P2E00_01 DTC → LvrPosSta=LVR_Flt (CtAp_SBWSigSet.c:153-158)")
A[('PosMgt','PositionSensor_PosSnrRaw2','ALL')] = (
    "위치 센서 유효 범위 검사 (100~900); Snr1/Snr2 이중 채널 교차 검증",
    "범위 외 → PosSnrFlt=ON → P2E00_01 DTC → LvrPosSta=LVR_Flt")
A[('PosMgt','PositionSensorInfo_PButtonFltSta','ALL')] = (
    "P버튼 이중 채널 (P_SW1_Raw, P_SW2_Raw) 교차 검증",
    "불일치 → PButtonFaultSta=ON → C1181_96 DTC; LvrPSta=P_FAULT (CtAp_SBWSigSet.c:214-217)")
A[('PosMgt','PositionSensorInfo_PButtonSta','ALL')] = (
    "P버튼 양쪽 채널 일치 확인",
    "PButtonFault 감지 → C1181_96 DTC; P위치 진입 불가 (CtAp_SBWSigSet.c:214)")

# ─────────────────────────────── CstAp_MovingMgt ──────────────────
_MOV_PREV = "PWM 입력 캡처 신호 유효 범위 검증; 모터 상태 머신에서 입력값 필터링"
_MOV_DET  = "범위 외 값 → Moving_Position 미갱신 → 500ms Stuck 판정 → P2E00_01 DTC (CtAp_MotorControl.c)"
for _v in ['MovSnr1.ActiveTime','MovSnr1.PeriodTime','MovSnr2.ActiveTime','MovSnr2.PeriodTime',
           'PosSnr1.ActiveTime','PosSnr1.PeriodTime','PosSnr2.ActiveTime','PosSnr2.PeriodTime']:
    A[('MovingMgt', _v, 'ALL')] = (_MOV_PREV, _MOV_DET)
# 변수명에 점이 있으므로 dot 없는 버전도 등록
for _v in ['MovSnr1','MovSnr2','PosSnr1','PosSnr2']:
    A[('MovingMgt', _v, 'ALL')] = (_MOV_PREV, _MOV_DET)

# ─────────────────────────────── CstAp_IdtMgt ─────────────────────
_IDT_PWM_PREV = "PWM 듀티 범위 검증 (0~100%); 전압 모니터링 회로 이중화"
_IDT_PWM_DET  = "CtAp_IdtFltChk: MntrVolt 임계값 초과 → IdtFltSta=ON → C1182_96 DTC; PWM 출력 OFF 강제"
for _g in ['DIdtCntl','NIdtCntl','PIdtCntl','RIdtCntl']:
    A[('IdtMgt', _g, 'ALL')] = (_IDT_PWM_PREV, _IDT_PWM_DET)

_VOLT_PREV = "인디케이터 PWM 출력 후 전압 피드백 모니터링 (ADC 기반)"
_VOLT_DET  = "전압 임계값 초과/미달 → IdtFltSta=ON → C1182_96 DTC → PWM 출력 OFF (CtAp_IdtFltChk.c)"
for _v in ['DMntrVolt','NMntrVolt','PMntrVolt','RMntrVolt',
           'DMntrVolt2','NMntrVolt12','NMntrVolt2','PMntrVolt12','PMntrVolt2','RMntrVolt2']:
    A[('IdtMgt', _v, 'ALL')] = (_VOLT_PREV, _VOLT_DET)

A[('IdtMgt','GearPosSta','ALL')] = (
    "E2E 검증된 SActSig 기반 기어 위치 판단; NOT_DISPLAY 상태 게이팅",
    "GearPosSta=NOT_DISPLAY → 인디케이터 표시 불가 → 모든 인디케이터 OFF 또는 오류 표시")
A[('IdtMgt','AutoBrightSta','ALL')] = (
    "CLU_01 메시지 정상 수신 조건 게이팅",
    "CLU_01 타임아웃 시 AutoBrightSta=OFF → 수동 밝기 모드로 전환")
A[('IdtMgt','BltDimLvl','ALL')] = (
    "BLT 밝기 레벨 유효 범위 검증; DimLvlSet 함수 내 범위 클램프",
    "범위 외 → 기본 밝기 레벨 적용 → 인디케이터 정상 표시 유지")
A[('IdtMgt','HltDimLvl','ALL')] = (
    "HLT 밝기 레벨 유효 범위 검증; DimLvlSet 함수 내 범위 클램프",
    "범위 외 → 기본 밝기 레벨 적용 → 인디케이터 정상 표시 유지")
A[('IdtMgt','Indi_tmp_Pos','ALL')] = (
    "MovingSnr 입력 기반 위치값 유효 범위 검사",
    "위치값 이상 → 인디케이터 표시 위치 오류 → 현재 위치 유지")
A[('IdtMgt','ECUSta','ALL')] = (
    "ECU 상태 기반 인디케이터 제어 조건 게이팅",
    "ECUSta≠WAKEUP → 인디케이터 OFF; ECU 상태 이상 시 인디케이터 비활성화")
A[('IdtMgt','Ldo2FltSta','ALL')] = (
    "LDO2 결함 상태 모니터링; CtAp_LdoStaChk에서 감지",
    "Ldo2FltSta=ON → 인디케이터 전원 불안정 → IdtCntlCdtSta=ALLOFF → 전체 인디케이터 OFF")
A[('IdtMgt','MainCANBusOffSta','ALL')] = (
    "Main CAN BusOff 상태 모니터링",
    "MainCANBusOffSta=ON → CAN 신호 무효화 → 인디케이터 기본값 표시")
A[('IdtMgt','MotorStopSig','ALL')] = (
    "모터 정지 신호 기반 인디케이터 상태 제어",
    "MotorStopSig=ON → 인디케이터 오류 표시 → 운전자 경고")
A[('IdtMgt','SysPwrSta','ALL')] = (
    "시스템 전원 상태 게이팅 (POWER_ON 조건)",
    "SysPwrSta≠POWER_ON → 인디케이터 제어 비활성화")
A[('IdtMgt','BatStbSta','ALL')] = (
    "배터리 안정화 상태 모니터링",
    "BatStbSta=OFF → 인디케이터 전원 불안정 → 인디케이터 비활성화")
A[('IdtMgt','CanBatNormalSta.DebounceBatNorSta','ALL')] = (
    "CAN 배터리 정상 상태 디바운스 플래그 모니터링",
    "배터리 비정상 → 인디케이터 전원 조건 불충족 → 인디케이터 OFF")
A[('IdtMgt','CanBatNormalSta.ImmediateBatNorSta','ALL')] = (
    "CAN 배터리 즉시 정상 상태 플래그 모니터링",
    "배터리 이상 즉시 감지 → 인디케이터 비활성화")
A[('IdtMgt','TotalAlvCntFlt','ALL')] = (
    "E2E AliveCounter 오류 통합 플래그 모니터링",
    "TotalAlvCntFlt=ON → 관련 신호 무효화 → 인디케이터 기본값 표시")
A[('IdtMgt','TotalCrcFlt','ALL')] = (
    "E2E CRC 오류 통합 플래그 모니터링",
    "TotalCrcFlt=ON → 관련 신호 무효화 → 인디케이터 기본값 표시")
A[('IdtMgt','PwrOnModeSta','ALL')] = (
    "SMK PwrOnModeSta 정상 수신 조건 게이팅",
    "PwrOnModeSta≠READY(2) → 인디케이터 표시 제한 → 이전 상태 유지")
A[('IdtMgt','CLU01MsgTo','ALL')] = (
    "CLU_01 메시지 타임아웃 감지",
    "CLU_01 타임아웃 → AutoBrightSta=OFF → 수동 밝기 모드 전환")
A[('IdtMgt','BDC05MsgTo','ALL')] = (
    "BDC_05 메시지 타임아웃 감지",
    "BDC_05 타임아웃 → 관련 신호 OFF 강제 → 인디케이터 기본 상태 유지")
A[('IdtMgt','PowerOn12','ALL')] = (
    "12V 전원 ON 상태 모니터링",
    "PowerOn12=OFF → 인디케이터 BLT/HLT 비활성화")
A[('IdtMgt','PowerOn2','ALL')] = (
    "2V 전원 ON 상태 모니터링",
    "PowerOn2=OFF → 인디케이터 전원 불안정 → 인디케이터 비활성화")

# ─────────────────────────────── CstAp_ButtonMgt ──────────────────
A[('ButtonMgt','P_SW1_Raw','ALL')] = (
    "P버튼 ADC 유효 범위 검사 (InterSwOnMin_1~InterSwOnMax_1: 2356~2604, InterSwOffMin_1~InterSwOffMax_1: 804~889)",
    "범위 외 → CONTACT_FAULT(2) 판정; 3 사이클 디바운스 후 오류 카운터 증가 → C1181_96 DTC (CtAp_PButtonSet.c)")
A[('ButtonMgt','P_SW2_Raw','ALL')] = (
    "P버튼 ADC 유효 범위 검사 (InterSwOnMin_2~InterSwOnMax_2: 3156~3488, InterSwOffMin_2~InterSwOffMax_2: 1608~1778)",
    "범위 외 → CONTACT_FAULT(2) 판정; PSw1/PSw2 불일치 시 MechaErrCnt 증가 → PButtonFault=ON → C1181_96 DTC")

# ─────────────────────────────── CstAp_HapticControlMgt ───────────
A[('HapticControlMgt','lvrPosInfo','ALL')] = (
    "기어 레버 위치 정보 유효 범위 검증 (PosSta 기반)",
    "lvrPosInfo 이상 → 햅틱 피드백 패턴 매핑 실패 → 기본 패턴(OFF) 적용")
A[('HapticControlMgt','gearPositionVcu','ALL')] = (
    "VCU 기어 위치 신호 E2E 검증 기반 유효성 확인",
    "gearPositionVcu 이상 → 햅틱 패턴 매핑 실패 → 기본 패턴(OFF) 적용")
A[('HapticControlMgt','sysPwrSta','ALL')] = (
    "시스템 전원 상태 게이팅 (POWER_ON 조건 확인)",
    "sysPwrSta≠POWER_ON → 햅틱 I2C 통신 비활성화 → 햅틱 피드백 OFF")
A[('HapticControlMgt','DataBufferPtr','ALL')] = (
    "I2C 전송 데이터 버퍼 포인터 유효성 검증",
    "포인터 이상 → I2C 통신 실패 → 햅틱 피드백 명령 전달 불가 (CtAp_HapticControl.c)")
A[('HapticControlMgt','SlaveAddress','ALL')] = (
    "I2C 슬레이브 주소 유효성 검증 (Generated RTE 설정값)",
    "주소 불일치 → I2C ACK 없음 → 통신 실패 → 햅틱 피드백 동작 불가")
A[('HapticControlMgt','TransmitLength','ALL')] = (
    "I2C 전송 길이 유효성 검증 (고정 프로토콜 길이)",
    "Length 이상 → I2C 통신 실패 또는 데이터 오염 → 햅틱 피드백 동작 불가")

# ─────────────────────────────── CstAp_MoodControlMgt ─────────────
A[('MoodControlMgt','BDC05MsgTimeout','ALL')] = (
    "BDC_05 메시지 타임아웃 감지; CAN Management 레벨 타임아웃 플래그",
    "BDC_05 타임아웃 → 무드램프 PWM 출력 OFF 강제; BDC03 타임아웃과 OR 조건 (CtAp_MoodControl.c:241)")
A[('MoodControlMgt','MdLmpFadeSta','ALL')] = (
    "BDC 메시지 정상 수신 조건 게이팅; 0 또는 1/2 유효값 검증",
    "BDC 타임아웃 시 MdLmpFadeSta=OFF 강제 → 페이드 기능 비활성화 → 즉시 ON/OFF 전환")
A[('MoodControlMgt','MoodLed_BPWM','ALL')] = (
    "PWM 값 유효 범위 검증 (0~255); BDC 타임아웃 시 0 강제",
    "BDC 타임아웃 또는 PwrOnModeSta≠READY → BPWM=0 → 파란색 LED OFF")
A[('MoodControlMgt','MoodLed_GPWM','ALL')] = (
    "PWM 값 유효 범위 검증 (0~255); BDC 타임아웃 시 0 강제",
    "BDC 타임아웃 또는 PwrOnModeSta≠READY → GPWM=0 → 초록색 LED OFF")
A[('MoodControlMgt','MoodLed_RPWM','ALL')] = (
    "PWM 값 유효 범위 검증 (0~255); BDC 타임아웃 시 0 강제",
    "BDC 타임아웃 또는 PwrOnModeSta≠READY → RPWM=0 → 빨간색 LED OFF")
A[('MoodControlMgt','PwrOnModeSta','ALL')] = (
    "SMK PwrOnModeSta 정상 수신 조건 게이팅",
    "PwrOnModeSta≠READY(2) 또는 ==CRANKING(3) → 무드램프 PWM 출력 OFF 강제 (CtAp_MoodControl.c:241)")
A[('MoodControlMgt','SlvBrgtnsVal','ALL')] = (
    "BDC 메시지 정상 수신 조건 게이팅; 범위 클램프 (0~250)",
    "SlvBrgtnsVal=0 또는 ≥251 → BrgtStep=0 (최소 밝기); 범위 외 → 기본 밝기 적용 (CtAp_MoodControl.c:160-166)")
A[('MoodControlMgt','SlvFadInTimetVal','ALL')] = (
    "BDC 메시지 정상 수신 조건 게이팅; 0 입력 시 1로 클램프",
    "0 입력 → FadeInTimeVal=1 강제 (최소값 적용) → 즉시 페이드인 처리 (CtAp_MoodControl.c:195-200)")
A[('MoodControlMgt','SlvFadOutTimetVal','ALL')] = (
    "BDC 메시지 정상 수신 조건 게이팅; 0 입력 시 1로 클램프",
    "0 입력 → FadeOutTimeVal=1 강제 (최소값 적용) (CtAp_MoodControl.c:204-209)")
A[('MoodControlMgt','SlvXVal','ALL')] = (
    "BDC 메시지 정상 수신 조건 게이팅; 색상 좌표 유효 범위 검증",
    "BDC 타임아웃 시 색상 좌표 기본값 유지 → 색상 변경 불가")
A[('MoodControlMgt','SlvYVal','ALL')] = A[('MoodControlMgt','SlvXVal','ALL')]
A[('MoodControlMgt','UtilMode','ALL')] = (
    "Utility Mode 활성화 조건 검증",
    "UtilMode 이상 → Utility Mode 비활성화 유지 → 무드램프 안전 상태 유지")
A[('MoodControlMgt','ValGearSlctDis','ALL')] = (
    "기어 선택 표시 유효성 검증 (미구현 가능성 있음 - 코드 확인 필요)",
    "ValGearSlctDis 이상 → 기어 선택 표시 비활성화 → 기본값 표시")

# ─────────────────────────────── CstAp_DtcMgt ─────────────────────
# SnapShot 변수들 (DTC 발생 시 시스템 상태 캡처)
_SS_PREV = "DTC 발생 시 시스템 상태 자동 스냅샷 캡처 (CtAp_SnapShotSet.c:UpdateSnapShot 호출)"
_SS_DET  = "DTC 저장 시 스냅샷 자동 기록; 진단 도구(OBD)를 통해 사후 분석 가능"
for _ss in ['SnapShot0200','SnapShot0202','SnapShot0203','SnapShot0204','SnapShot0205',
            'SnapShot0206','SnapShot0207','SnapShot0208','SnapShot0209','SnapShot020A',
            'SnapShot020B','SnapShot020D','SnapShot020E']:
    A[('DtcMgt', _ss, 'ALL')] = (_SS_PREV, _SS_DET)

# AlvCnt / CRC 관련
A[('DtcMgt','AlvCntFlt','ALL')] = (
    "E2E AliveCounter 오류 누적 카운터 기반 DTC 설정 조건 게이팅",
    "AlvCntFlt=ON → P0A1D_82 DTC 설정; GearPosSta=NOT_DISPLAY")
A[('DtcMgt','AlvCntRep','ALL')] = (
    "E2E AliveCounter 반복 오류 감지",
    "AlvCntRep=ON → P0A1D_82 DTC 설정 조건 충족")
A[('DtcMgt','AlvCntDiff','ALL')] = (
    "E2E AliveCounter 불일치 감지",
    "AlvCntDiff=ON → P0A1D_82 DTC 설정 조건 충족")
A[('DtcMgt','SubAlvCntFlt','ALL')] = (
    "Sub CAN E2E AliveCounter 오류 누적 카운터",
    "SubAlvCntFlt=ON → Sub CAN 신호 무효화; Main CAN 폴백")
A[('DtcMgt','SubAlvCntRep','ALL')] = (
    "Sub CAN E2E AliveCounter 반복 오류 감지",
    "SubAlvCntRep=ON → Sub CAN 신호 무효화")
A[('DtcMgt','SubAlvCntDiff','ALL')] = (
    "Sub CAN E2E AliveCounter 불일치 감지",
    "SubAlvCntDiff=ON → Sub CAN 신호 무효화")
A[('DtcMgt','CrcFltInfo.CrcFit','ALL')] = (
    "E2E CRC 오류 플래그 기반 DTC 설정 조건 게이팅",
    "CrcFit=ON → P0A1D_83 DTC 설정; GearPosSta=NOT_DISPLAY")
A[('DtcMgt','CrcFltInfo.SubCrcFit','ALL')] = (
    "Sub CAN E2E CRC 오류 플래그",
    "SubCrcFit=ON → Sub CAN 신호 무효화; P0A1D_83 DTC")
A[('DtcMgt','HallSnrFltInfo.Alpha','ALL')] = (
    "Hall 센서 Alpha 채널 결함 검사 (Generated BSW/DEM 설정)",
    "Alpha 채널 결함 → HallSnrFltInfo.TotalFlt=ON → P2E00_01 DTC 설정")
A[('DtcMgt','HallSnrFltInfo.Beta','ALL')] = (
    "Hall 센서 Beta 채널 결함 검사",
    "Beta 채널 결함 → HallSnrFltInfo.TotalFlt=ON → P2E00_01 DTC 설정")
A[('DtcMgt','HallSnrFltInfo.TotalFlt','ALL')] = (
    "Hall 센서 전체 결함 통합 플래그",
    "TotalFlt=ON → P2E00_01 DTC 설정 → 모터 정지")
A[('DtcMgt','HallSnrFltInfo.VG','ALL')] = (
    "Hall 센서 VG 채널 결함 검사",
    "VG 채널 결함 → P2E00_01 DTC 설정 조건 충족")
A[('DtcMgt','HallSnrFltVal.AlphaVal','ALL')] = (
    "Hall 센서 Alpha 채널 실측값 범위 검사",
    "AlphaVal 이상 → Hall 센서 결함으로 판정 → P2E00_01 DTC")
A[('DtcMgt','HallSnrFltVal.BetaVal','ALL')] = (
    "Hall 센서 Beta 채널 실측값 범위 검사",
    "BetaVal 이상 → Hall 센서 결함으로 판정 → P2E00_01 DTC")

# 전원/배터리 관련 DtcMgt 변수
A[('DtcMgt','BatVolt','ALL')] = (
    "ADC 기반 배터리 전압 모니터링; DTC 진입 조건 검사용",
    "BatVolt 이상값 → U3003_A2/A3 DTC 조건 판단에 사용 (CtAp_VBatStaChk.c)")
A[('DtcMgt','BatOverSta','ALL')] = (
    "200ms 디바운스 후 과전압 판정 플래그",
    "BatOverSta=ON → U3003_A3 DTC 설정 (CtAp_VBatStaChk.c)")
A[('DtcMgt','BatUnderSta','ALL')] = (
    "200ms 디바운스 후 저전압 판정 플래그",
    "BatUnderSta=ON → U3003_A2 DTC 설정")
A[('DtcMgt','BatStbSta','ALL')] = (
    "배터리 안정화 상태 모니터링; LDO2/SBC 검사 활성화 조건",
    "BatStbSta=OFF → LDO2/SBC 결함 검사 비활성화; DTC 조건 판단 보류")
A[('DtcMgt','IgnVolt','ALL')] = (
    "ADC 기반 점화 전압 모니터링",
    "IgnVolt 이상 → PowerOnSta 판단에 영향 → ECU 모드 전환")
A[('DtcMgt','Ldo2FltSta','ALL')] = (
    "LDO2 저전압 결함 플래그",
    "Ldo2FltSta=ON → 전원 불안정 → 시스템 동작 제한")
A[('DtcMgt','Ldo2OnVolt','ALL')] = (
    "LDO2 출력 전압 ADC 모니터링",
    "Ldo2OnVolt≤2866 → 100ms 후 Ldo2FltSta=ON")
A[('DtcMgt','SbcFlt','ALL')] = (
    "SBC 결함 핀 모니터링",
    "SbcFlt=FAULT → SbcFltSta=ON → SysPwrSta=POWER_OFF 강제")
A[('DtcMgt','SysPwrSta','ALL')] = (
    "시스템 전원 상태 DTC 진입 조건 게이팅",
    "SysPwrSta≠POWER_ON → DTC 인에이블 조건 불충족 → DTC 설정 보류")
A[('DtcMgt','ECUSta','ALL')] = (
    "ECU 상태 DTC 진입 조건 게이팅",
    "ECUSta≠WAKEUP → DTC 인에이블 조건 불충족")
A[('DtcMgt','DriveSta','ALL')] = (
    "주행 준비 상태 DTC 진입 조건 게이팅",
    "DriveSta=OFF → 일부 DTC 인에이블 조건 불충족")
A[('DtcMgt','DriveSigAllTo','ALL')] = (
    "Main/Sub CAN 이중화 타임아웃 통합 플래그",
    "DriveSigAllTo=ON → P0A1D_8C DTC 설정 조건 충족")
A[('DtcMgt','DriveSigMainTo','ALL')] = (
    "Main CAN 타임아웃 플래그",
    "DriveSigMainTo=ON → Sub CAN 폴백; DriveSigAllTo 판단에 사용")
A[('DtcMgt','DriveSigSubTo','ALL')] = (
    "Sub CAN 타임아웃 플래그",
    "DriveSigSubTo=ON → Main CAN 유지; DriveSigAllTo 판단에 사용")
A[('DtcMgt','MainCANBusOffSta','ALL')] = (
    "Main CAN BusOff 상태 DTC 조건",
    "MainCANBusOffSta=ON → U0028_88 DTC 설정 (CtAp_CANBusOffChk.c)")
A[('DtcMgt','SubCanBusOffSta','ALL')] = (
    "Sub CAN BusOff 상태 DTC 조건",
    "SubCanBusOffSta=ON → U0028_88 DTC 설정 (Sub CAN)")
A[('DtcMgt','MainCanRxSta','ALL')] = (
    "Main CAN 수신 상태 모니터링",
    "MainCanRxSta 이상 → CAN 수신 오류 DTC 설정 조건 충족")
A[('DtcMgt','MainCanTxSta','ALL')] = (
    "Main CAN 송신 상태 모니터링",
    "MainCanTxSta 이상 → CAN 송신 오류 DTC 설정 조건 충족")
A[('DtcMgt','SubCanRxSta','ALL')] = (
    "Sub CAN 수신 상태 모니터링",
    "SubCanRxSta 이상 → Sub CAN 수신 오류 DTC 설정 조건 충족")
A[('DtcMgt','SubCanTxSta','ALL')] = (
    "Sub CAN 송신 상태 모니터링",
    "SubCanTxSta 이상 → Sub CAN 송신 오류 DTC 설정 조건 충족")
A[('DtcMgt','MainShiftActMsgTo','ALL')] = (
    "Main CAN ShiftAct 메시지 타임아웃 플래그",
    "MainShiftActMsgTo=ON → P0A1D_8C DTC 설정; Sub CAN 폴백")
A[('DtcMgt','SubShiftActMsgTo','ALL')] = (
    "Sub CAN ShiftAct 메시지 타임아웃 플래그",
    "SubShiftActMsgTo=ON → Main CAN 유지; 양쪽 타임아웃 시 P0A1D_8C DTC")
A[('DtcMgt','BDC02MsgTo','ALL')] = (
    "BDC_02 메시지 타임아웃 플래그",
    "BDC02MsgTo=ON → U0840_8C DTC 설정 조건 충족 (CtAp_NoPubDtcSet.c)")
A[('DtcMgt','BDC05MsgTo','ALL')] = (
    "BDC_05 메시지 타임아웃 플래그",
    "BDC05MsgTo=ON → U0840_8C DTC 설정 조건 충족")
A[('DtcMgt','CLUMsgTo','ALL')] = (
    "CLU_01 메시지 타임아웃 플래그",
    "CLUMsgTo=ON → U0855_8C DTC 설정 조건 충족 (CtAp_NoPubDtcSet.c)")
A[('DtcMgt','CLU_01_20ms_Timeout','ALL')] = (
    "CAN 스택 레벨 CLU_01 타임아웃 원시 플래그",
    "CLU_01_20ms_Timeout=ON → CLUMsgTo=ON → U0855_8C DTC")
A[('DtcMgt','SMK03MsgTo','ALL')] = (
    "SMK_03 메시지 타임아웃 플래그 (Main+Sub 이중화)",
    "SMK03MsgTo=ON → U1065_8C DTC 설정 (CtAp_NoPubDtcSet.c)")
A[('DtcMgt','PDC03MsgTo','ALL')] = (
    "PDC_03 메시지 타임아웃 플래그",
    "PDC03MsgTo=ON → DTC 설정 조건 충족; DrvDrSwSta=OFF 강제")
A[('DtcMgt','SactSig','ALL')] = (
    "E2E 검증된 SActSig 상태 플래그 (DTC 조건용)",
    "SActSig E2E 오류 → P0A1D_82/83 DTC 설정; GearPosSta=NOT_DISPLAY")
A[('DtcMgt','SactSigTo','ALL')] = (
    "SActSig 타임아웃 플래그",
    "SactSigTo=ON → P0A1D_8C DTC 설정 조건 충족")
A[('DtcMgt','SubSActSig','ALL')] = (
    "Sub CAN E2E 검증된 SActSig",
    "Sub SActSig E2E 오류 → P0A1D_82/83 DTC; Main CAN 폴백")
A[('DtcMgt','SubSActSigTo','ALL')] = (
    "Sub CAN SActSig 타임아웃 플래그",
    "SubSActSigTo=ON → Main CAN 유지; 양쪽 타임아웃 시 P0A1D_8C DTC")
A[('DtcMgt','PosSnrFlt','ALL')] = (
    "위치 센서 결함 플래그 (범위 검사 결과)",
    "PosSnrFlt=ON → P2E00_01 DTC 설정 → 모터 정지 (CtAp_PubDtcSet.c)")
A[('DtcMgt','PButtonStuck','ALL')] = (
    "P버튼 stuck 감지 플래그 (18000ms 기준)",
    "PButtonStuck=ON → C1181_96 DTC 설정 조건 충족")
A[('DtcMgt','IdtFltSta','ALL')] = (
    "인디케이터 회로 결함 플래그",
    "IdtFltSta=ON → C1182_96 DTC 설정 (CtAp_PubDtcSet.c)")
A[('DtcMgt','IdtSta','ALL')] = (
    "인디케이터 상태 플래그 DTC 조건 확인",
    "IdtSta 이상 → C1182_96 DTC 조건 판단에 사용")
A[('DtcMgt','DFlReadErrSta','ALL')] = (
    "NVM 읽기 오류 상태 모니터링",
    "DFlReadErrSta=ON → NVM 읽기 실패 DTC 설정 또는 기본값 사용")
A[('DtcMgt','DFlWriteErrSta','ALL')] = (
    "NVM 쓰기 오류 상태 모니터링",
    "DFlWriteErrSta=ON → NVM 쓰기 실패 DTC 설정")
A[('DtcMgt','DMntrVolt','ALL')] = (
    "D단 인디케이터 모니터링 전압 측정값",
    "전압 이상 → IdtFltSta=ON → C1182_96 DTC (CtAp_IdtFltChk.c)")
A[('DtcMgt','NMntrVolt','ALL')] = (
    "N단 인디케이터 모니터링 전압 측정값",
    "전압 이상 → IdtFltSta=ON → C1182_96 DTC")
A[('DtcMgt','PMntrVolt','ALL')] = (
    "P단 인디케이터 모니터링 전압 측정값",
    "전압 이상 → IdtFltSta=ON → C1182_96 DTC")
A[('DtcMgt','RMntrVolt','ALL')] = (
    "R단 인디케이터 모니터링 전압 측정값",
    "전압 이상 → IdtFltSta=ON → C1182_96 DTC")
A[('DtcMgt','AutoBrightSta','ALL')] = (
    "자동 밝기 상태 DTC 조건 게이팅",
    "AutoBrightSta 이상 → 인디케이터 밝기 이상 → C1182_96 DTC 조건 판단")
A[('DtcMgt','GetPSigSta','ALL')] = (
    "P위치 신호 상태 모니터링",
    "GetPSigSta 이상 → P위치 전환 실패 DTC 설정 조건 충족")
A[('DtcMgt','GetPSigTo','ALL')] = (
    "P위치 신호 타임아웃 플래그",
    "GetPSigTo=ON → P위치 신호 타임아웃 DTC 설정")
A[('DtcMgt','RKESig','ALL')] = (
    "RKE 신호 유효성 검증",
    "RKESig 이상 → 원격 P위치 전환 차단; 관련 DTC 설정")
A[('DtcMgt','SMKSig','ALL')] = (
    "SMK 신호 유효성 검증",
    "SMKSig 이상 → 모터 활성화 조건 불충족 → 관련 DTC 설정")
A[('DtcMgt','SleepModeFlag','ALL')] = (
    "슬립 모드 진입 조건 검증",
    "SleepModeFlag 이상 → 슬립 모드 진입 불가 → 전력 소모 증가 모니터링")
A[('DtcMgt','CanBatNormalSta.DebounceBatNorSta','ALL')] = (
    "CAN 배터리 정상 상태 디바운스 플래그 DTC 조건",
    "배터리 비정상 → DTC 인에이블 조건 불충족")
A[('DtcMgt','CanBatNormalSta.ImmediateBatNorSta','ALL')] = (
    "CAN 배터리 즉시 정상 상태 플래그 DTC 조건",
    "배터리 즉시 이상 → DTC 즉시 설정 조건 충족")
A[('DtcMgt','DIdtCntl.BltPwm','ALL')] = (
    "D단 인디케이터 BLT PWM 제어값 DTC 스냅샷 캡처",
    "DTC 발생 시 해당 PWM 값 스냅샷에 기록")
A[('DtcMgt','DIdtCntl.HltPwm','ALL')] = (
    "D단 인디케이터 HLT PWM 제어값 DTC 스냅샷 캡처",
    "DTC 발생 시 해당 PWM 값 스냅샷에 기록")
A[('DtcMgt','DIdtCntl.IdtOnsta','ALL')] = (
    "D단 인디케이터 ON 상태 DTC 스냅샷 캡처",
    "DTC 발생 시 해당 ON/OFF 상태 스냅샷에 기록")
A[('DtcMgt','NIdtCntl.BltPwm','ALL')] = A[('DtcMgt','DIdtCntl.BltPwm','ALL')]
A[('DtcMgt','NIdtCntl.HltPwm','ALL')] = A[('DtcMgt','DIdtCntl.HltPwm','ALL')]
A[('DtcMgt','NIdtCntl.IdtOnsta','ALL')] = A[('DtcMgt','DIdtCntl.IdtOnsta','ALL')]
A[('DtcMgt','PIdtCntl.BltPwm','ALL')] = A[('DtcMgt','DIdtCntl.BltPwm','ALL')]
A[('DtcMgt','PIdtCntl.HltPwm','ALL')] = A[('DtcMgt','DIdtCntl.HltPwm','ALL')]
A[('DtcMgt','PIdtCntl.IdtOnsta','ALL')] = A[('DtcMgt','DIdtCntl.IdtOnsta','ALL')]
A[('DtcMgt','RIdtCntl.BltPwm','ALL')] = A[('DtcMgt','DIdtCntl.BltPwm','ALL')]
A[('DtcMgt','RIdtCntl.HltPwm','ALL')] = A[('DtcMgt','DIdtCntl.HltPwm','ALL')]
A[('DtcMgt','RIdtCntl.IdtOnsta','ALL')] = A[('DtcMgt','DIdtCntl.IdtOnsta','ALL')]

# ══════════════════════════════════════════════════════════════════
# 변수명에서 키워드 추출
# ══════════════════════════════════════════════════════════════════
def extract_var_key(var_str):
    """변수명 첫 줄에서 핵심 키워드 추출 (괄호/줄바꿈 제거)"""
    if not var_str:
        return ''
    first = str(var_str).split('\n')[0].strip()
    first = re.sub(r'\s*\(.*', '', first).strip()
    return first

def extract_unit_key(unit_str):
    if not unit_str:
        return ''
    s = str(unit_str)
    for k in ['PwrMGT','ECUModeMgt','CANMGT','MotorControlMgt','PosMgt',
              'MovingMgt','IdtMgt','ButtonMgt','HapticControlMgt','MoodControlMgt','DtcMgt']:
        if k in s:
            return k
    return s

def lookup(unit_key, var_key, fm):
    """ACTION_DB에서 매칭 검색: 정확히 → fm=ALL 순"""
    # 정확히 매칭
    r = A.get((unit_key, var_key, fm))
    if r: return r
    r = A.get((unit_key, var_key, 'ALL'))
    if r: return r
    # var_key가 더 긴 경우 (예: MovSnr1.ActiveTime → MovSnr1)
    short_var = var_key.split('.')[0]
    r = A.get((unit_key, short_var, fm))
    if r: return r
    r = A.get((unit_key, short_var, 'ALL'))
    if r: return r
    # 부분 매칭: var_key가 DB 키를 포함하거나 DB 키가 var_key를 포함
    for (uk, vk, fmk), val in A.items():
        if uk != unit_key:
            continue
        if (vk in var_key or var_key in vk) and vk:
            if fmk == fm or fmk == 'ALL':
                return val
    return None

# ══════════════════════════════════════════════════════════════════
# Effect on Module / System 패턴 생성기
# ══════════════════════════════════════════════════════════════════
FAILURE_MODE_KO = {
    'MORE':       '정상 범위 초과',
    'LESS':       '정상 범위 미달',
    'CORRUPT':    '데이터 오염/변조',
    'OMISSION':   '신호 누락',
    'COMMISSION': '불필요한 신호 발생',
    'WRONG':      '잘못된 신호',
    'EARLY':      '조기 발생',
    'LATE':       '지연 발생',
    'STUCK':      '고착(Stuck)',
    'REVERSE':    '역전(Reverse)',
    'NO':         '신호 없음',
    'AS WELL AS': '추가 신호 발생',
}

def make_effect_module(unit_key, var_key, fm):
    fm_ko = FAILURE_MODE_KO.get(fm, fm)
    return f"{var_key} {fm_ko} → 해당 SW Unit 처리 오류"

def make_effect_system(unit_key, var_key, fm):
    fm_ko = FAILURE_MODE_KO.get(fm, fm)
    # 유닛별 시스템 영향 패턴
    patterns = {
        'PwrMGT':           '전원 관리 오류 → 시스템 전원 이상 상태 진입 가능',
        'ECUModeMgt':       'ECU 모드 판단 오류 → 잘못된 동작 모드 진입 가능',
        'CANMGT':           'CAN 신호 이상 → 기어 위치 정보 오류 또는 모터 제어 이상',
        'MotorControlMgt':  '모터 제어 오류 → 레버 동작 이상 또는 목표 위치 도달 실패',
        'PosMgt':           '위치 감지 오류 → 기어 위치 판단 불가 → P2E00_01 DTC',
        'MovingMgt':        '이동 센서 오류 → 모터 위치 제어 정확도 저하',
        'IdtMgt':           '인디케이터 이상 → 기어 위치 표시 오류',
        'ButtonMgt':        'P버튼 인식 오류 → P위치 전환 실패 또는 오작동',
        'HapticControlMgt': '햅틱 피드백 이상 → 레버 조작감 오류',
        'MoodControlMgt':   '무드램프 제어 이상 → 조명 품질 저하 (안전에 직접 영향 없음)',
        'DtcMgt':           'DTC 기록 오류 → 고장 진단 정보 손실 또는 오기록',
    }
    base = patterns.get(unit_key, '시스템 기능 이상')
    return base

# ══════════════════════════════════════════════════════════════════
# 메인 처리
# ══════════════════════════════════════════════════════════════════
print("Excel 시작...")
excel = win32.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

print(f"파일 복사: {SRC} → {DEST}")
import shutil
shutil.copy2(SRC, DEST)

wb = excel.Workbooks.Open(DEST)
ws = wb.Worksheets('SW_FMEA')
max_row = ws.UsedRange.Rows.Count
print(f"총 행: {max_row}")

filled   = 0
skipped  = 0
s_cleared = 0
current_unit = ''

for r in range(14, max_row + 1):
    u = ws.Cells(r, 2).Value
    if u:
        current_unit = str(u)

    fm_raw = ws.Cells(r, 6).Value
    if not fm_raw:
        continue

    fm = str(fm_raw).strip().upper()
    unit_key = extract_unit_key(current_unit)
    var_raw  = ws.Cells(r, 4).Value
    var_key  = extract_var_key(var_raw)

    # ── S/O/D/RPN 모두 비우기 ──────────────────────────────────
    for col in [11, 13, 16, 17]:   # S, O, D, RPN
        if ws.Cells(r, col).Value is not None:
            ws.Cells(r, col).Value = None
            s_cleared += 1

    # ── Effect on Module (col8) ────────────────────────────────
    if not ws.Cells(r, 8).Value:
        ws.Cells(r, 8).Value = make_effect_module(unit_key, var_key, fm)
        filled += 1

    # ── Effect on System (col9) ────────────────────────────────
    if not ws.Cells(r, 9).Value:
        ws.Cells(r, 9).Value = make_effect_system(unit_key, var_key, fm)
        filled += 1

    # ── Preventive + Detection 조회 ───────────────────────────
    result = lookup(unit_key, var_key, fm)

    if result:
        prev, det = result
        if not ws.Cells(r, 12).Value:
            ws.Cells(r, 12).Value = prev
            filled += 1
        if not ws.Cells(r, 14).Value:
            ws.Cells(r, 14).Value = det
            filled += 1
    else:
        # 매칭 없는 경우: 패턴 기반 기본값
        if not ws.Cells(r, 12).Value:
            ws.Cells(r, 12).Value = f"{var_key} 입력값 유효성 검증; 범위/타입 확인"
            filled += 1
        if not ws.Cells(r, 14).Value:
            ws.Cells(r, 14).Value = f"{var_key} {FAILURE_MODE_KO.get(fm,fm)} 감지 → 안전 기본값 적용"
            filled += 1
        skipped += 1

    if r % 200 == 0:
        print(f"  진행: {r}/{max_row} (filled={filled}, s_cleared={s_cleared})")

print(f"\n저장 중...")
wb.Save()
wb.Close(False)
excel.Quit()

print(f"\n완료!")
print(f"  채운 셀: {filled}")
print(f"  S/O/D/RPN 삭제: {s_cleared}")
print(f"  DB 미매칭(기본값 적용): {skipped}")
print(f"  출력 파일: {DEST}")
