"""
NQ6e SBW HEV SW FMEA 생성기
코드 분석 기반으로 각 SW Unit의 Interface별 Failure Mode를 작성한다.
"""
import win32com.client

# ──────────────────────────────────────────────────────────────────────────────
# FMEA 데이터 정의
# 각 항목: (No, SWUnit, Category, IntfName, IntfType, FailMode, Detail, EffectModule, EffectSystem, EffectSG)
# Category: Ext=External(HW/CAN), Int=Internal(RTE)
# ──────────────────────────────────────────────────────────────────────────────

FMEA_ROWS = [
    # ───────────────────────────────────────────────────────────────────────────
    # 1. CtAp_VBatStaChk – Battery Voltage State Check
    # ───────────────────────────────────────────────────────────────────────────
    ("1.0", "CtAp_VBatStaChk", "External", "BatVolt\n(ADC 12bit, 0~4095u\n8.5V Under: BatUnder\n9~16V Normal\n16.5V Over: BatOver)",
     "uint16\nRange: 0~4095",
     "MORE",
     "배터리 전압 ADC값이 실제보다 높게 읽힘\n→ V_BAT_OVER 임계값 초과로 잘못 판단",
     "BatOverSta가 잘못 STD_ON으로 설정됨\n→ SysPwrSta에 영향 없으나 DTC 오진단 가능",
     "불필요한 OBD DTC 저장\n배터리 과전압 경고 오표시",
     ""),
    ("", "CtAp_VBatStaChk", "External", "", "", "LESS",
     "배터리 전압 ADC값이 실제보다 낮게 읽힘\n→ V_BAT_UNDER 임계값 미만으로 잘못 판단",
     "BatUnderSta가 잘못 STD_ON으로 설정됨",
     "배터리 저전압 경고 오표시\nHall 센서 BatStbSta 판정에 영향",
     ""),
    ("", "CtAp_VBatStaChk", "External", "", "", "CORRUPT",
     "ADC 노이즈 등으로 BatVolt 값이 불규칙하게 변동\n(Over↔Under 반복)",
     "BatOverSta/BatUnderSta가 교번 Set/Clear됨\n→ Debounce 카운터가 리셋되어 판정 지연",
     "전압 상태 판정 불안정으로 Hall 센서 동작 조건 오판정\n→ 위치 감지 지연 가능",
     ""),
    ("", "CtAp_VBatStaChk", "Internal", "BatStbSta\n(0: None Stable\n1: Stable Voltage)",
     "uint8\nRange: 0~1",
     "NO",
     "BatStbSta가 STD_ON으로 설정되지 않음\n(배터리 전압 안정 조건 미충족 상태 지속)",
     "Hall 센서 Fault Check에서 BatStbSta==OFF로 판정\n→ HallFltChk에서 Fault 판정 억제 실패 가능",
     "Hall 센서 정상 동작 전 Fault 오판정\n→ 위치 감지 불가 상태로 진입",
     ""),

    # ───────────────────────────────────────────────────────────────────────────
    # 2. CtAp_LdoStaChk – SBC Fault State Check
    # ───────────────────────────────────────────────────────────────────────────
    ("2.0", "CtAp_LdoStaChk", "External", "SbcFlt\n(0: Normal\n1: SBC Fault)",
     "uint8\nRange: 0~1",
     "MORE",
     "SBC Fault 핀이 순간적 글리치로 STD_ON 입력\n→ SBC_FLT_TIME 카운터 누적 시 SbcFltSta=ON",
     "SbcFltSta=STD_ON → SysStaChk에서 SysPwrSta=IGN_OFF 강제",
     "IGN 전원이 정상인데도 시스템이 OFF 상태로 판정\n→ 모든 SBW 기능 비활성화",
     "주차 제동 기능 불가 (SG-1 관련)"),
    ("", "CtAp_LdoStaChk", "External", "", "", "NO",
     "SBC 실제 Fault 발생 시 SbcFlt 핀이 STD_ON을 출력하지 않음\n(하드웨어 고장)",
     "SbcFltSta=STD_OFF 유지 → SysStaChk는 정상으로 판단",
     "SBC 공급 전압 이상 상태에서도 Hall 센서 동작 지속\n→ 잘못된 위치 감지 데이터 사용",
     "오류 감지 불가 상태에서 잘못된 기어 위치 정보 송출"),
    ("", "CtAp_LdoStaChk", "Internal", "SbcFltSta\n(0: Normal\n1: SBC Fault After Filter)",
     "uint8\nRange: 0~1",
     "MORE",
     "SBC Fault 없는데 SbcFltSta가 STD_ON으로 유지\n(필터 카운터 초기화 오류)",
     "SysStaChk: SysPwrSta=VEHICLE_POWER_IGN_OFF 유지",
     "정상 IGN 상태에서 SBW 시스템 비활성화",
     ""),

    # ───────────────────────────────────────────────────────────────────────────
    # 3. CtAp_IgnStaChk – Ignition State Check
    # ───────────────────────────────────────────────────────────────────────────
    ("3.0", "CtAp_IgnStaChk", "External", "IgnVolt\n(ADC 12bit, 0~4095u\n4V↓: IgnOFF / 7V↑: IgnON)",
     "uint16\nRange: 0~4095",
     "MORE",
     "IGN 전압 ADC값이 실제보다 높게 읽힘\n→ V_IGN_ON 임계값 초과로 IgnHwSta=STD_ON 오판정",
     "IgnOnSta=STD_ON → SysPwrSta=IGN_ON\n(실제 IGN OFF 상태)",
     "IGN OFF 시 SBW 시스템이 활성화된 상태 유지\n→ 불필요한 위치 감지 동작",
     ""),
    ("", "CtAp_IgnStaChk", "External", "", "", "LESS",
     "IGN 전압 ADC값이 실제보다 낮게 읽힘\n→ V_IGN_OFF 이하로 IgnHwSta=STD_OFF 오판정",
     "IgnOnSta는 CAN IgnSwSta에 의존\n(CAN도 OFF이면 SysPwrSta=IGN_OFF)",
     "실제 IGN ON 상태에서 HW 전압 기반 IGN 감지 실패\n→ CAN IGN 신호만으로 동작 (단일 포인트)",
     ""),
    ("", "CtAp_IgnStaChk", "Internal", "IgnSwSta\n(0: IGN_OFF\n1: IGN_ON)",
     "uint8\nRange: 0~1",
     "MORE",
     "BDC CAN Ign1InSta=ON 신호가 타임아웃/BusOff 후에도\nIgnSwSta=ON으로 유지됨",
     "IgnOnSta=STD_ON으로 잘못 유지\n→ 실제 IGN OFF 시에도 SysPwrSta=IGN_ON",
     "IGN OFF 후 SBW 시스템이 계속 활성화\n→ 불필요한 CAN 메시지 송출 및 전류 소모",
     ""),
    ("", "CtAp_IgnStaChk", "Internal", "IgnOnSta\n(0: STD_OFF\n1: STD_ON)",
     "uint8\nRange: 0~1",
     "NO",
     "IGN 전압도 정상, CAN IGN도 ON이지만\nIgnOnSta=STD_OFF 출력",
     "SysPwrSta=IGN_OFF → 전체 SBW 비활성화",
     "운전자 IGN ON 상태에서 SBW 기어 감지 불가\n→ 기어 위치 CAN 신호 미송출",
     "차량 주행 중 기어 위치 정보 손실 (SG 관련)"),

    # ───────────────────────────────────────────────────────────────────────────
    # 4. CtAp_SysStaChk – System Power State Check
    # ───────────────────────────────────────────────────────────────────────────
    ("4.0", "CtAp_SysStaChk", "Internal", "SbcFltSta\n(0: Normal\n1: SBC Fault)",
     "uint8\nRange: 0~1",
     "MORE",
     "SbcFltSta=STD_ON 오입력\n→ SysPwrStaChk 함수에서 else 분기 진입\n→ SysPwrSta=IGN_OFF 강제 설정",
     "IGN ON 상태임에도 SysPwrSta=VEHICLE_POWER_IGN_OFF",
     "Hall 센서 동작 중단, CAN 신호 송출 중단\n→ 기어 위치 정보 미전달",
     ""),
    ("", "CtAp_SysStaChk", "Internal", "IgnOnSta\n(0: STD_OFF\n1: STD_ON)",
     "uint8\nRange: 0~1",
     "LESS",
     "IgnOnSta=STD_OFF 오입력\n(실제 IGN ON 상태)\n→ SysPwrSta=VEHICLE_POWER_IGN_OFF",
     "SysPwrSta=IGN_OFF → 전체 SBW 기능 비활성화",
     "운전 중 기어 위치 송출 불가\nIndi LED 소등",
     ""),
    ("", "CtAp_SysStaChk", "Internal", "SysPwrSta\n(0: VEHICLE_POWER_IGN_OFF\n1: VEHICLE_POWER_IGN_ON)",
     "uint8\nRange: 0~1",
     "MORE",
     "SysPwrSta=IGN_ON 오출력\n(실제 IGN OFF 상태)",
     "Hall 센서 구동, CAN 송출 계속 동작\n→ 불필요한 전력 소모",
     "IGN OFF 시 SBW 계속 활성화\n→ 배터리 방전 위험",
     ""),
    ("", "CtAp_SysStaChk", "Internal", "", "", "LESS",
     "SysPwrSta=IGN_OFF 오출력\n(실제 IGN ON 상태)",
     "LvrPosInfo, ShiftActSigChk 등 IGN ON 조건 기능 전체 비활성화",
     "기어 위치 감지 및 표시 불가\n→ 운전자 조작 불가",
     "운전 중 SBW 기능 손실 (Safety Goal 관련)"),

    # ───────────────────────────────────────────────────────────────────────────
    # 5. CtAp_CANBusOffChk – CAN Bus-Off Check
    # ───────────────────────────────────────────────────────────────────────────
    ("5.0", "CtAp_CANBusOffChk", "External", "MainCANBusOff\n(0: STD_OFF – 정상\n1: STD_ON – BusOff)",
     "uint8\nRange: 0~1",
     "MORE",
     "Main CAN 정상인데 MainCANBusOff=STD_ON 오판정",
     "CGWSigChk: BDC/CLU/PDC 신호 전부 디폴트값 사용\nShiftActSigChk: GearPosSta=P 고정",
     "BDC IGN 신호, Shift 신호 무효 처리\n→ 기어 위치 P 고정, Indicator 전체 소등",
     ""),
    ("", "CtAp_CANBusOffChk", "External", "", "", "NO",
     "Main CAN BusOff 발생했으나 MainCANBusOff=STD_OFF 유지",
     "CGWSigChk: 타임아웃 전 유효하지 않은 마지막 값 유지\nShiftActSigChk: E2E 에러 미감지",
     "BusOff 상태에서 이전 기어 위치 값으로 동작\n→ 실제 기어 위치와 다른 정보 유지",
     "잘못된 기어 위치 정보로 P 잠금 제어 오동작 가능"),
    ("", "CtAp_CANBusOffChk", "External", "SubCANBusOff\n(0: STD_OFF – 정상\n1: STD_ON – BusOff)",
     "uint8\nRange: 0~1",
     "MORE",
     "Sub CAN 정상인데 SubCANBusOff=STD_ON 오판정",
     "ShiftActSigChk: Sub Shift 신호 무효 처리\n→ TotalAlvCntFlt/TotalCrcFlt 영향",
     "Sub CAN Shift 신호 수신 불가로 판정\n→ Indicator Control 조건 오판단",
     ""),
    ("", "CtAp_CANBusOffChk", "External", "", "", "NO",
     "Sub CAN BusOff 발생했으나 SubCANBusOff=STD_OFF 유지",
     "SubShiftSigErrChk에서 E2E 오류 감지 불가\n→ TotalAlvCntFlt/TotalCrcFlt 미설정",
     "Sub CAN 이상 시 Indicator 정상 동작 유지\n→ Fault 숨김",
     "Indicator 고장 미감지"),

    # ───────────────────────────────────────────────────────────────────────────
    # 6. CtAp_CGWSigChk – CGW Signal Check
    # ───────────────────────────────────────────────────────────────────────────
    ("6.0", "CtAp_CGWSigChk", "External", "Ign1InSta [BDC_02]\n(0: IGN1 Off\n1: IGN1 On\n2: Not Used\n3: Error Indicator)",
     "uint8\nRange: 0~3",
     "MORE",
     "BDC_02 메시지 정상 수신 중 Ign1InSta=1(On) 오수신\n(실제 IGN OFF 상태)",
     "IgnSwStaFlag=STD_ON → IgnStaChk에서 IgnOnSta=STD_ON\n→ SysPwrSta=IGN_ON",
     "IGN OFF 시 SBW 계속 활성화",
     ""),
    ("", "CtAp_CGWSigChk", "External", "", "", "NO",
     "BDC_02 메시지 타임아웃\n→ BDC_02_Timeout=STD_ON\n→ IgnSwStaFlag=STD_OFF 디폴트 처리",
     "IgnSwSta=STD_OFF → IgnStaChk가 HW전압만 참조\n(단일 포인트로 IGN 판정)",
     "CAN IGN 신호 손실 시 HW 전압만으로 IGN 판정\n→ HW 전압도 낮으면 SBW 비활성화",
     ""),
    ("", "CtAp_CGWSigChk", "External", "GearPosPSta [BDC_02]\n(0: Not P\n1: P Position)",
     "uint8\nRange: 0~1",
     "MORE",
     "BDC의 VCU GearPosPSta=1(P) 오수신\n(실제 비-P 상태)",
     "GearPosPStaFlag=P로 잘못 설정\n→ P 위치 신호 오판정",
     "P 잠금 해제 조건 오판단 가능",
     "P 잠금 불필요하게 활성화 또는 해제 오동작"),
    ("", "CtAp_CGWSigChk", "External", "DrvDrSwSta [PDC_03]\n(0: Close\n1: Open\n2: Not Use\n3: Error Indicator)",
     "uint8\nRange: 0~3",
     "MORE",
     "운전석 도어 닫힌 상태에서 DrvDrSwSta=Open(1) 오수신",
     "DrvDrSwStaFlag=Open으로 잘못 설정",
     "P 버튼 오동작 방지 로직(PreventPButtonMisoperation)이 올바르지 않게 동작",
     ""),
    ("", "CtApCGWSigChk", "External", "RKE_BtnReq [RKE]\n(0: None\n1: Lock\n2: Unlock\n3: Trunk\n...)",
     "uint8\nRange: 0~F",
     "CORRUPT",
     "RKE 신호 값이 임의 값으로 수신\n(특히 Lock=1 또는 Unlock=2)",
     "DoorLockStaFlag가 잘못 설정\n→ 도어락 상태 오판단",
     "P 버튼 Misoperation 방지 로직이 오작동\n→ P 버튼 입력 무시 또는 잘못 허용",
     ""),
    ("", "CtAp_CGWSigChk", "External", "AutoBrightSta [CLU_01]\n(1~200: Brightness Step)",
     "uint8\nRange: 0~255",
     "MORE",
     "AutoBrightSta 값이 최대(255)로 수신",
     "DimLvlSet에서 최대 밝기로 설정\n→ Indicator LED 과밝기 구동",
     "Indicator LED 과부하 가능\n(IdtFltChk에서 감지 필요)",
     ""),
    ("", "CtAp_CGWSigChk", "External", "", "", "NO",
     "CLU_01 메시지 타임아웃\n→ CLUMsgToFlag=STD_ON",
     "AutoBrightSta 마지막 수신값 유지\n(CLU 타임아웃 시 디폴트 없음)",
     "조도 변화 시 밝기 자동 조절 불가\n→ 고정 밝기로 동작",
     ""),

    # ───────────────────────────────────────────────────────────────────────────
    # 7. CtAp_DrvRdySigChk – Drive Ready Signal Check (HEV)
    # ───────────────────────────────────────────────────────────────────────────
    ("7.0", "CtAp_DrvRdySigChk", "External", "DrvRdySig [VCU/HCU]\n(HEV_NOT_READY: 0\nHEV_DRIVABLE: 1\n2: Not Used\n3: Error Indicator)",
     "uint8\nRange: 0~3",
     "MORE",
     "HEV_DRIVABLE(1) 신호가 실제 준비 안된 상태에서 수신",
     "DriveSta=STD_ON으로 잘못 설정",
     "HEV 주행 불가 상태에서 주행 가능으로 판정\n→ Indicator P 표시 유지 안 됨",
     "HEV 미준비 시 P 위치 강제 유지 로직 미동작"),
    ("", "CtAp_DrvRdySigChk", "External", "", "", "NO",
     "VCU HEV_DRIVABLE 신호 미수신\n(타임아웃 또는 BusOff)",
     "DriveSigAllTo=STD_ON\n→ DriveSta=STD_OFF (Fail-safe: 주행 불가)",
     "HEV 정상 주행 가능 상태에서 DriveSta=OFF\n→ 기어 위치 제어 제한",
     ""),
    ("", "CtAp_DrvRdySigChk", "External", "", "", "CORRUPT",
     "E2E AlvCnt 반복(REPEATED) 또는 CRC 오류(ERROR)\n→ E2E ErrorCode 이상",
     "AlvCntFlt/CrcFlt 카운터 증가\n→ TotalErrSta=STD_ON",
     "DriveSigAllTo/SubTo 에러 플래그 설정\n→ IdtCntlCdtChk에서 Indicator 소등 조건 진입",
     ""),

    # ───────────────────────────────────────────────────────────────────────────
    # 8. CtAp_ShiftActSigChk – Shift Actuator Signal Check
    # ───────────────────────────────────────────────────────────────────────────
    ("8.0", "CtAp_ShiftActSigChk", "External", "SactSig [Main CAN Shift]\n(E2E Protected: SBW_GearSelSta)\n(0:P / 1:B / 5:D / 6:N / 7:R / ...)",
     "uint8\nRange: 0~E",
     "CORRUPT",
     "Main CAN Shift 신호 CRC 오류\n→ E2E_P_ERROR 검출\n→ CrcFltCnt 누적 → TotalCrcFlt=STD_ON",
     "TotalCrcFlt=ON → IdtCntlCdtChk: INDICDT_HIGHOFF\n→ Indicator 전체 소등",
     "기어 위치 Indicator 소등\n→ 운전자 현재 기어 위치 확인 불가",
     "Indicator 오류로 인한 운전자 혼란 (SG관련)"),
    ("", "CtAp_ShiftActSigChk", "External", "", "", "NO",
     "Main Shift CAN 메시지 타임아웃\n→ MainShiftActMsgTo=STD_ON\n→ GearPosSta=P 디폴트",
     "GearPosSta=P 강제 설정\n→ IdtCntl에서 P LED만 점등",
     "실제 기어 위치와 무관하게 P 표시\n→ 운전자 혼란",
     "주행 중 P 강제 표시"),
    ("", "CtAp_ShiftActSigChk", "External", "SubSActSig [Sub CAN Shift]\n(E2E Protected: SCU_FF_PosTarSta)\n(0:P / 5:D / 6:N / 7:R / ...)",
     "uint8\nRange: 0~E",
     "CORRUPT",
     "Sub CAN AlvCnt 시퀀스 이상\n→ E2E_P_WRONGSEQUENCE\n→ AlvCntFltCnt 누적 → TotalAlvCntFlt=STD_ON",
     "TotalAlvCntFlt=ON → IdtCntlCdtChk: INDICDT_HIGHOFF\n→ Indicator 전체 소등",
     "기어 위치 Indicator 소등",
     ""),
    ("", "CtAp_ShiftActSigChk", "Internal", "GearPosSta\n(0:P / 5:D / 6:N / 7:R / ...)",
     "uint8\nRange: 0~F",
     "MORE",
     "GearPosSta가 실제 기어 위치보다 높은 값으로 출력\n(예: N→D 오판정)",
     "IdtCntl에서 D LED 점등\n(실제는 N 위치)",
     "운전자에게 잘못된 기어 위치 표시\n→ 혼란으로 인한 조작 오류",
     "잘못된 기어 위치 표시 (SG 관련)"),

    # ───────────────────────────────────────────────────────────────────────────
    # 9. CtAp_HallDataSet – Hall Sensor Data Acquisition
    # ───────────────────────────────────────────────────────────────────────────
    ("9.0", "CtAp_HallDataSet", "External", "TmagANGLE_Die1 [TMAG5170 I2C]\n(sint16, Die1 각도값)",
     "sint16\nRange: -32768~32767",
     "MORE",
     "Die1 ANGLE 값이 실제보다 크게 읽힘\n(I2C 노이즈 또는 센서 이상)",
     "LvrPosChk에서 칼리브레이션 범위 벗어남\n→ R/Nr/Null/Nd/D 위치 판정 오류",
     "기어 레버 위치 오판정\n→ SBWSigSet에서 잘못된 위치 CAN 송출",
     "잘못된 기어 위치 정보 전달 (SG관련)"),
    ("", "CtAp_HallDataSet", "External", "", "", "NO",
     "I2C 통신 오류로 Die1 ANGLE 값 미수신\n(0 또는 이전값 유지)",
     "HallFltChk: AngleFltChk에서 Die1-Die2 차이 > 55도\n→ HallSnrFlt.TotalFlt=ON",
     "LvrPosInfo: LvrPos=LVR_Flt 강제 설정\n→ 위치 감지 불가",
     "위치 감지 중단 → SBW 기능 안전 중단"),
    ("", "CtAp_HallDataSet", "External", "TmagMAG_Die1 [TMAG5170 I2C]\n(sint16, 자기 세기)",
     "sint16\nRange: 300~1700 (정상)",
     "LESS",
     "Die1 자기 세기가 300 미만으로 측정\n(마그넷 이탈 또는 센서 고장)",
     "HallFltChk: MagnitudeFltChk → HallSnrFlt.VG=ON\n→ TotalFlt=ON",
     "위치 감지 불가 상태로 Fail-safe 진입",
     ""),
    ("", "CtAp_HallDataSet", "External", "TmagANGLE_Die2 [TMAG5170 I2C]\n(sint16, Die2 각도값)",
     "sint16\nRange: -32768~32767",
     "CORRUPT",
     "Die2 ANGLE 값이 임의 값으로 변화\n(Die1-Die2 각도 차이 > 55도 발생)",
     "HallFltChk: AngleFltChk → ANGLE_DIFF_FLT_VAL(55) 초과\n→ HallSnrFlt.Alpha 또는 Beta=ON",
     "TotalFlt=ON → LvrPos=LVR_Flt\n→ 위치 감지 중단",
     ""),

    # ───────────────────────────────────────────────────────────────────────────
    # 10. CtAp_HallFltChk – Hall Sensor Fault Check
    # ───────────────────────────────────────────────────────────────────────────
    ("10.0", "CtAp_HallFltChk", "Internal", "HallSnrDie1/Die2 ANGLE\n(AngleDiff 임계: 55deg\nERR_TIME: 100 cycle)",
     "sint16\nRange: -32768~32767",
     "MORE",
     "Die1-Die2 각도 차이가 ANGLE_DIFF_FLT_VAL(55) 초과\n→ HALL_ERR_TIME(100cy) 경과 후 TotalFlt=ON",
     "HallSnrFltSta.TotalFlt=STD_ON",
     "LvrPosInfo: LvrPos=LVR_Flt\n→ SBWSigSet: PosSnrFlt=ON 송출\n→ 위치 감지 중단",
     ""),
    ("", "CtAp_HallFltChk", "Internal", "HallSnrDie1/Die2 MAG\n(정상 범위: 300~1700\nMAG_FLT_VAL_LOW:300 / HIGH:1700)",
     "sint16\nRange: 300~1700",
     "LESS",
     "자기 세기가 MAG_FLT_VAL_LOW(300) 미만\n→ HALL_ERR_TIME 경과 후 TotalFlt=ON",
     "HallSnrFltSta.TotalFlt=STD_ON\n(MagnitudeFlt)",
     "위치 감지 Fail-safe 진입\n마그넷 이탈 또는 조립 불량 감지",
     ""),
    ("", "CtAp_HallFltChk", "Internal", "HallSnrFltSta\n(TotalFlt, Alpha, Beta, VG 포함)",
     "struct",
     "NO",
     "실제 Hall 센서 이상이지만 TotalFlt=STD_OFF 오판정\n(EcuSta 또는 BatStbSta 조건 미충족으로 Fault 억제)",
     "LvrPosInfo: Fault 억제로 잘못된 LvrPos 사용 지속",
     "잘못된 위치에서 SBW 동작 계속\n→ 잘못된 기어 위치 CAN 송출",
     "잘못된 기어 위치 정보 지속 송출 (SG관련)"),

    # ───────────────────────────────────────────────────────────────────────────
    # 11. CtAp_LvrPosChk – Lever Position Check
    # ───────────────────────────────────────────────────────────────────────────
    ("11.0", "CtAp_LvrPosChk", "Internal", "Die1/Die2 ANGLE + CalData\n(R/Nr/Null/Nd/D 범위 비교)",
     "sint16 + struct",
     "MORE",
     "Hall 각도 값이 CalData 범위보다 크게 측정\n→ 상위 위치(D 방향)로 오판정",
     "LvrPos가 실제보다 높은 기어 위치로 판정\n(예: Nr→D 오인식)",
     "SBWSigSet에서 잘못된 기어 위치 CAN 송출\n→ Indicator 오표시",
     "잘못된 기어 표시로 운전자 혼란"),
    ("", "CtAp_LvrPosChk", "Internal", "", "", "LESS",
     "Hall 각도 값이 CalData 범위보다 작게 측정\n→ 하위 위치(R 방향)로 오판정",
     "LvrPos가 실제보다 낮은 기어 위치로 판정",
     "잘못된 기어 위치 CAN 송출",
     ""),
    ("", "CtAp_LvrPosChk", "Internal", "", "", "NO",
     "모든 위치 조건 미충족 (R/Nr/Null/Nd/D 모두 불일치)\n→ LvrPos = Null(중간 위치)로 판정",
     "LvrPosInfo에서 LVR_Null 상태로 처리",
     "기어 레버가 특정 위치에 있는데 Null로 처리\n→ 위치 검출 실패 상태 CAN 송출",
     ""),

    # ───────────────────────────────────────────────────────────────────────────
    # 12. CtAp_LvrPosInfo – Lever Position Info
    # ───────────────────────────────────────────────────────────────────────────
    ("12.0", "CtAp_LvrPosInfo", "Internal", "LvrPos (from LvrPosChk)\n(1:R / 2:Nr / 3:Null / 4:Nd / 5:D / LVR_Flt)",
     "uint8\nRange: 1~6",
     "MORE",
     "LvrPos가 실제보다 높은 값으로 입력\n→ LvrPosDecTimeChk에서 위치 변경으로 감지\n→ PosDecCnt 카운터 동작",
     "Debounce 시간 경과 후 잘못된 위치 확정",
     "잘못된 LvrPosInfo 출력\n→ SBWSigSet에서 잘못된 기어 위치 CAN 전송",
     ""),
    ("", "CtAp_LvrPosInfo", "Internal", "HallSnrFlt.TotalFlt\n(0: Normal\n1: Fault)",
     "uint8\nRange: 0~1",
     "NO",
     "Hall 센서 실제 Fault이지만 TotalFlt=STD_OFF\n→ LvrPosDecCndChk에서 Fault 분기 미진입",
     "잘못된 Hall 데이터로 LvrPos 계속 계산",
     "Fault 상태에서 위치 감지 지속\n→ 잘못된 기어 위치 정보 사용",
     ""),
    ("", "CtAp_LvrPosInfo", "Internal", "LvrPosInfo.PosStuck\n(0: STD_OFF\n1: STD_ON – Stuck 감지)",
     "uint8\nRange: 0~1",
     "NO",
     "기어 레버 Stuck 발생했으나 PosStuck=STD_OFF 유지\n(Stuck 감지 시간 조건 미충족)",
     "SBWSigSet에서 PosStuck 미송출\n→ VCU에 Stuck 상태 미전달",
     "기어 레버 고착 상태에서 정상 동작 오판단",
     "기어 레버 고착 감지 실패"),
    ("", "CtAp_LvrPosInfo", "Internal", "LvrPosInfo.PosSta\n(1:Lvr_R / 2:Lvr_Nr / 3:Lvr_Null / 4:Lvr_Nd / 5:Lvr_D / 6:Lvr_Flt)",
     "uint8\nRange: 1~6",
     "CORRUPT",
     "메모리 손상 등으로 PosSta가 임의 값으로 변경",
     "SBWSigSet에서 잘못된 기어 위치 CAN 전송",
     "VCU/Cluster에 잘못된 기어 위치 전달",
     "잘못된 기어 위치 표시/제어 (SG관련)"),

    # ───────────────────────────────────────────────────────────────────────────
    # 13. CtAp_PButtonSet – P Button State Set
    # ───────────────────────────────────────────────────────────────────────────
    ("13.0", "CtAp_PButtonSet", "External", "P_SW1_Raw [ADC]\n(정상 ON: 2187~2776u\n정상 OFF: 581~1106u)",
     "uint16\nRange: 0~4095",
     "MORE",
     "P_SW1_Raw 값이 InterSwOnMax_1(2776) 초과\n→ CONTACT_FAULT(2) 판정",
     "PSw1State=CONTACT_FAULT\n→ P 버튼 입력 무효 처리",
     "운전자 P 버튼 입력이 무시됨\n→ P 잠금 요청 미전달",
     "P 잠금 요청 불가"),
    ("", "CtAp_PButtonSet", "External", "", "", "LESS",
     "P_SW1_Raw 값이 InterSwOffMin_1(581) 미만\n→ CONTACT_FAULT(2) 판정",
     "PSw1State=CONTACT_FAULT\n→ P 버튼 입력 무효 처리",
     "P 버튼 단선 오판정\n→ P 잠금 요청 불가",
     ""),
    ("", "CtAp_PButtonSet", "External", "P_SW2_Raw [ADC]\n(정상 ON: 2850~3800u\n정상 OFF: 1253~2138u)",
     "uint16\nRange: 0~4095",
     "AS WELL AS",
     "P_SW1과 P_SW2가 동시에 ON 조건 만족\n(정상: 하나씩 순차적으로 눌림)\n→ 두 스위치 동시 활성화",
     "FilterPSwStuck에서 동시 ON = Stuck 조건 판정 가능",
     "PButtonStuck=STD_ON 설정\n→ SBWSigSet에서 Stuck 상태 송출",
     ""),
    ("", "CtAp_PButtonSet", "Internal", "PButtonSta\n(0: STD_OFF\n1: STD_ON)",
     "uint8\nRange: 0~1",
     "MORE",
     "P 버튼 누르지 않았는데 PButtonSta=STD_ON 출력\n(ADC 오판정 또는 메모리 오류)",
     "SBWSigSet: LvrPStaSet에서 P 잠금 요청 신호 생성",
     "의도하지 않은 P 잠금 활성화\n→ 주행 중 P 잠금 요청",
     "주행 중 비의도적 P 잠금 (SG 관련)"),
    ("", "CtAp_PButtonSet", "Internal", "PButtonStuck\n(0: STD_OFF\n1: Stuck 감지)",
     "uint8\nRange: 0~1",
     "NO",
     "P 버튼 실제 Stuck이지만 PButtonStuck=STD_OFF\n(PButtonSet_PSwSutckTime(18000cy) 미도달)",
     "Stuck 상태에서 지속적 P 잠금 요청 발생",
     "의도하지 않은 지속적 P 잠금 활성화",
     ""),

    # ───────────────────────────────────────────────────────────────────────────
    # 14. CtAp_IdtFltChk – Indicator Fault Check
    # ───────────────────────────────────────────────────────────────────────────
    ("14.0", "CtAp_IdtFltChk", "External", "PMntrVolt [ADC]\n(P LED 모니터링 전압\n정상 ON: HIGHLIGHT_P_LED_ON_HIGH 이상)",
     "uint16\nRange: 0~4095",
     "LESS",
     "P LED 구동 중 PMntrVolt가 HIGHLIGHT_P_LED_ON_HIGH 미만\n→ P LED 단선 또는 쇼트 감지",
     "IdtPHltChk에서 P LED Fault 판정\n→ IdtFltSta=STD_ON",
     "IdtCntlCdtChk: INDICDT_HIGHOFF → Highlight 소등\n→ 운전자 P 위치 확인 불가",
     ""),
    ("", "CtAp_IdtFltChk", "External", "NMntrVolt / DMntrVolt / RMntrVolt [ADC]\n(N/D/R LED 모니터링 전압)",
     "uint16\nRange: 0~4095",
     "NO",
     "LED 점등 명령 없는 상태에서 모니터링 전압 측정 생략\n→ LED Fault 감지 불가",
     "IdtFltSta=STD_OFF 유지 (정상으로 판단)",
     "LED 단선/쇼트 미감지 상태로 동작\n→ Fault LED 계속 사용",
     ""),
    ("", "CtAp_IdtFltChk", "Internal", "IdtFltSta\n(0: STD_OFF\n1: Fault)",
     "uint8\nRange: 0~1",
     "MORE",
     "LED 정상인데 IdtFltSta=STD_ON 오출력",
     "IdtCntlCdtChk: INDICDT_HIGHOFF → Highlight LED 전체 소등",
     "정상 상태에서 Indicator Highlight 소등\n→ 운전자 기어 위치 확인 불가",
     ""),

    # ───────────────────────────────────────────────────────────────────────────
    # 15. CtAp_IdtCntl – Indicator Control
    # ───────────────────────────────────────────────────────────────────────────
    ("15.0", "CtAp_IdtCntl", "Internal", "GearPosSta\n(0:P / 5:D / 6:N / 7:R / ...)",
     "uint8\nRange: 0~F",
     "MORE",
     "GearPosSta가 실제 N(6)인데 D(5)로 입력\n→ IdtDOnSet() 호출",
     "D LED 점등 (실제는 N 위치)",
     "운전자에게 잘못된 기어 위치 표시\n→ D인 줄 알고 조작 오류",
     "잘못된 기어 표시"),
    ("", "CtAp_IdtCntl", "Internal", "", "", "NO",
     "GearPosSta 수신 실패\n(RTE Read 오류 또는 데이터 미갱신)\n→ 이전 값 유지",
     "이전 기어 위치 LED 계속 점등",
     "기어 변경 시 LED 미갱신\n→ 잘못된 기어 표시 지속",
     ""),
    ("", "CtAp_IdtCntl", "Internal", "IdtCntlCdtSta\n(0: INDICDT_NORMAL\n1: INDICDT_HIGHOFF)",
     "uint8\nRange: 0~1",
     "MORE",
     "INDICDT_HIGHOFF(1) 오입력\n(실제 정상 조건)",
     "IdtHltOffSet() 호출 → Highlight 전체 소등",
     "정상 상태에서 Indicator Highlight 소등",
     ""),
    ("", "CtAp_IdtCntl", "Internal", "HltDimLvl / BltDimLvl\n(0~65535u)",
     "uint16\nRange: 0~65535",
     "MORE",
     "HltDimLvl 값이 최대(65535)로 입력\n→ Highlight LED 최고 밝기 구동",
     "LED에 과전류 공급 가능\n→ LED 손상 또는 수명 단축",
     "IdtFltChk에서 과전압 감지 후 FaultSta=ON 가능",
     ""),

    # ───────────────────────────────────────────────────────────────────────────
    # 16. CtAp_SBWSigSet – SBW CAN Signal Set
    # ───────────────────────────────────────────────────────────────────────────
    ("16.0", "CtAp_SBWSigSet", "Internal", "LvrPosInfo.PosSta\n(1:R / 2:Nr / 3:Null / 4:Nd / 5:D / 6:Flt)",
     "uint8\nRange: 1~6",
     "MORE",
     "PosSta가 실제보다 높은 값으로 입력\n→ LvrPosInfoSet에서 잘못된 위치 CAN 데이터 생성",
     "CtApSBWSigSet_O_u1_LvrMsg에 잘못된 위치 데이터 설정",
     "Main/Sub CAN에 잘못된 기어 위치 전송\n→ VCU, Cluster 오판단",
     "잘못된 기어 위치 제어/표시 (SG관련)"),
    ("", "CtAp_SBWSigSet", "Internal", "PButtonSta\n(0: STD_OFF\n1: STD_ON)",
     "uint8\nRange: 0~1",
     "MORE",
     "PButtonSta=STD_ON 오입력 (P 버튼 안 눌렸음)\n→ LvrPStaSet에서 P 잠금 요청 생성",
     "LvrPSta=1(P 잠금 요청) CAN 송출",
     "비의도적 P 잠금 CAN 신호 송출\n→ VCU에서 P 잠금 활성화",
     "주행 중 비의도적 P 잠금 (SG 관련)"),
    ("", "CtAp_SBWSigSet", "Internal", "", "", "NO",
     "PButtonSta=STD_OFF (P 버튼 눌렸음에도 미감지)\n→ LvrPStaSet에서 P 잠금 요청 미생성",
     "LvrPSta=0(P 잠금 미요청) CAN 송출",
     "운전자 P 버튼 입력이 VCU에 전달되지 않음\n→ P 잠금 미동작",
     "P 잠금 요청 미전달"),
    ("", "CtAp_SBWSigSet", "Internal", "IdtSta\n(INIT/P_ON/R_ON/N_ON/D_ON 등)",
     "uint8\nRange: 0~FF",
     "CORRUPT",
     "IdtSta 값이 임의 값으로 설정\n→ LvrIdtStaSet에서 잘못된 Indicator 상태 CAN 송출",
     "잘못된 IdtSta CAN 메시지 전송",
     "VCU/Cluster에 잘못된 Indicator 상태 전달",
     ""),
]

# ──────────────────────────────────────────────────────────────────────────────
# Excel 파일 생성
# ──────────────────────────────────────────────────────────────────────────────

OUT_FILE = r'E:\claude\FMEA\NQ6e_SBW_SW_FMEA.xlsx'

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

wb = excel.Workbooks.Add()
ws = wb.Sheets(1)
ws.Name = "SW_FMEA"

# ── 헤더 ──────────────────────────────────────────────────────────────────────
HDR_ROW = 1
headers = [
    "No", "SW Unit Name", "Interface Category",
    "Interface\n(Variable) name", "Interface\n(Variable) type",
    "Failure Mode\n(HAZOP Keyword)",
    "Detail of the Failure Mode",
    "Effect on Module",
    "Effect on System",
    "Effect on SG",
    "S", "Preventive Action", "O",
    "Detection Action\n(Safety Mechanism)", "Test Method",
    "D", "RPN",
]

# 헤더 스타일
for ci, h in enumerate(headers, start=1):
    c = ws.Cells(HDR_ROW, ci)
    c.Value = h
    c.Interior.Color = 0x4472C4       # 파란색
    c.Font.Color = 0xFFFFFF           # 흰색 텍스트
    c.Font.Bold = True
    c.WrapText = True

# ── 데이터 입력 ───────────────────────────────────────────────────────────────
SKYBLUE   = 0xFFE4B5   # Moccasin (External)
LIGHTGRAY = 0xF2F2F2   # 연회색 (Internal)
HAZOP_COLORS = {
    "MORE":       0xE2EFDA,  # 연두
    "LESS":       0xFFF2CC,  # 연노랑
    "NO":         0xFCE4D6,  # 연주황
    "CORRUPT":    0xFFE6E6,  # 연빨강
    "REVERSE":    0xE6E0FF,  # 연보라
    "AS WELL AS": 0xE0F7FA,  # 연시안
}

row = HDR_ROW + 1
prev_unit = None

for item in FMEA_ROWS:
    (no, unit, cat, intf, itype, fail, detail, eff_mod, eff_sys, eff_sg) = item

    vals = [no, unit, cat, intf, itype, fail, detail, eff_mod, eff_sys, eff_sg,
            "", "", "", "", "", "", ""]

    bg = HAZOP_COLORS.get(fail, 0xFFFFFF)

    for ci, v in enumerate(vals, start=1):
        c = ws.Cells(row, ci)
        c.Value = v
        c.WrapText = True
        c.VerticalAlignment = -4160  # xlTop
        if ci <= 3:
            # 구분 열 배경
            if cat == "External":
                c.Interior.Color = 0xDDEBF7  # 연파랑
            elif cat == "Internal":
                c.Interior.Color = 0xE2EFDA  # 연초록
        elif ci in (6,):  # Failure Mode 열
            c.Interior.Color = bg
            c.Font.Bold = True

    row += 1

# ── 열 너비 ───────────────────────────────────────────────────────────────────
col_widths = [6, 22, 14, 40, 18, 16, 50, 35, 35, 20,
              5, 20, 5, 25, 20, 5, 8]
for ci, w in enumerate(col_widths, start=1):
    ws.Columns(ci).ColumnWidth = w

# ── 행 높이 자동 (헤더 제외) ──────────────────────────────────────────────────
ws.Rows(HDR_ROW).RowHeight = 36
for r in range(HDR_ROW + 1, row):
    ws.Rows(r).RowHeight = 60

# ── 틀 고정 ───────────────────────────────────────────────────────────────────
ws.Cells(HDR_ROW + 1, 4).Select()
excel.ActiveWindow.FreezePanes = True

# ── 저장 ──────────────────────────────────────────────────────────────────────
wb.SaveAs(OUT_FILE)
wb.Close(False)
excel.Quit()

print(f"FMEA 생성 완료: {OUT_FILE}")
print(f"총 {row - HDR_ROW - 1}개 항목")
