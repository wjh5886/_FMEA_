"""
LQ2 SBW ICE SW FMEA 생성기
Static_Code/Source 분석 기반 - 각 SW Unit의 Interface별 Failure Mode
ISO 26262 / A-SPICE 기준 적용
"""
import xlsxwriter

# ──────────────────────────────────────────────────────────────────────────────
# FMEA 데이터 정의
# (No, SWUnit, Category, IntfName, IntfType, FailMode, Detail, EffectModule, EffectSystem, EffectSG)
# Category: External(HW/CAN입력), Internal(RTE포트)
# FailMode: MORE / LESS / NO / CORRUPT / REVERSE / AS WELL AS
# ──────────────────────────────────────────────────────────────────────────────

FMEA_ROWS = [

    # ═══════════════════════════════════════════════════════════════════════════
    # 1. CtAp_VBatStaChk – Battery Voltage State Check
    # ═══════════════════════════════════════════════════════════════════════════
    ("1.0", "CtAp_VBatStaChk", "External",
     "BatVolt\n(ADC 12bit, 0~4095\nV_BAT_UNDER: 8.5V 이하\nV_BAT_NORMAL: 9~16V\nV_BAT_OVER: 16.5V 이상)",
     "uint16\nRange: 0~4095",
     "MORE",
     "배터리 전압 ADC값이 실제보다 높게 읽힘\n→ V_BAT_OVER 임계값 초과로 BatOverSta=ON 오판정",
     "BatOverSta=STD_ON 잘못 설정\n→ 배터리 과전압 DTC 오진단",
     "불필요한 DTC C1101 저장\n배터리 과전압 경고 오표시",
     ""),
    ("", "CtAp_VBatStaChk", "External", "", "", "LESS",
     "배터리 전압 ADC값이 실제보다 낮게 읽힘\n→ V_BAT_UNDER 임계값 미만으로 BatUnderSta=ON 오판정",
     "BatUnderSta=STD_ON 잘못 설정\n→ BatStbSta 판정에 영향",
     "DTC C1102 저전압 오진단\nHall 센서 BatStbSta 조건 불충족",
     ""),
    ("", "CtAp_VBatStaChk", "External", "", "", "CORRUPT",
     "ADC 노이즈로 BatVolt값이 불규칙 변동\n(Over↔Under 반복)",
     "BatOverSta/BatUnderSta 교번 Set/Clear\n→ Debounce 카운터 리셋으로 판정 지연",
     "전압 상태 판정 불안정\n→ Hall 센서 동작 조건 오판정",
     ""),
    ("", "CtAp_VBatStaChk", "Internal",
     "BatStbSta\n(0: Unstable\n1: Stable)",
     "uint8\nRange: 0~1",
     "NO",
     "BatStbSta가 STD_ON으로 설정되지 않음\n(배터리 전압 안정 조건 미충족 지속)",
     "LdoStaChk: Ldo2FltChk 활성화 안됨\nHallFltChk: Fault 판정 억제 실패",
     "Hall 센서 정상 동작 전 Fault 오판정\n→ LvrPosInfo 위치 감지 불가",
     "SG01"),
    ("", "CtAp_VBatStaChk", "Internal",
     "BatOverSta\n(0: Normal\n1: Over Voltage)",
     "uint8\nRange: 0~1",
     "MORE",
     "실제 과전압 아닌 상황에서 BatOverSta=ON 지속",
     "DTC C1101 오발생",
     "배터리 과전압 경고 오표시",
     ""),

    # ═══════════════════════════════════════════════════════════════════════════
    # 2. CtAp_LdoStaChk – LDO2/SBC Fault State Check
    # ═══════════════════════════════════════════════════════════════════════════
    ("2.0", "CtAp_LdoStaChk", "External",
     "SbcFlt\n(0: Normal\n1: SBC Fault)",
     "uint8\nRange: 0~1",
     "MORE",
     "SBC Fault 핀이 순간 글리치로 STD_ON 입력\n→ SBC_FLT_TIME 카운터 누적 시 SbcFltSta=ON",
     "SbcFltSta=ON → SysStaChk에서 SysPwrSta=IGN_OFF 강제",
     "정상 IGN 상태에서 SBW 시스템 비활성화\n→ 기어 변속 불가",
     "SG01, SG02, SG03"),
    ("", "CtAp_LdoStaChk", "External", "", "", "NO",
     "SBC 실제 Fault 발생 시 SbcFlt 핀이 STD_ON 미출력\n(HW 고장)",
     "SbcFltSta=OFF 유지 → SysStaChk 정상으로 판단",
     "SBC 전압 이상 상태에서 Hall 센서 동작 지속\n→ 잘못된 위치 감지 데이터 사용",
     ""),
    ("", "CtAp_LdoStaChk", "External",
     "Ldo2OnVolt\n(ADC 12bit\nLDO2 정상: 3.0~3.6V)",
     "uint16\nRange: 0~4095",
     "LESS",
     "LDO2 공급 전압 ADC값이 임계값 이하로 읽힘\n→ Ldo2UnderFlag=ON",
     "Ldo2FltSta=ON → 인디케이터 Fault 상태 진입\nIdtFltChk에서 Ldo2FltSta 참조",
     "인디케이터 오동작 DTC 발생\n위치 표시 LED 소등",
     "SG03"),
    ("", "CtAp_LdoStaChk", "External", "", "", "MORE",
     "LDO2 공급 전압 ADC값이 정상 범위 초과\n→ 과전압 상태",
     "Ldo2FltSta 판정 로직에서 이상 없음으로 처리\n(과전압 미감지 케이스)",
     "인디케이터 LED 과전압으로 소손 위험\n(DTC 미발생)",
     ""),
    ("", "CtAp_LdoStaChk", "Internal",
     "SbcFltSta\n(0: Normal\n1: SBC Fault After Filter)",
     "uint8\nRange: 0~1",
     "MORE",
     "SBC Fault 없는데 SbcFltSta=ON 유지\n(필터 카운터 초기화 오류)",
     "SysStaChk: SysPwrSta=IGN_OFF 유지",
     "정상 IGN 상태에서 SBW 시스템 비활성화",
     ""),
    ("", "CtAp_LdoStaChk", "Internal",
     "Ldo2FltSta\n(0: Normal\n1: LDO2 Fault)",
     "uint8\nRange: 0~1",
     "NO",
     "LDO2 실제 Fault 상태에서 Ldo2FltSta=OFF 유지",
     "IdtFltChk: Ldo2FltSta=OFF → 인디케이터 Fault 미감지",
     "인디케이터 이상 상태에서도 정상 동작 가정\n→ 오위치 표시 지속",
     "SG03"),

    # ═══════════════════════════════════════════════════════════════════════════
    # 3. CtAp_IgnStaChk – Ignition State Check
    # ═══════════════════════════════════════════════════════════════════════════
    ("3.0", "CtAp_IgnStaChk", "External",
     "IgnVolt\n(ADC 12bit\n4V 이하: IGN_OFF\n7V 이상: IGN_ON)",
     "uint16\nRange: 0~4095",
     "MORE",
     "IGN 전압 ADC값이 실제보다 높게 읽힘\n→ V_IGN_ON 임계값 초과로 IgnHwSta=ON 오판정",
     "IgnOnSta=ON → SysPwrSta=IGN_ON\n(실제 IGN OFF 상태)",
     "IGN OFF 시 SBW 시스템 활성화 유지\n→ 불필요한 CAN 메시지 송출",
     ""),
    ("", "CtAp_IgnStaChk", "External", "", "", "LESS",
     "IGN 전압 ADC값이 실제보다 낮게 읽힘\n→ V_IGN_OFF 이하로 IgnHwSta=OFF 오판정",
     "IgnOnSta: CAN IgnSwSta에만 의존\n(HW IGN 감지 단일 포인트 실패)",
     "실제 IGN ON 상태에서 HW 전압 기반 감지 실패\n→ CAN 신호만으로 동작",
     ""),
    ("", "CtAp_IgnStaChk", "Internal",
     "IgnSwSta\n(CAN BDC Ign1InSta\n0: IGN_OFF\n1: IGN_ON)",
     "uint8\nRange: 0~1",
     "MORE",
     "BDC CAN Ign1InSta=ON 신호가 타임아웃 후에도\nIgnSwSta=ON으로 유지됨",
     "IgnOnSta=ON 잘못 유지\n→ IGN OFF 시에도 SysPwrSta=IGN_ON",
     "IGN OFF 후 SBW 시스템 계속 활성화\n→ 불필요한 전류 소모 및 CAN 송출",
     ""),
    ("", "CtAp_IgnStaChk", "Internal", "", "", "NO",
     "BDC CAN 신호 수신 불가\n→ IgnSwSta=OFF 고착",
     "IgnHwSta만으로 IgnOnSta 판정\n(BDC CAN Timeout 상태)",
     "IGN ON 시 BDC CAN 미수신 시 IgnOnSta 불안정\n→ SysPwrSta 오판정 가능",
     ""),
    ("", "CtAp_IgnStaChk", "Internal",
     "IgnOnSta\n(0: OFF\n1: ON)",
     "uint8\nRange: 0~1",
     "NO",
     "IgnVolt 정상, CAN IGN ON 상태에서도\nIgnOnSta=OFF 출력",
     "SysPwrSta=IGN_OFF → 전체 SBW 비활성화",
     "운전자 IGN ON 상태에서 기어 위치 감지 불가\n→ 기어 위치 CAN 신호 미송출",
     "SG01, SG02, SG03"),

    # ═══════════════════════════════════════════════════════════════════════════
    # 4. CtAp_SysStaChk – System Power State Check
    # ═══════════════════════════════════════════════════════════════════════════
    ("4.0", "CtAp_SysStaChk", "Internal",
     "SbcFltSta\n(0: Normal\n1: SBC Fault)",
     "uint8\nRange: 0~1",
     "MORE",
     "SBC 정상인데 SbcFltSta=ON이 입력됨\n→ SysPwrSta=IGN_OFF 강제",
     "모든 SBW 기능 비활성화",
     "IGN ON 상태에서 SBW 동작 불가\n→ 기어 변속 불가능",
     "SG01, SG02, SG03"),
    ("", "CtAp_SysStaChk", "Internal", "", "", "NO",
     "SBC 실제 Fault 상태인데 SbcFltSta=OFF 입력\n→ SysPwrSta=IGN_ON 유지",
     "SBC 이상 상태에서도 SBW 동작 계속",
     "잘못된 전원 상태에서 Hall 센서/인디케이터 동작\n→ 오위치 감지 위험",
     ""),
    ("", "CtAp_SysStaChk", "Internal",
     "IgnOnSta\n(0: OFF\n1: ON)",
     "uint8\nRange: 0~1",
     "NO",
     "IGN 실제 ON 상태인데 IgnOnSta=OFF 입력됨\n→ SysPwrSta=IGN_OFF",
     "SBW 전체 기능 비활성화",
     "주행 중 기어 위치 정보 손실\n→ 변속 요청 무시",
     "SG01, SG02"),
    ("", "CtAp_SysStaChk", "Internal",
     "SysPwrSta\n(0: IGN_OFF\n1: IGN_ON)",
     "uint8\nRange: 0~1",
     "CORRUPT",
     "SysPwrSta가 IGN_ON↔IGN_OFF 간 비정상 전환\n(진동/노이즈로 인한 랜덤 변경)",
     "HallMGT, PositionMGT 등 전체 모듈의\n동작 조건 불안정",
     "기어 위치 감지/변속 동작 불안정\n→ 간헐적 기능 이상",
     "SG01, SG02, SG03"),

    # ═══════════════════════════════════════════════════════════════════════════
    # 5. CtAp_HallModeCntl – Hall IC Mode Control (3D Mode Write/Read)
    # ═══════════════════════════════════════════════════════════════════════════
    ("5.0", "CtAp_HallModeCntl", "External",
     "SPI_Die1/Die2\n(Hall IC SPI 통신\nRead/Write Flow)",
     "SPI Protocol\n8byte 송수신",
     "NO",
     "SPI 통신 응답 없음\n→ Hall IC 응답 타임아웃\nRead/Write 시퀀스 진행 불가",
     "HallModeSta=STD_OFF 유지\n→ HallDataSet에서 SPI GetPos 미호출",
     "Hall 센서 위치 데이터 수집 불가\n→ 기어 위치 감지 전체 실패",
     "SG01, SG02, SG03"),
    ("", "CtAp_HallModeCntl", "External", "", "", "CORRUPT",
     "SPI 통신 데이터 손상 (CRC 오류, 비트 반전)\n→ Hall IC 레지스터 설정값 이상",
     "3D 모드 설정 실패\n→ K_Alpha/K_Beta 값 오설정",
     "Hall 센서 각도 계산 기준값 오설정\n→ 위치 감지 정확도 저하",
     "SG01"),
    ("", "CtAp_HallModeCntl", "Internal",
     "HallModeSta\n(0: OFF\n1: 3D Mode Set 완료)",
     "uint8\nRange: 0~1",
     "NO",
     "3D 모드 Write/Read 시퀀스 완료 후에도\nHallModeSta=STD_OFF 유지",
     "HallDataSet: HallModeSta==OFF로 판단\n→ SPI GetPos 미호출",
     "Hall 센서 Die1/Die2 데이터 갱신 안됨\n→ 이전 값 또는 초기값 사용",
     "SG01"),
    ("", "CtAp_HallModeCntl", "Internal", "", "", "MORE",
     "3D 모드 설정이 아직 완료되지 않았는데\nHallModeSta=STD_ON 설정됨",
     "HallDataSet에서 SPI GetPos 조기 호출\n→ 부정확한 Hall 센서 데이터 수집",
     "초기화 미완료 상태에서 위치 감지 시작\n→ P/R/N/D 오판정 가능",
     "SG01, SG02"),

    # ═══════════════════════════════════════════════════════════════════════════
    # 6. CtAp_HallDataSet – Hall Sensor Data Set
    # ═══════════════════════════════════════════════════════════════════════════
    ("6.0", "CtAp_HallDataSet", "External",
     "SPI_GetPos_Die1/Die2\n(0x53 Command\nAlpha, Beta, VG 데이터)",
     "uint8[8]\nSPI Response",
     "NO",
     "SPI GetPos 명령 응답 없음\n→ Die1/Die2 Alpha, Beta, VG값 미업데이트",
     "HallFltChk 입력값 고착\n→ 이전 사이클 값 그대로 사용",
     "Hall 센서 위치 데이터 정체\n→ 위치 변화 미감지",
     "SG01, SG02"),
    ("", "CtAp_HallDataSet", "External", "", "", "CORRUPT",
     "SPI 데이터 손상으로 Alpha, Beta, VG값 이상\n(비트 오류, 통신 노이즈)",
     "HallFltChk에서 잘못된 값으로 Fault 판정\n→ 정상 센서를 Fault로 오진단",
     "HallFltSta=ON 오발생\n→ LvrPosInfo에서 위치 감지 중단",
     "SG01"),
    ("", "CtAp_HallDataSet", "Internal",
     "HallSnrDie1Data\nHallSnrDie2Data\n(Alpha, Beta, VG)",
     "HallSnrData struct",
     "CORRUPT",
     "Die1/Die2 데이터가 정상 범위 밖의 값으로\nRTE를 통해 LvrPosChk에 전달됨",
     "LvrPosChk: 잘못된 각도값으로 P/R/N/D 위치 판정\n→ 오위치 감지",
     "기어 위치 오판정으로 잘못된 변속 명령 허용\n→ 안전 위험",
     "SG02"),
    ("", "CtAp_HallDataSet", "Internal", "", "", "NO",
     "HallModeSta=OFF로 SPI GetPos 미호출\n→ Die1/Die2 데이터 미업데이트",
     "LvrPosChk에 이전 값 또는 초기값 전달",
     "위치 감지 정체로 기어 변환 미감지",
     "SG01"),

    # ═══════════════════════════════════════════════════════════════════════════
    # 7. CtAp_HallFltChk – Hall Sensor Fault Check
    # ═══════════════════════════════════════════════════════════════════════════
    ("7.0", "CtAp_HallFltChk", "Internal",
     "HallSnrDie1/Die2\n(Alpha: 0~16383\nBeta: 0~16383\nVG: 0~63)",
     "HallSnrData struct",
     "MORE",
     "Alpha 합산값 > ALPHA_FLT_VAL_HIGH(8417) 초과\n→ AlphaFlt=ON",
     "HallFltSta.TotalFlt=ON\n→ LvrPosInfo에서 위치 감지 중단",
     "Hall 센서 범위 초과 오진단\n→ 정상 위치인데 Fault 처리",
     "SG01"),
    ("", "CtAp_HallFltChk", "Internal", "", "", "LESS",
     "Alpha 합산값 < ALPHA_FLT_VAL_LOW(7964) 미만\n→ AlphaFlt=ON",
     "HallFltSta.TotalFlt=ON",
     "Hall 센서 범위 미달 오진단",
     "SG01"),
    ("", "CtAp_HallFltChk", "Internal", "", "", "NO",
     "실제 Hall 센서 Fault 상태인데\nBatStbSta=OFF로 Fault 판정 억제됨",
     "HallFltSta=OFF 유지\n→ Fault 미감지 상태로 LvrPosInfo 동작 계속",
     "고장난 Hall 센서 데이터로 위치 감지 계속\n→ 오위치 정보 CAN 송출",
     "SG01, SG02"),
    ("", "CtAp_HallFltChk", "Internal",
     "HallSnrFltSta\n(Alpha/Beta/VG/Total Fault)",
     "HallSnrFltInfo struct",
     "CORRUPT",
     "Fault 플래그가 랜덤하게 Set/Clear 반복\n(HALL_ERR_TIME 타이머 동작 이상)",
     "LvrPosInfo: TotalFlt 간헐적 발생\n→ 위치 감지 불안정",
     "기어 위치 CAN 신호 간헐적 이상\n→ TCU 오동작 유발 가능",
     "SG01, SG02"),

    # ═══════════════════════════════════════════════════════════════════════════
    # 8. CtAp_CGWSigChk – CGW CAN Signal Check
    # ═══════════════════════════════════════════════════════════════════════════
    ("8.0", "CtAp_CGWSigChk", "External",
     "BDC CAN Msg\n(BDC02: Ign1InSta\nBDC05: AutoBrightSta)",
     "CanMsg_BDC struct\n(Timeout 감지 포함)",
     "NO",
     "BDC CAN 메시지 타임아웃\n→ BDC02MsgToFlag=ON, IgnSwSta 미업데이트",
     "IgnStaChk: IgnSwSta=OFF → CAN IGN 정보 상실\nHW 전압만으로 IgnOnSta 판정",
     "BDC CAN 단절 시 IgnOnSta 불안정\n→ SysPwrSta 오판정 위험",
     ""),
    ("", "CtAp_CGWSigChk", "External", "", "", "CORRUPT",
     "BDC CAN 데이터 손상\n→ Ign1InSta 값 이상",
     "IgnSwSta 오설정\n→ IgnOnSta 오판정",
     "IGN 상태 오인식으로 SBW 비정상 동작",
     ""),
    ("", "CtAp_CGWSigChk", "External",
     "CLU CAN Msg\n(DtentOut, RhstLvlSta)",
     "CanMsg_CLU struct",
     "NO",
     "CLU CAN 메시지 타임아웃\n→ CLUMsgToFlag=ON",
     "DimLvlSet: 밝기 조절 정보 상실\n→ 인디케이터 기본 밝기 유지",
     "인디케이터 밝기 조절 불가\n(안전 영향 없음)",
     ""),
    ("", "CtAp_CGWSigChk", "External",
     "PDC CAN Msg\n(PDC_ResetPreWrng,\nPDC_ResetReq)",
     "CanMsg_PDC struct",
     "NO",
     "PDC CAN 메시지 타임아웃\n→ PDC03MsgToFlag=ON",
     "PDC Reset 요청 처리 불가",
     "P버튼 관련 PDC 리셋 동작 불가",
     "SG03"),
    ("", "CtAp_CGWSigChk", "Internal",
     "IgnSwSta\n(0: IGN_OFF\n1: IGN_ON)",
     "uint8\nRange: 0~1",
     "MORE",
     "BDC CAN Ign1InSta가 실제 OFF인데\nIgnSwSta=ON으로 고착",
     "IgnStaChk: IgnOnSta=ON 유지\n→ SysPwrSta=IGN_ON 잘못 유지",
     "IGN OFF 후 SBW 시스템 계속 활성화",
     ""),

    # ═══════════════════════════════════════════════════════════════════════════
    # 9. CtAp_ShiftActSigChk – Shift Actuator Signal Check (SCU)
    # ═══════════════════════════════════════════════════════════════════════════
    ("9.0", "CtAp_ShiftActSigChk", "External",
     "Main SCU SA CAN Signal\n(E2E CRC + AliveCount\n변속 위치 요청)",
     "uint8[6]\nE2E 보호",
     "NO",
     "Main SCU CAN SA 메시지 타임아웃\n→ MainSActMsgToFlag=ON",
     "GearPosSta 업데이트 불가\n→ 변속 위치 정보 상실",
     "TCU 변속 요청 수신 불가\n→ 기어 위치 명령 처리 중단",
     "SG01, SG02"),
    ("", "CtAp_ShiftActSigChk", "External", "", "", "CORRUPT",
     "Main SCU CAN 데이터 손상\n→ CRC 오류 발생\nCrcFltInfo.MainCrcFlt=ON",
     "TotalCrcFlt=ON\n→ 변속 신호 유효성 상실",
     "E2E 보호 위반으로 변속 명령 무시\n→ 기어 위치 변환 불가",
     "SG02"),
    ("", "CtAp_ShiftActSigChk", "External",
     "Sub SCU SA CAN Signal\n(E2E CRC + AliveCount\n변속 위치 요청 이중화)",
     "uint8[6]\nE2E 보호",
     "NO",
     "Sub SCU CAN SA 메시지 타임아웃\n→ SubSActMsgToFlag=ON",
     "Main만으로 동작 (이중화 기능 손실)",
     "Sub CAN 단일 고장 시 이중화 기능 상실\n(Main 정상이면 동작 가능)",
     ""),
    ("", "CtAp_ShiftActSigChk", "External", "", "", "CORRUPT",
     "Sub SCU CAN E2E AliveCount 불일치\n→ AlvCntFltInfo.SubAlvCntFlt=ON",
     "TotalAlvCntFlt=ON\n→ Sub 채널 유효성 상실",
     "Sub CAN E2E 실패로 이중화 기능 저하",
     ""),
    ("", "CtAp_ShiftActSigChk", "Internal",
     "GearPosSta\n(P/R/N/D 현재 위치)",
     "uint8\nRange: 0~3",
     "CORRUPT",
     "Main/Sub CAN 신호 불일치로\nGearPosSta가 실제 위치와 다른 값 출력",
     "SBWSigSet: LvrPosInfo와 GearPosSta 불일치\n→ CAN TX에 잘못된 기어 위치 송출",
     "TCU/CLU에 잘못된 기어 위치 정보 전달\n→ 오변속 위험",
     "SG02"),
    ("", "CtAp_ShiftActSigChk", "Internal",
     "GearPosTarSta\n(목표 변속 위치)",
     "uint8\nRange: 0~3",
     "CORRUPT",
     "목표 변속 위치가 실제 요청과 다른 값으로 설정",
     "잘못된 목표 위치로 기어 제어 시도",
     "의도하지 않은 기어 위치로 변환\n→ 급발진/급정거 위험",
     "SG02"),

    # ═══════════════════════════════════════════════════════════════════════════
    # 10. CtAp_CANBusOffChk – CAN Bus-Off Check
    # ═══════════════════════════════════════════════════════════════════════════
    ("10.0", "CtAp_CANBusOffChk", "External",
     "MainCAN BusOff Status\n(CAN 컨트롤러 오류 상태)",
     "uint8\nRange: 0~1",
     "MORE",
     "실제 BusOff 아닌 상태에서\nMainCANBusOffSta=ON 오진단",
     "CGWSigChk/ShiftActSigChk에서\nCAN 신호 무효 처리",
     "정상 CAN 통신 중 신호 무시\n→ BDC/SCU 메시지 처리 중단",
     "SG01, SG02"),
    ("", "CtAp_CANBusOffChk", "External", "", "", "NO",
     "실제 Main CAN BusOff 발생 시 감지 실패\n→ MainCANBusOffSta=OFF 유지",
     "CGWSigChk: BusOff 상태에서 CAN 수신 계속 시도",
     "BusOff 미감지로 오래된 CAN 데이터 계속 사용\n→ 타임아웃 처리 지연",
     ""),
    ("", "CtAp_CANBusOffChk", "Internal",
     "MainCANBusOffSta\n(0: Normal\n1: Bus-Off)",
     "uint8\nRange: 0~1",
     "MORE",
     "BusOff 아닌데 MainCANBusOffSta=ON 유지",
     "ShiftActSigChk에서 SA 신호 유효성 검증 실패\n→ 변속 신호 무시",
     "정상 상태에서 변속 불가",
     ""),

    # ═══════════════════════════════════════════════════════════════════════════
    # 11. CtAp_LvrPosChk – Lever Position Check
    # ═══════════════════════════════════════════════════════════════════════════
    ("11.0", "CtAp_LvrPosChk", "Internal",
     "HallSnrDie1/Die2Data\n(Alpha, Beta, VG\n각도 값)",
     "HallSnrData struct",
     "CORRUPT",
     "손상된 Hall 데이터로 P/R/N/D 판정 임계값과\n비교 시 오판정 발생",
     "LvrPos가 실제와 다른 위치로 출력\n→ P=R 또는 N=D 오판정",
     "TCU에 잘못된 기어 위치 정보 전달\n→ 오변속",
     "SG02"),
    ("", "CtAp_LvrPosChk", "Internal", "", "", "NO",
     "Hall 데이터가 모든 위치 임계값 밖에 있음\n→ LvrPos = NULL (미판정)",
     "LvrPosInfo: LvrPos=NULL → 위치 미결정 상태",
     "기어 위치 미감지\n→ CAN TX에 무효 위치 송출",
     "SG01"),
    ("", "CtAp_LvrPosChk", "Internal",
     "CalDataCtx\n(보정 데이터 맥락)",
     "CalDataCtx struct",
     "CORRUPT",
     "Cal 데이터 손상으로 위치 임계값 오설정",
     "모든 위치 판정 오류\n→ P/R/N/D 전체 오판정",
     "기어 위치 전체 오인식",
     "SG01, SG02"),
    ("", "CtAp_LvrPosChk", "Internal",
     "LvrPos\n(0:NULL/1:P/2:R\n3:Nr/4:N/5:Nd/6:D)",
     "uint8\nRange: 0~6",
     "CORRUPT",
     "LvrPos 출력값이 실제 위치와 다른 값으로 고착",
     "LvrPosInfo: 잘못된 LvrPos로 LvrPosInfo 생성",
     "SBWSigSet에서 잘못된 기어 위치 CAN 송출",
     "SG01, SG02"),

    # ═══════════════════════════════════════════════════════════════════════════
    # 12. CtAp_LvrPosInfo – Lever Position Information
    # ═══════════════════════════════════════════════════════════════════════════
    ("12.0", "CtAp_LvrPosInfo", "Internal",
     "HallSnrFltSta\n(TotalFlt 포함)",
     "HallSnrFltInfo struct",
     "MORE",
     "HallFltSta.TotalFlt=ON 오발생\n→ LvrPosInfo 위치 감지 비활성화",
     "LvrPosInfo 출력에 Fault 상태 반영\n→ SBWSigSet: LvrUnitSta=FAULT 송출",
     "TCU에 Hall Fault 정보 전달\n→ 변속 제한",
     "SG01"),
    ("", "CtAp_LvrPosInfo", "Internal", "", "", "NO",
     "실제 Hall TotalFlt 상태인데 HallFltSta=OFF 입력\n→ Fault 미반영",
     "LvrPosInfo: 고장 상태에서도 정상 위치 정보 출력",
     "고장 상태의 위치 정보 CAN 송출\n→ TCU 오동작",
     "SG01, SG02"),
    ("", "CtAp_LvrPosInfo", "Internal",
     "CalGapSta\n(0: Cal 무효\n1: Cal 유효)",
     "uint8\nRange: 0~1",
     "NO",
     "Cal 데이터 미완료 상태에서 CalGapSta=OFF\n→ 위치 감지 조건 미충족",
     "LvrPosDecCndChk에서 감지 조건 불만족\n→ LvrPosInfo 갱신 안됨",
     "Cal 미완료 시 위치 감지 불가\n→ 초기 기어 위치 미확인",
     "SG01"),
    ("", "CtAp_LvrPosInfo", "Internal",
     "LvrPosInfo\n(LvrPosSta, LvrUnitSta\nStuckSta 등)",
     "LvrPosInfo struct",
     "CORRUPT",
     "LvrPosInfo 구조체 일부 필드 손상\n(LvrPosSta와 LvrUnitSta 불일치)",
     "SBWSigSet: 불일치 정보 CAN TX에 반영\n→ CAN 신호 신뢰성 저하",
     "TCU에 불일치 기어 위치/상태 정보 전달",
     "SG01, SG02"),
    ("", "CtAp_LvrPosInfo", "Internal", "", "", "NO",
     "Stuck 조건 충족 시 LvrPosInfo.StuckSta 미갱신",
     "기어 Stuck 상태 미감지\n→ 강제 이탈 로직 미동작",
     "기어 레버 물리적 고착 시 비정상 상태 지속",
     "SG01"),

    # ═══════════════════════════════════════════════════════════════════════════
    # 13. CtAp_SBWSigSet – SBW CAN Signal Set (TX)
    # ═══════════════════════════════════════════════════════════════════════════
    ("13.0", "CtAp_SBWSigSet", "Internal",
     "LvrPosInfo\n(LvrPosSta, LvrUnitSta 등)",
     "LvrPosInfo struct",
     "CORRUPT",
     "LvrPosInfo 구조체 손상된 값이 입력됨\n→ CAN TX 메시지에 잘못된 기어 위치 매핑",
     "LvrMsg[8] 잘못된 값 설정\n→ TxMainCAN/TxSubCAN으로 오전송",
     "TCU/CLU에 잘못된 기어 위치 정보 전달\n→ 오변속 위험",
     "SG01, SG02"),
    ("", "CtAp_SBWSigSet", "Internal", "", "", "NO",
     "LvrPosInfo 미수신 또는 초기값 유지\n→ CAN TX에 기본값 송출",
     "기어 위치 미업데이트 상태로 CAN 송출",
     "TCU에 고착된 기어 위치 정보 전달",
     "SG01"),
    ("", "CtAp_SBWSigSet", "Internal",
     "HallSnrFltSta\n(TotalFlt)",
     "HallSnrFltInfo struct",
     "MORE",
     "HallFlt 오발생으로 LvrUnitSta=FAULT 설정\n→ CAN TX에 Fault 상태 송출",
     "TCU가 Hall Fault로 인식\n→ 변속 제한 동작",
     "정상 상태에서 변속 불가",
     ""),
    ("", "CtAp_SBWSigSet", "External",
     "LvrMsg[8]\n(SBW TX CAN Message\n기어위치, 상태, 경고)",
     "uint8[8]",
     "NO",
     "CAN TX 메시지가 송출되지 않음\n(TxMainCAN 미동작 또는 BusOff)",
     "TCU: SBW 상태 정보 수신 불가\n→ 타임아웃 처리",
     "TCU에서 SBW 고장으로 판단\n→ 변속 제한 또는 페일세이프 진입",
     "SG01, SG02"),
    ("", "CtAp_SBWSigSet", "External", "", "", "CORRUPT",
     "CAN TX 데이터 일부 비트 손상\n→ 기어 위치 또는 상태 비트 오전송",
     "TCU에서 수신 데이터 이상 감지",
     "TCU CRC/AliveCount 오류로 신호 무효 처리\n→ 변속 명령 거부",
     "SG02"),

    # ═══════════════════════════════════════════════════════════════════════════
    # 14. CtAp_IdtFltChk – Indicator Fault Check
    # ═══════════════════════════════════════════════════════════════════════════
    ("14.0", "CtAp_IdtFltChk", "External",
     "PMntrVolt / RMntrVolt\nNMntrVolt / DMntrVolt\n(P/R/N/D LED 모니터링\nADC 0~4095)",
     "uint16\nRange: 0~4095",
     "NO",
     "P/R/N/D 인디케이터 LED 단선\n→ MntrVolt=0 (오픈)",
     "IdtFltSta=ON\n→ SBWSigSet: IdtFltSta=FAULT 설정",
     "인디케이터 단선 DTC C1548 발생\n→ 기어 위치 표시 불가",
     "SG03"),
    ("", "CtAp_IdtFltChk", "External", "", "", "MORE",
     "LED 단락으로 MntrVolt 임계값 초과\n→ 과전류 상태",
     "IdtFltSta=ON\n→ DTC C1548 발생",
     "인디케이터 LED 손상\n기어 위치 표시 불가",
     "SG03"),
    ("", "CtAp_IdtFltChk", "External", "", "", "CORRUPT",
     "ADC 노이즈로 MntrVolt 값 불규칙 변동\n→ 간헐적 Fault 판정",
     "IdtFltSta 간헐적 ON/OFF",
     "인디케이터 Fault DTC 간헐적 발생\n→ 기어 위치 표시 불안정",
     "SG03"),
    ("", "CtAp_IdtFltChk", "Internal",
     "Ldo2FltSta\n(0: Normal\n1: LDO2 Fault)",
     "uint8\nRange: 0~1",
     "MORE",
     "Ldo2FltSta=ON 오입력\n→ 인디케이터 Fault 조건 충족",
     "IdtFltSta=ON 오설정",
     "정상 LDO2 상태에서 인디케이터 Fault 처리",
     ""),
    ("", "CtAp_IdtFltChk", "Internal",
     "IdtFltSta\n(0: Normal\n1: Indicator Fault)",
     "uint8\nRange: 0~1",
     "NO",
     "실제 인디케이터 Fault인데 IdtFltSta=OFF 유지",
     "SBWSigSet: IdtFltSta 반영 안됨\n→ CAN TX에 Fault 미송출",
     "인디케이터 고장 상태 미보고\n→ 운전자 기어 위치 오인식 위험",
     "SG03"),
    ("", "CtAp_IdtFltChk", "Internal",
     "IdtSta\n(0: OFF\n1: Normal ON)",
     "uint8\nRange: 0~1",
     "NO",
     "인디케이터 정상 동작 중인데 IdtSta=OFF 출력",
     "SBWSigSet: LvrIdtSta=OFF 설정\n→ CAN TX에 Indicator OFF 상태 송출",
     "TCU/CLU에서 인디케이터 미동작으로 판단",
     ""),

    # ═══════════════════════════════════════════════════════════════════════════
    # 15. CtAp_IdtCntl – Indicator Control
    # ═══════════════════════════════════════════════════════════════════════════
    ("15.0", "CtAp_IdtCntl", "Internal",
     "LvrPosInfo\n(LvrPosSta)",
     "LvrPosInfo struct",
     "CORRUPT",
     "LvrPosSta 오값 입력\n→ 잘못된 위치에 해당하는 LED 점등",
     "P 위치인데 D LED 점등 등 오표시",
     "운전자 기어 위치 오인식\n→ 오조작 위험",
     "SG03"),
    ("", "CtAp_IdtCntl", "Internal", "", "", "NO",
     "LvrPosInfo 미수신\n→ 인디케이터 기본값(OFF) 유지",
     "모든 LED 소등",
     "기어 위치 미표시\n→ 운전자 위치 확인 불가",
     "SG03"),
    ("", "CtAp_IdtCntl", "External",
     "IoHwAb LED PWM Output\n(P/R/N/D LED 출력)",
     "PWM Duty\n0~100%",
     "NO",
     "PWM 출력 신호 미발생\n→ 해당 LED 소등",
     "인디케이터 해당 위치 LED 미점등",
     "기어 위치 표시 불가",
     "SG03"),
    ("", "CtAp_IdtCntl", "External", "", "", "MORE",
     "PWM Duty 최대값 고착\n→ LED 상시 점등",
     "해당 위치 LED 항상 ON\n→ 실제 위치와 무관하게 표시",
     "기어 위치 오표시\n운전자 혼동 유발",
     "SG03"),

    # ═══════════════════════════════════════════════════════════════════════════
    # 16. CtAp_HapticCntl – Haptic Control
    # ═══════════════════════════════════════════════════════════════════════════
    ("16.0", "CtAp_HapticCntl", "Internal",
     "LvrPosInfo\n(기어 변환 이벤트)",
     "LvrPosInfo struct",
     "NO",
     "기어 위치 변환 이벤트 미감지\n→ Haptic 동작 미발생",
     "HapticCntl: 햅틱 피드백 미동작",
     "운전자 햅틱 피드백 없음\n(안전 영향 없음)",
     ""),
    ("", "CtAp_HapticCntl", "Internal", "", "", "MORE",
     "기어 변환 이벤트 오감지\n→ 불필요한 Haptic 동작",
     "HapticCntl: 불필요한 진동 발생",
     "운전자 불필요한 햅틱 피드백\n(혼동 유발 가능)",
     ""),
    ("", "CtAp_HapticCntl", "External",
     "Haptic Motor Output\n(진동 모터 PWM)",
     "PWM Output",
     "MORE",
     "Haptic Motor PWM 출력 고착(항상 ON)",
     "진동 모터 상시 동작",
     "불필요한 진동 지속\n운전자 불쾌감 및 모터 과열",
     ""),

    # ═══════════════════════════════════════════════════════════════════════════
    # 17. CtAp_ECUModeChk – ECU Mode Check
    # ═══════════════════════════════════════════════════════════════════════════
    ("17.0", "CtAp_ECUModeChk", "Internal",
     "SysPwrSta\n(0: IGN_OFF\n1: IGN_ON)",
     "uint8\nRange: 0~1",
     "NO",
     "SysPwrSta=OFF 상태에서 EcuSta=ACTIVE로\n잘못 설정됨",
     "LdoStaChk, HallFltChk에 잘못된 EcuSta 전달\n→ 비정상 동작 조건",
     "ECU 슬립 상태에서 불필요한 모듈 동작",
     ""),
    ("", "CtAp_ECUModeChk", "Internal", "", "", "MORE",
     "SysPwrSta=ON 상태에서 EcuSta=STANDBY로\n잘못 설정됨",
     "Ldo2FltChk 활성화 조건 미충족\n→ LDO2 Fault 감지 억제",
     "LDO2 Fault 미감지로 인디케이터 이상 방치",
     "SG03"),
    ("", "CtAp_ECUModeChk", "Internal",
     "EcuSta\n(STANDBY/ACTIVE 등)",
     "uint8 enum",
     "CORRUPT",
     "EcuSta 값이 정의되지 않은 열거값으로 출력",
     "LdoStaChk, HallFltChk에서 EcuSta 조건 분기 오동작",
     "여러 Fault Check 모듈의 동작 불안정",
     ""),

    # ═══════════════════════════════════════════════════════════════════════════
    # 18. CtAp_DrvRdySigChk – Drive Ready Signal Check
    # ═══════════════════════════════════════════════════════════════════════════
    ("18.0", "CtAp_DrvRdySigChk", "External",
     "EMS Drive Ready Signal\n(CAN: 엔진 구동 준비 완료)",
     "uint8\nRange: 0~1",
     "NO",
     "EMS CAN DrvRdy 신호 타임아웃\n→ DrvRdySta=OFF",
     "DrvRdySta=OFF → P→D 등 변속 허가 조건 미충족\n(일부 변속 제한 가능)",
     "엔진 구동 준비 확인 불가\n→ 특정 기어 변환 불가",
     "SG01"),
    ("", "CtAp_DrvRdySigChk", "External", "", "", "MORE",
     "EMS DrvRdy 신호가 실제 준비 안됐는데 ON 입력",
     "DrvRdySta=ON 오설정\n→ 변속 허가 조건 충족으로 오판정",
     "엔진 미준비 상태에서 기어 변환 허가\n→ 비정상 변속",
     "SG02"),

]

# ──────────────────────────────────────────────────────────────────────────────
# Excel 파일 생성 (xlsxwriter)
# ──────────────────────────────────────────────────────────────────────────────

OUT_FILE = r'E:\claude\FMEA\SBW_FMEA\LQ2\LQ2_SBW_SW_FMEA_auto.xlsx'

wb = xlsxwriter.Workbook(OUT_FILE)
ws = wb.add_worksheet("SW_FMEA")

# ── 포맷 정의 ─────────────────────────────────────────────────────────────────
def mk_fmt(bg='#FFFFFF', bold=False, font_color='#000000', wrap=True, valign='top', border=1):
    return wb.add_format({
        'bg_color': bg, 'bold': bold, 'font_color': font_color,
        'text_wrap': wrap, 'valign': valign, 'border': border,
        'font_size': 9, 'align': 'left',
    })

hdr_fmt   = mk_fmt('#4472C4', bold=True, font_color='#FFFFFF')
ext_fmt   = mk_fmt('#DDEDF7')
int_fmt   = mk_fmt('#E2EFDA')
base_fmt  = mk_fmt()

HAZOP_FMTS = {
    "MORE":       mk_fmt('#E2EFDA', bold=True),
    "LESS":       mk_fmt('#FFF2CC', bold=True),
    "NO":         mk_fmt('#FCE4D6', bold=True),
    "CORRUPT":    mk_fmt('#FFE6E6', bold=True),
    "REVERSE":    mk_fmt('#E6E0FF', bold=True),
    "AS WELL AS": mk_fmt('#E0F7FA', bold=True),
}

# ── 열 너비 / 행 높이 ─────────────────────────────────────────────────────────
col_widths = [6, 22, 14, 40, 18, 16, 55, 38, 38, 22, 5, 20, 5, 28, 20, 5, 8]
for ci, w in enumerate(col_widths):
    ws.set_column(ci, ci, w)

# ── 헤더 ──────────────────────────────────────────────────────────────────────
headers = [
    "No", "SW Unit Name", "Interface Category",
    "Interface\n(Variable) name", "Interface\n(Variable) type",
    "Failure Mode\n(HAZOP Keyword)",
    "Detail of the Failure Mode",
    "Effect on Module",
    "Effect on System",
    "Effect on SG",
    "S", "Preventive Action", "O",
    "Safety Mechanism", "Test Method",
    "D", "RPN",
]
ws.set_row(0, 40)
for ci, h in enumerate(headers):
    ws.write(0, ci, h, hdr_fmt)

ws.freeze_panes(1, 3)
ws.autofilter(0, 0, 0, len(headers) - 1)

# ── 데이터 입력 ───────────────────────────────────────────────────────────────
row = 1
for item in FMEA_ROWS:
    (no, unit, cat, intf, itype, fail, detail, eff_mod, eff_sys, eff_sg) = item

    cat_fmt = ext_fmt if cat == "External" else (int_fmt if cat == "Internal" else base_fmt)
    fail_fmt = HAZOP_FMTS.get(fail, mk_fmt(bold=True))

    vals = [no, unit, cat, intf, itype, fail, detail, eff_mod, eff_sys, eff_sg,
            "", "Design Review", "", "", "", "", ""]

    ws.set_row(row, 65)
    for ci, v in enumerate(vals):
        if ci < 3:
            ws.write(row, ci, v, cat_fmt)
        elif ci == 5:
            ws.write(row, ci, v, fail_fmt)
        else:
            ws.write(row, ci, v, base_fmt)

    # RPN 수식 (S=col K=10, O=col M=12, D=col P=15, RPN=col Q=16)
    ws.write_formula(row, 16,
        f'=IF(K{row+1}*M{row+1}*P{row+1}>0,K{row+1}*M{row+1}*P{row+1},"")',
        base_fmt)

    row += 1

wb.close()

total = row - 1
print(f"LQ2 SW FMEA 생성 완료: {OUT_FILE}")
print(f"총 {total}개 항목 ({len([r for r in FMEA_ROWS if r[0] and r[0] != ''])}개 SW Unit 그룹)")
