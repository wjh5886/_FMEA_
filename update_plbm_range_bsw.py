"""
PLBM_SV1 BSW/소프트웨어 인터페이스 signal_range 채우기
- variable_type 기반 AUTOSAR 표준 범위 매핑
- Gr_MsgGr_E2E_FD_LOCAL_CAN_*_<period>ms → DBC 재매핑 (주기 접미사 제거)
- IoHwAb / DataServices / DCM / DEM / NvM / ComM 등 BSW 전체 커버
"""
import re, sys, requests, urllib3
import lxml.etree as ET
from collections import defaultdict

urllib3.disable_warnings()
sys.stdout.reconfigure(encoding='utf-8')

URL      = "https://itzgdbeiyvodhfhmvrfw.supabase.co"
KEY      = ("eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Iml0"
            "emdkYmVpeXZvZGhmaG12cmZ3Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzY3NDE2MzcsImV4cCI"
            "6MjA5MjMxNzYzN30.iqzQr-3Lqf1O4UHFe9euTLyeIyBXreLPoSzUdtEaNP8")
H        = {"apikey": KEY, "Authorization": f"Bearer {KEY}", "Content-Type": "application/json"}
PROJ_ID  = "fa631b49-e52a-4834-a019-3175009b2ddf"
DBC_FILE = "E:/claude/FMEA/PLBM_SV_V01_BSWv36-4/References/DB/20250528_STD_LOCAL_PDC_2023_FD_Local_v25.05.01.dbc"

# ── 1. DBC 재파싱 (메시지명 → range 문자열, 주기 접미사 매핑용) ─────────
SG_RE = re.compile(
    r'^\s*SG_\s+(\S+)\s*:.*?\(([^,]+),([^)]+)\)\s*\[([^\|]+)\|([^\]]+)\]\s*"([^"]*)"'
)
BO_RE = re.compile(r'^BO_\s+\d+\s+(\S+)\s*:')

def fmt(v): return str(int(v)) if v == int(v) else f"{v:.4g}"
def range_str(lo, hi, unit=""):
    s = f"{fmt(lo)} ~ {fmt(hi)}"
    return f"{s} {unit}".strip() if unit else s

msg_signals = defaultdict(list)
cur_msg = None
for line in open(DBC_FILE, encoding='utf-8', errors='replace'):
    bm = BO_RE.match(line)
    if bm: cur_msg = bm.group(1).rstrip(':'); continue
    if cur_msg:
        m = SG_RE.match(line)
        if m:
            name, factor, offset, rmin, rmax, unit = m.groups()
            try: f, o, lo, hi = float(factor), float(offset), float(rmin), float(rmax)
            except: continue
            if not (lo == 0 and hi == 0):
                pl, ph = lo * f + o, hi * f + o
                if pl != ph:
                    msg_signals[cur_msg].append(f"{name}: {range_str(pl, ph, unit.strip())}")

print(f"DBC 메시지: {len(msg_signals)}개 (범위 있는 것)")

# ── 2. variable_type → signal_range 매핑 테이블 ─────────────────────────

# 2-A. AUTOSAR 타입별 표준 범위
TYPE_RANGE = {
    # NvM
    "NvMNotifyJobFinished":    "E_OK(0) / E_NOT_OK(1)",
    "NvMService":              "ReadBlock / WriteBlock / EraseNvBlock / InvalidateNvBlock",
    # DEM
    "DiagnosticMonitor":       "E_OK(0) / E_NOT_OK(1)",
    "DiagnosticInfo":          "E_OK(0) / E_NOT_OK(1)",
    "CallbackEventStatusChange": (
        "DEM_EVENT_STATUS_PASSED(0x00) / DEM_EVENT_STATUS_FAILED(0x01) / "
        "DEM_EVENT_STATUS_PREPASSED(0x02) / DEM_EVENT_STATUS_PREFAILED(0x03)"),
    "CallbackInitMonitorForEvent": "E_OK(0) / E_NOT_OK(1)",
    "CallbackClearEventAllowed":   "E_OK(0) / E_NOT_OK(1)",
    "OperationCycle":          "DEM_CYCLE_STATE_START(0) / DEM_CYCLE_STATE_END(1)",
    "EnableCondition":         "ENABLE(0) / DISABLE(1)",
    # DCM
    "DCMServices":             "E_OK(0) / E_NOT_OK(1)",
    "CallbackDCMRequestServices": "E_OK(0) / E_NOT_OK(1)",
    "ServiceRequestNotification": "E_OK(0) / E_NOT_OK(1) / DCM_E_PENDING(10)",
    "SecurityAccess_L9":       "E_OK(0) / E_NOT_OK(1) / DCM_E_COMPARE_KEY_FAILED(11)",
    "DcmDiagnosticSessionControl": (
        "DCM_DEFAULT_SESSION / DCM_PROGRAMMING_SESSION / "
        "DCM_EXTENDED_DIAGNOSTIC_SESSION / DCM_SAFETY_SYSTEM_DIAGNOSTIC_SESSION"),
    "DcmControlDTCSetting":    "ENABLEDTCSETTING / DISABLEDTCSETTING",
    "DcmModeRapidPowerShutDown": "ENABLE_RAPIDPOWERSHUTDOWN / DISABLE_RAPIDPOWERSHUTDOWN",
    "DcmApplicationUpdated":   "APP_NOT_UPDATED / APP_UPDATED",
    "DcmCommunicationControl_0": (
        "DCM_ENABLE_RX_TX_NORM / DCM_ENABLE_RX_DISABLE_TX_NORM / "
        "DCM_DISABLE_RX_ENABLE_TX_NORM / DCM_DISABLE_RX_TX_NORMAL"),
    "RoutineServices":         "E_OK(0) / E_NOT_OK(1) / DCM_E_PENDING(10)",
    # Csm
    "CsmRandomSeed":           "CRYPTO_E_OK(0) / CRYPTO_E_BUSY(2) / CRYPTO_E_KEY_NOT_VALID(9)",
    "CsmRandomGenerate":       "CRYPTO_E_OK(0) / CRYPTO_E_BUSY(2) / CRYPTO_E_SMALL_BUFFER(3)",
    "CsmHash":                 "CRYPTO_E_OK(0) / CRYPTO_E_BUSY(2) / CRYPTO_E_SMALL_BUFFER(3)",
    "CsmCallback":             "CRYPTO_E_OK(0) / CRYPTO_E_BUSY(2) / CRYPTO_E_SMALL_BUFFER(3)",
    # WdgM
    "WdgM_AliveSupervision":   "E_OK(0) / E_NOT_OK(1)",
    "WdgM_API":                (
        "WDGM_SUPERVISION_OK(0) / WDGM_SUPERVISION_FAILED(1) / "
        "WDGM_SUPERVISION_EXPIRED(2) / WDGM_SUPERVISION_STOPPED(3)"),
    "WdgM_GlobalMode":         (
        "WDGM_GLOBAL_STATUS_OK / WDGM_GLOBAL_STATUS_FAILED / "
        "WDGM_GLOBAL_STATUS_EXPIRED / WDGM_GLOBAL_STATUS_STOPPED"),
    # EcuM
    "EcuM_StateRequest":       (
        "ECUM_STATE_APP_RUN(0x32) / ECUM_STATE_APP_POST_RUN(0x33) / "
        "ECUM_STATE_SHUTDOWN(0x40) / ECUM_STATE_SLEEP(0x50)"),
    "EcuModeInterface":        "RUN / POST_RUN / SLEEP / WAKE_SLEEP",
    # ComM
    "ComM_UserRequest":        "COMM_NO_COMMUNICATION(0) / COMM_FULL_COMMUNICATION(2)",
    # CanSM / LinSM
    "CanSMBORState":           "COMPLETE / START",
    "CanSMState":              "NO_COM / SILENT_COM / FULL_COM / BUS_OFF / CHANGE_BAUDRATE",
    "LinSMState":              "NO_COM / FULL_COM",
    # PduGroup / Schedule / Wakeup / Init
    "PduGroup":                "STOP / START",
    "WakeupEvent":             "POWER / RESET / INTERNAL_RESET / INTERNAL_WDG / EXTERNAL_WDG / GPT",
    "InitState":               "START / COMPLETE",
    # IoHwAb (접두어 매칭)
    "IoHwAb_If_AnaInDir":      "0 ~ 65535 (ADC raw)",
    "IoHwAb_If_DigDir":        "0 / 1 (digital)",
    "IoHwAb_If_Pwm":           "0 ~ 100 %",
    # SPI / DET / FoD / RomTst / LIN callback
    "Cdd_If_Spi":              "0 ~ 255 (byte, SPI raw)",
    "DETService":              "E_OK(0) / E_NOT_OK(1)",
    "FoD_Service_Interface":   "E_OK(0) / E_NOT_OK(1)",
    "RomTstService":           "E_OK(0) / E_NOT_OK(1)",
    "CallbackAfterSchedule":   "E_OK(0) / E_NOT_OK(1)",
    "CallbackError":           "E_OK(0) / E_NOT_OK(1)",
}

# 2-B. DataServices 물리 범위 (variable_type 기반)
DS_RANGE = {
    "DataServices_X_12V_Lithium_Battery_Voltage":
        "0 ~ 18000 mV (uint16, 12V 리튬 배터리 전압)",
    "DataServices_X_12V_Lithium_Battery_Current":
        "-3276.8 ~ 3276.7 A (sint16 × 0.1, 충/방전 전류)",
    "DataServices_X_12V_Lithium_Battery_SOC":
        "0 ~ 100 % (uint8, State of Charge)",
    "DataServices_X_12V_Lithium_Battery_Used_Time":
        "0 ~ 65535 h (uint16, 누적 사용 시간)",
    "DataServices_X_12V_Lithium_Battery_Integrated_Charging_Capacity":
        "0 ~ 65535 Ah (uint16, 누적 충전 용량)",
    "DataServices_X_12V_Lithium_Battery_Integrated_Discharging_Capacity":
        "0 ~ 65535 Ah (uint16, 누적 방전 용량)",
    "DataServices_Cell_1_Battery_Voltage":
        "0 ~ 4200 mV (uint16, Cell 1 전압)",
    "DataServices_Cell_2_Battery_Voltage":
        "0 ~ 4200 mV (uint16, Cell 2 전압)",
    "DataServices_Cell_3_Battery_Voltage":
        "0 ~ 4200 mV (uint16, Cell 3 전압)",
    "DataServices_B_Voltage":
        "0 ~ 18000 mV (uint16, B+ 단자 전압)",
    "DataServices_IGN2_Voltage":
        "0 ~ 18000 mV (uint16, IGN2 전압)",
    "DataServices_Cell_1_Temperature_Sensor1":
        "-40 ~ 125 °C (sint8 + 40 offset, Cell 1 온도)",
    "DataServices_Cell_2_Temperature_Sensor2":
        "-40 ~ 125 °C (sint8 + 40 offset, Cell 2 온도)",
    "DataServices_BMS_Circuit_Temperature_DC_DC_Converter":
        "-40 ~ 125 °C (sint8 + 40 offset, DC-DC 회로 온도)",
    "DataServices_IPS_1_Current_Built_in_CAM":
        "0 ~ 65535 mA (uint16, IPS 1 출력 전류)",
    "DataServices_IPS_2_Current_ACU":
        "0 ~ 65535 mA (uint16, IPS 2 출력 전류)",
    "DataServices_SOH":
        "0 ~ 100 % (uint8, State of Health)",
    "DataServices_LBM_Capacity":
        "0 ~ 65535 mAh (uint16, 배터리 용량)",
    "DataServices_DTC_Count_Time":
        "0 ~ 65535 (uint16, DTC 발생 횟수 또는 시간)",
    "DataServices_FET_Status_Read":
        "0x00 ~ 0xFF (uint8 bit flags, FET 스위치 상태)",
    "DataServices_Converter_Mode_State":
        "0 ~ 255 (uint8 enum, DC-DC 컨버터 동작 모드)",
    "DataServices_Built_in_CAM_Mode_State":
        "0 ~ 255 (uint8 enum, 내장 CAM 동작 모드)",
    "DataServices_DC_DC_Converter":
        "0 ~ 65535 (uint16, DC-DC 컨버터 데이터)",
    "DataServices_DC_DC_Converter_For_FCT":
        "0 ~ 65535 (uint16, FCT용 DC-DC 데이터)",
    "DataServices_IPS_1st_Discharging":
        "0 ~ 65535 mA (uint16, IPS 1차 방전 전류)",
    "DataServices_IPS_2nd_Discharging":
        "0 ~ 65535 mA (uint16, IPS 2차 방전 전류)",
    "DataServices_Cell_1_Balancing_FET":
        "0 / 1 (boolean, Cell 1 밸런싱 FET 상태)",
    "DataServices_Cell_2_Balancing_FET":
        "0 / 1 (boolean, Cell 2 밸런싱 FET 상태)",
    "DataServices_Cell_3_Balancing_FET":
        "0 / 1 (boolean, Cell 3 밸런싱 FET 상태)",
    "DataServices_IVD":
        "0 ~ 65535 (uint16, IVD 데이터)",
    "DataServices_Data_0xFD05h_B2BVoltageMonitor":
        "0 ~ 18000 mV (uint16, B2B 전압 모니터)",
    "DataServices_Data_0xFD06h_DCDCConverterDriverADC":
        "0 ~ 65535 (uint16, DC-DC 드라이버 ADC raw)",
    "DataServices_Forced_Recalibration_ON":
        "0 / 1 (boolean, 강제 재보정 활성화)",
    "DataServices_Data_0xFD07h_SBCDevelopmentMode":
        "0 / 1 (boolean, SBC 개발 모드)",
    # 문자열/식별자 (ASCII string)
    "DataServices_VehicleManufacturerECUSoftwareVersionNumber":
        "ASCII string (20 bytes, SW 버전 번호)",
    "DataServices_SystemSupplierECUHardwareVersionNumber":
        "ASCII string (11 bytes, HW 버전)",
    "DataServices_SystemSupplierECUSoftwareNumber":
        "ASCII string (11 bytes, SW 번호)",
    "DataServices_System_Name":
        "ASCII string (시스템 이름)",
    "DataServices_VehicleManufacturerSparePartNumber":
        "ASCII string (11 bytes, 부품 번호)",
    "DataServices_ECUSerialNumber":
        "ASCII string (16 bytes, ECU 시리얼 번호)",
    "DataServices_ECU_Manufacturing_Date":
        "YYYYMMDD (4 bytes BCD, 제조일자)",
    "DataServices_ECU_Software_UNIT_number_Data_Identifie":
        "ASCII string (SW 유닛 번호)",
    "DataServices_DID_F1B1h_Data":
        "byte array (DID F1B1h raw data)",
    "DataServices_CAN_Data_Base_Version_Number_for_Local_CAN":
        "ASCII string (Local CAN DB 버전)",
    "DataServices_CAN_Data_Base_Version_Number_for_LIN":
        "ASCII string (LIN DB 버전)",
    "DataServices_ECUSupplierCodeDataIdentifier":
        "ASCII string (공급업체 코드)",
    "DataServices_RXSWIN":
        "ASCII string (규제 확장 SW 식별자)",
    "DataServices_RXSWIN_VehicleName":
        "ASCII string (RXSWIN 차종명)",
    "DataServices_VehicleManufacturer_ECUHWVersion":
        "ASCII string (완성차 HW 버전)",
    "DataServices_VehicleManufacturer_ECUSWVersion":
        "ASCII string (완성차 SW 버전)",
    "DataServices_SystemSupplier_ECUSoftwareVersionNumber":
        "ASCII string (공급업체 SW 버전)",
    "DataServices_DIDFD70_Data":
        "byte array (DID FD70h raw data)",
    "DataServices_Data_0xFD01h_ASSYPartNumber_SL":
        "ASCII string (ASSY 부품 번호)",
    "DataServices_Data_0xFD02h_ASSYproductionDateAndSequenceProductNumber_SL":
        "BCD date + seq (ASSY 생산 일련 번호)",
    "DataServices_Data_0xFD03h_PCBproductionDateAndSequence_JinYoon":
        "BCD date + seq (PCB 생산 일련 번호)",
    "DataServices_Data_0xFD04h_PCBunitNumber_Yunjin":
        "ASCII string (PCB 유닛 번호)",
}

# ── 3. variable_type → range 결정 함수 ────────────────────────────────────
PERIOD_RE = re.compile(r'_\d+ms$', re.IGNORECASE)

def get_range(vname: str, vtype: str) -> str | None:
    # (A) DataServices 물리 범위
    if vtype in DS_RANGE:
        return DS_RANGE[vtype]

    # (B) DataServices_LogFuncData_* → raw byte array
    if vtype.startswith("DataServices_LogFuncData"):
        return "byte array (raw log data, 64 bytes)"

    # (C) CSDataServices_DemDataElementClass_* → uint8
    if vtype.startswith("CSDataServices_DemDataElementClass"):
        return "0 ~ 255 (uint8, DEM data element)"

    # (D) CSDataServices_DE_* → uint16
    if vtype.startswith("CSDataServices_DE"):
        return "0 ~ 65535 (uint16, DEM data element)"

    # (E) RoutineServices_* → DCM routine result
    if vtype.startswith("RoutineServices"):
        return "E_OK(0) / E_NOT_OK(1) / DCM_E_PENDING(10)"

    # (F) IoHwAb 접두어 매칭
    for prefix, rng in [("IoHwAb_If_AnaInDir", "0 ~ 65535 (ADC raw)"),
                        ("IoHwAb_If_DigDir",   "0 / 1 (digital)"),
                        ("IoHwAb_If_Pwm",      "0 ~ 100 %")]:
        if vtype.startswith(prefix) or vname.startswith(prefix):
            return rng

    # (G) ComM mode 인터페이스 (채널별 접미사 붙음)
    if vtype.startswith("ComMModeInterface"):
        if "PNC" in vtype:
            return ("PNC_REQUESTED / PNC_READY_SLEEP / PNC_PREPARE_SLEEP / "
                    "PNC_NO_COMMUNICATION / PNC_FULL_COMMUNICATION")
        if "LIN" in vtype or "ICULIN" in vtype:
            return "NO_COM / FULL_COM"
        return "NO_COM / SILENT_COM / FULL_COM"
    if vtype.startswith("ComMModeRequestInterface"):
        if "PNC" in vtype: return "PNC_REQUESTED / PNC_NO_COMMUNICATION"
        return "NO_COM / SILENT_COM / FULL_COM"

    # (H) CanSM 인터페이스
    if vtype.startswith("CanSMStateInterface"):
        return "NO_COM / SILENT_COM / FULL_COM / BUS_OFF / CHANGE_BAUDRATE"
    if vtype.startswith("CanSMBORStateInterface"):
        return "COMPLETE / START"

    # (I) LinSM / LinSchedule 인터페이스
    if vtype.startswith("LinSMStateInterface"):
        return "NO_COM / FULL_COM"
    if vtype.startswith("LinScheduleInterface") or vtype.startswith("LinScheduleRequestInterface"):
        return "NULL / Always / DiagnoseReq / DiagnoseResp"

    # (J) PduGroup 인터페이스
    if vtype.startswith("PduGroupTxInterface") or vtype.startswith("PduGroupRxInterface"):
        return "STOP / START"

    # (K) Gr_MsgGr_E2E_FD_LOCAL_CAN_*_<period>ms → DBC (주기 접미사 제거)
    if vname.startswith("Gr_MsgGr_E2E_FD_LOCAL_CAN_") or vname.startswith("Gr_MsgGr_FD_LOCAL_CAN_"):
        for prefix in ["Gr_MsgGr_E2E_FD_LOCAL_CAN_", "Gr_MsgGr_FD_LOCAL_CAN_"]:
            if vname.startswith(prefix):
                msg = vname[len(prefix):]
                msg_stripped = PERIOD_RE.sub('', msg)   # _10ms 등 제거
                parts = msg_signals.get(msg_stripped) or msg_signals.get(msg)
                if parts:
                    return " | ".join(parts)
        return None

    # (L) 직접 타입 테이블 조회
    if vtype in TYPE_RANGE:
        return TYPE_RANGE[vtype]

    # (M) Cdd_If_Spi 접두어
    if vtype.startswith("Cdd_If_Spi"):
        return "0 ~ 255 (byte, SPI raw)"

    return None

# ── 4. Supabase에서 NULL signal_range 항목 로드 ───────────────────────────
print("미채운 FMEA 항목 로드 중...")
all_missing = []
offset = 0
while True:
    r = requests.get(
        f"{URL}/rest/v1/fmea_items?project_id=eq.{PROJ_ID}"
        f"&signal_range=is.null&select=id,variable_name,variable_type"
        f"&offset={offset}&limit=1000",
        headers=H, verify=False,
    )
    batch = r.json()
    if not batch: break
    all_missing.extend(batch)
    offset += len(batch)
    if len(batch) < 1000: break

print(f"미채운 항목: {len(all_missing)}개")

# ── 5. 범위 결정 ─────────────────────────────────────────────────────────
updates = []
skipped_types = defaultdict(int)

for item in all_missing:
    vname = item.get("variable_name") or ""
    vtype = item.get("variable_type") or ""
    rng = get_range(vname, vtype)
    if rng:
        updates.append({"id": item["id"], "signal_range": rng})
    else:
        skipped_types[vtype or "(없음)"] += 1

print(f"\n업데이트 가능: {len(updates)}개 / 전체 미채운 {len(all_missing)}개")
if skipped_types:
    print(f"매핑 없음: {sum(skipped_types.values())}개")
    for t, cnt in sorted(skipped_types.items(), key=lambda x: -x[1])[:10]:
        print(f"  [{cnt:4d}개] {t}")

# ── 6. 샘플 출력 ─────────────────────────────────────────────────────────
print("\n[매핑 샘플 15개]")
shown = set()
for u in updates:
    item = next(i for i in all_missing if i["id"] == u["id"])
    vn = item["variable_name"]
    if vn not in shown:
        shown.add(vn)
        rng_preview = u["signal_range"][:65]
        print(f"  {vn:50s} → {rng_preview}")
    if len(shown) >= 15: break

if not updates:
    print("업데이트할 항목이 없습니다.")
    import sys; sys.exit(0)

# ── 7. 배치 업데이트 ─────────────────────────────────────────────────────
print(f"\n업데이트 시작 ({len(updates)}개)...")
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
print(f"  남은 미채운 항목: {len(all_missing) - ok}개")
