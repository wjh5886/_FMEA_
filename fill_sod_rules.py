"""
PLBM_SV1 FMEA S/O/D 규칙 기반 자동 채우기 (API 불필요)
AUTOSAR BSW 타입별 표준 S/O/D 정의 + 실패 모드별 보정
"""
import sys, re, requests, urllib3
urllib3.disable_warnings()
sys.stdout.reconfigure(encoding='utf-8')

URL      = "https://itzgdbeiyvodhfhmvrfw.supabase.co"
SUPA_KEY = ("eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Iml0"
            "emdkYmVpeXZvZGhmaG12cmZ3Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzY3NDE2MzcsImV4cCI"
            "6MjA5MjMxNzYzN30.iqzQr-3Lqf1O4UHFe9euTLyeIyBXreLPoSzUdtEaNP8")
H        = {"apikey": SUPA_KEY, "Authorization": f"Bearer {SUPA_KEY}", "Content-Type": "application/json"}
PROJ_ID  = "fa631b49-e52a-4834-a019-3175009b2ddf"

# ── 타입별 기본 S/O/D 및 설명 ────────────────────────────────────────
# (base_s, base_o, base_d, effect_module_ko, potential_cause_ko, preventive_ko, detection_ko)
TYPE_RULES = {
    # NvM
    "NvMService":              (6, 3, 4, "NvM 블록 읽기/쓰기 오류로 영구 데이터 손실 가능", "NvM 메모리 CRC 오류 또는 쓰기 실패", "NvM 블록 CRC 보호 및 이중화 설계", "NvM_JobResultType 반환값 모니터링"),
    "NvMNotifyJobFinished":    (5, 3, 4, "NvM 작업 완료 알림 오류로 후속 처리 누락", "NvM 작업 콜백 호출 시점 오류", "NvM 작업 완료 상태 폴링 병행", "NvM 작업 완료 타임아웃 감시"),
    # DEM
    "DiagnosticMonitor":       (5, 2, 2, "DEM 이벤트 감시 오류로 고장 진단 누락", "DEM 이벤트 상태 전이 로직 오류", "DEM 이벤트 조건 명확한 설계 및 검토", "DEM 이벤트 상태 DTC 등록으로 검출"),
    "CallbackEventStatusChange":(4, 2, 2, "DEM 이벤트 상태 변화 콜백 오류로 후속 처리 누락", "DEM 콜백 등록 오류 또는 상태 전이 누락", "DEM 이벤트 콜백 등록 검증", "DEM 이벤트 플래그 모니터링"),
    "DiagnosticInfo":          (4, 2, 2, "DEM 진단 정보 오류로 잘못된 DTC 정보 제공", "DEM 진단 정보 업데이트 오류", "DEM 진단 정보 유효성 검증", "DEM 진단 정보 일관성 체크"),
    "OperationCycle":          (4, 2, 3, "운전 사이클 관리 오류로 DEM 이벤트 리셋 누락", "운전 사이클 시작/종료 조건 오류", "운전 사이클 조건 명확한 정의", "운전 사이클 상태 감시"),
    "InitEvt":                 (4, 2, 2, "초기화 이벤트 오류로 DEM 이벤트 미등록", "DEM 초기화 이벤트 설정 오류", "초기화 순서 검증", "DEM 이벤트 상태 확인"),
    # DCM
    "DCMServices":             (3, 2, 3, "UDS 진단 서비스 응답 오류로 진단 기능 저하", "DCM 서비스 처리 로직 오류", "DCM 서비스 처리 코드 검토 및 단위 테스트", "UDS 응답 코드 확인"),
    "CallbackDCMRequestServices":(3, 2, 3, "DCM 요청 콜백 오류로 진단 서비스 처리 실패", "DCM 콜백 등록 또는 처리 로직 오류", "DCM 콜백 처리 코드 검토", "DCM 서비스 응답 모니터링"),
    "DataServices":            (3, 2, 3, "UDS DID 데이터 읽기/쓰기 오류", "DID 데이터 처리 로직 오류", "DID 처리 코드 단위 테스트", "UDS 응답 NRC 코드 확인"),
    "RoutineServices":         (3, 2, 3, "UDS 루틴 서비스 실행 오류", "루틴 서비스 처리 로직 오류", "루틴 서비스 사전/사후 조건 검증", "루틴 결과 코드 확인"),
    "CSDataServices":          (3, 2, 3, "DEM 데이터 요소 읽기 오류로 진단 정보 오류", "DEM 데이터 요소 접근 오류", "DEM 데이터 요소 유효성 검증", "DEM 데이터 요소 반환값 확인"),
    # ComM / CanSM / LinSM
    "ComMModeInterface":       (6, 3, 3, "통신 모드 전환 오류로 CAN/LIN 통신 장애", "ComM 모드 전이 조건 오류 또는 채널 상태 불일치", "ComM 모드 전이 조건 명확한 정의 및 검토", "ComM 모드 상태 감시 및 DTC 등록"),
    "ComMModeRequestInterface":(5, 2, 3, "통신 모드 요청 오류로 통신 시작/종료 실패", "ComM 모드 요청 로직 오류", "ComM 모드 요청 조건 검토", "ComM 모드 전이 타임아웃 감시"),
    "ComM_UserRequest":        (5, 2, 3, "통신 모드 사용자 요청 오류로 통신 제어 실패", "ComM 사용자 요청 처리 오류", "ComM 사용자 요청 처리 코드 검토", "ComM 모드 상태 확인"),
    "CanSMStateInterface":     (6, 3, 3, "CAN SM 상태 오류로 CAN 통신 장애 또는 버스 오프", "CanSM 상태 전이 오류 또는 하드웨어 오류", "CanSM 상태 전이 로직 검토 및 버스 오프 처리 강화", "CanSM 상태 DTC 등록 및 CAN 버스 오프 감지"),
    "CanSMBORStateInterface":  (5, 3, 3, "CAN 버스 오프 복구 상태 오류로 통신 복구 실패", "CanSM 버스 오프 복구 로직 오류", "버스 오프 복구 절차 검토", "버스 오프 복구 타임아웃 감시"),
    "LinSMStateInterface":     (5, 3, 3, "LIN SM 상태 오류로 LIN 통신 장애", "LinSM 상태 전이 오류", "LinSM 상태 전이 로직 검토", "LinSM 상태 타임아웃 감시"),
    "LinScheduleInterface":    (4, 2, 3, "LIN 스케줄 오류로 LIN 프레임 송수신 실패", "LIN 스케줄 설정 오류", "LIN 스케줄 설정 검토", "LIN 스케줄 완료 확인"),
    "LinScheduleRequestInterface":(4, 2, 3, "LIN 스케줄 요청 오류로 LIN 동작 제어 실패", "LIN 스케줄 요청 처리 오류", "LIN 스케줄 요청 처리 코드 검토", "LIN 스케줄 상태 확인"),
    # PduGroup
    "PduGroupTxInterface":     (5, 2, 3, "PDU 그룹 송신 제어 오류로 CAN 송신 중단/과다", "PDU 그룹 Enable/Disable 로직 오류", "PDU 그룹 제어 조건 검토", "PDU 그룹 상태 감시"),
    "PduGroupRxInterface":     (5, 2, 3, "PDU 그룹 수신 제어 오류로 CAN 수신 차단", "PDU 그룹 수신 설정 오류", "PDU 그룹 수신 조건 검토", "PDU 그룹 수신 상태 확인"),
    # EcuM
    "EcuModeInterface":        (6, 2, 3, "ECU 모드 전환 오류로 전원 관리 실패", "EcuM 모드 전이 조건 오류", "EcuM 모드 전이 로직 검토", "EcuM 모드 상태 감시"),
    "EcuM_StateRequest":       (6, 2, 3, "ECU 상태 요청 오류로 전원 관리 제어 실패", "EcuM 상태 요청 처리 오류", "EcuM 상태 요청 처리 코드 검토", "EcuM 상태 전이 타임아웃 감시"),
    "InitStateInterface":      (5, 2, 3, "초기화 상태 오류로 ECU 부팅 시퀀스 실패", "EcuM 초기화 상태 전이 오류", "초기화 순서 및 조건 검토", "초기화 완료 감시 타이머"),
    "WakeupEventInterface":    (5, 2, 4, "웨이크업 이벤트 오류로 ECU 불필요한 기동 또는 미기동", "EcuM 웨이크업 소스 설정 오류", "웨이크업 소스 유효성 검증 설계", "웨이크업 소스 마스크 및 DET 에러 감지"),
    # WdgM
    "WdgMCheckpointInterface": (7, 2, 2, "WdgM 체크포인트 오류로 소프트웨어 감시 실패", "태스크 실행 시간 초과 또는 체크포인트 누락", "태스크 워치독 체크포인트 설계 검토", "WdgM 로컬 상태 감시 및 시스템 리셋"),
    "WdgMGlobalStatusInterface":(7, 2, 2, "WdgM 전역 상태 오류로 시스템 감시 실패", "WdgM 전역 상태 전이 오류", "WdgM 전역 상태 전이 조건 검토", "WdgM 전역 상태 DTC 등록"),
    # IoHwAb
    "IoHwAb_If_AnaInDir":      (6, 2, 3, "ADC 아날로그 입력 오류로 센서 측정값 부정확", "ADC 채널 노이즈 또는 단선/단락", "ADC 입력 범위 검사 및 필터링 설계", "ADC 범위 체크 및 소프트웨어 진단"),
    "IoHwAb_If_DigDir":        (5, 2, 3, "디지털 I/O 오류로 HW 제어 신호 오동작", "GPIO 하드웨어 오류 또는 설정 오류", "GPIO 설정 검토 및 Pull-up/down 설계", "디지털 I/O 상태 읽기 및 진단"),
    "IoHwAb_If_Pwm":           (5, 2, 3, "PWM 출력 오류로 액추에이터 제어 실패", "PWM 타이머 설정 오류 또는 하드웨어 오류", "PWM 듀티/주파수 범위 검증 설계", "PWM 출력 피드백 모니터링"),
    # CAN 그룹 메시지
    "Gr_MsgGr_E2E":            (6, 3, 3, "CAN 메시지 수신 오류로 외부 ECU 정보 손실", "CAN 버스 오류 또는 송신 ECU 오류", "E2E 프로파일 적용 및 수신 타임아웃 설정", "E2E CRC/카운터 에러 및 수신 타임아웃 감시"),
    "Gr_MsgGr":                (5, 3, 4, "CAN 메시지 수신 오류로 외부 ECU 정보 손실", "CAN 버스 오류 또는 송신 ECU 오류", "수신 타임아웃 설정 및 유효성 검사", "수신 타임아웃 및 신호 유효성 감시"),
    # SPI
    "Cdd_If_Spi":              (5, 2, 3, "SPI 통신 오류로 외부 디바이스 접근 실패", "SPI 버스 오류 또는 CS 신호 오류", "SPI 통신 오류 감지 로직 설계", "SPI 전송 완료 및 에러 플래그 확인"),
    # FoD
    "FoD_Service_Interface":   (4, 2, 3, "FoD 서비스 오류로 기능 잠금/해제 제어 실패", "FoD 인증 또는 서비스 처리 오류", "FoD 서비스 인증 로직 검토", "FoD 서비스 결과 코드 확인"),
    # Csm
    "CsmJob":                  (5, 2, 3, "암호화 작업 오류로 보안 기능 실패", "Csm 잡 설정 오류 또는 키 유효성 오류", "Csm 잡 파라미터 유효성 검증", "Csm 잡 결과 코드 확인"),
    "CsmKeyInterface":         (4, 2, 3, "암호화 키 관리 오류로 보안 기능 저하", "Csm 키 설정 오류", "Csm 키 관리 절차 검토", "Csm 키 상태 확인"),
    # Callbacks
    "CallbackAfterSchedule":   (3, 2, 3, "스케줄러 후처리 콜백 오류로 주기적 처리 누락", "스케줄러 콜백 등록 오류", "콜백 등록 검증", "콜백 실행 완료 모니터링"),
    "CallbackClearEventAllowed":(4, 2, 2, "DTC 클리어 허용 콜백 오류로 DTC 관리 오류", "DTC 클리어 조건 판단 로직 오류", "DTC 클리어 조건 검토", "DTC 클리어 결과 확인"),
    "CallbackError":           (4, 2, 2, "오류 콜백 처리 오류로 오류 전파 누락", "오류 콜백 등록 또는 처리 오류", "오류 콜백 처리 로직 검토", "오류 상태 감시"),
    # DET
    "DETServiceInterface":     (3, 2, 2, "DET 에러 리포트 오류로 개발 오류 감지 실패", "DET 설정 오류", "DET 에러 코드 정의 검토", "DET 에러 로그 확인"),
    # RomTst
    "RomTst":                  (6, 2, 2, "ROM 무결성 검사 실패로 SW 코드 오류 미감지", "플래시 메모리 비트 반전 또는 검사 로직 오류", "CRC 기반 ROM 무결성 검사 설계", "ROM 테스트 결과 DTC 등록"),
}

# ── 실패 모드별 S/O/D 보정 ────────────────────────────────────────────
MODE_DELTA = {
    # (delta_s, delta_o, delta_d, detail_suffix)
    "CORRUPT": ( 1,  0,  0, "잘못된 값으로 데이터 변조"),
    "MORE":    ( 1,  0,  0, "허용 상한값 초과"),
    "LESS":    ( 0,  0,  0, "허용 하한값 미달"),
    "STUCK":   ( 0,  0,  2, "이전 값에 고착됨"),
    "ERRATIC": ( 0,  0,  1, "비정상적인 간헐적 변동"),
    "EARLY":   ( 0,  0,  0, "예상보다 이른 시점에 발생"),
    "LATE":    ( 0,  0,  1, "예상보다 늦은 시점에 발생"),
    "NO":      ( 1,  0,  0, "신호/기능이 전혀 발생하지 않음"),
    "AS_WELL_AS":( 0, 0,  0, "불필요한 추가 동작 발생"),
    "PART_OF": ( 0,  0,  0, "부분적으로만 동작"),
}

def clamp(v, lo=1, hi=10):
    return max(lo, min(hi, v))

def get_base_rule(vtype: str):
    """variable_type에서 base rule 매핑"""
    if not vtype:
        return None
    # 정확한 매칭 우선
    for key, rule in TYPE_RULES.items():
        if vtype == key:
            return rule
    # prefix 매칭
    for key, rule in TYPE_RULES.items():
        if vtype.startswith(key):
            return rule
    # 부분 포함 매칭
    for key, rule in TYPE_RULES.items():
        if key in vtype:
            return rule
    return None

def make_payload(vtype: str, vname: str, category: str, mode: str, signal_range: str | None) -> dict | None:
    rule = get_base_rule(vtype)
    if rule is None:
        return None

    base_s, base_o, base_d, effect_mod, cause, preventive, detection = rule
    delta = MODE_DELTA.get(mode, (0, 0, 0, mode.lower() + " 오류"))
    ds, do, dd, detail_suffix = delta

    s = clamp(base_s + ds)
    o = clamp(base_o + do)
    d = clamp(base_d + dd)

    cat_ko = "외부 입력 신호" if category == "External" else "내부 출력 신호"
    range_str = f" (범위: {signal_range})" if signal_range else ""

    failure_detail = f"{cat_ko} '{vtype}'{range_str} — {detail_suffix}"
    effect_system_map = {
        "NvMService": "배터리 캘리브레이션 데이터 손실로 충전/방전 제어 오류 가능",
        "NvMNotifyJobFinished": "NvM 데이터 영속성 불보장으로 설정 손실 가능",
        "DiagnosticMonitor": "배터리 이상 상태 DTC 미기록으로 진단 누락",
        "CallbackEventStatusChange": "고장 이벤트 전파 누락으로 상위 제어기 오판단",
        "DCMServices": "UDS 진단 통신 오류로 정비 기능 저하",
        "ComMModeInterface": "CAN 통신 모드 오류로 배터리 데이터 송수신 장애",
        "CanSMStateInterface": "CAN 버스 오프로 배터리 상태 정보 차량 네트워크 전달 불가",
        "IoHwAb_If_AnaInDir": "센서 측정값 오류로 배터리 전압/온도 모니터링 실패",
        "Gr_MsgGr_E2E": "외부 ECU 데이터 수신 오류로 배터리 제어 판단 오류 가능",
        "WdgMCheckpointInterface": "소프트웨어 감시 기능 상실로 오류 상태에서 ECU 미리셋",
    }
    # 기본 effect_system: 타입에서 유추
    effect_sys = None
    for key, es in effect_system_map.items():
        if vtype.startswith(key) or key in vtype:
            effect_sys = es
            break
    if not effect_sys:
        effect_sys = f"배터리 관리 시스템 기능 저하 — {effect_mod}"

    return {
        "failure_detail":    failure_detail,
        "effect_module":     effect_mod,
        "potential_cause":   cause,
        "effect_system":     effect_sys,
        "severity":          s,
        "occurrence":        o,
        "detection":         d,
        "preventive_action": preventive,
        "detection_action":  detection,
        "ai_generated":      True,
        "status":            "in_review",
    }

# ── 1. FMEA 항목 로드 ─────────────────────────────────────────────
print("FMEA 항목 로드 중...")
all_items = []
offset = 0
while True:
    r = requests.get(
        f"{URL}/rest/v1/fmea_items"
        f"?project_id=eq.{PROJ_ID}&severity=is.null"
        f"&select=id,category,variable_name,variable_type,failure_mode,signal_range"
        f"&offset={offset}&limit=1000",
        headers=H, verify=False,
    )
    batch = r.json()
    if not batch: break
    all_items.extend(batch)
    offset += len(batch)
    if len(batch) < 1000: break

print(f"대상 항목: {len(all_items)}개")

# ── 2. 매핑 및 업데이트 ──────────────────────────────────────────────
ok = 0
skipped = 0
skipped_types = {}

for i, item in enumerate(all_items):
    vtype = item.get("variable_type") or ""
    vname = item.get("variable_name") or ""
    mode  = item.get("failure_mode") or ""
    cat   = item.get("category") or ""
    rng   = item.get("signal_range")

    payload = make_payload(vtype, vname, cat, mode, rng)
    if payload is None:
        skipped += 1
        skipped_types[vtype] = skipped_types.get(vtype, 0) + 1
        continue

    r2 = requests.patch(
        f"{URL}/rest/v1/fmea_items?id=eq.{item['id']}",
        json=payload, headers=H, verify=False,
    )
    if r2.status_code < 300:
        ok += 1
    else:
        print(f"  ⚠ DB 오류: {r2.text[:60]}")

    if (i + 1) % 200 == 0:
        print(f"  진행: {i+1}/{len(all_items)} ✓{ok} skip{skipped}")

print(f"\n✓ 완료: {ok}개 업데이트")
if skipped:
    print(f"  매핑 없음: {skipped}개")
    for vt, cnt in sorted(skipped_types.items(), key=lambda x: -x[1])[:15]:
        print(f"    [{cnt:3d}개] {vt}")
