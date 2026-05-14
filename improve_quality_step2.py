"""
improve_quality_step2.py — 규칙 기반 템플릿 개선 (2단계)
신호명 키워드 × failure_mode × category 조합으로 구체적인 텍스트 생성

사용법:
  python improve_quality_step2.py
  python improve_quality_step2.py "LQ2"
"""

import sys, re, requests, urllib3, concurrent.futures
urllib3.disable_warnings()

SB_URL = "https://itzgdbeiyvodhfhmvrfw.supabase.co"
SB_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Iml0emdkYmVpeXZvZGhmaG12cmZ3Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzY3NDE2MzcsImV4cCI6MjA5MjMxNzYzN30.iqzQr-3Lqf1O4UHFe9euTLyeIyBXreLPoSzUdtEaNP8"
SB_H = {
    "apikey": SB_KEY,
    "Authorization": f"Bearer {SB_KEY}",
    "Content-Type": "application/json",
    "Prefer": "return=minimal",
}

TARGET_PROJECTS = [
    ("32cc5148-96b5-42c5-a3dc-2b8fa3b9f93e", "GN7_FL"),
    ("89dc5818-2435-4d09-a1a9-36aea664d11d", "LQ2"),
    ("70f74c19-66b6-4b2e-a2b3-1ee04dd1b101", "TK1"),
    ("a43d7d4e-b104-4e90-ab55-04c7b31aa3e7", "SX3"),
]

GENERIC_PATTERNS = ["설계 리뷰 및 코드 인스펙션", "리뷰 및 코드 인스펙션", "코드 인스펙션 수행"]

# ── 신호 카테고리 키워드 ──────────────────────────────────────────────────────
SIG_CATEGORIES = [
    ("GEAR",    ("gearpos", "shiftpos", "lvrpos", "parkpos", "shiftact", "lvr", "gear", "shift", "park")),
    ("POWER",   ("volt", "vbat", "batvolt", "ldovolt", "igvolt", "ignvolt", "power", "ldo", "bat", "vcc")),
    ("FAULT",   ("flt", "fault", "err", "diag", "dem", "idtflt", "dtc", "warn")),
    ("STATUS",  ("sta", "status", "mode", "ect", "ecu", "drv", "drvr", "idtsta", "ecusta", "ecumodeect")),
    ("CAN",     ("can", "rx", "tx", "msg", "grp", "busoff", "msggr", "canbus", "crc", "alive")),
    ("SENSOR",  ("hall", "sensor", "pos", "angle", "temp", "tempr", "nvm", "cal", "calib")),
    ("SBW",     ("sbw", "act", "motor", "ctrl", "torq", "steer", "brake", "accel")),
    ("SAFETY",  ("asil", "safety", "sg", "sm", "wdg", "watchdog", "checksum", "e2e")),
    ("IGN",     ("ign", "ignon", "ignoff", "ignvolt", "ignsw")),
]

# ── 신호 카테고리별 effect_system 템플릿 (failure_mode → 텍스트) ──────────────
EFFECT_TEMPLATES: dict[str, dict[str, str]] = {
    "GEAR": {
        "MORE":    "{vn} 신호가 유효 범위를 초과하여 SBW 제어기가 잘못된 기어/레버 위치로 판단, 오변속 또는 변속 거부 발생",
        "LESS":    "{vn} 신호가 유효 범위 미만으로 수신되어 실제 레버 위치를 감지하지 못하고 변속 미작동 또는 잘못된 기어 유지",
        "CORRUPT": "{vn} 신호값이 논리적으로 유효하지 않아 SBW 제어 로직이 잘못된 변속 명령을 수행하여 오변속 또는 P락 오동작",
        "EARLY":   "{vn} 신호가 예상 시점보다 일찍 변환되어 변속 시퀀스 오류 발생, 주행 중 예기치 않은 변속단 전환 가능",
        "LATE":    "{vn} 신호 응답 지연으로 인해 SBW 제어기가 변속 명령을 늦게 처리하여 변속 타이밍 오류 및 충격 발생",
        None:      "{vn} 신호 이상으로 SBW 제어기가 기어/레버 위치를 잘못 판단하여 변속 오동작 또는 기능 상실 발생",
    },
    "POWER": {
        "MORE":    "{vn} 신호가 과전압 범위를 나타내어 전원 관리 모듈이 과전압 보호 동작을 수행, SBW 시스템 강제 비활성화 가능",
        "LESS":    "{vn} 저전압 신호로 인해 SBW 구동 회로의 전원이 부족하여 모터 제어 불능 및 변속 기능 상실",
        "CORRUPT": "{vn} 전압 값이 손상되어 전원 관리 모듈이 잘못된 전원 상태를 판단하고 SBW 시스템 비정상 동작 유발",
        "EARLY":   "{vn} 전압 측정이 안정화 전 조기 수행되어 부정확한 값 기반으로 전원 상태 판단, 초기화 오류 유발",
        "LATE":    "{vn} 전압 모니터링 지연으로 과전압/저전압 상황에 즉각 대응하지 못해 시스템 손상 위험 증가",
        None:      "{vn} 전원/전압 신호 이상으로 SBW 구동 전원 공급에 오류 발생, 변속 기능 저하 또는 시스템 비활성화",
    },
    "FAULT": {
        "MORE":    "{vn} 고장 신호가 과도하게 발생하여 SBW 시스템이 불필요한 페일세이프 모드로 전환, 정상 변속 기능 제한",
        "LESS":    "{vn} 고장 신호가 실제보다 낮게 보고되어 실제 결함을 감지하지 못하고 페일세이프 미작동",
        "CORRUPT": "{vn} 고장 코드가 손상되어 진단 시스템이 실제 결함 위치를 잘못 판단, 잘못된 수리 조치 유발",
        "EARLY":   "{vn} 고장 신호가 조기 발생하여 정상 상태임에도 페일세이프 전환, 불필요한 변속 제한 발생",
        "LATE":    "{vn} 고장 감지 지연으로 실제 결함이 발생해도 즉각적인 안전 조치가 이루어지지 않아 2차 손상 위험",
        None:      "{vn} 고장 진단 신호 이상으로 SBW 시스템의 결함 감지 및 페일세이프 동작에 오류 발생",
    },
    "STATUS": {
        "MORE":    "{vn} 상태값이 정의된 범위를 벗어난 값으로 수신되어 SBW 제어 모드 판단 오류 및 비정상 동작 유발",
        "LESS":    "{vn} 상태값이 최솟값 미만으로 수신되어 SBW 제어 로직이 잘못된 동작 모드로 전환",
        "CORRUPT": "{vn} 상태 정보가 손상되어 SBW 제어기가 잘못된 시스템 상태로 판단하고 부적절한 제어 수행",
        "EARLY":   "{vn} 상태 전환 신호가 조기 발생하여 SBW 초기화 시퀀스 오류 또는 잘못된 제어 모드 진입",
        "LATE":    "{vn} 상태 업데이트 지연으로 SBW 제어기가 이전 상태 기반으로 계속 동작하여 제어 불일치 발생",
        None:      "{vn} 시스템 상태 신호 이상으로 SBW 제어 모드 판단에 오류 발생, 의도치 않은 제어 동작 유발",
    },
    "CAN": {
        "MORE":    "{vn} CAN 메시지가 정상보다 높은 빈도로 수신되어 버스 부하 증가 및 SBW 제어 주기 오류 발생",
        "LESS":    "{vn} CAN 메시지 수신 빈도 저하로 SBW 제어기가 최신 신호값을 획득하지 못하고 이전값으로 제어 지속",
        "CORRUPT": "{vn} CAN 메시지 데이터가 손상되어 SBW 제어기가 잘못된 신호값으로 변속 제어를 수행",
        "EARLY":   "{vn} CAN 메시지가 예상보다 일찍 수신되어 동기화 오류 발생, SBW 제어 시퀀스 이상",
        "LATE":    "{vn} CAN 메시지 수신 지연(Timeout)으로 SBW 제어기가 통신 오류를 감지하고 페일세이프 모드로 전환",
        None:      "{vn} CAN 통신 신호 이상으로 SBW 제어에 필요한 정보가 정확히 전달되지 않아 변속 오동작 발생",
    },
    "SENSOR": {
        "MORE":    "{vn} 센서 출력값이 측정 범위를 초과하여 SBW 제어기가 비정상 위치/상태로 판단하고 잘못된 제어 수행",
        "LESS":    "{vn} 센서 출력값이 측정 범위 미만으로 하락하여 실제 위치/상태를 정확히 반영하지 못하고 제어 오류 발생",
        "CORRUPT": "{vn} 센서 데이터가 손상되어 SBW 위치 제어 로직이 잘못된 피드백 값을 사용하여 제어 정밀도 저하",
        "EARLY":   "{vn} 센서 신호가 실제 물리량보다 일찍 변화하여 SBW 제어기가 잘못된 시점에 제어 동작 수행",
        "LATE":    "{vn} 센서 응답 지연으로 SBW 위치 피드백이 늦게 전달되어 제어 응답성 저하 및 오버슈트 발생",
        None:      "{vn} 센서 신호 이상으로 SBW 위치/상태 측정에 오류 발생, 정밀 제어 불가 및 기능 저하",
    },
    "SBW": {
        "MORE":    "{vn} 제어 출력이 과도하여 SBW 액추에이터에 과도한 토크/전류가 인가되고 기계적 손상 또는 안전 위험 발생",
        "LESS":    "{vn} 제어 출력 부족으로 SBW 액추에이터가 목표 위치에 도달하지 못하여 변속 불완전 또는 변속 실패",
        "CORRUPT": "{vn} 제어 명령값 손상으로 SBW 액추에이터가 의도치 않은 방향/크기로 동작하여 오변속 또는 충격 발생",
        "EARLY":   "{vn} 제어 명령이 조기 발생하여 변속 시퀀스 오류 및 예기치 않은 기어 변환 발생",
        "LATE":    "{vn} 제어 명령 지연으로 변속 응답이 느려지고 드라이버 조작과 실제 변속 간 불일치 발생",
        None:      "{vn} SBW 제어 신호 이상으로 액추에이터 동작에 오류 발생, 변속 기능 저하 또는 안전 위험",
    },
    "SAFETY": {
        "MORE":    "{vn} 안전 신호가 과도하게 발생하여 불필요한 안전 메커니즘 활성화, 정상 SBW 기능 제한",
        "LESS":    "{vn} 안전 신호 부족으로 실제 위험 상황에서 안전 메커니즘이 작동하지 않아 ASIL 요구사항 미충족",
        "CORRUPT": "{vn} 안전 관련 데이터 손상으로 E2E 보호 실패 및 잘못된 안전 상태 진입 위험",
        "EARLY":   "{vn} 안전 메커니즘이 조기 활성화되어 정상 동작 중 불필요한 페일세이프 전환 발생",
        "LATE":    "{vn} 안전 메커니즘 반응 지연으로 위험 상황 발생 후 안전 상태 전환이 늦어져 안전 목표 미달성",
        None:      "{vn} 안전 관련 신호 이상으로 SBW 안전 메커니즘 작동에 오류 발생, ASIL 요구사항 미충족 위험",
    },
    "IGN": {
        "MORE":    "{vn} 이그니션 신호가 고전압 상태로 지속되어 전원 관리 모듈이 IG ON 상태를 잘못 유지, 불필요한 전력 소모",
        "LESS":    "{vn} 이그니션 전압 저하로 SBW 시스템 활성화 조건 미충족, 변속 기능 초기화 실패",
        "CORRUPT": "{vn} 이그니션 신호 손상으로 SBW 전원 관리 모듈이 IG 상태를 잘못 판단하고 부적절한 전원 시퀀스 수행",
        "EARLY":   "{vn} 이그니션 신호 조기 감지로 SBW 초기화가 시스템 준비 전에 시작되어 초기화 시퀀스 오류",
        "LATE":    "{vn} 이그니션 신호 감지 지연으로 IG ON 후 SBW 초기화가 늦어져 변속 가능 상태 진입 지연",
        None:      "{vn} 이그니션 신호 이상으로 SBW 전원 관리 및 초기화 시퀀스에 오류 발생",
    },
    "DEFAULT": {
        "MORE":    "{vn} 신호가 정상 범위를 초과하여 SBW 제어 로직이 잘못된 값을 처리하고 의도치 않은 동작 발생",
        "LESS":    "{vn} 신호값이 정상 범위 미만으로 하락하여 SBW 기능 수행에 필요한 정보가 부족하고 기능 저하 발생",
        "CORRUPT": "{vn} 신호 데이터 손상으로 SBW 제어기가 유효하지 않은 값을 처리하여 오동작 또는 기능 상실",
        "EARLY":   "{vn} 신호가 예상 시점보다 일찍 발생하여 SBW 제어 시퀀스 오류 및 의도치 않은 동작 유발",
        "LATE":    "{vn} 신호 응답 지연으로 SBW 제어기가 최신 상태를 반영하지 못하고 타이밍 오류 발생",
        None:      "{vn} 신호 이상으로 SBW 제어 기능에 오류 발생, 변속 성능 저하 또는 기능 상실 위험",
    },
}

# ── preventive_action 템플릿 ──────────────────────────────────────────────────
PREVENT_TEMPLATES: dict[str, dict[str, str]] = {
    "GEAR": {
        "MORE":    "레버/기어 위치 신호 상한 범위 검사(range check) 구현, 범위 초과 값 수신 시 이전 유효값 유지 및 오류 플래그 설정",
        "LESS":    "레버/기어 위치 신호 하한 범위 검사 구현, 범위 위반 시 기본 안전 기어(P/N) 유지 및 진단 DTC 기록",
        "CORRUPT": "기어/레버 위치 신호 E2E 보호 또는 CRC 검사 적용, 논리적 유효성 검증(허용 상태 전환 표) 구현",
        "EARLY":   "변속 요청 신호 타이밍 윈도우 검증 로직 구현, 유효 타이밍 외 신호 무시 및 시퀀스 상태 머신 적용",
        "LATE":    "변속 신호 타임아웃 모니터링 구현(alive counter 기반), 타임아웃 발생 시 최종 유효값 유지 및 DTC 기록",
        None:      "기어/레버 위치 신호 유효성 검증(범위·논리 체크) 구현, 이상 감지 시 안전 기어 유지 및 DTC 기록",
    },
    "POWER": {
        "MORE":    "전압 상한 임계값 검사 구현, 과전압 감지 시 SBW 시스템 보호 모드 전환 및 하드웨어 OVP 회로 설계 검토",
        "LESS":    "저전압 감지 임계값 설정 및 히스테리시스 적용, 저전압 시 변속 동작 중단 후 안전 상태 유지",
        "CORRUPT": "ADC 입력 이중 샘플링 비교 로직 적용, 샘플 간 편차 초과 시 이상값으로 처리 및 이전 유효값 사용",
        "EARLY":   "전원 안정화 대기 시간 확보 후 전압 측정 시작, 초기화 시퀀스 내 전압 샘플링 타이밍 검증 로직 추가",
        "LATE":    "전압 모니터링 주기 단축 및 인터럽트 기반 즉각 감지 구현, 과전압/저전압 즉시 대응 로직 적용",
        None:      "전원/전압 신호 범위 검사 및 안정화 검증 로직 구현, 이상 감지 시 보호 모드 전환 및 DTC 기록",
    },
    "FAULT": {
        "MORE":    "고장 신호 디바운스 필터 적용(최소 연속 감지 횟수 설정), 과도한 DTC 기록 방지를 위한 이벤트 쿨다운 로직",
        "LESS":    "고장 감지 임계값 하향 조정 및 자가진단(BIST) 주기적 실행, 고장 미감지 방지를 위한 교차 검증 로직",
        "CORRUPT": "고장 코드 데이터 무결성 검사(체크섬/CRC) 구현, 손상된 DTC 무시 및 재진단 트리거 로직 적용",
        "EARLY":   "고장 판정 이전 타당성(plausibility) 검사 구현, 단발성 노이즈와 실제 고장 구분을 위한 debounce 적용",
        "LATE":    "고장 감지 응답 시간 단축을 위한 샘플링 주기 최적화, 실시간 모니터링 인터럽트 기반 구현 검토",
        None:      "고장 진단 신호 유효성 검증 및 자가진단 로직 구현, 이상 감지 시 DTC 기록 및 페일세이프 전환",
    },
    "STATUS": {
        "MORE":    "상태값 허용 범위 상한 검사 구현, 비정의 상태값 수신 시 이전 유효 상태 유지 및 오류 DTC 기록",
        "LESS":    "상태값 허용 범위 하한 검사 구현, 유효 범위 이탈 시 기본 안전 상태로 전환 및 진단 기록",
        "CORRUPT": "상태 신호 논리적 유효성 검증(허용 상태 전환 표 기반), 유효하지 않은 상태 전환 시 무시 및 재확인",
        "EARLY":   "상태 전환 타이밍 검증 로직 구현, 초기화 시퀀스 완료 확인 후 상태값 처리 시작",
        "LATE":    "상태 업데이트 타임아웃 감시 구현, 타임아웃 시 이전 상태 유지 또는 안전 상태 전환",
        None:      "시스템 상태 신호 유효성 검증 및 허용 상태 전환 로직 구현, 이상 상태 감지 시 안전 상태 전환",
    },
    "CAN": {
        "MORE":    "CAN 메시지 수신 빈도 상한 모니터링 구현, 과도한 메시지 수신 시 버스 과부하 감지 및 오류 처리",
        "LESS":    "CAN 타임아웃 모니터링 구현(alive counter 기반), 수신 빈도 저하 감지 시 오류 플래그 설정 및 최종 유효값 사용",
        "CORRUPT": "CAN 메시지 CRC/체크섬 검증 구현, E2E 보호 프로토콜 적용으로 데이터 무결성 보장",
        "EARLY":   "CAN 메시지 타이밍 윈도우 검증 로직 구현, 예상 수신 시점 외 메시지 필터링 처리",
        "LATE":    "CAN 타임아웃 임계값 최적화 및 타임아웃 발생 시 즉각 페일세이프 전환, 버스 재초기화 로직 구현",
        None:      "CAN 통신 신호 타임아웃·CRC 검증 구현, 통신 이상 감지 시 오류 플래그 설정 및 안전 상태 전환",
    },
    "SENSOR": {
        "MORE":    "센서 출력 상한 범위 검사 구현, 범위 초과 시 센서 오류로 판정하고 이전 유효값 또는 기본값 사용",
        "LESS":    "센서 출력 하한 범위 검사 구현, 유효 범위 이탈 시 이전 유효값 유지 및 센서 고장 DTC 기록",
        "CORRUPT": "이중 센서 교차 검증 로직 구현, 두 센서 간 편차 임계값 이상 시 신호 손상으로 판정 및 페일세이프",
        "EARLY":   "센서 초기화 안정화 시간 확보 후 데이터 사용 시작, 초기 측정값 유효성 검증 후 제어 반영",
        "LATE":    "센서 샘플링 주기 최적화 및 인터럽트 기반 즉각 처리 구현, 지연 감지를 위한 타임스탬프 검증",
        None:      "센서 신호 범위 검사 및 이중화 검증 로직 구현, 센서 이상 감지 시 대체 신호 사용 및 DTC 기록",
    },
    "SBW": {
        "MORE":    "SBW 제어 출력 상한 클리핑 및 과전류 보호 로직 구현, 하드웨어 OCP 회로 설계 검토로 이중 보호",
        "LESS":    "SBW 제어 출력 최솟값 보정 로직 구현, 목표 위치 미달 시 재시도 및 위치 오차 임계값 초과 시 오류 처리",
        "CORRUPT": "SBW 제어 명령 무결성 검사 구현, 유효하지 않은 명령 필터링 및 안전 범위 내 클리핑 후 적용",
        "EARLY":   "SBW 제어 명령 발행 전 시스템 준비 상태 확인, 초기화 완료 플래그 검증 후 제어 명령 허용",
        "LATE":    "SBW 제어 응답 시간 모니터링 및 타임아웃 설정, 응답 지연 시 재시도 후 한계 초과 시 페일세이프 전환",
        None:      "SBW 제어 신호 유효성 검증 및 출력 범위 클리핑 구현, 이상 동작 감지 시 안전 모드 전환",
    },
    "SAFETY": {
        "MORE":    "안전 신호 과발생 방지를 위한 debounce 및 임계 횟수 설정, 불필요한 페일세이프 전환 방지 로직 구현",
        "LESS":    "안전 메커니즘 작동 임계값 하향 조정, 독립 채널 교차 검증으로 안전 신호 미감지 방지",
        "CORRUPT": "안전 데이터 E2E 보호 프로토콜(CRC + 카운터) 적용, ASIL 요구사항에 따른 데이터 무결성 보장",
        "EARLY":   "안전 메커니즘 활성화 조건 타당성 검사 구현, 단발성 이벤트와 실제 위험 상황 구분 로직 적용",
        "LATE":    "안전 메커니즘 응답 시간 요구사항 검증, 인터럽트 기반 즉각 처리 구현으로 ASIL 응답 시간 충족",
        None:      "ASIL 요구사항에 따른 안전 신호 E2E 보호 구현, 안전 메커니즘 작동 검증 및 주기적 자가진단 실행",
    },
    "IGN": {
        "MORE":    "이그니션 전압 상한 검사 구현, IG ON 유지 조건 타당성 검증(차속·기어 상태 교차 확인)",
        "LESS":    "이그니션 전압 하한 임계값 설정 및 히스테리시스 적용, 저전압 시 SBW 안전 종료 시퀀스 실행",
        "CORRUPT": "이그니션 신호 이중화(하드웨어 + 소프트웨어) 감지 구현, 불일치 시 더 안전한 상태(IG OFF) 적용",
        "EARLY":   "이그니션 ON 감지 후 전원 안정화 대기 시간 설정, 안정화 완료 후 SBW 초기화 시퀀스 시작",
        "LATE":    "이그니션 신호 감지 우선순위 상향 설정, 인터럽트 기반 즉각 처리로 초기화 지연 최소화",
        None:      "이그니션 신호 유효성 검증 및 전원 시퀀스 타이밍 검증 로직 구현, 이상 감지 시 안전 종료 시퀀스 실행",
    },
    "DEFAULT": {
        "MORE":    "신호 상한 범위 검사 구현, 범위 초과 감지 시 이전 유효값 유지 및 오류 플래그 설정·DTC 기록",
        "LESS":    "신호 하한 범위 검사 구현, 유효 범위 이탈 시 기본 안전값 적용 및 진단 코드 기록",
        "CORRUPT": "신호 데이터 무결성 검사(CRC/체크섬) 구현, 손상 감지 시 이전 유효값 사용 및 DTC 기록",
        "EARLY":   "신호 타이밍 윈도우 검증 로직 구현, 예상 타이밍 외 신호 발생 시 필터링 및 시퀀스 재검토",
        "LATE":    "신호 타임아웃 모니터링 구현, 타임아웃 발생 시 최종 유효값 유지 및 오류 상태 진입",
        None:      "신호 유효성 검증(범위·논리 체크) 및 타임아웃 모니터링 구현, 이상 감지 시 안전 상태 전환 및 DTC 기록",
    },
}


def classify_signal(variable_name: str) -> str:
    low = variable_name.lower()
    for cat, keywords in SIG_CATEGORIES:
        if any(kw in low for kw in keywords):
            return cat
    return "DEFAULT"


def is_generic(text: str | None) -> bool:
    if not text:
        return True
    return any(p in text for p in GENERIC_PATTERNS)


def make_text(templates: dict, sig_cat: str, fm: str | None, vn: str) -> str:
    cat_tmpl = templates.get(sig_cat, templates["DEFAULT"])
    tmpl = cat_tmpl.get(fm) or cat_tmpl.get(None) or ""
    return tmpl.replace("{vn}", vn)


def sb_get(table, params):
    rows, offset = [], 0
    while True:
        r = requests.get(f"{SB_URL}/rest/v1/{table}",
                         headers={**SB_H, "Range": f"{offset}-{offset+999}"},
                         params=params, verify=False)
        data = r.json()
        if not isinstance(data, list) or not data:
            break
        rows.extend(data)
        if len(data) < 1000:
            break
        offset += 1000
    return rows


def sb_patch(item_id: str, patch: dict) -> bool:
    r = requests.patch(f"{SB_URL}/rest/v1/fmea_items",
                       headers=SB_H,
                       params={"id": f"eq.{item_id}"},
                       json=patch, verify=False)
    return r.status_code in (200, 204)


def process_project(pid: str, name: str) -> int:
    print(f"\n[{name}] 형식적 텍스트 항목 조회...")
    todo = sb_get("fmea_items", {
        "project_id": f"eq.{pid}",
        "severity": "not.is.null",
        "select": "id,variable_name,failure_mode,category,preventive_action,effect_system",
    })
    todo = [i for i in todo if is_generic(i.get("preventive_action"))]
    print(f"  개선 대상: {len(todo)}개")
    if not todo:
        return 0

    improved = 0

    def update_one(item):
        vn = item["variable_name"]
        fm = item.get("failure_mode")
        sig_cat = classify_signal(vn)

        patch = {}
        new_pa = make_text(PREVENT_TEMPLATES, sig_cat, fm, vn)
        if new_pa:
            patch["preventive_action"] = new_pa

        # effect_system도 짧으면 개선
        if not item.get("effect_system") or len(item.get("effect_system", "")) < 25:
            new_es = make_text(EFFECT_TEMPLATES, sig_cat, fm, vn)
            if new_es:
                patch["effect_system"] = new_es

        if not patch:
            return False
        return sb_patch(item["id"], patch)

    with concurrent.futures.ThreadPoolExecutor(max_workers=20) as ex:
        futures = {ex.submit(update_one, item): item for item in todo}
        for i, f in enumerate(concurrent.futures.as_completed(futures), 1):
            if f.result():
                improved += 1
            if i % 300 == 0:
                print(f"  진행: {i}/{len(todo)} (개선:{improved})", end="\r")

    print(f"  완료: {improved}개 개선")
    return improved


def main():
    filter_names = [a.upper() for a in sys.argv[1:]] if len(sys.argv) > 1 else []
    targets = [(pid, name) for pid, name in TARGET_PROJECTS
               if not filter_names or name.upper() in filter_names]

    print("=" * 60)
    print("FMEA 텍스트 품질 개선 2단계 - 신호유형별 구체 템플릿")
    print(f"대상: {[n for _, n in targets]}")
    print("=" * 60)

    total = sum(process_project(pid, name) for pid, name in targets)
    print(f"\n전체 완료: {total}개 항목 개선")


if __name__ == "__main__":
    main()
