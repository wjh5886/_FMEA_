# JG1 SBW Software FMEA 소스코드 대조 분석

분석일: 2026-04-10  
대상 FMEA: JG1_SBW-Software_FMEA_1.xlsx  
대상 코드: HKMC_JG1_SBW_R44_4.2109.00_26210_Mood_R1 1/Static_Code/Source/

---

## 1. FMEA 전체 현황

| 항목 | 값 |
|------|-----|
| 총 FMEA 항목 | 1,460개 |
| SW Unit 수 | 11개 |
| 분석 대상 변수 | 291개 |
| S/O/D/RPN 입력 | **0개 (미입력)** |
| Preventive Action 입력 | **0개 (미입력)** |
| Detection Action 입력 | **0개 (미입력)** |

> ⚠️ **중요**: S, O, D, RPN 값이 전혀 입력되지 않아 위험도 평가가 되지 않은 상태

---

## 2. SW Unit별 소스코드 대조 결과

| SW Unit | 변수 수 | 코드에서 발견 | 미발견 | 커버리지 |
|---------|---------|------------|------|--------|
| CstAp_ButtonMgt | 2 | 2 | 0 | 100% ✅ |
| CstAp_IdtMgt | 31 | 31 | 0 | 100% ✅ |
| CstAp_MotorControlMgt | 21 | 21 | 0 | 100% ✅ |
| CstAp_MovingMgt | 4 | 4 | 0 | 100% ✅ |
| CstAp_PwrMGT | 6 | 6 | 0 | 100% ✅ |
| CstAp_MoodControlMgt | 13 | 12 | 1 | 92% 🟡 |
| CstAp_CANMGT | 85 | 71 | 14 | 83% 🟡 |
| CstAp_DtcMgt | 75 | 58 | 17 | 77% 🟡 |
| CstAp_HapticControlMgt | 6 | 4 | 2 | 66% 🔴 |
| CstAp_ECUModeMgt | 25 | 14 | 11 | 56% 🔴 |
| CstAp_PosMgt | 4 | 2 | 2 | 50% 🔴 |

---

## 3. 미발견 변수 상세 및 원인 분류

### 3-1. 명명 규칙 불일치 (코드에 존재하나 이름 다름)

| FMEA 변수명 | 실제 코드 변수명 | 파일 |
|------------|--------------|------|
| `BDC02Timeout` | `BDC_02_Timeout` | CtAp_CGWSigChk.c |
| `BDC05Timeout` | `BDC_05_Timeout` | CtAp_CGWSigChk.c |
| `CLU01Timeout` | `CLU_01_Timeout` | CAN_Management/ |
| `SMK03Timeout` | `SMK_03_Timeout` | CAN_Management/ |
| `SactSig` | `SActSig` (대문자 A) | CtAp_ShiftActSigChk.c |
| `SactSigTo` | `SActSigTo` | CtAp_ShiftActSigChk.c |
| `LeverWarning_Dial` | `LvrWrngMsg` (enum 값) | CtAp_SBWSigSet.c |
| `LeverWarning_Sphere` | `LvrWrngMsg` (enum 값) | CtAp_SBWSigSet.c |
| `RetryWarning_Dial` | `SBW_MotorWarn.RetryWarning` | CtAp_SBWSigSet.c |
| `RetryWarning_Sphere` | `SBW_MotorWarn.RetryWarning` | CtAp_SBWSigSet.c |
| `MotorFaultWarning` | `MotorFaultPinChek` | CtIoHwAb_IntfOut.c |
| `PositionSensorInfo_PButtonFltSta` | `PButtonFltSta` | Position_Management/ |
| `PositionSensorInfo_PButtonSta` | `PButtonSta` | Position_Management/ |

### 3-2. Generated 코드에 있는 변수 (Static Code에 없음)

| FMEA 변수명 | 생성 코드 파일 |
|------------|-------------|
| `SnapShot0200`~`SnapShot020E` | `Dtc_Management/CtAp_SnapShot0xFD50~FD60.c` (DTC ID 다름) |
| `HallSnrFltInfo` | Generated/Bsw_Output 또는 Dem 설정 |
| `HallSnrFltVal` | Generated/Bsw_Output 또는 Dem 설정 |
| `SlaveAddress` | Generated RTE (I2C 설정) |
| `TransmitLength` | Generated RTE (I2C 설정) |

### 3-3. 미구현 가능성 (추가 확인 필요)

| FMEA 변수명 | SW Unit | Effect on System |
|------------|---------|-----------------|
| `ValGearSlctDis` | CstAp_MoodControlMgt | - |
| `CGWCLU_01_20ms_Timeout` | CstAp_CANMGT | 가용성 관련 |
| `BDC_04_Timeout` | CstAp_CANMGT | - |
| `Balarm_BglrAlrmSta` | CstAp_CANMGT | - |
| `DrLockSta` | CstAp_ECUModeMgt | - |
| `MainCanSigSet_LvrMsg` | CstAp_CANMGT | - |
| `SubCanSigSet_LvrMsg` | CstAp_CANMGT | - |
| `Main_Vcu_E2E_Error_Return` | CstAp_CANMGT | - |
| `Naccept` | CstAp_CANMGT | - |
| `DeVCU01Timeout`, `DeVCU04Timeout` | CstAp_ECUModeMgt | - |
| `ICC02Timeout`, `PDC01Timeout`, `SBCMDRV01Timeout` | CstAp_ECUModeMgt | - |

---

## 4. 핵심 이슈 요약

### 🔴 Critical
1. **S/O/D/RPN 전체 미입력** — 위험도 평가가 전혀 이루어지지 않은 상태
2. **Preventive/Detection Action 전체 미입력** — 안전 조치 근거가 FMEA에 없음

### 🟡 Important
3. **변수 명명 불일치** — FMEA 변수명과 실제 코드 변수명이 상이  
   (언더스코어 유무, 대소문자, 접두어 차이)
4. **SnapShot DTC ID 불일치** — FMEA: `SnapShot020x`, 코드: `0xFD5x`
5. **CstAp_ECUModeMgt 커버리지 56%** — timeout 신호 11개 FMEA-코드 불일치

### 🟢 양호
6. **CstAp_PwrMGT, ButtonMgt, IdtMgt, MotorControlMgt, MovingMgt** — 100% 매칭

---

## 5. 권장 조치

| 우선순위 | 조치 | 담당 |
|---------|------|------|
| 1 | S/O/D/RPN 위험도 평가 진행 | FMEA 담당자 |
| 2 | Preventive/Detection Action 기술 | SW 개발팀 |
| 3 | 변수 명명 규칙 통일 (FMEA ↔ 코드) | SW 개발팀 |
| 4 | SnapShot DTC ID 매핑 정정 | DTC 담당자 |
| 5 | 미구현 변수 구현 여부 결정 | 팀장 검토 |
