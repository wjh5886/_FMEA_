# FMEA Automation — 파싱 스크립트

SW FMEA 자동화를 위한 Python 스크립트 모음.  
웹 앱 소스는 → [fmea-web/](./fmea-web/) 또는 https://github.com/wjh5886/fmea-automation

## 스크립트 목록

| 파일 | 설명 |
|---|---|
| `parse_sx3_arxml.py` | SX3 ARXML 파일 파싱 → Supabase 업로드 (SX3_ARXML 프로젝트) |
| `parse_arxml_fmea.py` | ARXML 일반 파싱 유틸 |
| `build_fmea.py` | JG1 FMEA Excel → Supabase 업로드 |
| `build_lq2_fmea.py` | LQ2 프로젝트 FMEA 빌드 |
| `build_nq6e_fmea.py` | NQ6e 프로젝트 FMEA 빌드 |
| `build_pwrmgt_fmea.py` | PwrMGT 프로젝트 FMEA 빌드 |
| `import_fmea.py` | Excel FMEA → Supabase import |
| `import_sx3_excel.py` | SX3 Excel → Supabase import |
| `fill_all_fields.py` | FMEA 필드 일괄 채우기 |
| `fill_fmea_full.py` | FMEA 전체 필드 채우기 |
| `fill_ghi_from_code.py` | 소스코드 기반 G/H/I 필드 채우기 |
| `fill_signal_range.py` | 신호 범위 필드 채우기 |
| `fill_sx3_sod.py` | SX3 S/O/D 점수 채우기 |
| `generate_sx3_fmea.py` | SX3 FMEA 생성 |
| `organize_fmea.py` | FMEA 항목 정리 |
| `cross_ref.py` | 신호 교차 참조 |
| `update_dbc_values.py` | DBC 파일 값 업데이트 |

## 환경 설정

```bash
pip install supabase openpyxl lxml anthropic
```

`.env` 파일 생성:
```
SUPABASE_URL=...
SUPABASE_KEY=...
ANTHROPIC_API_KEY=...
```

## ARXML 파싱 실행 예시

```bash
# SX3 ARXML → Supabase (SX3_ARXML 프로젝트)
python parse_sx3_arxml.py

# 원본 파일 위치 (로컬 필요):
# SBW_FMEA/SX3/HKMC_SX3_SBW_.../Configuration/System/Swcd_App/*.arxml
# SBW_FMEA/SX3/HKMC_SX3_SBW_.../RootComposition.arxml
```
