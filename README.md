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

## RAG 유사 사례 검색

| 파일 | 설명 |
|---|---|
| `rag_text.py` | 임베딩 텍스트 정규화 공용 모듈 (약어 확장, camelCase 분리, canonical 앵커) |
| `rag_embed.py` | FMEA 항목 임베딩 생성 → Supabase 저장 |
| `rag_search.py` | 유사 FMEA 항목 검색 (CLI / `RagSearcher` 모듈) |

```bash
# 임베딩 재생성 (rag_text.py 변경 시 필수, 회사망에서는 오프라인 모드로)
HF_HUB_OFFLINE=1 TRANSFORMERS_OFFLINE=1 python -u rag_embed.py            # 전체
HF_HUB_OFFLINE=1 TRANSFORMERS_OFFLINE=1 python -u rag_embed.py JG1 LQ2    # 특정 프로젝트

# 유사 항목 검색
python rag_search.py "BDC_02_Timeout" "LATE" --top 3
```

> `rag_embed.py`와 `rag_search.py`는 반드시 같은 `rag_text.build_embed_text()`를
> 사용해야 함 (코퍼스와 쿼리의 임베딩 텍스트 형식이 다르면 검색 품질 급락).

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
