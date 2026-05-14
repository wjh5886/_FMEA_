"""
rag_search.py — FMEA RAG 유사 항목 검색
임베딩된 FMEA 데이터에서 유사 신호 항목을 찾아 반환

사용법 (단독 실행):
  python rag_search.py "GearPosSta" "CORRUPT"
  python rag_search.py "LvrPosInfo" "LATE" --top 3

모듈 사용:
  from rag_search import RagSearcher
  searcher = RagSearcher()
  results = searcher.search("GearPosSta", "STUCK", top_k=5)
"""

import sys, os, requests, urllib3
import numpy as np
urllib3.disable_warnings()

# 회사 네트워크 SSL 우회
import ssl, httpx
ssl._create_default_https_context = ssl._create_unverified_context
os.environ.setdefault("CURL_CA_BUNDLE", "")
os.environ.setdefault("REQUESTS_CA_BUNDLE", "")
_orig_client = httpx.Client.__init__
def _ssl_client(self, *args, **kwargs): kwargs["verify"] = False; _orig_client(self, *args, **kwargs)
httpx.Client.__init__ = _ssl_client
_orig_async = httpx.AsyncClient.__init__
def _ssl_async(self, *args, **kwargs): kwargs["verify"] = False; _orig_async(self, *args, **kwargs)
httpx.AsyncClient.__init__ = _ssl_async

from sentence_transformers import SentenceTransformer

SB_URL = "https://itzgdbeiyvodhfhmvrfw.supabase.co"
SB_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Iml0emdkYmVpeXZvZGhmaG12cmZ3Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzY3NDE2MzcsImV4cCI6MjA5MjMxNzYzN30.iqzQr-3Lqf1O4UHFe9euTLyeIyBXreLPoSzUdtEaNP8"
SB_H = {
    "apikey": SB_KEY,
    "Authorization": f"Bearer {SB_KEY}",
    "Content-Type": "application/json",
}

MODEL_NAME = "paraphrase-multilingual-MiniLM-L12-v2"

# 양질 소스 프로젝트 (검색 대상) — JG1만 사용 (SBW FMEA 표준 5종 failure_mode)
# SX3_ICE_TEST는 STUCK/ERRATIC 포함(비표준)이므로 제외
SRC_PROJECT_IDS = [
    "0715f883-d3a1-4ddd-8a3b-d3071da9ed3e",  # JG1
]


class RagSearcher:
    def __init__(self):
        self._model = None

    @property
    def model(self):
        if self._model is None:
            print(f"  임베딩 모델 로드: {MODEL_NAME}")
            self._model = SentenceTransformer(MODEL_NAME)
        return self._model

    def embed(self, variable_name: str, failure_mode: str | None) -> list[float]:
        base = variable_name.split("(")[0].strip()
        mode = failure_mode or "ANY"
        text = f"{base} {mode}"
        vec = self.model.encode(text, normalize_embeddings=True)
        return vec.tolist()

    def search(
        self,
        variable_name: str,
        failure_mode: str | None = None,
        top_k: int = 5,
        src_only: bool = True,
    ) -> list[dict]:
        """
        유사 FMEA 항목 검색

        Returns:
            [{id, variable_name, failure_mode, effect_system,
              preventive_action, severity, occurrence, detection, rpn, similarity}]
        """
        embedding = self.embed(variable_name, failure_mode)

        payload = {
            "query_embedding": embedding,
            "match_count": top_k,
            "source_project_ids": SRC_PROJECT_IDS if src_only else None,
            "filter_failure_mode": failure_mode,
        }

        r = requests.post(
            f"{SB_URL}/rest/v1/rpc/match_fmea_items",
            headers=SB_H,
            json=payload,
            verify=False,
        )

        if r.status_code != 200:
            print(f"  검색 오류: {r.status_code} {r.text[:300]}")
            return []

        results = r.json()

        # failure_mode 일치 결과 없으면 ANY로 재검색
        if not results and failure_mode:
            payload["filter_failure_mode"] = None
            r2 = requests.post(
                f"{SB_URL}/rest/v1/rpc/match_fmea_items",
                headers=SB_H,
                json=payload,
                verify=False,
            )
            if r2.status_code == 200:
                results = r2.json()

        return results


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    top_k = 5
    for i, a in enumerate(sys.argv[1:]):
        if a == "--top" and i + 2 < len(sys.argv):
            top_k = int(sys.argv[i + 2])

    if not args:
        print("사용법: python rag_search.py <variable_name> [failure_mode] [--top N]")
        sys.exit(1)

    vn = args[0]
    fm = args[1] if len(args) > 1 else None

    print(f"\n검색: '{vn}'  고장모드: {fm or 'ANY'}  top_k={top_k}")
    print("-" * 60)

    searcher = RagSearcher()
    results = searcher.search(vn, fm, top_k=top_k)

    if not results:
        print("  결과 없음 (임베딩이 생성되지 않았거나 유사 항목 없음)")
        return

    for i, r in enumerate(results, 1):
        sim = r.get("similarity", 0)
        print(f"\n[{i}] 유사도 {sim:.3f}  {r['variable_name'][:60]}  [{r['failure_mode']}]")
        print(f"    effect_system    : {(r.get('effect_system') or '-')[:120]}")
        print(f"    preventive_action: {(r.get('preventive_action') or '-')[:120]}")
        print(f"    S/O/D: {r.get('severity')}/{r.get('occurrence')}/{r.get('detection')}  RPN={r.get('rpn')}")


if __name__ == "__main__":
    main()
