"""
rag_embed.py — FMEA 항목 임베딩 생성 & Supabase 저장
sentence-transformers paraphrase-multilingual-MiniLM-L12-v2 사용 (384차원, 무료)

사용법:
  python rag_embed.py              # 전체 프로젝트 임베딩
  python rag_embed.py --src-only   # JG1, SX3_ICE_TEST (소스 프로젝트)만
  python rag_embed.py "LQ2"        # 특정 프로젝트만
"""

import sys, os, requests, urllib3
import numpy as np
urllib3.disable_warnings()

# 회사 네트워크 SSL 우회 (자체 서명 인증서 환경)
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

from rag_text import build_embed_text

SB_URL = "https://itzgdbeiyvodhfhmvrfw.supabase.co"
SB_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Iml0emdkYmVpeXZvZGhmaG12cmZ3Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzY3NDE2MzcsImV4cCI6MjA5MjMxNzYzN30.iqzQr-3Lqf1O4UHFe9euTLyeIyBXreLPoSzUdtEaNP8"
SB_H = {
    "apikey": SB_KEY,
    "Authorization": f"Bearer {SB_KEY}",
    "Content-Type": "application/json",
    "Prefer": "return=minimal",
}

ALL_PROJECTS = [
    ("0715f883-d3a1-4ddd-8a3b-d3071da9ed3e", "JG1"),
    ("1ba10b41-4717-4d1b-b0df-1b91fe2c4870", "SX3_ICE_TEST"),
    ("32cc5148-96b5-42c5-a3dc-2b8fa3b9f93e", "GN7_FL"),
    ("89dc5818-2435-4d09-a1a9-36aea664d11d", "LQ2"),
    ("70f74c19-66b6-4b2e-a2b3-1ee04dd1b101", "TK1"),
    ("a43d7d4e-b104-4e90-ab55-04c7b31aa3e7", "SX3"),
]

SRC_ONLY = [
    ("0715f883-d3a1-4ddd-8a3b-d3071da9ed3e", "JG1"),
    ("1ba10b41-4717-4d1b-b0df-1b91fe2c4870", "SX3_ICE_TEST"),
]

MODEL_NAME = "paraphrase-multilingual-MiniLM-L12-v2"
BATCH_SIZE = 256
UPSERT_BATCH = 100


# 임베딩 텍스트 구성은 rag_text.build_embed_text로 통일 (검색 쿼리와 동일해야 함)
embed_text = build_embed_text


def sb_get_items(pid: str) -> list:
    rows, offset = [], 0
    while True:
        r = requests.get(
            f"{SB_URL}/rest/v1/fmea_items",
            headers={**SB_H, "Range": f"{offset}-{offset+999}"},
            params={"project_id": f"eq.{pid}",
                    "select": "id,variable_name,failure_mode"},
            verify=False,
        )
        data = r.json()
        if not isinstance(data, list) or not data:
            break
        rows.extend(data)
        if len(data) < 1000:
            break
        offset += 1000
    return rows


def sb_patch_embedding(item_id: str, embedding: list) -> bool:
    r = requests.patch(
        f"{SB_URL}/rest/v1/fmea_items",
        headers=SB_H,
        params={"id": f"eq.{item_id}"},
        json={"embedding": embedding},
        verify=False,
    )
    return r.status_code in (200, 204)


def sb_upsert_embeddings(records: list[dict]) -> bool:
    """records: [{"id": uuid, "embedding": [float, ...]}] — 병렬 PATCH"""
    import concurrent.futures
    ok_count = 0
    with concurrent.futures.ThreadPoolExecutor(max_workers=30) as ex:
        futures = {ex.submit(sb_patch_embedding, r["id"], r["embedding"]): r for r in records}
        for i, f in enumerate(concurrent.futures.as_completed(futures), 1):
            if f.result():
                ok_count += 1
            if i % 200 == 0:
                print(f"    저장 진행: {i}/{len(records)}", end="\r")
    return ok_count == len(records)


def process_project(pid: str, name: str, model: SentenceTransformer) -> int:
    print(f"\n[{name}] 항목 로드...")
    items = sb_get_items(pid)
    print(f"  {len(items)}개 항목")

    texts = [embed_text(it["variable_name"], it.get("failure_mode")) for it in items]

    print(f"  임베딩 생성 중 (배치 {BATCH_SIZE})...")
    embeddings = model.encode(
        texts,
        batch_size=BATCH_SIZE,
        show_progress_bar=True,
        normalize_embeddings=True,
    )

    records = [
        {"id": it["id"], "embedding": emb.tolist()}
        for it, emb in zip(items, embeddings)
    ]

    print(f"  Supabase 저장 중...")
    ok = sb_upsert_embeddings(records)
    if ok:
        print(f"  완료: {len(records)}개 저장")
    return len(records) if ok else 0


def main():
    args = sys.argv[1:]
    src_only = "--src-only" in args
    filter_names = [a.upper() for a in args if not a.startswith("--")]

    if src_only:
        targets = SRC_ONLY
    elif filter_names:
        targets = [(pid, name) for pid, name in ALL_PROJECTS
                   if name.upper() in filter_names]
    else:
        targets = ALL_PROJECTS

    print("=" * 60)
    print("FMEA RAG 임베딩 생성")
    print(f"모델: {MODEL_NAME}  대상: {[n for _, n in targets]}")
    print("=" * 60)

    print(f"\n모델 로드 중: {MODEL_NAME}")
    model = SentenceTransformer(MODEL_NAME)
    print("  모델 로드 완료")

    total = 0
    for pid, name in targets:
        total += process_project(pid, name, model)

    print(f"\n전체 완료: {total}개 임베딩 저장")


if __name__ == "__main__":
    main()
