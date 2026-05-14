"""FMEA Pipeline API — Railway 배포용 FastAPI 앱"""

import uuid, os, requests as _req

from fastapi import FastAPI, UploadFile, File, Form, BackgroundTasks, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from pipeline import run_pipeline

SB_URL = "https://itzgdbeiyvodhfhmvrfw.supabase.co"
SB_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Iml0emdkYmVpeXZvZGhmaG12cmZ3Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzY3NDE2MzcsImV4cCI6MjA5MjMxNzYzN30.iqzQr-3Lqf1O4UHFe9euTLyeIyBXreLPoSzUdtEaNP8"
SB_H = {"apikey": SB_KEY, "Authorization": f"Bearer {SB_KEY}", "Content-Type": "application/json"}
SRC_PROJECT_IDS = ["0715f883-d3a1-4ddd-8a3b-d3071da9ed3e"]  # JG1

app = FastAPI(title="FMEA Pipeline API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# 메모리 내 잡 저장소 (재시작 시 초기화됨 — 장기 운영 시 Redis/DB 권장)
jobs: dict = {}


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/jobs")
async def create_job(
    background_tasks: BackgroundTasks,
    project_name: str = Form(...),
    vehicle_model: str = Form(...),
    arxml_zip: UploadFile = File(...),
    dbc_files: list[UploadFile] = File(default=[]),
):
    job_id = str(uuid.uuid4())
    jobs[job_id] = {"status": "running", "progress": 0, "logs": [], "project_id": None}

    arxml_data = await arxml_zip.read()
    dbc_data = [(f.filename or "upload.dbc", await f.read()) for f in dbc_files]

    background_tasks.add_task(
        run_pipeline, job_id, jobs, project_name, vehicle_model, arxml_data, dbc_data
    )
    return {"job_id": job_id}


@app.get("/jobs")
def list_jobs():
    return [{"job_id": k, **v} for k, v in jobs.items()]


@app.get("/jobs/{job_id}")
def get_job(job_id: str):
    job = jobs.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job not found")
    return job


class SimilarRequest(BaseModel):
    item_id: str
    failure_mode: str | None = None
    top_k: int = 5


@app.post("/rag/similar")
def find_similar(req: SimilarRequest):
    """유사 FMEA 항목 검색 — 항목의 저장된 임베딩을 쿼리로 사용 (모델 불필요)"""
    # 1. 해당 항목의 임베딩 조회
    r = _req.get(
        f"{SB_URL}/rest/v1/fmea_items",
        headers=SB_H,
        params={"id": f"eq.{req.item_id}", "select": "embedding"},
        verify=False,
    )
    rows = r.json()
    if not rows or not rows[0].get("embedding"):
        raise HTTPException(status_code=404, detail="임베딩 없음 — rag_embed.py 실행 필요")

    embedding = rows[0]["embedding"]

    # 2. 유사 항목 검색
    payload = {
        "query_embedding": embedding,
        "match_count": req.top_k,
        "source_project_ids": SRC_PROJECT_IDS,
        "filter_failure_mode": req.failure_mode,
    }
    r2 = _req.post(
        f"{SB_URL}/rest/v1/rpc/match_fmea_items",
        headers=SB_H,
        json=payload,
        verify=False,
    )
    if r2.status_code != 200:
        raise HTTPException(status_code=502, detail=f"Supabase 오류: {r2.text[:200]}")

    results = r2.json()

    # failure_mode 일치 결과 없으면 ANY로 재검색
    if not results and req.failure_mode:
        payload["filter_failure_mode"] = None
        r3 = _req.post(f"{SB_URL}/rest/v1/rpc/match_fmea_items", headers=SB_H, json=payload, verify=False)
        if r3.status_code == 200:
            results = r3.json()

    return {"results": results}
