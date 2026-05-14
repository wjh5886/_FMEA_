"""FMEA Pipeline API — Railway 배포용 FastAPI 앱"""

import uuid, sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

from fastapi import FastAPI, UploadFile, File, Form, BackgroundTasks, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from pipeline import run_pipeline

# RAG 검색기 (지연 초기화)
_searcher = None
def get_searcher():
    global _searcher
    if _searcher is None:
        try:
            from rag_search import RagSearcher
            _searcher = RagSearcher()
        except Exception:
            pass
    return _searcher

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
    variable_name: str
    failure_mode: str | None = None
    top_k: int = 5


@app.post("/rag/similar")
def find_similar(req: SimilarRequest):
    """유사 FMEA 항목 검색 (RAG)"""
    searcher = get_searcher()
    if searcher is None:
        raise HTTPException(status_code=503, detail="RAG 모델 미초기화 (rag_embed.py 먼저 실행)")
    results = searcher.search(req.variable_name, req.failure_mode, top_k=req.top_k)
    return {"results": results}
