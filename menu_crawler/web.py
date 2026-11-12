import secrets
from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import FastAPI, HTTPException, Request, Response
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from typing import Literal

from pydantic import BaseModel, Field
from starlette.middleware.trustedhost import TrustedHostMiddleware

from .crawler import Job
from .export import make_workbook
from .models import CrawlConfig

STATIC = Path(__file__).with_name("static")


class Selection(BaseModel):
    id: str = Field(max_length=64)
    depths: list[str] = Field(min_length=3, max_length=3)


class Control(BaseModel):
    phase: Literal["before", "after"] | None = None


class LoginFill(BaseModel):
    # No length constraints here: a validation error would echo the value back.
    username: str = ""
    password: str = ""


class ExportRequest(BaseModel):
    selections: list[Selection] | None = Field(default=None, max_length=20000)


def create_app(token=None):
    token = token or secrets.token_urlsafe(32)

    @asynccontextmanager
    async def lifespan(app):
        yield
        if app.state.job:
            app.state.job.close()

    app = FastAPI(lifespan=lifespan, docs_url=None, redoc_url=None, openapi_url=None)
    app.state.token = token
    app.state.job = None
    app.add_middleware(TrustedHostMiddleware, allowed_hosts=["127.0.0.1", "localhost", "testserver"])

    @app.middleware("http")
    async def local_session(request: Request, call_next):
        if request.url.path.startswith("/api/"):
            if not secrets.compare_digest(request.headers.get("x-session-token", ""), token):
                return Response("프로그램의 실행 링크에서 다시 접속하세요.", status_code=403)
            if request.headers.get("origin") and request.headers["origin"] != str(request.base_url).rstrip("/"):
                return Response("요청 출처가 일치하지 않습니다.", status_code=403)
        response = await call_next(request)
        response.headers["Cache-Control"] = "no-store"
        response.headers["X-Content-Type-Options"] = "nosniff"
        response.headers["Referrer-Policy"] = "no-referrer"
        response.headers["Content-Security-Policy"] = "default-src 'self'; script-src 'self'; style-src 'self'; connect-src 'self'; img-src 'self' data:; frame-ancestors 'none'; object-src 'none'; base-uri 'none'"
        return response

    @app.get("/")
    def index():
        return FileResponse(STATIC / "index.html")

    app.mount("/static", StaticFiles(directory=STATIC), name="static")

    @app.get("/api/status")
    def status():
        return app.state.job.snapshot() if app.state.job else {"state": "idle", "menus": [], "issues": [], "pages": []}

    @app.post("/api/open")
    def open_browser(config: CrawlConfig):
        if app.state.job:
            if app.state.job.snapshot()["state"] in {"starting", "capturing", "crawling"}:
                raise HTTPException(409, "진행 중인 수집을 먼저 중지하세요.")
            app.state.job.close()
        app.state.job = Job(config)
        return {"ok": True}

    def allowed_phase(job, phase):
        mode = job.config.mode
        phase = phase or ("after" if mode == "login" else "before")
        if mode == "public" and phase == "after":
            raise HTTPException(409, "로그인 전 수집 방식입니다. 로그인 후 수집은 수집 방식을 바꿔 새로 시작하세요.")
        if mode == "login" and phase == "before":
            raise HTTPException(409, "로그인 후 수집 방식입니다.")
        return phase

    @app.post("/api/control/{action}")
    def action(action: str, payload: Control | None = None):
        job = app.state.job
        if not job:
            raise HTTPException(409, "브라우저를 먼저 열어 주세요.")
        try:
            if action in {"crawl", "capture"}:
                job.command(action, phase=allowed_phase(job, payload.phase if payload else None))
            elif action == "login":
                if job.config.mode == "public":
                    raise HTTPException(409, "로그인 전 수집 방식에서는 로그인 화면을 쓰지 않습니다.")
                job.command("goto", url=job.config.login_url or job.config.url)
            elif action == "stop":
                job.stop.set()
            elif action == "close":
                job.close()
                app.state.job = None
            else:
                raise HTTPException(404, "없는 기능입니다.")
        except ValueError as exc:
            raise HTTPException(409, str(exc)) from exc
        return {"ok": True}

    @app.post("/api/login/fill")
    def fill(payload: LoginFill):
        job = app.state.job
        if not job or job.config.mode == "public":
            raise HTTPException(409, "로그인 단계에서만 사용할 수 있습니다.")
        if not payload.password or len(payload.password) > 500 or len(payload.username) > 300:
            raise HTTPException(422, "비밀번호를 입력하세요. 아이디는 300자, 비밀번호는 500자 이내입니다.")
        try:
            job.command("fill", username=payload.username, password=payload.password)
        except ValueError as exc:
            raise HTTPException(409, str(exc)) from exc
        return {"ok": True}

    @app.post("/api/export")
    def export(payload: ExportRequest):
        if not app.state.job:
            raise HTTPException(409, "수집 결과가 없습니다.")
        report = app.state.job.snapshot()
        if report["state"] in {"starting", "crawling", "capturing"}:
            raise HTTPException(409, "수집이 끝나거나 중지된 후 내려받으세요.")
        try:
            data = make_workbook(report, None if payload.selections is None else [s.model_dump() for s in payload.selections])
        except ValueError as exc:
            raise HTTPException(422, str(exc)) from exc
        return Response(data, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        headers={"Content-Disposition": 'attachment; filename="menu-inventory.xlsx"'})

    return app
