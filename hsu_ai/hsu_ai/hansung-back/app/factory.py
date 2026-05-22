from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from app.routers import documents, intranet, prompts, reports, rules


def create_app() -> FastAPI:
    app = FastAPI(title="Hansung AI Report API")
    app.add_middleware(
        CORSMiddleware,
        allow_origins=["http://localhost:3000"],
        allow_methods=["*"],
        allow_headers=["*"],
    )
    app.include_router(intranet.router)
    app.include_router(documents.router)
    app.include_router(rules.router)
    app.include_router(prompts.router)
    app.include_router(reports.router)
    return app
