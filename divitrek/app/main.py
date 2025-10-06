from fastapi import FastAPI

from app.api.routes import assets
from app.core.db import init_db


def create_app() -> FastAPI:
    init_db()
    application = FastAPI(title="DiviTrek API")
    application.include_router(assets.router)
    return application


app = create_app()
