from fastapi import FastAPI
from src.config import settings
from src.pdf_to_docx.router import router as pdf_to_docx_router
from src.html_to_docx.router import router as html_to_docx_router
from contextlib import asynccontextmanager
import logging

access_logger = logging.getLogger("uvicorn.access")


class HealthEndpointFilter(logging.Filter):
    def filter(self, record):
        return not (record.getMessage().find("GET /health") != -1)


access_logger.addFilter(HealthEndpointFilter())

logger = logging.getLogger("uvicorn")


@asynccontextmanager
async def lifespan(app: FastAPI):
    logger.info("FastAPI app initialized")
    logger.info(f"\ttitle: {settings.TITLE}")
    logger.info(f"\tversion: {settings.VERSION}")
    yield
    logger.info("FastAPI app shutdown complete.")


app = FastAPI(
    title=settings.TITLE,
    version=settings.VERSION,
    openapi_url="/openapi.json",
    docs_url="/docs",
    lifespan=lifespan,
)

# Mantém o router original sem prefix para compatibilidade
app.include_router(
    pdf_to_docx_router,
    prefix="/pdf-to-docx",
    tags=["PDF to DOCX"],
)

# Novo endpoint para conversão direta de HTML (evita problemas com containers coloridos)
app.include_router(
    html_to_docx_router,
    prefix="/html-to-docx",
    tags=["HTML to DOCX"],
)
