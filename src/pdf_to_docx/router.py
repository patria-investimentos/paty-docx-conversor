from fastapi import APIRouter, Depends, status
from fastapi.responses import StreamingResponse, JSONResponse
from src.pdf_to_docx.service import convert_pdf_to_docx
from src.constants import DOCX_MIME_TYPE
from src.pdf_to_docx.dependencies import pdf_dependency
from src.pdf_to_docx.examples import get_response_examples

router = APIRouter()


@router.get(
    "/health",
    tags=["Health"],
    status_code=status.HTTP_200_OK,
    description="Health check endpoint",
    summary="Check service health",
    response_class=JSONResponse,
)
async def health_check():
    return {"status": "healthy"}


@router.post(
    "/",
    tags=["PDF to DOCX"],
    status_code=status.HTTP_200_OK,
    description="Recebe um arquivo PDF e converte para DOCX",
    summary="Converter PDF em DOCX",
    response_class=StreamingResponse,
    responses=get_response_examples(),
)
async def convert_pdf_file(
    pdf: dict = Depends(pdf_dependency),
) -> StreamingResponse:

    docx_buffer = await convert_pdf_to_docx(pdf.get("data"))

    filename = f"{(pdf.get('filename') if 'filename' in pdf else 'download').rsplit('.', 1)[0]}.docx"

    return StreamingResponse(
        docx_buffer,
        media_type=DOCX_MIME_TYPE,
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"; charset=utf-8'
        },
    )
