from fastapi import APIRouter, Depends, status
from fastapi.responses import StreamingResponse
from src.html_to_docx.service import convert_html_to_docx
from src.constants import DOCX_MIME_TYPE
from src.html_to_docx.dependencies import html_dependency

router = APIRouter()


@router.post(
    "/",
    tags=["HTML to DOCX"],
    status_code=status.HTTP_200_OK,
    description="Recebe um arquivo HTML e converte para DOCX. "
                "Distingue tabelas de dados de tabelas de layout automaticamente.",
    summary="Converter HTML em DOCX",
    response_class=StreamingResponse,
)
async def convert_html_file(
    html: dict = Depends(html_dependency),
) -> StreamingResponse:
    """
    Converte HTML diretamente para DOCX.
    
    Este endpoint converte HTML para DOCX com suporte a:
    - Headings (h1-h6) com estilos apropriados
    - Parágrafos e formatação inline (bold, italic, underline)
    - Listas ordenadas e não-ordenadas
    - Tabelas de dados (com bordas) vs tabelas de layout (sem bordas)
    - Imagens base64 inline
    - Hyperlinks
    - Estilos CSS de classes e inline
    """
    docx_buffer = await convert_html_to_docx(html.get("content"))

    filename = html.get("filename", "document")
    if filename.endswith(".html"):
        filename = filename[:-5]
    filename = f"{filename}.docx"

    return StreamingResponse(
        docx_buffer,
        media_type=DOCX_MIME_TYPE,
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"; charset=utf-8'
        },
    )
