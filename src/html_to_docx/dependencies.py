from fastapi import UploadFile
from src.constants import HTML_MIME_TYPE
import src.html_to_docx.exceptions as exceptions


async def html_dependency(html: UploadFile) -> dict[str, bytes | str]:
    """Valida e extrai dados do arquivo HTML."""
    if not html:
        raise exceptions.HTMLNotProvided()
    
    content_type = html.content_type or ""
    
    # Aceita text/html ou text/plain (alguns editores salvam como plain)
    if content_type not in [HTML_MIME_TYPE, "text/plain"]:
        raise exceptions.InvalidHTMLMimeType(content_type=content_type)

    html_bytes = await html.read()
    if not html_bytes:
        raise exceptions.HTMLEmpty()
    
    # Decodifica para string
    try:
        html_content = html_bytes.decode('utf-8')
    except UnicodeDecodeError:
        html_content = html_bytes.decode('latin-1')
    
    return {
        "content": html_content,
        "filename": html.filename or "document",
        "content_type": content_type
    }
