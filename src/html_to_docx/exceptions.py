from src.exceptions import BadRequest, UnsupportedMediaType


class HTMLNotProvided(BadRequest):
    def __init__(self):
        super().__init__(detail="Conteúdo HTML não fornecido")


class HTMLEmpty(BadRequest):
    def __init__(self):
        super().__init__(detail="Conteúdo HTML vazio")


class InvalidHTMLMimeType(UnsupportedMediaType):
    def __init__(self, content_type: str):
        super().__init__(detail=f"Tipo de mídia inválido: {content_type}. Esperado: text/html")


class HTMLConversionError(BadRequest):
    def __init__(self, message: str = "Erro ao converter HTML para DOCX"):
        super().__init__(detail=message)
