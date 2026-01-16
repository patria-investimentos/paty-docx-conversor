from pydantic_settings import BaseSettings


class Settings(BaseSettings):
    VERSION: str = "1.0.0"
    TITLE: str = "PDF to DOCX Converter"
    PORT: int = 8000

    class Config:
        env_file = ".env"


settings = Settings()
