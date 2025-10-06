from pydantic_settings import BaseSettings


class Settings(BaseSettings):
    database_url: str = (
        "postgresql://divitrek_user:divitrek_password@localhost:5432/divitrek_db"
    )
    api_host: str = "0.0.0.0"
    api_port: int = 8000
    api_reload: bool = True
    streamlit_host: str = "0.0.0.0"
    streamlit_port: int = 8501

    class Config:
        env_prefix = ""
        case_sensitive = False


settings = Settings()  # reads from environment if provided
