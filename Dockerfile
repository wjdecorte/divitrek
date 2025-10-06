FROM ghcr.io/astral-sh/uv:python3.13-bookworm-slim AS builder

WORKDIR /app
COPY pyproject.toml ./
RUN uv pip install --system .[dev]

FROM ghcr.io/astral-sh/uv:python3.13-bookworm-slim
WORKDIR /app

ENV PYTHONUNBUFFERED=1

COPY --from=builder /usr/local /usr/local
COPY app ./app
COPY frontend ./frontend

EXPOSE 8000 8501

CMD ["bash", "-lc", "uvicorn app.main:app --host 0.0.0.0 --port 8000 & streamlit run frontend/streamlit_app.py --server.port 8501 --server.address 0.0.0.0"]


