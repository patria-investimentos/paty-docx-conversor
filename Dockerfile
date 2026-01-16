FROM python:3.12-slim-bookworm AS builder

ENV PIP_DISABLE_PIP_VERSION_CHECK 1
ENV DEBIAN_FRONTEND noninteractive
ENV PATH "/usr/local/bin:/usr/bin:/bin:${HOME}/.local/bin:${PATH}"
WORKDIR /app

RUN pip install --no-cache-dir uv

COPY pyproject.toml ./
COPY requirements.txt ./

RUN uv pip install --system -r requirements.txt


FROM python:3.12-slim-bookworm AS runner

ENV PIP_DISABLE_PIP_VERSION_CHECK 1
ENV DEBIAN_FRONTEND noninteractive
ENV PATH "/usr/local/bin:/usr/bin:/bin:${HOME}/.local/bin:${PATH}"
ENV PYTHONUNBUFFERED=1
ENV PYTHONPATH=/app/src
ENV PORT=8000

WORKDIR /app

RUN useradd -m -u 1000 appuser && \
  chown -R appuser:appuser /app

COPY --from=builder /usr/local /usr/local
COPY --from=builder /usr/local/bin /usr/local/bin

RUN find /bin /sbin /usr/bin /usr/sbin /usr/local/bin -type f -executable \
  ! -name "python*" ! -name "uvicorn" ! -name "sh" -exec chmod -x {} +

COPY ./src ./src

USER appuser

EXPOSE ${PORT}
CMD sh -c "uvicorn src.main:app --host 0.0.0.0 --port ${PORT}"
