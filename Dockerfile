FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1

WORKDIR /app

COPY requirements.txt ./
RUN pip install --upgrade pip \
    && pip install -r requirements.txt

COPY app.py ./
COPY static ./static
COPY templates ./templates
COPY 模板文件 ./模板文件

RUN mkdir -p /app/data/tmp /app/data/projects /app/data/market_skill /app/data/feishu

VOLUME ["/app/data"]

EXPOSE 8008

CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8008"]
