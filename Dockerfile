FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1 \
    PYTHONIOENCODING=utf-8 \
    DATA_DIR=/tmp/receipt_data

WORKDIR /app

COPY webapp/requirements.txt webapp/requirements.txt
RUN pip install --no-cache-dir -r webapp/requirements.txt

# 기존 도구(영수증_정리_도구.py)와 웹 껍데기(webapp/)만 담는다.
# 샘플 영수증, 키 파일, exe 는 .dockerignore / .gitignore 로 제외된다.
COPY 영수증_정리_도구.py ./
COPY webapp ./webapp

WORKDIR /app/webapp
EXPOSE 10000
# Render 는 PORT 환경변수를 넣어 준다 (없으면 10000)
CMD ["sh", "-c", "uvicorn app:app --host 0.0.0.0 --port ${PORT:-10000}"]
