FROM python:3.12-slim

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY bot.py answers.json ./

# Baza /data papkasida saqlanadi (serverda shu papkani doimiy qiling)
ENV DB_PATH=/data/bot.db
VOLUME /data

CMD ["python", "bot.py"]
