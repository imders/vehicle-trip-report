FROM python:3.10-slim

WORKDIR /app

# Установка системных зависимостей (если потребуется для pandas/openpyxl)
RUN apt-get update && apt-get install -y --no-install-recommends \
    && rm -rf /var/lib/apt/lists/*

# Копирование и установка python пакетов
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Копирование исходного кода
COPY . .

# Проброс порта (по умолчанию для FastAPI)
EXPOSE 8000

# Запуск приложения
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]
