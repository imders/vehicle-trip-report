import os
import shutil
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import HTMLResponse, StreamingResponse
import io

from logic import run_pipeline
from exporter import export_report_to_excel

app = FastAPI(title="Анализатор Рейсов")

BASE_FILE_PATH = "base.xlsx"

HTML_CONTENT = """
<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Анализатор Рейсов ТС</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #fce4ec; color: #333; margin: 0; padding: 20px; }
        .container { max-width: 800px; margin: 0 auto; background: white; padding: 30px; border-radius: 12px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
        h1 { color: #d81b60; text-align: center; }
        .card { background: #fdf2f6; padding: 20px; border-radius: 8px; margin-bottom: 20px; border-left: 4px solid #d81b60; }
        .btn { display: inline-block; background: #d81b60; color: white; padding: 10px 20px; border: none; border-radius: 6px; cursor: pointer; font-size: 16px; font-weight: bold; width: 100%; transition: background 0.3s; }
        .btn:hover { background: #ad1457; }
        input[type="file"] { display: block; margin-bottom: 15px; width: 100%; padding: 10px; border: 1px solid #ccc; border-radius: 4px; }
        .instructions { font-size: 14px; color: #555; background: #fff; padding: 15px; border: 1px dashed #d81b60; border-radius: 6px; }
        .spinner { display: none; border: 4px solid rgba(0, 0, 0, 0.1); border-left: 4px solid #d81b60; border-radius: 50%; width: 30px; height: 30px; animation: spin 1s linear infinite; margin: 10px auto; }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
    </style>
</head>
<body>
    <div class="container">
        <h1>🚚 Автоматический расчет рейсов ТС</h1>
        
        <div class="card">
            <h2>1. Инструкция по заполнению Справочника (base.xlsx)</h2>
            <div class="instructions">
                <p>Справочник используется для обогащения данных или исправления ошибок OCR распознавания номеров. Чтобы система корректно применяла справочник, файл <strong>base.xlsx</strong> должен иметь следующую структуру:</p>
                <ul>
                    <li><strong>Первая строка:</strong> Заголовки столбцов.</li>
                    <li><strong>Обязательно:</strong> Столбец с точным названием <code>Номер</code> (например, А123ВВ77). Без пробелов и в правильной раскладке.</li>
                    <li><strong>Дополнительные столбцы (по желанию):</strong> <code>Марка</code>, <code>Контрагент</code>. Если алгоритм найдет номер машины в справочнике, он возьмет эти значения оттуда.</li>
                </ul>
            </div>
            
            <form action="/upload-base" method="post" enctype="multipart/form-data">
                <label><strong>Обновить справочник на сервере (base.xlsx):</strong></label>
                <input type="file" name="file" accept=".xlsx, .xls" required>
                <button type="submit" class="btn" style="background: #880e4f;">Загрузить справочник</button>
            </form>
        </div>

        <div class="card">
            <h2>2. Генерация Отчета</h2>
            <p>Загрузите выгрузку системы видеонаблюдения для обработки и расчета рейсов.</p>
            <form id="processForm" action="/process" method="post" enctype="multipart/form-data">
                <label><strong>Файл выгрузки (.xlsx):</strong></label>
                <input type="file" name="file" accept=".xlsx, .xls" required>
                <button type="button" class="btn" onclick="submitProcessForm()">🚀 Сгенерировать отчет</button>
                <div class="spinner" id="spinner"></div>
            </form>
        </div>
    </div>

    <script>
        function submitProcessForm() {
            const form = document.getElementById('processForm');
            const spinner = document.getElementById('spinner');
            
            if (!form.file.value) {
                alert("Пожалуйста, выберите файл выгрузки.");
                return;
            }
            
            spinner.style.display = 'block';
            form.submit();
            // Скрываем спиннер через небольшую задержку, так как начнется скачивание файла
            setTimeout(() => { spinner.style.display = 'none'; }, 3000);
        }
    </script>
</body>
</html>
"""

@app.get("/", response_class=HTMLResponse)
async def read_root():
    """Отдает интерфейс."""
    return HTMLContent()

def HTMLContent():
    return HTML_CONTENT

@app.post("/upload-base")
async def upload_base(file: UploadFile = File(...)):
    """Обновляет файл base.xlsx на диске."""
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="Только файлы Excel (.xlsx, .xls) поддерживаются для базы.")
        
    try:
        with open(BASE_FILE_PATH, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        return HTMLResponse(content="<script>alert('Справочник успешно обновлен!'); window.location.href='/';</script>")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/process")
async def process_file(file: UploadFile = File(...)):
    """Обработка выгрузки с возвратом итогового EXCEL отчета."""
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="Только файлы Excel (.xlsx, .xls) поддерживаются.")

    contents = await file.read()
    
    # Чтение base.xlsx в байты, если он существует
    base_contents = None
    if os.path.exists(BASE_FILE_PATH):
        with open(BASE_FILE_PATH, "rb") as f:
            base_contents = f.read()
            
    try:
        # Выполнение бизнес-логики в памяти
        df_details, df_summary_corp, df_summary_cars, date_min, date_max = run_pipeline(io.BytesIO(contents), io.BytesIO(base_contents) if base_contents else None)
        
        if df_details.empty and df_summary_corp.empty:
            raise HTTPException(status_code=400, detail="Нет данных для формирования отчета после очистки и парсинга.")
            
        # Экспрот в Excel в памяти
        excel_bytes = export_report_to_excel(df_details, df_summary_corp, df_summary_cars, date_min, date_max)
        
        return StreamingResponse(
            io.BytesIO(excel_bytes),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename=report_trips.xlsx"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Ошибка обработки: {str(e)}")
