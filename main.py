import os
import shutil
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import HTMLResponse, StreamingResponse
import io

from logic import run_pipeline
from exporter import export_report_to_excel

from pathlib import Path

app = FastAPI(title="Анализатор Рейсов")

BASE_FILE_PATH = "data/base.xlsx"
Path("data").mkdir(exist_ok=True)

HTML_CONTENT = """
<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Система анализа рейсов</title>
    <style>
        :root {
            --primary: #1e40af;
            --primary-hover: #1e3a8a;
            --bg: #f3f4f6;
            --surface: #ffffff;
            --text-main: #1f2937;
            --text-muted: #6b7280;
            --border: #e5e7eb;
            --border-hover: #9ca3af;
        }
        body {
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
            background-color: var(--bg);
            color: var(--text-main);
            margin: 0;
            padding: 40px 20px;
            display: flex;
            flex-direction: column;
            align-items: center;
            min-height: 100vh;
            box-sizing: border-box;
        }
        .container {
            width: 100%;
            max-width: 640px;
            background: var(--surface);
            padding: 40px;
            border-radius: 16px;
            box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.05), 0 8px 10px -6px rgba(0, 0, 0, 0.01);
        }
        h1 {
            font-size: 24px;
            font-weight: 700;
            margin-top: 0;
            margin-bottom: 8px;
            color: var(--text-main);
            text-align: center;
        }
        p.subtitle {
            text-align: center;
            color: var(--text-muted);
            margin-bottom: 32px;
            font-size: 15px;
        }
        .drop-zone {
            border: 2px dashed var(--border);
            border-radius: 12px;
            padding: 60px 20px;
            text-align: center;
            background: #fdfdfd;
            transition: all 0.2s ease;
            cursor: pointer;
            position: relative;
        }
        .drop-zone.dragover {
            border-color: var(--primary);
            background: #eff6ff;
        }
        .drop-zone input[type="file"] {
            position: absolute;
            width: 100%;
            height: 100%;
            top: 0;
            left: 0;
            opacity: 0;
            cursor: pointer;
        }
        .drop-zone-icon {
            font-size: 48px;
            color: var(--text-muted);
            margin-bottom: 16px;
            display: block;
        }
        .drop-zone-text {
            font-size: 16px;
            font-weight: 500;
            color: var(--text-main);
        }
        .drop-zone-subtext {
            font-size: 14px;
            color: var(--text-muted);
            margin-top: 8px;
        }
        .file-info {
            display: none;
            margin-top: 16px;
            padding: 12px;
            background: #f3f4f6;
            border-radius: 8px;
            font-size: 14px;
            color: var(--primary);
            font-weight: 500;
            text-align: center;
        }
        .btn {
            display: block;
            width: 100%;
            background: var(--primary);
            color: white;
            padding: 14px;
            border: none;
            border-radius: 8px;
            cursor: pointer;
            font-size: 16px;
            font-weight: 600;
            transition: background 0.2s;
            margin-top: 24px;
        }
        .btn:hover { background: var(--primary-hover); }
        .btn:disabled { background: #9ca3af; cursor: not-allowed; }
        
        .spinner {
            display: none;
            border: 3px solid rgba(255, 255, 255, 0.3);
            border-top: 3px solid white;
            border-radius: 50%;
            width: 20px;
            height: 20px;
            animation: spin 1s linear infinite;
            margin: 0 auto;
        }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }

        /* Hidden Settings for Base File */
        .settings-btn {
            position: fixed;
            bottom: 20px;
            left: 20px;
            background: none;
            border: none;
            color: var(--text-muted);
            cursor: pointer;
            opacity: 0.5;
            transition: opacity 0.2s;
            padding: 8px;
            font-size: 12px;
            text-decoration: underline;
        }
        .settings-btn:hover { opacity: 1; }
        .base-upload-modal {
            display: none;
            position: fixed;
            bottom: 60px;
            left: 20px;
            background: white;
            padding: 20px;
            border-radius: 12px;
            box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.1), 0 10px 10px -5px rgba(0, 0, 0, 0.04);
            border: 1px solid var(--border);
            width: 320px;
            z-index: 100;
        }
        .base-upload-modal.active { display: block; }
        .base-upload-modal h3 { margin-top: 0; font-size: 16px; margin-bottom: 10px; color: var(--text-main); }
        .base-upload-modal p { font-size: 12px; color: var(--text-muted); margin-bottom: 16px; }
        .base-upload-modal input[type="file"] { margin-bottom: 12px; width: 100%; font-size: 12px; }
        .base-btn {
            background: var(--text-main); color: white; padding: 8px 12px; border-radius: 6px; 
            border: none; cursor: pointer; width: 100%; font-size: 13px; font-weight: 500;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>Анализ рейсов ТС</h1>
        <p class="subtitle">Генерация аналитических отчетов по данным камер</p>
        
        <form id="processForm" action="/process" method="post" enctype="multipart/form-data">
            <div class="drop-zone" id="dropZone">
                <input type="file" name="file" id="fileInput" accept=".xlsx, .xls" required>
                <span class="drop-zone-icon">📄</span>
                <div class="drop-zone-text">Перетащите файл сюда или нажмите для выбора</div>
                <div class="drop-zone-subtext">Поддерживаются форматы Excel (.xlsx, .xls)</div>
            </div>
            <div class="file-info" id="fileInfo"></div>
            
            <button type="button" class="btn" id="submitBtn" onclick="submitProcessForm()">
                <span id="btnText">Сгенерировать отчет</span>
                <div class="spinner" id="spinner"></div>
            </button>
        </form>
    </div>

    <!-- Обновление справочника -->
    <button class="settings-btn" onclick="toggleBaseModal()">⚙️ Данные справочника</button>
    <div class="base-upload-modal" id="baseModal">
        <h3>Обновление справочника (base.xlsx)</h3>
        <p>Файл используется для обогащения данных. Заголовки (строка 1): <b>Номер</b>, <i>Марка</i>, <i>Контрагент</i>.</p>
        <form action="/upload-base" method="post" enctype="multipart/form-data">
            <input type="file" name="file" accept=".xlsx, .xls" required>
            <button type="submit" class="base-btn">Загрузить справочник</button>
        </form>
    </div>

    <script>
        const dropZone = document.getElementById('dropZone');
        const fileInput = document.getElementById('fileInput');
        const fileInfo = document.getElementById('fileInfo');

        // Подсветка зоны при drag & drop
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            dropZone.addEventListener(eventName, preventDefaults, false);
        });
        function preventDefaults(e) { e.preventDefault(); e.stopPropagation(); }

        ['dragenter', 'dragover'].forEach(eventName => {
            dropZone.addEventListener(eventName, () => dropZone.classList.add('dragover'), false);
        });
        ['dragleave', 'drop'].forEach(eventName => {
            dropZone.addEventListener(eventName, () => dropZone.classList.remove('dragover'), false);
        });

        // Обработка файла при сбросе
        dropZone.addEventListener('drop', handleDrop, false);
        function handleDrop(e) {
            let dt = e.dataTransfer;
            let files = dt.files;
            fileInput.files = files;
            updateFileInfo();
        }

        fileInput.addEventListener('change', updateFileInfo);

        function updateFileInfo() {
            if (fileInput.files.length > 0) {
                fileInfo.textContent = 'Выбран файл: ' + fileInput.files[0].name;
                fileInfo.style.display = 'block';
            } else {
                fileInfo.style.display = 'none';
            }
        }

        function submitProcessForm() {
            if (!fileInput.files || fileInput.files.length === 0) {
                alert("Пожалуйста, выберите файл выгрузки.");
                return;
            }
            
            const btn = document.getElementById('submitBtn');
            const btnText = document.getElementById('btnText');
            const spinner = document.getElementById('spinner');
            
            btn.disabled = true;
            btnText.style.display = 'none';
            spinner.style.display = 'block';
            
            document.getElementById('processForm').submit();
            
            // Восстанавливаем кнопку(примерно) т.к. файл будет загружен как attachment
            setTimeout(() => { 
                btn.disabled = false;
                btnText.style.display = 'block';
                spinner.style.display = 'none';
            }, 5000);
        }

        function toggleBaseModal() {
            const modal = document.getElementById('baseModal');
            modal.classList.toggle('active');
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
