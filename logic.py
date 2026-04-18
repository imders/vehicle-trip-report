import logging
import datetime
import io
from pathlib import Path

import pandas as pd

logger = logging.getLogger(__name__)

# --- Константы столбцов ---
COL_DATETIME = "Дата и время"
COL_PLATE = "Номер"
COL_CAMERA = "Камера"  # может быть аналогично Точке учета
COL_DIRECTION = "Направление"
COL_POINT = "Точка учета"
COL_MAKE = "Марка"
COL_CORP = "Контрагент"
COL_VOLUME = "Объем"
COL_ARCHIVE_LINK = "Ссылка на архив"

# --- Настройки рейсов ---
TRIP_GAP_MINUTES = 20             
MIN_TRIP_DURATION_MINUTES = 3     
MAX_TRIP_DURATION_MINUTES = 720   

RUSSIAN_MONTHS = {
    "янв.": "01", "февр.": "02", "мар.": "03", "апр.": "04",
    "май": "05", "июн.": "06", "июл.": "07", "авг.": "08",
    "сент.": "09", "окт.": "10", "нояб.": "11", "дек.": "12",
}

def clean_and_validate_data(df: pd.DataFrame) -> pd.DataFrame:
    """Очистка и валидация сырых данных."""
    df = df.copy()
    
    # Гарантируем наличие базовых колонок, если их нет, создаем с NaN
    for col in [COL_POINT, COL_MAKE, COL_CORP, COL_VOLUME, COL_ARCHIVE_LINK]:
        if col not in df.columns:
             # Попробуем маппинги (например Камера -> Точка учета)
             if col == COL_POINT and COL_CAMERA in df.columns:
                 df[COL_POINT] = df[COL_CAMERA]
             else:
                 df[col] = pd.NA
                 
    # 1. Очистка 'Объема' (в float, запятые в точки, NaN в 0.0)
    def clean_volume(val):
        if pd.isna(val):
            return 0.0
        if isinstance(val, str):
            val = val.replace(",", ".").replace(" ", "")
        try:
            return float(val)
        except ValueError:
            return 0.0
            
    df[COL_VOLUME] = df[COL_VOLUME].apply(clean_volume)
    
    # 2. Очистка Контрагента (NaN -> "Неизвестно")
    df[COL_CORP] = df[COL_CORP].fillna("Неизвестно").astype(str)
    # Замена пустых строк и "nan"
    df.loc[df[COL_CORP] == 'nan', COL_CORP] = "Неизвестно"
    df.loc[df[COL_CORP].str.strip() == '', COL_CORP] = "Неизвестно"
    
    # Очистка Марки
    df[COL_MAKE] = df[COL_MAKE].fillna("Неизвестно").astype(str)
    df.loc[df[COL_MAKE] == 'nan', COL_MAKE] = "Неизвестно"
    
    # Очистка Точки учета
    df[COL_POINT] = df[COL_POINT].fillna("Без точки учета").astype(str)
    
    return df

def parse_russian_datetime(df: pd.DataFrame) -> pd.DataFrame:
    """Парсинг дат с русскими названиями месяцев."""
    df = df.copy()
    if COL_DATETIME not in df.columns:
        logger.error(f"Столбец '{COL_DATETIME}' не найден в данных.")
        return df

    def replace_months(val):
        if pd.isna(val):
            return val
        s = str(val).lower()
        for ru_mo, num_mo in RUSSIAN_MONTHS.items():
            if ru_mo in s:
                s = s.replace(ru_mo, num_mo)
                break
        return s
    
    df['_temp_date_str'] = df[COL_DATETIME].apply(replace_months)
    df[COL_DATETIME] = pd.to_datetime(df['_temp_date_str'], format="%d %m %Y %H:%M:%S", errors='coerce')
    df = df.drop(columns=['_temp_date_str'])
    
    nat_count = df[COL_DATETIME].isna().sum()
    if nat_count > 0:
        logger.warning(f"Пропущено {nat_count} дат.")
        df = df.dropna(subset=[COL_DATETIME])
        
    return df

def fix_ocr_errors(df: pd.DataFrame, df_base: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Исправление OCR ошибок по временным окнам (Левенштейн)."""
    df = df.copy()
    df_base = df_base.copy()
    
    cyrillic_to_latin = str.maketrans('АВЕКМНОРСТУХ', 'ABEKMHOPCTYX')
    df[COL_PLATE] = df[COL_PLATE].astype(str).str.upper().str.translate(cyrillic_to_latin).str.replace(' ', '')
    if COL_PLATE in df_base.columns:
        df_base[COL_PLATE] = df_base[COL_PLATE].astype(str).str.upper().str.translate(cyrillic_to_latin).str.replace(' ', '')
    
        base_plates = set(df_base[COL_PLATE].unique())
    else:
        base_plates = set()
        
    plate_counts = df[COL_PLATE].value_counts().to_dict()
    df = df.sort_values(COL_DATETIME)
    
    new_plates = df[COL_PLATE].tolist()
    times = df[COL_DATETIME].tolist()
    fixes_applied = 0
    
    for i in range(len(new_plates)):
        if pd.isna(times[i]):
            continue
            
        best_plate = new_plates[i]
        best_score = 0
        if best_plate in base_plates: best_score += 1000
        best_score += len(best_plate) * 10
        best_score += plate_counts.get(best_plate, 0)
        
        similar_indices = [i]
        j = i + 1
        while j < len(new_plates) and pd.notna(times[j]) and (times[j] - times[i]).total_seconds() <= 120:
            similar_indices.append(j)
            j += 1
            
        j = i - 1
        while j >= 0 and pd.notna(times[j]) and (times[i] - times[j]).total_seconds() <= 120:
            similar_indices.append(j)
            j -= 1
            
        for idx in similar_indices:
            p2 = new_plates[idx]
            if best_plate == p2:
                continue
                
            is_similar = False
            def levenshtein(s1, s2):
                if len(s1) < len(s2): return levenshtein(s2, s1)
                if len(s2) == 0: return len(s1)
                prev_row = range(len(s2) + 1)
                for i1, c1 in enumerate(s1):
                    curr_row = [i1 + 1]
                    for j1, c2 in enumerate(s2):
                        curr_row.append(min(prev_row[j1 + 1] + 1, curr_row[j1] + 1, prev_row[j1] + (c1 != c2)))
                    prev_row = curr_row
                return prev_row[-1]

            if (len(best_plate) >= 4 and best_plate in p2) or (len(p2) >= 4 and p2 in best_plate):
                is_similar = True
            elif levenshtein(best_plate, p2) <= 2:
                is_similar = True
                
            if is_similar:
                if best_plate in base_plates and p2 in base_plates and best_plate != p2:
                    continue
                score_p2 = 0
                if p2 in base_plates: score_p2 += 1000
                score_p2 += len(p2) * 10
                score_p2 += plate_counts.get(p2, 0)
                
                if score_p2 > best_score:
                    best_score = score_p2
                    best_plate = p2
                    
        if new_plates[i] != best_plate:
            fixes_applied += 1
        new_plates[i] = best_plate

    df[COL_PLATE] = new_plates
    logger.info(f"Сшито ошибок: {fixes_applied}")
    return df.sort_index(), df_base

def assign_trips(df: pd.DataFrame) -> pd.DataFrame:
    """Выделение рейсов."""
    if df.empty: return df
    df = df.copy()
    
    df = df.sort_values(by=[COL_PLATE, COL_DATETIME]).reset_index(drop=True)
    df['prev_time'] = df.groupby(COL_PLATE)[COL_DATETIME].shift(1)
    df['delta_minutes'] = (df[COL_DATETIME] - df['prev_time']).dt.total_seconds() / 60
    
    if COL_DIRECTION in df.columns:
        prev_dir = df.groupby(COL_PLATE)[COL_DIRECTION].shift(1).str.lower().fillna("")
        curr_dir = df[COL_DIRECTION].str.lower().fillna("")
        is_on_site = prev_dir.str.contains('въезд') & curr_dir.str.contains('выезд')
        gap_exceeded = df['delta_minutes'] > TRIP_GAP_MINUTES
        df['new_trip'] = (
            df['prev_time'].isna() | 
            (gap_exceeded & ~is_on_site) | 
            (df['delta_minutes'] > MAX_TRIP_DURATION_MINUTES)
        )
    else:
        df['new_trip'] = df['prev_time'].isna() | (df['delta_minutes'] > TRIP_GAP_MINUTES)
    
    df['Рейс'] = df.groupby(COL_PLATE)['new_trip'].cumsum().astype(int)
    return df

def _format_duration(td: pd.Timedelta) -> str:
    """Вспомогательная функция: Конвертирует Timedelta в строку HH:MM."""
    if pd.isna(td):
        return "00:00"
    ts = int(td.total_seconds())
    return f"{ts // 3600:02d}:{(ts % 3600) // 60:02d}"

def calculate_trip_metrics(df: pd.DataFrame) -> pd.DataFrame:
    """Сборка рейсов по метрикам, включая бизнес-поля (Точка учета, Объем, и др.)."""
    if df.empty: return pd.DataFrame()
    
    # Для бизнес показателей (Объем берем как сумму за рейс, остальные поля как first())
    # Так как за рейс точка учета вряд ли меняется (если меняется - берем первую)
    
    agg_funcs = {
        COL_DATETIME: ['min', 'max', 'count'],
        COL_VOLUME: 'sum',
        COL_POINT: 'first',
        COL_MAKE: 'first',
        COL_CORP: 'first',
        COL_ARCHIVE_LINK: ['first', 'last']
    }
    
    # Оставляем только те поля, что реально есть (заглушки сделали в clean_and_validate)
    grouped = df.groupby([COL_PLATE, 'Рейс'])
    metrics = grouped.agg(agg_funcs)
    
    metrics.columns = [
        'Дата начала', 'Дата окончания', 'Всего событий',
        'Суммарный Объем', 'Точка учета (первая)', 'Марка (первая)', 'Контрагент (первый)',
        'Ссылка начало', 'Ссылка конец'
    ]
    
    if COL_DIRECTION in df.columns:
        direction_counts = df.groupby([COL_PLATE, 'Рейс', COL_DIRECTION]).size().unstack(fill_value=0)
        in_col = 'Въезд'
        out_col = 'Выезд'
        
        metrics['Въездов (кол-во)'] = direction_counts[in_col] if in_col in direction_counts.columns else 0
        metrics['Выездов (кол-во)'] = direction_counts[out_col] if out_col in direction_counts.columns else 0
    else:
        metrics['Въездов (кол-во)'] = 0
        metrics['Выездов (кол-во)'] = 0
    
    metrics['Время на объекте_td'] = metrics['Дата окончания'] - metrics['Дата начала']
    metrics['Время на объекте'] = metrics['Время на объекте_td'].apply(_format_duration)
    metrics['duration_minutes'] = metrics['Время на объекте_td'].dt.total_seconds() / 60
    metrics = metrics.drop(columns=['Время на объекте_td'])
    
    # Статус
    def assign_status(row):
        dur = row['duration_minutes']
        if dur < MIN_TRIP_DURATION_MINUTES:
            return "Валидный рейс (одиночный)" if row['Всего событий'] == 1 else f"Валидный рейс (< {MIN_TRIP_DURATION_MINUTES} мин)"
        if dur > MAX_TRIP_DURATION_MINUTES:
            return f"Отбраковка: > {MAX_TRIP_DURATION_MINUTES} мин"
        return "Валидный рейс"
        
    metrics['Статус'] = metrics.apply(assign_status, axis=1)
    
    metrics = metrics.reset_index().rename(columns={'Рейс': 'Рейс №'})
    
    valid_mask = metrics['Статус'].str.startswith('Валидный')
    metrics['Рейс №_orig'] = metrics['Рейс №']
    
    # Pre-allocate object column to avoid FutureWarning on incompatible dtypes
    new_reis = pd.Series("—", index=metrics.index, dtype=object)
    new_reis[valid_mask] = metrics[valid_mask].groupby(COL_PLATE).cumcount() + 1
    metrics['Рейс №'] = new_reis
    
    return metrics

def enrich_with_base(df_trips: pd.DataFrame, df_base: pd.DataFrame) -> pd.DataFrame:
    """Комбинирование с базой, приоритет у полей из базы."""
    if df_trips.empty or df_base.empty:
        return df_trips
        
    if COL_PLATE not in df_base.columns:
        return df_trips
        
    df_base = df_base.drop_duplicates(subset=[COL_PLATE], keep='first')
    
    df_merged = pd.merge(df_trips, df_base, on=COL_PLATE, how='left')
    
    # Если в базе была Марка, Контрагент или Объем, и они не NaN -> используем их
    mapping = {
        COL_MAKE: COL_MAKE + " (первая)",
        COL_CORP: COL_CORP + " (первый)",
        COL_VOLUME: "Суммарный Объем"
    }
    
    for base_col, target_col in mapping.items():
        if base_col in df_merged.columns:
            mask = df_merged[base_col].notna() & (df_merged[base_col] != "")
            
            # Для Объема нужна конвертация во float
            if base_col == COL_VOLUME:
                 values_float = pd.to_numeric(
                     df_merged.loc[mask, base_col].astype(str).str.replace(",", ".").str.replace(" ", ""),
                     errors='coerce'
                 ).fillna(0).astype('float64')
                 df_merged.loc[mask, target_col] = values_float
            else:
                 df_merged.loc[mask, target_col] = df_merged.loc[mask, base_col]
                 
            df_merged = df_merged.drop(columns=[base_col])
            
    return df_merged

def build_final_dataframes(df_report: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, str, str]:
    """Формирование Детализации и Итого (по машинам и по контрагентам)."""
    # Детализация (плоская таблица)
    col_mapping = {
        COL_PLATE: 'Номер машины',
        'Марка (первая)': 'Марка',
        'Суммарный Объем': 'Объем',
        'Контрагент (первый)': 'Контрагент',
        'Точка учета (первая)': 'Точка учета',
        'Рейс №': 'Рейс №',
        'Дата начала': 'Дата начала',
        'Дата окончания': 'Дата окончания',
        'Время на объекте': 'Время на объекте',
        'Статус': 'Статус рейса',
        'Всего событий': 'Всего событий',
        'Въездов (кол-во)': 'Въездов (кол-во)',
        'Выездов (кол-во)': 'Выездов (кол-во)',
        'Ссылка начало': 'Ссылка на архив (начало)',
        'Ссылка конец': 'Ссылка на архив (конец)',
    }
    
    available_cols = {k: v for k, v in col_mapping.items() if k in df_report.columns}
    df_details = df_report[list(available_cols.keys())].rename(columns=available_cols)
    
    # Расчет дат для отчета
    if 'Дата начала' in df_report.columns and not df_report['Дата начала'].isna().all():
        date_min = df_report['Дата начала'].min().strftime('%d.%m.%Y %H:%M')
        date_max = df_report['Дата начала'].max().strftime('%d.%m.%Y %H:%M')
    else:
        date_min, date_max = "Нет данных", "Нет данных"

    valid_details = df_details[df_details['Статус рейса'].astype(str).str.startswith('Валидный')]
    
    if not valid_details.empty:
        valid_details = valid_details.copy()
        valid_details['Машина'] = valid_details['Номер машины'] + " (" + valid_details['Марка'] + ")"
        
        # Сводка по машинам
        df_summary_cars = valid_details.groupby(['Контрагент', 'Машина']).agg(
            Количество_рейсов=('Рейс №', 'count'),
            Сумма_объем=('Объем', 'sum')
        ).reset_index()
        df_summary_cars = df_summary_cars.rename(columns={
            'Количество_рейсов': 'Количество рейсов (валидных)',
            'Сумма_объем': 'Общий объем'
        })
        # Сортировка по убыванию объема
        df_summary_cars = df_summary_cars.sort_values(by='Общий объем', ascending=False).reset_index(drop=True)
        
        # Сводка по контрагентам (Общий итог)
        df_summary_corp = valid_details.groupby(['Контрагент']).agg(
            Количество_рейсов=('Рейс №', 'count'),
            Сумма_объем=('Объем', 'sum')
        ).reset_index()
        df_summary_corp = df_summary_corp.rename(columns={
            'Количество_рейсов': 'Всего рейсов',
            'Сумма_объем': 'Всего объем'
        })
        # Сортировка по убыванию объема
        df_summary_corp = df_summary_corp.sort_values(by='Всего объем', ascending=False).reset_index(drop=True)
        
        # Добавим финальный общий итог по всем
        df_summary_corp.loc[len(df_summary_corp)] = {
            'Контрагент': 'ИТОГО ПО ВСЕМ',
            'Всего рейсов': df_summary_corp['Всего рейсов'].sum(),
            'Всего объем': df_summary_corp['Всего объем'].sum()
        }
        
    else:
        df_summary_cars = pd.DataFrame(columns=['Контрагент', 'Машина', 'Количество рейсов (валидных)', 'Общий объем'])
        df_summary_corp = pd.DataFrame(columns=['Контрагент', 'Всего рейсов', 'Всего объем'])
        
    return df_details, df_summary_corp, df_summary_cars, date_min, date_max

def run_pipeline(raw_file: bytes | str, base_file: bytes | str | None) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, str, str]:
    """Основной пайплайн."""
    if isinstance(raw_file, bytes):
        raw_file = io.BytesIO(raw_file)
    if isinstance(base_file, bytes):
        base_file = io.BytesIO(base_file)

    # 1. Чтение файлов
    df_raw = pd.read_excel(raw_file, sheet_name=0, engine='openpyxl')
    
    # Извлечение реальных гиперссылок из столбца J (индекс 10 в openpyxl)
    try:
        import openpyxl
        if isinstance(raw_file, io.BytesIO):
            raw_file.seek(0)
            
        wb_raw = openpyxl.load_workbook(raw_file, data_only=True)
        ws_raw = wb_raw.worksheets[0]
        
        links = []
        for row in range(2, 2 + len(df_raw)):
            cell = ws_raw.cell(row=row, column=10)
            text_val = str(cell.value) if cell.value is not None else ""
            if cell.hyperlink and cell.hyperlink.target:
                links.append(f"{cell.hyperlink.target}|||{text_val}")
            else:
                links.append(text_val)
                
        df_raw[COL_ARCHIVE_LINK] = links
    except Exception as e:
        logger.error(f"Не удалось извлечь гиперссылки из столбца J: {e}")
        # Запасной вариант, берем текст
        if len(df_raw.columns) > 9:
            df_raw[COL_ARCHIVE_LINK] = df_raw.iloc[:, 9]
    
    df_base = pd.DataFrame()
    if base_file is not None:
        try:
            df_base = pd.read_excel(base_file, sheet_name=0, engine='openpyxl')
            if COL_PLATE not in df_base.columns:
                 # Пытаемся использовать первую колонку как Номер, если нет 'Номер'
                 df_base.rename(columns={df_base.columns[0]: COL_PLATE}, inplace=True)
        except Exception as e:
            logger.error(f"Ошибка чтения базы: {e}")
            
    # 2. Очистка
    df_raw = clean_and_validate_data(df_raw)
    df_cleaned = parse_russian_datetime(df_raw)
    
    if df_cleaned.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), "", ""
        
    # 3. OCR и рейсы
    df_cleaned, df_base = fix_ocr_errors(df_cleaned, df_base)
    df_trips = assign_trips(df_cleaned)
    df_metrics = calculate_trip_metrics(df_trips)
    
    # 4. Сборка и итого
    df_report = enrich_with_base(df_metrics, df_base)
    return build_final_dataframes(df_report)
