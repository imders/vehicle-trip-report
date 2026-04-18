import pandas as pd
import pytest
from logic import clean_and_validate_data, parse_russian_datetime, assign_trips

def test_clean_and_validate_volume():
    """Тест очистки столбца Объема (замена запятых, приведение к float)."""
    df = pd.DataFrame({
        "Объем": ["12,5", "10", None, "NaN", "текст"]
    })
    
    cleaned = clean_and_validate_data(df)
    
    assert cleaned["Объем"].iloc[0] == 12.5
    assert cleaned["Объем"].iloc[1] == 10.0
    assert cleaned["Объем"].iloc[2] == 0.0
    assert cleaned["Объем"].iloc[3] == 0.0
    assert cleaned["Объем"].iloc[4] == 0.0

def test_parse_russian_datetime():
    """Тест парсинга русских месяцев."""
    df = pd.DataFrame({
        "Дата и время": ["10 янв. 2024 15:30:00", "05 май 2024 08:15:00", "не дата"]
    })
    
    parsed = parse_russian_datetime(df)
    
    # "не дата" должна быть дропнута
    assert len(parsed) == 2
    assert parsed["Дата и время"].iloc[0] == pd.Timestamp("2024-01-10 15:30:00")
    assert parsed["Дата и время"].iloc[1] == pd.Timestamp("2024-05-05 08:15:00")

def test_assign_trips_simple_gap():
    """Тест разбивки на рейсы по прошествии TRIP_GAP_MINUTES (по умолчанию 20 мин)."""
    # 2 записи для машины А: разница 5 минут -> 1 рейс. Третья запись через 25 минут -> 2 рейс.
    df = pd.DataFrame({
        "Номер": ["A111AA", "A111AA", "A111AA"],
        "Дата и время": [
            pd.Timestamp("2024-01-01 10:00:00"),
            pd.Timestamp("2024-01-01 10:05:00"),
            pd.Timestamp("2024-01-01 10:30:00")
        ]
    })
    
    trips = assign_trips(df)
    
    # 1 запись: Рейс 1
    # 2 запись: Рейс 1
    # 3 запись: Рейс 2
    assert trips["Рейс"].iloc[0] == 1
    assert trips["Рейс"].iloc[1] == 1
    assert trips["Рейс"].iloc[2] == 2
