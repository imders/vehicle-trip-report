import sys
import logging
import datetime
from pathlib import Path

from logic import run_pipeline
from exporter import export_report_to_excel

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

BASE_FILENAME = "base.xlsx"

def main(input_filepath: str, base_filepath: str = BASE_FILENAME):
    """CLI Точка входа для запуска расчета из консоли."""
    logger.info(f"Запуск скрипта для файла: {input_filepath}")
    
    path = Path(input_filepath)
    if not path.exists():
        logger.error(f"Файл не найден: {input_filepath}")
        print(f"Ошибка: входной файл не найден ({input_filepath})")
        sys.exit(1)
        
    base_path = Path(base_filepath)
    if not base_path.exists():
        logger.warning(f"Файл внешней базы '{base_filepath}' не найден. Обогащение данных справочником не будет выполнено.")
        base_path = None
        
    try:
        with open(input_filepath, "rb") as f_raw:
            raw_bytes = f_raw.read()
            
        base_bytes = None
        if base_path:
            with open(base_path, "rb") as f_base:
                base_bytes = f_base.read()
                
        # Используем инкапсулированную логику расчета
        df_details, df_summary_corp, df_summary_cars, date_min, date_max = run_pipeline(raw_bytes, base_bytes)
        
        if df_details.empty and df_summary_corp.empty:
            logger.warning("После выполнения алгоритма данные для отчета отсутствуют.")
            sys.exit(0)
            
        timestamp_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"report_trips_{timestamp_str}.xlsx"
        
        # Используем экспортер для сохранения
        export_report_to_excel(df_details, df_summary_corp, df_summary_cars, date_min, date_max, output_file)
        
        logger.info(f"Успешно завершено! Файл сохранен: {output_file}")
        
    except KeyError as e:
        logger.error(f"Отсутствует требуемый элемент: {e}")
        sys.exit(1)
    except Exception as e:
        logger.exception(f"Непредвиденная ошибка: {e}")
        sys.exit(1)

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Использование: python script.py <путь/к/выгрузке.xlsx> [путь/к/базе.xlsx]")
        sys.exit(1)
        
    input_file = sys.argv[1]
    base_file = sys.argv[2] if len(sys.argv) > 2 else BASE_FILENAME
    
    main(input_file, base_file)
