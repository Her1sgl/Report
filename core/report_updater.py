import sys
import logging
import json
import os
import unicodedata
import re
from pathlib import Path
import pandas as pd
import xlwings as xw

CONFIG_FILE = "config.json"
DEFAULT_VALUE = 0

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    stream=sys.stdout,
)

def normalize_string(s: str) -> str:
    if not isinstance(s, str):
        return ""

    # Заменяем схожие символы (на кириллические заглавные, как в вашем примере)
    # Можно использовать и строчные, главное - единообразие
    char_map = {
        # Цифры и похожие буквы
        '3': 'З',
        '0': 'О', # Заменяем на заглавную О
        '6': 'б', # Пример для 6 и б
        # Латинские и похожие кириллические
        'a': 'а',
        'A': 'А', # Не забудьте про заглавные латинские
        'e': 'е',
        'E': 'Е',
        'o': 'о',
        'O': 'О',
        'p': 'р',
        'P': 'Р',
        'c': 'с',
        'C': 'С',
        'y': 'у',
        'Y': 'У',
        'x': 'х',
        'X': 'Х',
        'k': 'к',
        'K': 'К',
        't': 'т',
        'T': 'Т',
        'm': 'м',
        'M': 'М',
        'h': 'н',
        'H': 'Н',
        # Другие похожие буквы
        'ё': 'е', # Это тоже полезно сделать до lower()
    }

    for char, replacement in char_map.items():
        s = s.replace(char, replacement)

    # Теперь приводим к нижнему регистру
    s = s.lower()

    # Удаляем знаки препинания (или другие ненужные символы)
    s = re.sub(r'[^\w\s]', '', s)

    # Нормализуем пробелы
    s = re.sub(r'\s+', ' ', s).strip()

    # NFC нормализация Unicode
    s = unicodedata.normalize('NFC', s)

    return s

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [normalize_string(col) for col in df.columns]
    return df

def extract_surname(full_name: str) -> str:
    if not isinstance(full_name, str):
        return ""
    normalized = normalize_string(full_name)
    parts = normalized.split()
    if not parts:
        return ""
    surname = ''.join(c for c in parts[0] if c.isalpha())
    return surname.capitalize()

def column_letter_to_index(column_letter: str) -> int:
    index = 0
    for char in column_letter.upper():
        if not char.isalpha():
            raise ValueError(f"Недопустимый символ в обозначении столбца: {column_letter}")
        index = index * 26 + (ord(char) - ord('A') + 1)
    return index

def load_config() -> dict:
    config_paths = [
        Path(CONFIG_FILE),
        Path(__file__).parent / CONFIG_FILE,
        Path.home() / CONFIG_FILE
    ]
    
    for config_path in config_paths:
        try:
            if config_path.exists():
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                logging.info(f"Конфигурация загружена из: {config_path}")
                config["source_columns"] = {
                    key: normalize_string(value) 
                    for key, value in config["source_columns"].items()
                }
                return config
        except Exception as e:
            logging.warning(f"Ошибка при загрузке {config_path}: {e}")
            continue
    
    raise FileNotFoundError(f"Файл конфигурации '{CONFIG_FILE}' не найден")

def process_data(df: pd.DataFrame, config: dict) -> pd.DataFrame:
    df = normalize_columns(df)
    grouping_method = config.get("grouping_method", "region")
    source_cols = config["source_columns"]
    
    missing_columns = [col for col in source_cols.values() if col not in df.columns]
    if missing_columns:
        raise KeyError(f"Отсутствуют столбцы в исходных данных: {', '.join(missing_columns)}")
    
    if grouping_method == "manager":
        df["temp_surname"] = df[source_cols["manager"]].apply(extract_surname)
        normalized_mapping = {
            normalize_string(k): v 
            for k, v in config["manager_mapping"].items()
        }
        df["temp_sector"] = df["temp_surname"].apply(
            lambda x: normalized_mapping.get(normalize_string(x), None)
        )
        unknown = df[df["temp_sector"].isna()]["temp_surname"].unique()
        if len(unknown) > 0:
            logging.warning(f"Не распознаны менеджеры: {', '.join(unknown)}")
        
        result = df.groupby("temp_sector").agg({
            source_cols["bms_sales"]: "sum",
            source_cols["fms_sales"]: "sum"
        }).reset_index().rename(columns={"temp_sector": "Сектор"})
        
        return result
    
    elif grouping_method == "region":
        df["temp_region"] = df[source_cols["region"]].map(config["region_mapping"])
        unknown = df[df["temp_region"].isna()][source_cols["region"]].unique()
        if len(unknown) > 0:
            logging.warning(f"Не распознаны регионы: {', '.join(unknown)}")
        
        result = df.groupby("temp_region").agg({
            source_cols["bms_sales"]: "sum",
            source_cols["fms_sales"]: "sum"
        }).reset_index().rename(columns={"temp_region": "Регион"})
        
        return result
    
    else:
        raise ValueError(f"Неизвестный метод группировки: {grouping_method}")

def update_region_table(sheet, df, day, table_config, config):
    try:
        region_col = column_letter_to_index(table_config["region_col"])
        day_start = column_letter_to_index(table_config["day_start_col"])
        day_end = column_letter_to_index(table_config["day_end_col"])
        
        day_col = None
        for col in range(day_start, day_end + 1):
            cell_value = sheet.range((table_config["day_row"], col)).value
            if cell_value is not None and int(cell_value) == day:
                day_col = col
                break
        
        if not day_col:
            raise ValueError(f"Столбец для дня {day} не найден в таблице {table_config['name']}")
        
        sales_col = config["source_columns"]["bms_sales"] if table_config["type"] == "bms" else config["source_columns"]["fms_sales"]
        group_column = "Сектор" if config["grouping_method"] == "manager" else "Регион"
        df = df.set_index(group_column)
        
        for row in range(table_config["data_start_row"], table_config["data_end_row"] + 1):
            region_cell = sheet.range((row, region_col))
            region_name = str(region_cell.value).strip() if region_cell.value else None
            
            if not region_name or region_name == "None":
                continue
                
            if region_name in df.index:
                value = df.loc[region_name, sales_col]
                if pd.isna(value):
                    value = DEFAULT_VALUE
            else:
                value = DEFAULT_VALUE
                logging.warning(f"Не найден регион/сектор: {region_name}")
                
            sheet.range((row, day_col)).value = value
        
        logging.info(f"Таблица '{table_config['name']}' обновлена")
        return df.reset_index()
    except Exception as e:
        logging.error(f"Ошибка при обновлении таблицы '{table_config['name']}': {e}")
        return df.reset_index() if 'df' in locals() else None

def update_points_table(sheet, source_df, day, table_config, config):
    try:
        point_col = column_letter_to_index(table_config["point_col"])
        data_col = column_letter_to_index(table_config["data_col"])
        target_col = data_col + (day - 1)
        sales_col = config["source_columns"]["bms_sales"] if table_config["type"] == "bms" else config["source_columns"]["fms_sales"]
        
        source_df = source_df.copy()
        source_df["norm_point"] = source_df[config["source_columns"]["point"]].apply(normalize_string)
        
        # Получаем список пунктов из конфига
        point_names = table_config.get("point_names", [])
        normalized_point_names = [normalize_string(p) for p in point_names]
        
        for row in range(table_config["start_row"], table_config["end_row"] + 1):
            point_cell = sheet.range((row, point_col))
            point_name = str(point_cell.value).strip() if point_cell.value else None
            
            if not point_name or point_name == "None":
                continue
                
            normalized_point = normalize_string(point_name)
            value = DEFAULT_VALUE
            
            # Проверяем, есть ли пункт в списке из конфига
            if normalized_point in normalized_point_names:
                # Ищем в исходных данных
                matches = source_df[source_df["norm_point"] == normalized_point]
                if not matches.empty:
                    value = matches[sales_col].iloc[0]
                else:
                    # Проверяем вариант с "пп" перед названием
                    matches = source_df[source_df["norm_point"] == f"пп {normalized_point}"]
                    if not matches.empty:
                        value = matches[sales_col].iloc[0]
                    else:
                        logging.warning(f"Не найден пункт в данных: {point_name}")
            else:
                logging.warning(f"Пункт не найден в списке: {point_name}")
            
            if pd.isna(value):
                value = DEFAULT_VALUE
                
            sheet.range((row, target_col)).value = value
        
        logging.info(f"Таблица пунктов '{table_config['name']}' обновлена")
    except Exception as e:
        logging.error(f"Ошибка при обновлении таблицы пунктов '{table_config['name']}': {e}")

def update_report_sheet(report_path: str, sheet_name: str, input_file: str, day: int) -> bool:
    app = None
    try:
        config = load_config()
        source_df = pd.read_excel(input_file)
        source_df = normalize_columns(source_df)
        processed_df = process_data(source_df, config)
        
        app = xw.App(visible=False)
        wb = app.books.open(report_path)
        sheet = wb.sheets[sheet_name]
        
        if "region_tables" in config:
            for table in config["region_tables"]:
                processed_df = update_region_table(sheet, processed_df.copy(), day, table, config)
        
        if "new_points_tables" in config:
            for table in config["new_points_tables"]:
                update_points_table(sheet, source_df, day, table, config)
        
        wb.save()
        logging.info("Отчет успешно обновлен")
        return True
    except Exception as e:
        logging.error(f"Критическая ошибка: {e}", exc_info=True)
        return False
    finally:
        if app:
            app.quit()

def update_reports(input_file: str, report_file: str, sheet_name: str, day: int) -> bool:
    logging.info(f"Начало обновления отчета (день: {day})")
    result = update_report_sheet(report_file, sheet_name, input_file, day)
    if result:
        logging.info("Обновление завершено успешно")
    else:
        logging.error("Обновление завершено с ошибками")
    return result

def main():
    if len(sys.argv) != 5:
        print("Использование: python report_updater.py <входной_файл> <файл_отчета> <имя_листа> <день>")
        sys.exit(1)
    
    input_file = sys.argv[1]
    report_file = sys.argv[2]
    sheet_name = sys.argv[3]
    
    try:
        day = int(sys.argv[4])
        if day < 1 or day > 31:
            raise ValueError("День должен быть числом от 1 до 31")
    except ValueError as e:
        logging.error(f"Ошибка в параметре дня: {e}")
        sys.exit(1)
    
    try:
        logging.info(f"Начало обновления отчета (день: {day})")
        success = update_report_sheet(report_file, sheet_name, input_file, day)
        if success:
            logging.info("Обновление завершено успешно")
        else:
            logging.error("Обновление завершено с ошибками")
            sys.exit(1)
    except Exception as e:
        logging.error(f"Ошибка: {e}", exc_info=True)
        sys.exit(1)

if __name__ == "__main__":
    main()