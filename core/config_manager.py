import json
import logging
from pathlib import Path

CONFIG_FILE = "config.json"
CURRENT_CONFIG_PATH = None

def load_config():
    global CURRENT_CONFIG_PATH
    config_paths = [
        Path(CONFIG_FILE),
        Path(__file__).parent.parent / CONFIG_FILE,
        Path.home() / CONFIG_FILE
    ]
    
    for config_path in config_paths:
        try:
            if config_path.exists():
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                logging.info(f"Конфигурация загружена из: {config_path}")
                CURRENT_CONFIG_PATH = config_path
                return config
        except Exception as e:
            logging.warning(f"Ошибка при загрузке {config_path}: {e}")
            continue
    
    default_config = create_default_config()
    logging.warning("Используется конфигурация по умолчанию")
    return default_config

def save_config(config, path=None):
    global CURRENT_CONFIG_PATH
    if path is None:
        if CURRENT_CONFIG_PATH:
            path = CURRENT_CONFIG_PATH
        else:
            path = CONFIG_FILE
    
    try:
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=4, ensure_ascii=False)
        logging.info(f"Конфигурация сохранена в: {path}")
        CURRENT_CONFIG_PATH = path
        return True
    except Exception as e:
        logging.error(f"Ошибка при сохранении конфигурации: {e}")
        return False

def validate_config(config: dict) -> list:
    errors = []
    required_sections = ["region_tables", "new_points_tables", "source_columns"]
    for section in required_sections:
        if section not in config:
            errors.append(f"Отсутствует обязательная секция: {section}")
    
    if "source_columns" in config:
        required_columns = ["region", "manager", "point", "bms_sales", "fms_sales"]
        for column in required_columns:
            if column not in config["source_columns"]:
                errors.append(f"Отсутствует обязательный столбец в source_columns: {column}")
    
    for table_type in ["region_tables", "new_points_tables"]:
        if table_type in config:
            for i, table in enumerate(config[table_type]):
                required_fields = ["name", "type"]
                if table_type == "region_tables":
                    required_fields += ["day_row", "data_start_row", "data_end_row", "region_col", "day_start_col", "day_end_col"]
                else:
                    required_fields += ["start_row", "end_row", "point_col", "data_col"]
                
                for field in required_fields:
                    if field not in table:
                        errors.append(f"Таблица {table.get('name', f'#{i}')}: отсутствует поле {field}")
    
    return errors

def create_default_config():
    return {
  "region_tables": [
    {
      "name": "ЛЧМ по секторам",
      "type": "bms",
      "day_row": 9,
      "data_start_row": 10,
      "data_end_row": 15,
      "region_col": "D",
      "day_start_col": "E",
      "day_end_col": "AI"
    },
    {
      "name": "ЛЦМ по секторам",
      "type": "fms",
      "day_row": 19,
      "data_start_row": 20,
      "data_end_row": 25,
      "region_col": "D",
      "day_start_col": "E",
      "day_end_col": "AI"
    }
  ],
  "new_points_tables": [
    {
      "name": "ЛЧМ новые пункты",
      "type": "bms",
      "point_names": [
        "НМУС1", "БСВЗ1", "ИКНТ1", "ОМЛД1", "НДЕК1", "НПСЧ1", "НСПР1", "ТСТП1", "НКРВ1", "ПРПН1",
        "ОЛНГ1", "НКУБ1", "НСТЦ4", "ИСВК1", "НКВЛ1", "БТПР1", "БККА1", "БЗПС1", "ББГА3", "АСКЛ1",
        "ББУГ1", "БСПП1", "БСТП1", "НЧУЛ1", "НДВР1", "НВЛЧ2", "БГОГ1", "ББОН1", "ООКТ1", "АКРП1",
        "НКРП1", "НПДН1", "НБМС1", "НАВТ1", "БПЮЖ1", "ИКВК1", "БМКО1", "БТЛФ1", "АЖЛД1", "БРКО1",
        "НМОТ1", "БХПП1", "БППП1", "НББГ1", "ТКРШ1", "БККМ1", "НКРС1", "АУСТ1", "ББЦЛ1", "БПОП3",
        "БТРК1", "КТЕР1", "БЗВЛ1", "НТЮМ1", "НСТЦ5", "ББЛГ1", "БПРЮ1", "БАЛК1", "КТРУ1", "ББКР1",
        "НВСТ1", "НСГВ1", "ОМСТ1", "НАРГ1", "НТЮМ2", "ББАЛ1", "БСЛВ1", "БРПО1", "БРБХ1", "НПРТ3",
        "НСВТ2", "ТГРЦ1", "ОЧРК1", "БКАМ1", "БРОМ1"
      ],
      "start_row": 30,
      "end_row": 104,
      "point_col": "D",
      "data_col": "E"
    },
    {
      "name": "ЛЦМ новые пункты",
      "type": "fms",
      "point_names": [
        "НМУС1", "БСВЗ1", "ИКНТ1", "ОМЛД1", "НДЕК1", "НПСЧ1", "НСПР1", "ТСТП1", "НКРВ1", "ПРПН1",
        "ОЛНГ1", "НКУБ1", "НСТЦ4", "ИСВК1", "НКВЛ1", "БТПР1", "БККА1", "БЗПС1", "ББГА3", "АСКЛ1",
        "ББУГ1", "БСПП1", "БСТП1", "НЧУЛ1", "НДВР1", "НВЛЧ2", "БГОГ1", "ББОН1", "ООКТ1", "АКРП1",
        "НКРП1", "НПДН1", "НБМС1", "НАВТ1", "БПЮЖ1", "ИКВК1", "БМКО1", "БТЛФ1", "АЖЛД1", "БРКО1",
        "НМОТ1", "БХПП1", "БППП1", "НББГ1", "ТКРШ1", "БККМ1", "НКРС1", "АУСТ1", "ББЦЛ1", "БПОП3",
        "БТРК1", "КТЕР1", "БЗВЛ1", "НТЮМ1", "НСТЦ5", "ББЛГ1", "БПРЮ1", "БАЛК1", "КТРУ1", "ББКР1",
        "НВСТ1", "НСГВ1", "ОМСТ1", "НАРГ1", "НТЮМ2", "ББАЛ1", "БСЛВ1", "БРПО1", "БРБХ1", "НПРТ3",
        "НСВТ2", "ТГРЦ1", "ОЧРК1", "БКАМ1", "БРОМ1"
      ],
      "start_row": 111,
      "end_row": 185,
      "point_col": "D",
      "data_col": "E"
    }
  ],
  "source_columns": {
    "region": "Регион",
    "manager": "РегМенеджер",
    "point": "Приёмный пункт",
    "bms_sales": "ЛЧМ Приём",
    "fms_sales": "ЛЦМ Приём"
  },
  "grouping_method": "manager",
  "manager_mapping": {
    "Абдукундузов": "Новосибирский (НС)",
    "Горбанев": "Новосибирский (НС)",
    "Швецов": "Новосибирский (НС)",
    "Шишов": "Новосибирский (НС)",
    "Шершнев": "Новосибирский (НС)",
    "Федорович": "Томский (ТС)",
    "Быков": "Омский (ОС)",
    "Чичельник": "Кемеровский (КС)",
    "Кезин": "Алтайский (АС)",
    "Кряжев": "Алтайский (АС)",
    "Калиниченко": "Алтайский (АС)",
    "Маджара": "Алтайский (АС)",
    "Сотник": "Алтайский (АС)",
    "Чернов": "Абаканский (АбС)"
  },
  "region_mapping": {
    "Алтай Респ": "Алтайский (АС)",
    "Алтайский край": "Алтайский (АС)",
    "Кемеровская область - Кузбасс обл": "Кемеровский (КС)",
    "Новосибирская обл": "Новосибирский (НС)",
    "Омская обл": "Омский (ОС)",
    "Томская обл": "Томский (ТС)",
    "Хакасия Респ": "Абаканский (АбС)"
  }
}