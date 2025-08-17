import os
import sys
import logging
from gui.main_window import MainApp

def setup_logging():
    """Настройка системы логирования"""
    log_dir = "logs"
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(os.path.join(log_dir, "app.log"), encoding='utf-8'),
            logging.StreamHandler(sys.stdout)
        ]
    )
    logging.info("=" * 50)
    logging.info("Запуск приложения Report Updater")
    logging.info("=" * 50)

def main():
    """Главная функция приложения"""
    setup_logging()
    try:
        app = MainApp()
        app.mainloop()
        logging.info("Приложение успешно завершено")
    except Exception as e:
        logging.critical(f"Критическая ошибка в приложении: {e}", exc_info=True)
        sys.exit(1)

if __name__ == "__main__":
    main()