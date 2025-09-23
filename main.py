import pandas as pd

# Конфигурация (удобно вынести в отдельный блок)
INPUT_FILE = 'data.xlsx'
OUTPUT_FILE = 'report.xlsx'
CRITERIA = {'Status': 'Approved'}  # Можно усложнить

def main():
    # Шаг 1: Чтение данных
    try:
        # Чтение Excel-файла. sheet_name - название или индекс листа
        df = pd.read_excel(INPUT_FILE)
        print(f"Прочитано строк: {len(df)}")
    except FileNotFoundError:
        print(f"Ошибка: Файл {INPUT_FILE} не найден.")
        return

if __name__ == '__main__':
    main()