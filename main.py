import pandas as pd
from config import INPUT_FILE

# Конфигурация (удобно вынести в отдельный блок)


def main():
    # Шаг 1: Чтение данных
    try:
        # Чтение Excel-файла. sheet_name - название или индекс листа
        df = pd.read_excel(INPUT_FILE, skiprows=3)
        print(f"Прочитано строк: {len(df)}")
        print(f"Количество столбцов в файле: {len(df.columns)}")
      
        # # Основная информация о DataFrame
        # print(df.shape)      # размерность (строки, столбцы)
        # print(df.columns)    # названия столбцов
        # print(df.info())     # подробная информация
        # print(df.head())     # первые 5 строк
        # print(df.tail())     # последние 5 строк
        # # Статистика по числовым столбцам
        # print(df.describe())

        column_10 = df[9]
        print(f"10-й столбец (индекс 9):")
        print(f"Тип данных: {column_10.dtype}")
        print(f"Уникальные значения: {column_10.unique()[:20]}")  # Показываем первые 20 уникал

        filtered_df = df[df[9].astype(str).str.lower() == 'Белогорск']   
        print(f"Найдено записей с 'белогорск': {len(filtered_df)}")

    except FileNotFoundError:
        print(f"Ошибка: Файл {INPUT_FILE} не найден.")
        return



if __name__ == '__main__':
    main()