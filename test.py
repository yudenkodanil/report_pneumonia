import pandas as pd
from openpyxl import load_workbook
import config
import logging
from rich.console import Console
from rich.table import Table
from rich.logging import RichHandler

# --- Настройка логирования ---
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[RichHandler(console=Console(), rich_tracebacks=True)]
)
logger = logging.getLogger("population_analysis")
console = Console()


# --- Функции обработки ---
def preprocess_file(input_file: str) -> pd.DataFrame | None:
    """Загрузка и предобработка Excel-файла"""
    try:
        df = pd.read_excel(input_file, header=None)
        logger.info(f"Прочитано строк: {len(df)}, столбцов: {len(df.columns)}")

        # Удаляем первые три строки
        df = df.drop(index=[0, 1, 2]).reset_index(drop=True)

        # Ограничиваемся 39 столбцами
        df = df.iloc[:, :39]

        # Переименовываем столбцы
        if config.COLUMN_NAMES:
            df.columns = config.COLUMN_NAMES[:len(df.columns)]
        else:
            df.columns = [f"col_{i}" for i in range(1, len(df.columns) + 1)]

        # Удаляем столбцы с None
        df = df.loc[:, df.columns.notna()]

        # Заполняем соц. статус: если столбец L пустой, берем значение из столбца K
        if "Социальный статус" in df.columns and df.columns.get_loc("Социальный статус") >= 0:
            idx_status = df.columns.get_loc("Социальный статус")
            df.iloc[:, idx_status] = df.iloc[:, idx_status].fillna(df.iloc[:, idx_status-1])

        # Преобразуем даты в datetime
        df['Дата рождения'] = pd.to_datetime(df['Дата рождения'], errors='coerce')
        df['Дата подачи ЭИ'] = pd.to_datetime(df['Дата подачи ЭИ'], errors='coerce')

        # Считаем возраст в годах
        df['Возраст'] = (df['Дата подачи ЭИ'] - df['Дата рождения']).dt.days / 365.25

        df.to_excel("df_filtred.xlsx", sheet_name="main", index=False)
        logger.info("Файл успешно предобработан и экспортирован: df_filtred.xlsx")
        return df

    except FileNotFoundError:
        logger.error(f"Файл {input_file} не найден.")
        return None


def analyze_population_by_district(df: pd.DataFrame, districts: list) -> dict:
    """Анализ возрастной структуры по районам"""
    logger.info("Начало анализа возрастной структуры")
    results = {}

    for district in districts:
        df_district = df[df['Административная территория (город)'] == district]
        if df_district.empty:
            logger.warning(f"Нет данных для района: {district}")
            continue

        counts = {}
        for (low, high), label in zip(config.AGE_GROUP, config.AGE_GROUP_NAME):
            if label == "65 и старше":
                counts[label] = df_district[df_district['Возраст'] >= low].shape[0]
            else:
                counts[label] = df_district[(df_district['Возраст'] >= low) & (df_district['Возраст'] < high)].shape[0]

        results[district] = counts
        logger.info(f"Обработан район: {district}")

    return results


def classify_status(status: str, age: float | None = None) -> str:
    """Классифицирует социальный статус через конфиг"""
    status = str(status).strip().lower()

    # проверяем прямое соответствие через SOCIAL_GROUPS
    for group_name, keywords in config.SOCIAL_GROUPS.items():
        if status in [kw.lower() for kw in keywords]:
            return group_name

    # проверка по ключевым словам
    for keyword, group in config.KEYWORD_MAPPING.items():
        if keyword.lower() in status:
            return group

    # # проверка по возрасту
    # if age is not None:
    #     try:
    #         age = float(age)
    #         for group_def in config.AGE_GROUPS_BY_STATUS:
    #             min_age = group_def["min_age"]
    #             max_age = group_def["max_age"]
    #             if max_age is None:
    #                 if age >= min_age:
    #                     return group_def["group"]
    #             elif min_age <= age <= max_age:
    #                 return group_def["group"]
    #     except (ValueError, TypeError):
    #         pass

    return "Работающие взрослые"


def analyze_social_status_by_district(df: pd.DataFrame, districts: list) -> dict:
    """Анализ социальной структуры по районам"""
    logger.info("Начало анализа социальных статусов")
    results = {}

    df['Возраст'] = pd.to_numeric(df['Возраст'], errors='coerce')
    df['Социальный статус'] = df['Социальный статус'].astype(str).str.strip()

    for district in districts:
        df_district = df[df['Административная территория (город)'] == district]
        if df_district.empty:
            logger.warning(f"Нет данных для района: {district}")
            continue

        counts = {}
        for _, row in df_district.iterrows():
            status = row.get("Социальный статус", "")
            age = row.get("Возраст", None)
            group = classify_status(status, age)
            counts[group] = counts.get(group, 0) + 1

        results[district] = counts
        logger.info(f"Обработан район: {district}")

    return results


def print_structure_by_district(structure: dict, title: str):
    """Красивый вывод через Rich с цветами"""
    table = Table(title=title, title_style="bold magenta")
    table.add_column("Район", style="cyan", no_wrap=True)
    table.add_column("Статистика", style="magenta")

    for district, counts in structure.items():
        if not counts:
            table.add_row(district, "[red]Нет данных[/red]")
            continue
        counts_str = ", ".join(
            f"[green]{group}[/green]: [yellow]{count}[/yellow]" for group, count in counts.items()
        )
        table.add_row(district, counts_str)

    console.print(table)


# --- Главная функция ---
if __name__ == "__main__":
    logger.info("Программа запущена")

    df = preprocess_file(config.INPUT_FILE)
    if df is not None:
        structure_age = analyze_population_by_district(df, config.ADM_TERR)
        print_structure_by_district(structure_age, "Возрастная структура по районам")

        structure_socstatus = analyze_social_status_by_district(df, config.ADM_TERR)
        print_structure_by_district(structure_socstatus, "Социальная структура по районам")
