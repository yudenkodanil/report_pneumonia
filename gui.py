import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import config
from main import preprocess_file, analyze_population_full, analyze_by_med_org, fill_report, fill_report_by_med_org

def run_analysis(input_file):
    df = preprocess_file(input_file)
    if df is not None:
        result_regions = analyze_population_full(df)
        fill_report(result_regions, config.TEMPLATE_FILE_AO, config.OUTPUT_FILE_AO)
        result_med_orgs = analyze_by_med_org(df)
        fill_report_by_med_org(result_med_orgs, config.TEMPLATE_FILE_BLAG, config.OUTPUT_FILE_BLAG)
        messagebox.showinfo("Готово", "Отчёты успешно сохранены!")
    else:
        messagebox.showerror("Ошибка", "Не удалось обработать файл")

def choose_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        # Запускаем анализ в отдельном потоке, чтобы GUI не вис
        threading.Thread(target=run_analysis, args=(file_path,), daemon=True).start()

root = tk.Tk()
root.title("Анализ данных для отчета по пневмонии")

tk.Label(root, text="Выберите Excel-файл для анализа:").pack(pady=10)
tk.Button(root, text="Выбрать файл", command=choose_file).pack(pady=5)

root.geometry("400x200")
root.mainloop()
