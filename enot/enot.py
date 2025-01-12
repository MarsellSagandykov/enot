import pandas as pd
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import filedialog, messagebox


def load_data():
    # Получение пути к файлу
    file_path = filedialog.askopenfilename(title="Выберите файл Excel", filetypes=[("Excel files", "*.xlsx *.xls")])
    if not file_path:
        return

    # Получение данных с формы
    sheet_name = sheet_name_entry.get()
    group_name = group_name_entry.get()
    colum_name1 = colum_name1_entry.get()
    colum_name2 = colum_name2_entry.get()
    use_second_column = colum_name2_var.get()

    if not (sheet_name and group_name and colum_name1):
        messagebox.showwarning("Ошибка", "Пожалуйста, заполните все обязательные поля.")
        return

    # Попытка загрузить данные с выбранного листа
    try:
        data = pd.read_excel(file_path, sheet_name=sheet_name)
    except ValueError:
        messagebox.showerror("Ошибка", f"Лист '{sheet_name}' не найден.")
        return

    # Очистка названий колонок от пробелов
    data.columns = data.columns.str.strip()

    # Подсчет уникальных категорий
    if colum_name1 in data.columns and (not use_second_column or colum_name2 in data.columns):
        if use_second_column:
            # Объединяем данные из обеих колонок
            combined_counts = pd.concat([data[colum_name1], data[colum_name2]]).value_counts()
        else:
            # Используем только первую колонку
            combined_counts = data[colum_name1].value_counts()

        # Определение категорий и их цветов
        color_map = {
            'нет фокуса': '#636efa',
            'компетенция коммуникации': '#EF553B',
            'изменение действий': '#00cc96',
            'понимание себя': '#ab63fa',
            'положительная эмоция': '#FFA15A',
            'общение с одногруппниками': '#19d3f3',
            'рефлексивный дневник': '#ee64ca',
            'профессионализацию': '#30d5c8',
            'место ознакомительной практики': '#FFFF00'
        }

        # Приведение индексов к нижнему регистру
        combined_counts.index = combined_counts.index.str.lower()

        # Получение цветов для категорий, встречающихся в данных
        colors = [color_map.get(category, '#B6E880') for category in combined_counts.index]

        # Создание списка "explode", чтобы отсоединить категории с количеством от 1 до 20
        explode = [0.1 if 1 <= count <= 21 else 0 for count in combined_counts]

        # Функция для отображения процентов только для категорий с количеством более 20
        def autopct_func(pct, allvalues):
            absolute = int(pct / 100. * sum(allvalues))
            return f'{pct:.1f}%' if absolute > 20 else ''

        # Установка размера графика
        plt.figure(figsize=(11, 8))

        # Визуализация круговой диаграммы с параметром explode
        combined_counts.plot(
            kind='pie', autopct=lambda pct: autopct_func(pct, combined_counts), colors=colors, explode=explode,
            ylabel=''
        )

        # Добавление легенды с подсчетом
        legend_labels = [f"{category}: {count}" for category, count in zip(combined_counts.index, combined_counts)]
        plt.legend(legend_labels, title="Фокус внимания на", bbox_to_anchor=(0.80, 1))

        # Добавление заголовка с указанием группы
        plt.title(f"Результаты обратной связи: {group_name}", fontsize=14)

        # Сохранение графика с указанием имени листа и группы
        output_file = f'{sheet_name}_{group_name}.png'
        plt.savefig(output_file)
        plt.show()  # Показать график
        messagebox.showinfo("Успех", f"График сохранен как '{output_file}'")
    else:
        messagebox.showerror("Ошибка", "Не найдены указанные колонки в данных.")


# Создание основного окна
root = tk.Tk()
root.title("Программа для создания круговой диаграммы")

# Интерфейс для ввода данных
tk.Label(root, text="Название листа:").grid(row=0, column=0)
sheet_name_entry = tk.Entry(root)
sheet_name_entry.grid(row=0, column=1)

tk.Label(root, text="Название обратной связи:").grid(row=1, column=0)
group_name_entry = tk.Entry(root)
group_name_entry.grid(row=1, column=1)

tk.Label(root, text="Название 1 колонки:").grid(row=2, column=0)
colum_name1_entry = tk.Entry(root)
colum_name1_entry.grid(row=2, column=1)

tk.Label(root, text="Название 2 колонки (опционально):").grid(row=3, column=0)
colum_name2_entry = tk.Entry(root)
colum_name2_entry.grid(row=3, column=1)

# Добавление флажка для использования второй колонки
colum_name2_var = tk.BooleanVar()
use_colum_name2_checkbox = tk.Checkbutton(root, text="Использовать вторую колонку", variable=colum_name2_var)
use_colum_name2_checkbox.grid(row=4, column=0, columnspan=2)

# Кнопка для запуска анализа
load_button = tk.Button(root, text="Загрузить и построить график", command=load_data)
load_button.grid(row=5, column=0, columnspan=2, pady=10)

# Запуск приложения
root.mainloop()