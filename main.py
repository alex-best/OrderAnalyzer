import pandas as pd
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import filedialog, colorchooser
from tkinter import ttk

selected_files = []
selected_color = None

def get_time_data(file):
    """
    Читает данные из файла Excel и возвращает список временных меток.
    :param file: Путь к файлу Excel.
    :return: Список временных меток.
    """
    data = pd.read_excel(file)
    return data['Время создания'].tolist()

def load_data(files):
    """
    Загружает данные из нескольких файлов и вычисляет среднее количество событий по часам и дням недели.
    :param files: Список путей к файлам Excel.
    :return: DataFrame со средним количеством событий по часам и дням недели.
    """
    all_data = []
    for file in files:
        time_data = get_time_data(file)
        df = pd.DataFrame(time_data, columns=['timestamp'])
        df['timestamp'] = pd.to_datetime(df['timestamp'].str.strip(), format='%d.%m.%Y %H:%M:%S')
        df['hour'] = df['timestamp'].dt.hour
        df['day_of_week'] = df['timestamp'].dt.day_name()
        filtered_df = df
        hourly_counts = filtered_df.groupby(['day_of_week', 'hour']).size().reset_index(name='count')
        filtered_df['date'] = filtered_df['timestamp'].dt.date
        unique_days_per_weekday = filtered_df.groupby('day_of_week')['date'].nunique().reset_index(name='unique_days')
        hourly_counts = hourly_counts.merge(unique_days_per_weekday, on='day_of_week', how='left')
        hourly_counts['average_count'] = hourly_counts['count'] / hourly_counts['unique_days']
        average_hourly_counts_df = hourly_counts[['day_of_week', 'hour', 'average_count']]
        average_hourly_counts_df.columns = ['День недели', 'Час', 'Среднее количество']
        all_data.append(average_hourly_counts_df)
    return pd.concat(all_data).groupby(['День недели', 'Час']).mean().reset_index()

def plot_data(final_data, highlight, color):
    """
    Строит графики среднего количества заказов по часам и дням недели.
    :param final_data: DataFrame с финальными данными для построения графиков.
    :param highlight: Тип выделения (наибольшее, наименьшее, без выделения).
    :param color: Цвет выделения.
    """
    days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    days_rus = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота', 'Воскресенье']

    plt.figure(figsize=(15, 12))
    plt.suptitle('Среднее количество заказов в час с разбивкой по дням\n', fontsize=16)

    for i, day in enumerate(days):
        plt.subplot(4, 2, i + 1)
        day_data = final_data[final_data['День недели'] == day]
        plt.plot(day_data['Час'], day_data['Среднее количество'], marker='o')
        
        if highlight != 'none':
            if highlight == 'lowest':
                optimal_hours = day_data.nsmallest(6, 'Среднее количество')['Час'].tolist()
            elif highlight == 'highest':
                optimal_hours = day_data.nlargest(6, 'Среднее количество')['Час'].tolist()

            for hour in optimal_hours:
                plt.axvspan(hour, hour + 1, color=color, alpha=0.3)
                plt.text(hour + 0.5, max(day_data['Среднее количество']), f'{int(hour)}-{int(hour + 1)}', 
                         horizontalalignment='center', verticalalignment='bottom', fontsize=8, color='green')

        plt.xlabel('Час дня\n\n')
        # plt.ylabel('Среднее количество событий')
        plt.title(f'{days_rus[i]}')
        plt.grid(True)
        plt.xticks(range(0, 24))
        plt.legend()

    plt.tight_layout(rect=[0, 0, 1, 0.96])
    plt.show()

def select_files():
    """
    Открывает диалоговое окно для выбора файлов и обновляет метку с количеством выбранных файлов.
    """
    global selected_files
    selected_files = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xls *.xlsx")])
    files_label.config(text="Выбрано файлов: " + str(len(selected_files)))

def select_color():
    """
    Открывает диалоговое окно для выбора цвета и обновляет метку с выбранным цветом.
    """
    global selected_color
    color_code = colorchooser.askcolor(title="Выберите цвет")[1]
    selected_color = color_code
    color_label.config(text="Выбранный цвет: " + str(color_code), bg=color_code)

def run():
    """
    Загружает данные, строит графики и отображает их.
    """
    highlight = highlight_choice.get()
    final_data = load_data(selected_files)
    plot_data(final_data, highlight, selected_color if highlight != 'none' else None)

# Создание окна интерфейса
root = tk.Tk()
root.title("Анализ заказов")

# Элементы интерфейса
select_files_btn = ttk.Button(root, text="Выбрать файлы", command=select_files)
select_files_btn.grid(row=0, column=0, padx=10, pady=10)

files_label = ttk.Label(root, text="Выбрано файлов: 0")
files_label.grid(row=0, column=1, padx=10, pady=10)

highlight_choice = ttk.Combobox(root, values=["none", "lowest", "highest"], state="readonly")
highlight_choice.set("none")
highlight_choice.grid(row=1, column=0, padx=10, pady=10)

highlight_label = ttk.Label(root, text="Выделение значений")
highlight_label.grid(row=1, column=1, padx=10, pady=10)

select_color_btn = ttk.Button(root, text="Выбрать цвет выделения", command=select_color)
select_color_btn.grid(row=2, column=0, padx=10, pady=10)

color_label = ttk.Label(root, text="Выбранный цвет: None")
color_label.grid(row=2, column=1, padx=10, pady=10)

run_btn = ttk.Button(root, text="Построить график", command=run)
run_btn.grid(row=3, column=0, columnspan=2, padx=10, pady=10)

root.mainloop()
