import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from openpyxl import load_workbook
import os

path_file = None  # переменная для хранения пути к файлу
olimp_subjects = ['Выберите предмет']  # переменная для списка предметов из файла
sheet_names = []  # Список листов в Исходном файле

def open_file():
    """Функция для открытия файла и сохранения пути"""
    global path_file, olimp_subjects, sheet_names

    # Запрашиваем файл у пользователя
    file_path = filedialog.askopenfilename(title="Выберите Excel файл",
                                           filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])

    # Если пользователь выбрал файл (не нажал "Отмена")
    if file_path:
        path_file = file_path
        # Обновляем текст в метке
        file_label.config(text=f"Выбран файл: {os.path.basename(path_file)}")
        status_label.config(text="Файл успешно загружен", fg="green")
        # Активируем кнопку модификации
        modify_button.config(state=tk.NORMAL)
        print(f"Файл выбран: {path_file}")  # Для отладки

        wb = load_workbook(path_file)
        sheet_names = wb.sheetnames
        sheet_names = [x for x in sheet_names if any(y.isdigit() for y in x) and 'Лист' not in x]
        print("Список листов:", sheet_names)

        for x in sheet_names:
            collect_subjects(wb[x])
        olimp_subjects = [x for x in olimp_subjects if not('умма' in x or x == 'Выберите предмет')]
        subject_var.config(values=olimp_subjects)
        print(olimp_subjects)


def modify_file():
    """Функция для модификации файла"""
    global path_file, olimp_subjects, sheet_names

    if not path_file:
        messagebox.showerror("Ошибка", "Сначала выберите файл!")
        return

    try:
        # Загружаем workbook
        wb = load_workbook(path_file)


        if subject_var.get() in wb.sheetnames:
            del wb[subject_var.get()]

        wb.create_sheet(subject_var.get())
        sheet = wb[subject_var.get()]

        sheet.cell(row=1, column=1).value = '№'
        sheet.cell(row=1, column=2).value = 'Фамилия'
        sheet.cell(row=1, column=3).value = 'Имя'
        sheet.cell(row=1, column=4).value = 'Отчество'
        sheet.cell(row=1, column=5).value = 'Класс'
        row_pos = 2

        for y in sheet_names:
            col = 6
            row = 1
            for x in range(col, 100):
                if wb[y].cell(row=1, column=x).value == subject_var.get():
                    break

            row += 1
            print(wb[y].cell(row=1, column=x).value)

            for z in range(row, 100):
                if wb[y].cell(row=z, column=x).value:
                    wb[subject_var.get()].cell(row=row_pos, column=1).value = row_pos - 1
                    wb[subject_var.get()].cell(row=row_pos, column=2).value = wb[y].cell(row=z, column=2).value
                    wb[subject_var.get()].cell(row=row_pos, column=3).value = wb[y].cell(row=z, column=3).value
                    wb[subject_var.get()].cell(row=row_pos, column=4).value = wb[y].cell(row=z, column=4).value
                    wb[subject_var.get()].cell(row=row_pos, column=5).value = y
                    row_pos += 1


        # Сохраняем изменения
        wb.save(path_file)

        messagebox.showinfo("Успех", "Файл успешно модифицирован!")
        status_label.config(text="Файл модифицирован", fg="blue")

    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось модифицировать файл:\n{str(e)}")
        status_label.config(text="Ошибка модификации", fg="red")


# Функция сбора предметов с листа
def collect_subjects(sheet):
    global olimp_subjects

    col = 6
    row = 1
    for x in range(col, 100):
        aux = sheet.cell(row=row, column=x).value
        if aux:
            if aux not in olimp_subjects:
                olimp_subjects.append(aux)


def close_app():
    """Функция для закрытия приложения"""
    if messagebox.askokcancel("Выход", "Вы уверены, что хотите выйти?"):
        root.destroy()


# Создаем главное окно
root = tk.Tk()
root.title("Excel Modifier")
root.geometry("500x300")
root.resizable(False, False)

# Стилизация
button_style = {'font': ('Arial', 12), 'width': 15, 'height': 2, 'bg': '#f0f0f0'}

# Создаем фрейм для кнопок
button_frame = tk.Frame(root)
button_frame.pack(pady=20)

# Кнопка открытия файла
open_button = tk.Button(button_frame, text="Открыть файл", command=open_file, **button_style)
open_button.pack(pady=10)

# Фрейм для выпадающего списка и кнопки модификации
modify_frame = tk.Frame(button_frame)
modify_frame.pack(pady=10)

# Выпадающий список слева
subject_var = ttk.Combobox(modify_frame, values=olimp_subjects, state="readonly", height=15)
subject_var.set(olimp_subjects[0])  # значение по умолчанию


subject_var.config(font=('Arial', 10), width=15, height=15)
subject_var.pack(side=tk.LEFT, padx=(0, 10))

# Кнопка модификации (изначально неактивна)
modify_button = tk.Button(modify_frame, text="Модифицировать", command=modify_file, state=tk.DISABLED, **button_style)
modify_button.pack(side=tk.LEFT)

# Кнопка закрытия
close_button = tk.Button(button_frame, text="Закрыть", command=close_app, **button_style)
close_button.pack(pady=10)

# Метка для отображения выбранного файла
file_label = tk.Label(root, text="Файл не выбран", font=('Arial', 10), fg='gray')
file_label.pack(pady=10)

# Метка статуса
status_label = tk.Label(root, text="Готов к работе", font=('Arial', 10), fg='green')
status_label.pack(pady=5)

# Запускаем главный цикл
root.mainloop()