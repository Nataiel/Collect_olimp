import win32com.client as win32


def create_checkbox_with_win32(file_path):
    # Запускаем Excel
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False  # Скрыть Excel

    # Открываем файл
    workbook = excel.Workbooks.Open(file_path)
    worksheet = workbook.ActiveSheet

    # Создаем CheckBox в ячейке A1
    checkbox = worksheet.OLEObjects().Add(
        ClassType="Forms.CheckBox.1",
        Left=worksheet.Range("A1").Left,
        Top=worksheet.Range("A1").Top,
        Width=worksheet.Range("A1").Width,
        Height=worksheet.Range("A1").Height
    )

    # Настраиваем CheckBox
    checkbox.Object.Caption = ""  # Убираем текст
    checkbox.Name = "CheckBox1"

    # Привязываем значение к ячейке B1
    checkbox.LinkedCell = "B1"

    # Сохраняем и закрываем
    workbook.Save()
    workbook.Close()
    excel.Quit()


# Использование
create_checkbox_with_win32("your_file.xlsx")