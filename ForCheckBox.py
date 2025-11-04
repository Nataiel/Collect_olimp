import xlwings as xw


def create_checkbox(worksheet, cell_address, linked_cell, caption="", name=None):
    """
    Создает флажок в указанной ячейке и привязывает его к другой ячейке

    Parameters:
    - worksheet: лист xlwings
    - cell_address: адрес ячейки для размещения флажка (например, 'A1')
    - linked_cell: адрес ячейки для привязки значения (например, 'B1')
    - caption: текст подписи флажка
    - name: имя флажка (опционально)
    """
    # Получаем координаты ячейки
    target_cell = worksheet.range(cell_address)

    # Создаем флажок
    checkbox = worksheet.api.Shapes.AddFormControl(
        9,  # Тип элемента - Checkbox
        Left=target_cell.left,
        Top=target_cell.top,
        Width=target_cell.width,
        Height=target_cell.height
    )

    # Настраиваем флажок
    checkbox.ControlFormat.Caption = caption
    if name:
        checkbox.Name = name
    else:
        checkbox.Name = f"Checkbox_{cell_address}"

    # Привязываем к ячейке
    checkbox.ControlFormat.LinkedCell = worksheet.range(linked_cell).get_address(False, False)

    return checkbox


# Использование функции
wb = xw.Book()
ws = wb.sheets[0]

# Создаем флажок в A1, привязанный к B1
create_checkbox(ws, 'A1', 'B1', "", "MyCheckbox")

# Устанавливаем начальное значение
ws.range('B1').value = False

# Проверяем работу
print(f"Значение в B1: {ws.range('B1').value}")

# Можем изменить значение программно
ws.range('B1').value = True
print(f"Новое значение в B1: {ws.range('B1').value}")

# Сохраняем результат
# wb.save('workbook_with_checkbox.xlsx')
