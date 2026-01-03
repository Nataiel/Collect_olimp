from docx import Document
from docx.enum.text import WD_BREAK


def create_merged_application(template_path, names_path, output_path):
    """
    Создает объединенный документ с заявлениями для каждого ученика из списка.

    Args:
        template_path: путь к шаблону документа
        names_path: путь к файлу со списком ФИО
        output_path: путь для сохранения итогового документа
    """

    # Загружаем шаблон документа
    doc = Document(template_path)

    # Читаем список ФИО из файла
    with open(names_path, 'r', encoding='utf-8') as f:
        names = [line.strip() for line in f if line.strip()]

    # Обрабатываем каждого ученика
    for i, full_name in enumerate(names):
        if i == 0:
            # Для первого ученика используем существующий документ
            current_doc = doc
        else:
            # Для последующих учеников добавляем разрыв страницы и копируем шаблон
            # Добавляем разрыв страницы в конец документа
            if current_doc.paragraphs:
                last_paragraph = current_doc.paragraphs[-1]
                run = last_paragraph.add_run()
                run.add_break(WD_BREAK.PAGE)

            # Создаем новый документ из шаблона
            new_doc = Document(template_path)

            # Копируем все элементы из нового документа в текущий
            for element in new_doc.element.body:
                current_doc.element.body.append(element)

        # Заменяем заполнители в тексте абзацев
        replace_placeholders_in_paragraphs(current_doc, full_name)

    # Сохраняем итоговый документ
    current_doc.save(output_path)
    print(f"Создан объединенный документ с {len(names)} заявлениями: {output_path}")


def replace_placeholders_in_paragraphs(doc, full_name):
    """
    Заменяет заполнители в абзацах документа.

    Args:
        doc: объект документа
        full_name: полное ФИО ученика
    """
    # Разделяем ФИО на части
    name_parts = full_name.split()

    # Определяем, какое заявление обрабатываем
    paragraphs = doc.paragraphs
    start_index = 0

    # Находим начало текущего заявления
    for idx, para in enumerate(paragraphs):
        if "ЗАЯВЛЕНИЕ" in para.text and idx > start_index:
            # Это начало нового заявления после первого
            start_index = idx
            break

    # Обрабатываем абзацы текущего заявления
    for para in paragraphs[start_index:]:
        if '%ФИО%' in para.text:

            # Формируем полный текст для замены
            replacement = f"{full_name}"

            # Заменяем текст в абзаце
            for run in para.runs:
                if '%ФИО%' in run.text:
                    run.text = run.text.replace('%ФИО%', replacement)


def main():
    # Пути к файлам
    template_path = "Заявление ВСоШ.docx"
    names_path = "Список ФИО.txt"
    output_path = "Объединенные_заявления_ВСоШ.docx"

    try:
        # Создаем объединенный документ
        create_merged_application(template_path, names_path, output_path)
        print("Программа успешно завершена!")

    except FileNotFoundError as e:
        print(f"Ошибка: файл не найден - {e}")
    except Exception as e:
        print(f"Произошла ошибка: {e}")


if __name__ == "__main__":
    main()