import pandas as pd
import json
from collections import defaultdict

# Загрузка JSON файла
with open('olympiad_data.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

# Группируем данные по предметам (ТОЛЬКО УЧАСТНИКИ)
subject_data = defaultdict(list)

# Собираем данные по каждому предмету (только участники)
for class_name, students in data.items():
    for student_name, subjects in students.items():
        for subject, status in subjects.items():
            if status:  # Только если ученик участвует (status == True)
                subject_data[subject].append({
                    'Класс': class_name,
                    'Ученик': student_name,
                    'Участвует': 'Да'  # Всегда "Да", так как мы фильтруем
                })

# Создаем Excel файл
with pd.ExcelWriter('Отчёт.xlsx', engine='openpyxl') as writer:
    # Создаем лист для каждого предмета (только участники)
    for subject, records in sorted(subject_data.items()):
        if records:  # Если есть хотя бы один участник
            df = pd.DataFrame(records)
            df = df.sort_values(['Класс', 'Ученик'])

            # Создаем имя листа
            sheet_name = str(subject)[:31]
            # Заменяем запрещенные символы
            invalid_chars = ['\\', '/', '*', '?', ':', '[', ']']
            for char in invalid_chars:
                sheet_name = sheet_name.replace(char, '_')

            # Добавляем лист с участниками
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    # Общий лист со всеми данными (все участники)
    all_records = []
    for subject, records in subject_data.items():
        for record in records:
            all_records.append({
                'Предмет': subject,
                'Класс': record['Класс'],
                'Ученик': record['Ученик']
            })

    if all_records:
        df_all = pd.DataFrame(all_records)
        df_all.to_excel(writer, sheet_name='Все участники', index=False)

    # Лист со статистикой по предметам
    stats_records = []
    for subject, records in sorted(subject_data.items()):
        # Группируем участников по классам
        class_counts = defaultdict(int)
        for record in records:
            class_counts[record['Класс']] += 1

        # Формируем строку с распределением по классам
        class_distribution = ', '.join([f"{cls}: {count}" for cls, count in sorted(class_counts.items())])

        stats_records.append({
            'Предмет': subject,
            'Всего участников': len(records),
            'Распределение по классам': class_distribution
        })

    if stats_records:
        df_stats = pd.DataFrame(stats_records)
        df_stats.to_excel(writer, sheet_name='Статистика', index=False)

    # Лист со списком всех учеников и их олимпиад
    student_olympiads = defaultdict(list)

    for class_name, students in data.items():
        for student_name, subjects in students.items():
            student_participations = [subject for subject, status in subjects.items() if status]
            if student_participations:  # Если ученик участвует хотя бы в одной олимпиаде
                student_olympiads[(class_name, student_name)] = sorted(student_participations)

    # Создаем лист с учениками
    student_records = []
    for (class_name, student_name), olympiads in sorted(student_olympiads.items()):
        student_records.append({
            'Класс': class_name,
            'Ученик': student_name,
            'Количество олимпиад': len(olympiads),
            'Список олимпиад': ', '.join(olympiads)
        })

    if student_records:
        df_students = pd.DataFrame(student_records)
        df_students.to_excel(writer, sheet_name='Ученики', index=False)

print(f"Создан файл olympiad_data_by_subjects_participants_only.xlsx")
print(f"Количество листов: {len(subject_data) + 3}")
print("\nСодержание файла:")
print("1. Отдельные листы по предметам - только участники олимпиад")
print("2. 'Все участники' - полный список участников всех олимпиад")
print("3. 'Статистика' - количество участников по предметам")
print("4. 'Ученики' - список учеников и олимпиад, в которых они участвуют")