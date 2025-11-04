import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import pandas as pd
import json
import os
from flask import Flask, request, jsonify, render_template_string
import webbrowser


class OlympiadServer:
    def __init__(self):
        self.app = Flask(__name__)
        self.server_thread = None
        self.running = False
        self.data_file = 'olympiad_data.json'
        self.excel_file = None
        self.school_data = {}
        self.subjects = ['математика', 'русский язык']
        self.olympiad_data = {}
        self.first_run = True  # Флаг первого запуска

        self.setup_flask_routes()

    def load_students_from_excel(self, file_path):
        """Загрузка данных учеников из Excel файла"""
        try:
            self.school_data = {}
            excel_data = pd.read_excel(file_path, sheet_name=None)

            for sheet_name, df in excel_data.items():
                # Предполагаем, что ФИО в первом столбце
                if df.empty:
                    continue

                students = []
                for index, row in df.iterrows():
                    # Берем значение из первой колонки как ФИО
                    fio = str(row.iloc[0]).strip()
                    if fio and fio != 'nan':  # Игнорируем пустые строки
                        students.append(fio)

                if students:
                    self.school_data[sheet_name] = students
                    print(f"Загружен класс {sheet_name}: {len(students)} учеников")

            return True
        except Exception as e:
            print(f"Ошибка загрузки Excel: {e}")
            return False

    def reset_olympiad_data(self):
        """Сброс данных об участии (все чекбоксы не отмечены)"""
        self.olympiad_data = {}
        for class_name in self.school_data:
            self.olympiad_data[class_name] = {}
            for student in self.school_data[class_name]:
                self.olympiad_data[class_name][student] = {subject: False for subject in self.subjects}

        self.save_olympiad_data()
        print("✅ Данные олимпиад сброшены (все чекбоксы не отмечены)")

    def load_olympiad_data(self):
        """Загрузка данных об участии"""
        if os.path.exists(self.data_file):
            print("📁 Загружаем существующие данные из файла...")
            with open(self.data_file, 'r', encoding='utf-8') as f:
                self.olympiad_data = json.load(f)

            # Проверяем структуру данных и дополняем при необходимости
            data_updated = False
            for class_name in self.school_data:
                if class_name not in self.olympiad_data:
                    self.olympiad_data[class_name] = {}
                    data_updated = True

                for student in self.school_data[class_name]:
                    if student not in self.olympiad_data[class_name]:
                        self.olympiad_data[class_name][student] = {subject: False for subject in self.subjects}
                        data_updated = True
                    else:
                        # Проверяем наличие всех предметов
                        for subject in self.subjects:
                            if subject not in self.olympiad_data[class_name][student]:
                                self.olympiad_data[class_name][student][subject] = False
                                data_updated = True

            if data_updated:
                self.save_olympiad_data()
                print("✅ Структура данных обновлена")
        else:
            print("📝 Файл данных не найден, создаем новый...")
            self.reset_olympiad_data()

    def save_olympiad_data(self):
        """Сохранение данных об участии"""
        with open(self.data_file, 'w', encoding='utf-8') as f:
            json.dump(self.olympiad_data, f, ensure_ascii=False, indent=2)
        print("💾 Данные сохранены в файл")

    def setup_flask_routes(self):
        """Настройка маршрутов Flask"""

        MAIN_PAGE = """
        <!DOCTYPE html>
        <html>
        <head>
            <title>Участие в олимпиадах</title>
            <meta charset="utf-8">
            <style>
                body { 
                    font-family: Arial, sans-serif; 
                    max-width: 600px; 
                    margin: 50px auto; 
                    padding: 20px;
                    background: #f5f5f5;
                }
                .header {
                    text-align: center;
                    background: white;
                    padding: 30px;
                    border-radius: 10px;
                    box-shadow: 0 2px 10px rgba(0,0,0,0.1);
                    margin-bottom: 30px;
                }
                .class-grid {
                    display: grid;
                    grid-template-columns: repeat(2, 1fr);
                    gap: 20px;
                }
                .class-card {
                    background: white;
                    padding: 30px;
                    text-align: center;
                    border-radius: 10px;
                    text-decoration: none;
                    color: #333;
                    font-size: 24px;
                    font-weight: bold;
                    box-shadow: 0 2px 10px rgba(0,0,0,0.1);
                    transition: transform 0.2s, box-shadow 0.2s;
                }
                .class-card:hover {
                    transform: translateY(-5px);
                    box-shadow: 0 5px 20px rgba(0,0,0,0.15);
                }
                .stats-link {
                    display: block;
                    margin-top: 20px;
                    text-align: center;
                    color: #3498db;
                    text-decoration: none;
                    font-weight: bold;
                }
                .stats-link:hover {
                    text-decoration: underline;
                }
            </style>
        </head>
        <body>
            <div class="header">
                <h1>🏆 Участие в олимпиадах</h1>
                <div>Выберите класс для редактирования</div>

                <div class="class-grid">
                    {% for class in classes %}
                    <a href="/class/{{ class }}" class="class-card">{{ class }} класс</a>
                    {% endfor %}
                </div>

                <a href="/stats" class="stats-link">📊 Посмотреть статистику</a>
            </div>
        </body>
        </html>
        """

        CLASS_PAGE_TEMPLATE = """
        <!DOCTYPE html>
        <html>
        <head>
            <title>Класс {{ class_name }} - Олимпиады</title>
            <meta charset="utf-8">
            <style>
                body { 
                    font-family: Arial, sans-serif; 
                    max-width: 1000px; 
                    margin: 20px auto; 
                    padding: 20px;
                    background: #f5f5f5;
                }
                .header {
                    background: white;
                    padding: 20px;
                    border-radius: 10px;
                    box-shadow: 0 2px 10px rgba(0,0,0,0.1);
                    margin-bottom: 20px;
                }
                .back-link {
                    display: inline-block;
                    margin-bottom: 15px;
                    text-decoration: none;
                    color: #3498db;
                    font-weight: bold;
                }
                .back-link:hover {
                    text-decoration: underline;
                }
                table {
                    width: 100%;
                    border-collapse: collapse;
                    background: white;
                    border-radius: 10px;
                    overflow: hidden;
                    box-shadow: 0 2px 10px rgba(0,0,0,0.1);
                }
                th, td {
                    padding: 15px;
                    text-align: left;
                    border-bottom: 1px solid #ecf0f1;
                }
                th {
                    background: #34495e;
                    color: white;
                    font-weight: bold;
                }
                tr:hover {
                    background: #f8f9fa;
                }
                .student-name {
                    font-weight: bold;
                    width: 200px;
                }
                .subject-header {
                    text-align: center;
                    background: #2c3e50 !important;
                }
                .checkbox-cell {
                    text-align: center;
                    width: 120px;
                }
                input[type="checkbox"] {
                    transform: scale(1.5);
                    cursor: pointer;
                }
                .save-btn {
                    background: #27ae60;
                    color: white;
                    padding: 12px 30px;
                    border: none;
                    border-radius: 5px;
                    cursor: pointer;
                    font-size: 16px;
                    margin-top: 20px;
                }
                .save-btn:hover {
                    background: #219a52;
                }
                .reset-btn {
                    background: #e74c3c;
                    color: white;
                    padding: 12px 30px;
                    border: none;
                    border-radius: 5px;
                    cursor: pointer;
                    font-size: 16px;
                    margin-top: 20px;
                    margin-left: 10px;
                }
                .reset-btn:hover {
                    background: #c0392b;
                }
                .button-group {
                    text-align: center;
                }
                .message {
                    padding: 15px;
                    margin: 15px 0;
                    border-radius: 5px;
                    text-align: center;
                    font-weight: bold;
                }
                .success {
                    background: #d4edda;
                    color: #155724;
                    border: 1px solid #c3e6cb;
                }
                .info {
                    background: #d1ecf1;
                    color: #0c5460;
                    border: 1px solid #bee5eb;
                }
            </style>
        </head>
        <body>
            <div class="header">
                <a href="/" class="back-link">← Назад к выбору класса</a>
                <h1>🎓 {{ class_name }} класс - Участие в олимпиадах</h1>
                <div class="info message">
                    💡 Выберите предметы, в которых участвует каждый ученик. Изменения сохраняются автоматически.
                </div>
            </div>

            {% if message %}
            <div class="message success">{{ message }}</div>
            {% endif %}

            <form action="/save/{{ class_name }}" method="post">
                <table>
                    <thead>
                        <tr>
                            <th>Ученик</th>
                            {% for subject in subjects %}
                            <th class="subject-header">{{ subject.title() }}</th>
                            {% endfor %}
                        </tr>
                    </thead>
                    <tbody>
                        {% for student in students %}
                        <tr>
                            <td class="student-name">{{ student }}</td>
                            {% for subject in subjects %}
                            <td class="checkbox-cell">
                                <input type="checkbox" 
                                       name="{{ student }}_{{ subject }}" 
                                       {{ 'checked' if participation[student][subject] else '' }}>
                            </td>
                            {% endfor %}
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>

                <div class="button-group">
                    <button type="submit" class="save-btn">💾 Сохранить изменения</button>
                    <button type="button" class="reset-btn" onclick="resetClass()">🔄 Сбросить класс</button>
                </div>
            </form>

            <script>
                function resetClass() {
                    if (confirm('Вы уверены, что хотите сбросить ВСЕ выборы для этого класса? Это действие нельзя отменить.')) {
                        window.location.href = '/reset/{{ class_name }}';
                    }
                }
            </script>
        </body>
        </html>
        """

        STATS_PAGE = """
        <!DOCTYPE html>
        <html>
        <head>
            <title>Статистика олимпиад</title>
            <meta charset="utf-8">
            <style>
                body { font-family: Arial, sans-serif; max-width: 800px; margin: 50px auto; padding: 20px; }
                .stat-card { background: white; padding: 20px; margin: 15px 0; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
                .back-link { display: inline-block; margin-bottom: 20px; text-decoration: none; color: #3498db; font-weight: bold; }
                table { width: 100%; border-collapse: collapse; margin-top: 10px; }
                th, td { padding: 12px; text-align: center; border: 1px solid #ddd; }
                th { background: #34495e; color: white; }
                .total-row { background: #f8f9fa; font-weight: bold; }
            </style>
        </head>
        <body>
            <a href="/" class="back-link">← Назад к выбору класса</a>
            <h1>📊 Статистика участия в олимпиадах</h1>

            {% for class_name, subjects_data in stats.items() %}
            <div class="stat-card">
                <h2>Класс {{ class_name }}</h2>
                <table>
                    <tr>
                        <th>Предмет</th>
                        <th>Количество участников</th>
                        <th>Процент участия</th>
                    </tr>
                    {% for subject, count in subjects_data.items() %}
                    <tr>
                        <td>{{ subject.title() }}</td>
                        <td><strong>{{ count }}/{{ total_students[class_name] }}</strong></td>
                        <td>{{ (count / total_students[class_name] * 100) | round(1) }}%</td>
                    </tr>
                    {% endfor %}
                </table>
            </div>
            {% endfor %}
        </body>
        </html>
        """

        @self.app.route('/')
        def home():
            if not self.school_data:
                return "Данные не загружены. Загрузите Excel файл через GUI."
            return render_template_string(MAIN_PAGE, classes=list(self.school_data.keys()))

        @self.app.route('/class/<class_name>')
        def class_page(class_name):
            if class_name not in self.school_data:
                return "Класс не найден", 404

            students = self.school_data[class_name]
            participation = self.olympiad_data.get(class_name, {})

            return render_template_string(
                CLASS_PAGE_TEMPLATE,
                class_name=class_name,
                students=students,
                subjects=self.subjects,
                participation=participation,
                message=request.args.get('message')
            )

        @self.app.route('/save/<class_name>', methods=['POST'])
        def save_data(class_name):
            try:
                if class_name not in self.school_data:
                    return "Класс не найден", 404

                for student in self.school_data[class_name]:
                    for subject in self.subjects:
                        checkbox_name = f"{student}_{subject}"
                        self.olympiad_data[class_name][student][subject] = checkbox_name in request.form

                self.save_olympiad_data()
                return render_template_string(
                    CLASS_PAGE_TEMPLATE,
                    class_name=class_name,
                    students=self.school_data[class_name],
                    subjects=self.subjects,
                    participation=self.olympiad_data[class_name],
                    message='✅ Данные успешно сохранены!'
                )

            except Exception as e:
                return f"Ошибка: {e}", 500

        @self.app.route('/reset/<class_name>')
        def reset_class(class_name):
            """Сброс данных для конкретного класса"""
            if class_name in self.olympiad_data:
                for student in self.olympiad_data[class_name]:
                    for subject in self.subjects:
                        self.olympiad_data[class_name][student][subject] = False
                self.save_olympiad_data()

            return render_template_string(
                CLASS_PAGE_TEMPLATE,
                class_name=class_name,
                students=self.school_data[class_name],
                subjects=self.subjects,
                participation=self.olympiad_data[class_name],
                message='✅ Данные класса сброшены!'
            )

        @self.app.route('/stats')
        def stats_page():
            """Страница статистики"""
            stats_data = {}
            total_students = {}

            for class_name in self.school_data:
                stats_data[class_name] = {}
                total_students[class_name] = len(self.school_data[class_name])
                for subject in self.subjects:
                    count = sum(1 for student_data in self.olympiad_data.get(class_name, {}).values()
                                if student_data.get(subject, False))
                    stats_data[class_name][subject] = count

            return render_template_string(STATS_PAGE, stats=stats_data, total_students=total_students)

        @self.app.route('/api/stats')
        def api_stats():
            stats = {}
            for class_name in self.school_data:
                stats[class_name] = {
                    'total_students': len(self.school_data[class_name]),
                    'participation_by_subject': {}
                }
                for subject in self.subjects:
                    count = sum(1 for student_data in self.olympiad_data.get(class_name, {}).values()
                                if student_data.get(subject, False))
                    stats[class_name]['participation_by_subject'][subject] = count
            return jsonify(stats)

    def get_statistics(self):
        """Получение статистики для GUI"""
        stats = {}
        for class_name in self.school_data:
            stats[class_name] = {}
            for subject in self.subjects:
                count = sum(1 for student_data in self.olympiad_data.get(class_name, {}).values()
                            if student_data.get(subject, False))
                total = len(self.school_data[class_name])
                stats[class_name][subject] = f"{count}/{total}"
        return stats

    def run_server(self):
        """Запуск сервера"""
        self.running = True

        # При первом запуске спрашиваем о сбросе данных
        if self.first_run and os.path.exists(self.data_file):
            print("📋 Файл данных существует - загружаем предыдущие выборы")
        elif self.first_run:
            print("🆕 Первый запуск - создаем чистые данные")

        self.first_run = False
        self.app.run(host='0.0.0.0', port=5000, debug=False, use_reloader=False)

    def stop_server(self):
        """Остановка сервера"""
        self.running = False


class ServerGUI:
    def __init__(self):
        self.server = OlympiadServer()
        self.root = tk.Tk()
        self.root.title("Сервер олимпиад - Управление")
        self.root.geometry("800x600")

        self.setup_gui()

    def setup_gui(self):
        """Создание интерфейса"""
        # Основной фрейм
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Заголовок
        title_label = ttk.Label(main_frame, text="🏆 Управление сервером олимпиад",
                                font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))

        # Секция загрузки файла
        file_frame = ttk.LabelFrame(main_frame, text="Загрузка данных", padding="10")
        file_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))

        self.file_label = ttk.Label(file_frame, text="Файл не выбран")
        self.file_label.grid(row=0, column=0, sticky=(tk.W, tk.E))

        ttk.Button(file_frame, text="Выбрать Excel файл",
                   command=self.load_excel_file).grid(row=0, column=1, padx=(10, 0))

        ttk.Button(file_frame, text="Сбросить все данные",
                   command=self.reset_all_data).grid(row=0, column=2, padx=(10, 0))

        # Секция управления сервером
        server_frame = ttk.LabelFrame(main_frame, text="Управление сервером", padding="10")
        server_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))

        self.status_label = ttk.Label(server_frame, text="Статус: Сервер остановлен",
                                      foreground="red", font=("Arial", 10, "bold"))
        self.status_label.grid(row=0, column=0, columnspan=2, pady=(0, 10))

        ttk.Button(server_frame, text="Запустить сервер",
                   command=self.start_server).grid(row=1, column=0, padx=(0, 5))
        ttk.Button(server_frame, text="Остановить сервер",
                   command=self.stop_server).grid(row=1, column=1, padx=(5, 0))
        ttk.Button(server_frame, text="Открыть в браузере",
                   command=self.open_browser).grid(row=1, column=2, padx=(10, 0))

        # Секция статистики
        stats_frame = ttk.LabelFrame(main_frame, text="Статистика участия", padding="10")
        stats_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))

        # Таблица статистики
        columns = ("Класс", "Математика", "Русский язык")
        self.stats_tree = ttk.Treeview(stats_frame, columns=columns, show="headings", height=10)

        for col in columns:
            self.stats_tree.heading(col, text=col)
            self.stats_tree.column(col, width=150, anchor="center")

        self.stats_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Scrollbar для таблицы
        scrollbar = ttk.Scrollbar(stats_frame, orient="vertical", command=self.stats_tree.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.stats_tree.configure(yscrollcommand=scrollbar.set)

        # Кнопка обновления статистики
        ttk.Button(stats_frame, text="Обновить статистику",
                   command=self.update_statistics).grid(row=1, column=0, pady=(10, 0))

        # Настройка весов для растягивания
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(3, weight=1)
        file_frame.columnconfigure(0, weight=1)
        server_frame.columnconfigure(0, weight=1)
        stats_frame.columnconfigure(0, weight=1)
        stats_frame.rowconfigure(0, weight=1)

    def load_excel_file(self):
        """Загрузка Excel файла с данными учеников"""
        file_path = filedialog.askopenfilename(
            title="Выберите Excel файл с данными учеников",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )

        if file_path:
            try:
                if self.server.load_students_from_excel(file_path):
                    self.server.excel_file = file_path

                    # Проверяем, существует ли файл данных
                    if os.path.exists(self.server.data_file):
                        response = messagebox.askyesno(
                            "Загрузка данных",
                            "Обнаружен файл с предыдущими выборами.\n\nЗагрузить существующие данные? (Да)\nИли сбросить все выборы? (Нет)"
                        )
                        if response:
                            print("📁 Загружаем существующие данные...")
                            self.server.load_olympiad_data()
                        else:
                            print("🔄 Сбрасываем все данные...")
                            self.server.reset_olympiad_data()
                    else:
                        print("🆕 Создаем новые данные...")
                        self.server.load_olympiad_data()

                    self.file_label.config(text=f"Загружен: {os.path.basename(file_path)}")
                    messagebox.showinfo("Успех",
                                        f"Данные успешно загружены!\nКлассы: {', '.join(self.server.school_data.keys())}")
                    self.update_statistics()
                else:
                    messagebox.showerror("Ошибка", "Не удалось загрузить данные из файла")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка при загрузке файла: {e}")

    def reset_all_data(self):
        """Сброс всех данных олимпиад"""
        if messagebox.askyesno("Сброс данных",
                               "Вы уверены, что хотите сбросить ВСЕ данные олимпиад?\nВсе выборы будут удалены!"):
            self.server.reset_olympiad_data()
            self.update_statistics()
            messagebox.showinfo("Сброс данных", "Все данные по выбору олимпиад сброшены!")

    def start_server(self):
        """Запуск сервера"""
        if not self.server.school_data:
            messagebox.showwarning("Предупреждение", "Сначала загрузите Excel файл с данными")
            return

        if not self.server.running:
            self.server_thread = threading.Thread(target=self.server.run_server)
            self.server_thread.daemon = True
            self.server_thread.start()
            self.status_label.config(text="Статус: Сервер запущен на http://localhost:5000",
                                     foreground="green")
            messagebox.showinfo("Сервер запущен", "Сервер успешно запущен!\nВы можете открыть его в браузере.")

    def stop_server(self):
        """Остановка сервера"""
        if self.server.running:
            self.server.stop_server()
            self.status_label.config(text="Статус: Сервер остановлен", foreground="red")
            messagebox.showinfo("Сервер остановлен", "Сервер был остановлен.")
        else:
            messagebox.showinfo("Информация", "Сервер уже остановлен")

    def open_browser(self):
        """Открытие браузера"""
        if self.server.running:
            webbrowser.open("http://localhost:5000")
        else:
            messagebox.showwarning("Предупреждение", "Сервер не запущен")

    def update_statistics(self):
        """Обновление статистики в таблице"""
        # Очищаем таблицу
        for item in self.stats_tree.get_children():
            self.stats_tree.delete(item)

        # Загружаем статистику
        stats = self.server.get_statistics()

        # Заполняем таблицу
        for class_name, subjects_data in stats.items():
            math_data = subjects_data.get('математика', '0/0')
            rus_data = subjects_data.get('русский язык', '0/0')
            self.stats_tree.insert("", "end", values=(class_name, math_data, rus_data))

    def run(self):
        """Запуск GUI"""
        self.root.mainloop()


if __name__ == '__main__':
    print("🚀 Запуск приложения управления олимпиадами...")
    gui = ServerGUI()
    gui.run()