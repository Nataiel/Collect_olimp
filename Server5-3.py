import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import pandas as pd
import json
import os
from flask import Flask, request, jsonify, render_template_string
import webbrowser
import re
from docx import Document
from docx.enum.text import WD_BREAK
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.styles import PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import socket

# Глобальные переменные
server_thread = None
server_running = False
data_file = 'olympiad_data.json'  # Файл хранения выбора олимпиад для каждого участника
excel_file = None
school_data = {}
template_path = None  # Путь к фалу шаблона заявления
template_diploma_path = None  # Путь к фалу шаблона грамоты
progress_bar = None  # Прогресс бар для отслеживания создания заявлений
progress_bar_diploma = None  # Прогресс бар для отслеживания создания грамот
data_diploma = None  # Данные для формирования грамот ВсОШ ШЭ

# Предметы для разных классов
subjects_4_class = ['Математика', 'Русский язык']
subjects_5_11_class = [
    'Англ. язык', 'Астрономия', 'Биология', 'География',
    'Инф.без.', 'ИИ', 'Искусство', 'Исп.язык', 'История', 'Итал.язык', 'Кит.язык',
    'Литература', 'Математика', 'Нем.язык', 'ОБЗР', 'Обществознание',
    'Право', 'Программирование', 'Робототехника', 'Русский язык', 'Труд',
    'Физика', 'Физ.культура', 'Фран.язык', 'Химия', 'Экология', 'Экономика'
]

olympiad_data = {}
first_run = True
flask_app = None


def get_subjects_for_class(class_name):
    """Определяет набор предметов в зависимости от класса"""
    match = re.search(r'(\d+)', class_name)
    if match:
        grade = int(match.group(1))
        if grade == 4:
            return subjects_4_class
        else:
            return subjects_5_11_class
    return subjects_5_11_class


def setup_flask_app():
    """Настройка Flask приложения"""
    global flask_app
    flask_app = Flask(__name__)

    MAIN_PAGE = """
    <!DOCTYPE html>
    <html>
    <head>
        <title>Участие в олимпиадах</title>
        <meta charset="utf-8">
        <style>
            body { 
                font-family: Arial, sans-serif; 
                max-width: 900px; 
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
            .grade-section {
                background: white;
                padding: 25px;
                border-radius: 10px;
                box-shadow: 0 2px 10px rgba(0,0,0,0.1);
                margin-bottom: 20px;
                display: flex;
                align-items: center;
                gap: 25px;
            }
            .grade-title {
                font-size: 24px;
                font-weight: bold;
                color: #2c3e50;
                min-width: 80px;
                text-align: center;
            }
            .class-grid {
                display: flex;
                gap: 12px;
                flex-wrap: wrap;
                flex: 1;
            }
            .class-card {
                background: #3498db;
                color: white;
                padding: 18px 22px;
                text-align: center;
                border-radius: 8px;
                text-decoration: none;
                font-size: 18px;
                font-weight: bold;
                box-shadow: 0 2px 8px rgba(0,0,0,0.2);
                transition: transform 0.2s, box-shadow 0.2s, background 0.2s;
                min-width: 50px;
            }
            .class-card:hover {
                transform: translateY(-3px);
                box-shadow: 0 5px 15px rgba(0,0,0,0.3);
                background: #2980b9;
            }
            .stats-link {
                display: block;
                margin-top: 30px;
                text-align: center;
                color: #3498db;
                text-decoration: none;
                font-weight: bold;
                font-size: 18px;
                padding: 15px;
                background: white;
                border-radius: 8px;
                box-shadow: 0 2px 8px rgba(0,0,0,0.1);
                transition: all 0.2s;
            }
            .stats-link:hover {
                text-decoration: none;
                background: #3498db;
                color: white;
                box-shadow: 0 4px 12px rgba(52,152,219,0.3);
            }
            .no-classes {
                text-align: center;
                color: #7f8c8d;
                font-style: italic;
                padding: 20px;
                background: white;
                border-radius: 10px;
            }
        </style>
    </head>
    <body>
        <div class="header">
            <h1>🏆 Участие в олимпиадах</h1>
            <div>Выберите класс для редактирования</div>
        </div>

        {% if classes_by_grade %}
            {% for grade, class_letters in classes_by_grade.items() %}
            <div class="grade-section">
                <div class="grade-title">{{ grade }} класс</div>
                <div class="class-grid">
                    {% for class_letter in class_letters %}
                    <a href="/class/{{ grade }}{{ class_letter }}" class="class-card">{{ class_letter }}</a>
                    {% endfor %}
                </div>
            </div>
            {% endfor %}
        {% else %}
            <div class="no-classes">Классы не загружены</div>
        {% endif %}

        <a href="/stats" class="stats-link">📊 Посмотреть статистику</a>
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
                max-width: 98vw; 
                margin: 10px auto; 
                padding: 10px;
                background: #f5f5f5;
            }
            .header {
                background: white;
                padding: 5px;
                border-radius: 10px;
                box-shadow: 0 2px 10px rgba(0,0,0,0.1);
                margin-bottom: 10px;
                position: sticky;
                top: 0;
                z-index: 100;
            }
            .back-link { 
                display: inline-block; 
                text-decoration: none; 
                color: #3498db; 
                font-weight: bold; 
                padding: 10px 20px;
                background: white;
                border-radius: 5px;
                box-shadow: 0 2px 5px rgba(0,0,0,0.1);
            }
            .back-link:hover {
                background: #3498db;
                color: white;
            }
            .table-container {
                max-height: 70vh;
                overflow: auto;
                background: white;
                border-radius: 10px;
                box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            }
            table {
                width: 100%;
                border-collapse: collapse;
                background: white;
                font-size: 13px;
                table-layout: fixed;
            }
            thead {
                position: sticky;
                top: 0;
                z-index: 10;
            }
            th, td {
                padding: 2px 2px;

                border-bottom: 1px solid #ecf0f1;
                border-right: 1px solid #ecf0f1;
            }
            th {
                background: white;
                color: #34495e;
                position: relative;
                padding: 0;
            }
            .student-name {
                min-width: 120px;
                width: 120px;
                background: #f8f9fa;
                text-align: left;
                padding-: 8px;
                font-size: 14px;
                vertical-align: middle !important;
                border-right: 2px solid #ddd;
            }
            .subject-cell {
                display: flex;
                flex-direction: column;
                height: 100%;
                min-height: 100px;
                width: 100%;
                padding: 0;
            }
            .subject-name-wrapper {
                flex: 1;
                display: flex;
                align-items: center;
                justify-content: center;
                background: #fafafa;
                padding: 0;
                min-height: 150px;
                cursor: pointer;
                user-select: none;
                border-bottom: 1px solid #ecf0f1;
                transition: background 0.2s;
            }
            .subject-name-wrapper:hover {
                background: #e8f4fd;
            }
            .subject-name-wrapper .subject-text {
                writing-mode: vertical-lr;
                transform: rotate(180deg);
                font-weight: bold;
                font-size: 13px;
                letter-spacing: 1px;
                padding: 4px 0;
                text-orientation: mixed;
            }
            .subject-icons {
                display: flex;
                justify-content: center;
                align-items: center;
                background: white;
                min-height: 30px;
                flex-shrink: 0;
            }
            .subject-icons .icon-item {
                font-size: 18px;
                line-height: 1;
            }
            .subject-icons .icon-item.home-icon {
                color: #27ae60;
            }
            .subject-icons .icon-item.school-icon {
                color: #2980b9;
            }
            .radio-cell {
                text-align: center;
                padding: 2px 2px;
                vertical-align: middle;
            }
            .radio-group {
                display: flex;
                justify-content: center;
                align-items: center;
                gap: 4px;
            }
            .radio-group label {
                display: flex;
                align-items: center;
                gap: 1px;
                font-size: 10px;
                cursor: pointer;
                padding: 1px 2px;
                border-radius: 3px;
                transition: background 0.15s;
            }
            .radio-group label:hover {
                background: #f0f0f0;
            }
            .radio-group input[type="radio"] {
                margin: 0;
                cursor: pointer;
                transform: scale(1.0);
            }
            .radio-group .home-option {
                color: #27ae60;
            }
            .radio-group .school-option {
                color: #2980b9;
            }
            tr:hover {
                background: #f8f9fa;
            }
            tr:hover .student-name {
                background: #e8f4fd;
            }
            .save-btn {
                background: #27ae60;
                color: white;
                padding: 12px 30px;
                border: none;
                border-radius: 5px;
                cursor: pointer;
                font-size: 14px;
                margin-top: 5px;
                margin-right: 10px;
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
                font-size: 14px;
                margin-top: 5px;
            }
            .reset-btn:hover {
                background: #c0392b;
            }
            .button-group {
                text-align: center;
                position: sticky;
                bottom: 0;
                background: #f5f5f5;
                padding: 5px 0;
                margin-top: 10px;
            }
            .message {
                padding: 10px;
                margin: 10px 0;
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
            .student-row:hover {
                background: #f0f8ff;
            }
            .table-container::-webkit-scrollbar {
                width: 8px;
                height: 8px;
            }
            .table-container::-webkit-scrollbar-track {
                background: #f1f1f1;
                border-radius: 4px;
            }
            .table-container::-webkit-scrollbar-thumb {
                background: #c1c1c1;
                border-radius: 4px;
            }
            .table-container::-webkit-scrollbar-thumb:hover {
                background: #a8a8a8;
            }
            .legend {
                display: flex;
                gap: 20px;
                justify-content: center;
                margin: 5px 0 10px 0;
                font-size: 13px;
            }
            .legend-item {
                display: flex;
                align-items: center;
                gap: 5px;
            }
            .legend-item .home-dot {
                color: #27ae60;
                font-weight: bold;
            }
            .legend-item .school-dot {
                color: #2980b9;
                font-weight: bold;
            }
            .info-text {
                font-size: 13px;
                color: #7f8c8d;
                margin-top: 5px;
            }
            .col-student {
                width: 120px;
                min-width: 120px;
                max-width: 120px;
            }
            .col-subject {
                width: 50px;
                min-width: 50px;
                max-width: 50px;
            }

            /* Блок выбора направления Труд - табличный вид */
            .labor-section {
                background: white;
                padding: 20px 25px;
                border-radius: 10px;
                box-shadow: 0 2px 10px rgba(0,0,0,0.1);
                margin-top: 15px;
            }
            .labor-title {
                font-size: 16px;
                font-weight: bold;
                color: #2c3e50;
                padding-bottom: 10px;
            }
            .labor-table {
                width: 50%;
                border-collapse: collapse;
            }
            .labor-table th {
                background: #f8f9fa;
                color: #2c3e50;
                font-weight: bold;
                padding: 10px 12px;
                text-align: left;
                border-bottom: 2px solid #e67e22;
            }
            .labor-table td {
                padding: 8px 12px;
                border-bottom: 1px solid #ecf0f1;
                vertical-align: middle;
            }
            .labor-table tr:hover {
                background: #f8f9fa;
            }
            .labor-table .student-name-col {
                font-weight: 500;
                color: #2c3e50;
                width: 40%;
            }
            .labor-table select {
                width: 100%;
                padding: 6px 10px;
                border: 1px solid #ddd;
                border-radius: 4px;
                font-size: 13px;
                background: white;
                cursor: pointer;
            }
            .labor-table select:focus {
                border-color: #e67e22;
                outline: none;
                box-shadow: 0 0 5px rgba(230, 126, 34, 0.3);
            }
            .labor-note {
                font-size: 12px;
                color: #7f8c8d;
                margin-top: 12px;
                font-style: italic;
                padding-top: 10px;
                border-top: 1px solid #ecf0f1;
            }
            .subject-icons .icon-item {
                font-size: 18px;
                border-radius: 4px;
                transition: background 0.2s, transform 0.2s;
                cursor: pointer;
            }
            .subject-icons .icon-item:hover {
                background: #e8f4fd;
                transform: scale(1.15);
            }
            .subject-icons .icon-item.home-icon {
                color: #27ae60;
            }
            .subject-icons .icon-item.home-icon:hover {
                background: #d4edda;
            }
            .subject-icons .icon-item.school-icon {
                color: #2980b9;
            }
            .subject-icons .icon-item.school-icon:hover {
                background: #d1ecf1;
            }
        </style>
    </head>
    <body>
        <div class="header">
            <a href="/" class="back-link">← Назад к выбору класса</a>
            <h2>🎓 {{ class_name }} класс - Участие в олимпиадах</h2>
        </div>

        {% if message %}
        <div class="message success">{{ message }}</div>
        {% endif %}

        <form action="/save/{{ class_name }}" method="post">
            <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            <th class="student-name col-student" rowspan="2">Ученик</th>
                            {% for subject in subjects %}
                            <th class="col-subject">
                            <div class="subject-cell">
                                <div class="subject-name-wrapper">
                                    <span class="subject-text">{{ subject }}</span>
                                </div>
                                <div class="subject-icons">
                                    <span class="icon-item home-icon" 
                                          onclick="toggleColumn('{{ subject|escape }}', 'home')"
                                          data-subject="{{ subject|escape }}"
                                          title="Выбрать 'Дома' для всех учеников по предмету {{ subject }}">🏠</span>
                                    <span class="icon-item school-icon" 
                                          onclick="toggleColumn('{{ subject|escape }}', 'school')"
                                          data-subject="{{ subject|escape }}"
                                          title="Выбрать 'В школе' для всех учеников по предмету {{ subject }}">🏫</span>
                                </div>
                            </div>
                        </th>
                        {% endfor %}
                        </tr>
                        <tr>
                            <!-- Пустая строка для заголовков -->
                        </tr>
                    </thead>
                    <tbody>
                        {% for student in students %}
                        <tr class="student-row">
                            <td class="student-name col-student">{{ student }}</td>
                            {% for subject in subjects %}
                            <td class="radio-cell col-subject">
                                <div class="radio-group">
                                    <label class="home-option">
                                        <input type="radio" 
                                               name="{{ student }}_{{ subject }}" 
                                               value="home"
                                               {{ 'checked' if participation[student][subject] == 'home' else '' }}>
                                    </label>
                                    <label class="school-option">
                                        <input type="radio" 
                                               name="{{ student }}_{{ subject }}" 
                                               value="school"
                                               {{ 'checked' if participation[student][subject] == 'school' else '' }}>
                                    </label>
                                </div>
                            </td>
                            {% endfor %}
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>

            <!-- Блок выбора направления для предмета Труд - табличный вид -->
            {% if 'Труд' in subjects %}
            <div class="labor-section">
                <div class="labor-title">🔧 Выбор направления по предмету «Труд»</div>
                <table class="labor-table">
                    <thead>
                        <tr>
                            <th>ФИО ученика</th>
                            <th>Направление</th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for student in students %}
                        <tr>
                            <td class="student-name-col">{{ student }}</td>
                            <td>
                                <select name="labor_{{ student }}">
                                    <option value="">Не выбрано</option>
                                    <option value="Технологии обработки материалов" 
                                            {% if labor_data and labor_data[student] == 'Технологии обработки материалов' %}selected{% endif %}>
                                        Технологии обработки материалов
                                    </option>
                                    <option value="Конструирование и моделирование"
                                            {% if labor_data and labor_data[student] == 'Конструирование и моделирование' %}selected{% endif %}>
                                        Конструирование и моделирование
                                    </option>
                                    <option value="Дизайн и интерьер"
                                            {% if labor_data and labor_data[student] == 'Дизайн и интерьер' %}selected{% endif %}>
                                        Дизайн и интерьер
                                    </option>
                                    <option value="Декоративно-прикладное искусство"
                                            {% if labor_data and labor_data[student] == 'Декоративно-прикладное искусство' %}selected{% endif %}>
                                        Декоративно-прикладное искусство
                                    </option>
                                    <option value="Технологии обработки конструкционных материалов"
                                            {% if labor_data and labor_data[student] == 'Технологии обработки конструкционных материалов' %}selected{% endif %}>
                                        Технологии обработки конструкционных материалов
                                    </option>
                                    <option value="Машиноведение и элементы механики"
                                            {% if labor_data and labor_data[student] == 'Машиноведение и элементы механики' %}selected{% endif %}>
                                        Машиноведение и элементы механики
                                    </option>
                                    <option value="Электротехника и электроника"
                                            {% if labor_data and labor_data[student] == 'Электротехника и электроника' %}selected{% endif %}>
                                        Электротехника и электроника
                                    </option>
                                    <option value="Техническое моделирование и конструирование"
                                            {% if labor_data and labor_data[student] == 'Техническое моделирование и конструирование' %}selected{% endif %}>
                                        Техническое моделирование и конструирование
                                    </option>
                                </select>
                            </td>
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
                <div class="labor-note">📌 Выберите направление для каждого ученика, если он выбрал предмет «Труд»</div>
            </div>
            {% endif %}

            <div class="button-group">
                <button type="submit" class="save-btn">💾 Сохранить изменения</button>
                <button type="button" class="reset-btn" onclick="resetClass()">🔄 Сбросить класс</button>
                <a href="/print/{{ class_name }}" target="_blank"><button type="button" style="padding: 10px 20px; font-size: 16px; background-color: #007bff; color: white; border: none; cursor: pointer; border-radius: 5px; margin-left: 10px;">🖨️ Распечатать выборы</button></a>

            </div>
        </form>

        <script>
            function resetClass() {
                if (confirm('Вы уверены, что хотите сбросить ВСЕ выборы для этого класса? Это действие нельзя отменить.')) {
                    window.location.href = '/reset/{{ class_name }}';
                }
            }

            function toggleColumn(subjectName, value) {
                const allRadios = document.querySelectorAll('input[type="radio"]');
                const targetRadios = [];

                allRadios.forEach(radio => {
                    if (radio.name.endsWith('_' + subjectName) && radio.value === value) {
                        targetRadios.push(radio);
                    }
                });

                if (targetRadios.length === 0) {
                    console.warn('Не найдены radio для предмета:', subjectName);
                    return;
                }

                // Проверяем, все ли уже выбраны
                const allChecked = targetRadios.every(radio => radio.checked);
                const newState = !allChecked;

                targetRadios.forEach(radio => {
                    radio.checked = newState;
                });

                const label = value === 'home' ? '🏠 Дома' : '🏫 В школе';
                console.log(`Предмет "${subjectName}": установлено ${label} для ${targetRadios.length} учеников`);
            }

            document.addEventListener('DOMContentLoaded', function() {
                // Добавляем всплывающие подсказки для иконок
                document.querySelectorAll('.icon-item.home-icon').forEach(el => {
                    const subject = el.getAttribute('data-subject') || '';
                    el.setAttribute('title', `Выбрать "Дома" для всех учеников по предмету "${subject}"`);
                    el.style.cursor = 'pointer';
                });
                document.querySelectorAll('.icon-item.school-icon').forEach(el => {
                    const subject = el.getAttribute('data-subject') || '';
                    el.setAttribute('title', `Выбрать "В школе" для всех учеников по предмету "${subject}"`);
                    el.style.cursor = 'pointer';
                });
            });
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
            body { 
                font-family: Arial, sans-serif; 
                max-width: 1200px; 
                margin: 50px auto; 
                padding: 20px;
                background: #f5f5f5;
            }
            .back-link { 
                display: inline-block; 
                text-decoration: none; 
                color: #3498db; 
                font-weight: bold; 
                padding: 10px 20px;
                background: white;
                border-radius: 5px;
                box-shadow: 0 2px 5px rgba(0,0,0,0.1);
            }
            .back-link:hover {
                background: #3498db;
                color: white;
            }
            .quick-nav {
                background: white;
                padding: 15px 20px;
                border-radius: 10px;
                box-shadow: 0 2px 10px rgba(0,0,0,0.1);
                margin-bottom: 25px;
                display: flex;
                flex-wrap: wrap;
                align-items: center;
                gap: 8px 15px;
            }
            .quick-nav .nav-label {
                font-weight: bold;
                color: #2c3e50;
                margin-right: 5px;
            }
            .quick-nav .grade-group {
                display: inline-flex;
                align-items: center;
                gap: 5px;
                background: #f8f9fa;
                padding: 3px 10px 3px 8px;
                border-radius: 20px;
            }
            .quick-nav .grade-label {
                font-weight: bold;
                color: #34495e;
                font-size: 0.9em;
            }
            .quick-nav a {
                text-decoration: none;
                color: #3498db;
                padding: 2px 8px;
                border-radius: 12px;
                font-size: 0.9em;
                transition: background 0.2s, color 0.2s;
            }
            .quick-nav a:hover {
                background: #3498db;
                color: white;
            }
            .stat-card { 
                background: white; 
                padding: 25px; 
                margin: 20px 0; 
                border-radius: 10px; 
                box-shadow: 0 2px 10px rgba(0,0,0,0.1); 
            }
            .class-title {
                color: #2c3e50;
                border-bottom: 2px solid #3498db;
                padding-bottom: 5px;
            }
            .subjects-grid {
                display: grid;
                grid-template-columns: repeat(auto-fill, minmax(300px, 1fr));
                gap: 15px;
                margin-top: 15px;
            }
            .subject-item {
                background: #f8f9fa;
                padding: 15px;
                border-radius: 8px;
                border-left: 4px solid #3498db;
            }
            .subject-name {
                font-weight: bold;
                color: #2c3e50;
                margin-bottom: 5px;
            }
            .subject-stats {
                display: flex;
                justify-content: space-between;
                color: #666;
                flex-wrap: wrap;
                gap: 5px;
            }
            .subject-stats .stat-item {
                display: flex;
                align-items: center;
                gap: 3px;
            }
            .subject-count {
                font-weight: bold;
                color: #27ae60;
            }
            .subject-count-home {
                font-weight: bold;
                color: #27ae60;
            }
            .subject-count-school {
                font-weight: bold;
                color: #2980b9;
            }
            .subject-percent {
                font-weight: bold;
                color: #e74c3c;
            }
            .total-stats {
                background: white;
                color: #34495e;
                padding: 15px;
                border-radius: 8px;
                margin-top: 15px;
                text-align: center;
                display: flex;
                justify-content: center;
                gap: 30px;
                flex-wrap: wrap;
            }
            .total-stats .stat-block {
                display: flex;
                flex-direction: column;
                align-items: center;
            }
            .total-stats .stat-block .label {
                font-size: 12px;
                color: #7f8c8d;
            }
            .total-stats .stat-block .value {
                font-size: 18px;
                font-weight: bold;
            }
            .total-stats .stat-block .value.home-value {
                color: #27ae60;
            }
            .total-stats .stat-block .value.school-value {
                color: #2980b9;
            }
            .grade-section {
                margin-bottom: 30px;
            }
            .grade-title {
                font-size: 24px;
                color: #2c3e50;
                margin-bottom: 15px;
                padding: 10px;
                background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                color: white;
                border-radius: 8px;
                text-align: center;
            }
            .to-top-link {
                text-align: right;
                margin-top: 15px;
                padding-top: 10px;
                border-top: 1px solid #ecf0f1;
            }
            .to-top-link a {
                color: #3498db;
                text-decoration: none;
                font-weight: bold;
                padding: 5px 12px;
                border-radius: 15px;
                transition: background 0.2s;
            }
            .to-top-link a:hover {
                background: #e8f4fd;
                text-decoration: underline;
            }
            html {
                scroll-behavior: smooth;
            }
            #top {
                position: relative;
                top: -20px;
            }
        </style>
    </head>
    <body>
        <div id="top"></div>

        <a href="/" class="back-link">← Назад к выбору класса</a>

        <div class="quick-nav">
            <span class="nav-label">📌 Перейти к классу:</span>
            {% for grade, classes_data in stats_by_grade.items() %}
                <span class="grade-group">
                    <span class="grade-label">{{ grade }} кл.</span>
                    {% for class_data in classes_data %}
                        <a href="#class-{{ class_data.class_name }}">{{ class_data.class_name }}</a>
                    {% endfor %}
                </span>
            {% endfor %}
        </div>

        <h1>📊 Статистика участия в олимпиадах</h1>

        {% for grade, classes_data in stats_by_grade.items() %}
        <div class="grade-section">
            <div class="grade-title">{{ grade }} классы</div>
            {% for class_data in classes_data %}
            <div class="stat-card" id="class-{{ class_data.class_name }}">
                <h2 class="class-title">🎓 Класс {{ class_data.class_name }}</h2>

                <div class="total-stats">
                    <div class="stat-block">
                        <span class="label">Всего учеников</span>
                        <span class="value">{{ class_data.total_students }}</span>
                    </div>
                    <div class="stat-block">
                        <span class="label">🏠 Дома</span>
                        <span class="value home-value">{{ class_data.total_home }}</span>
                    </div>
                    <div class="stat-block">
                        <span class="label">🏫 В школе</span>
                        <span class="value school-value">{{ class_data.total_school }}</span>
                    </div>
                    <div class="stat-block">
                        <span class="label">Всего выборов</span>
                        <span class="value">{{ class_data.total_choices }}</span>
                    </div>
                </div>

                <div class="subjects-grid">
                    {% for subject_data in class_data.subjects %}
                    <div class="subject-item">
                        <div class="subject-name">{{ subject_data.name }}</div>
                        <div class="subject-stats">
                            <span class="stat-item">
                                🏠 <span class="subject-count-home">{{ subject_data.home }}</span>
                            </span>
                            <span class="stat-item">
                                🏫 <span class="subject-count-school">{{ subject_data.school }}</span>
                            </span>
                            <span class="stat-item">
                                Всего: <span class="subject-count">{{ subject_data.total }}</span>
                            </span>
                            <span class="stat-item">
                                <span class="subject-percent">{{ subject_data.percent }}%</span>
                            </span>
                        </div>
                    </div>
                    {% endfor %}
                </div>

                <div class="to-top-link">
                    <a href="#top">↑ Наверх</a>
                </div>
            </div>
            {% endfor %}
        </div>
        {% endfor %}
    </body>
    </html>
    """

    PRINT_PAGE_TEMPLATE = """
    <!DOCTYPE html>
    <html lang="ru">
    <head>
        <meta charset="UTF-8">
        <title>Выборы {{ class_name }}</title>
        <style>
            body { font-family: 'Times New Roman', Times, serif; margin: 20px; font-size: 12pt; }
            h2 { text-align: center; margin-bottom: 20px; }
            table { width: 100%; border-collapse: collapse; }
            th, td { border: 1px solid black; padding: 6px; text-align: left; vertical-align: top; }
            th { background-color: #f2f2f2; text-align: center; font-weight: bold; }
            .no-print { text-align: center; margin-bottom: 20px; }
            .no-print button, .no-print a { padding: 8px 15px; font-size: 14pt; cursor: pointer; text-decoration: none; color: white; background-color: #007BFF; border: none; border-radius: 4px; }
            .no-print a.back { background-color: #6c757d; }
            @media print {
                .no-print { display: none; }
                @page { margin: 1.5cm; size: A4 landscape; } /* Альбомная ориентация для широкой таблицы */
                body { margin: 0; }
            }
        </style>
    </head>
    <body>
        <div class="no-print">
            <a href="/class/{{ class_name }}" class="back">← Вернуться к редактированию</a>
            <button onclick="window.print()">🖨️ Распечатать</button>
        </div>
        <h2>Сведения о выборе олимпиад: {{ class_name }} класс</h2>
        <table>
            <thead>
                <tr>
                    <th style="width: 5%;">№</th>
                    <th style="width: 30%;">ФИО ученика</th>
                    <th style="width: 25%;">Предметы (на дому)</th>
                    <th style="width: 25%;">Предметы (в школе)</th>
                    <th style="width: 15%;">Направление «Труд»</th>
                </tr>
            </thead>
            <tbody>
                {% for row in print_data %}
                <tr>
                    <td style="text-align: center;">{{ row.num }}</td>
                    <td>{{ row.name }}</td>
                    <td>{{ row.home }}</td>
                    <td>{{ row.school }}</td>
                    <td>{{ row.labor }}</td>
                </tr>
                {% endfor %}
            </tbody>
        </table>
    </body>
    </html>
    """

    @flask_app.route('/')
    def home():
        if not school_data:
            return "Данные не загружены. Загрузите Excel файл через GUI."

        classes_by_grade = {}

        for class_name in school_data.keys():
            match = re.search(r'(\d+)\s*([А-ЯA-Z])', class_name)
            if match:
                grade = match.group(1)
                class_letter = match.group(2)

                if grade not in classes_by_grade:
                    classes_by_grade[grade] = []

                if class_letter not in classes_by_grade[grade]:
                    classes_by_grade[grade].append(class_letter)

        sorted_grades = sorted(classes_by_grade.keys(), key=lambda x: int(x))
        for grade in sorted_grades:
            classes_by_grade[grade] = sorted(classes_by_grade[grade])

        return render_template_string(MAIN_PAGE, classes_by_grade=classes_by_grade)

    @flask_app.route('/class/<class_name>')
    def class_page(class_name):
        if class_name not in school_data:
            return "Класс не найден", 404

        students = school_data[class_name]
        participation = olympiad_data.get(class_name, {})
        class_subjects = get_subjects_for_class(class_name)

        # Получаем данные по направлениям Труда
        labor_data = olympiad_data.get(class_name, {}).get('_labor', {}) if class_name in olympiad_data else {}

        return render_template_string(
            CLASS_PAGE_TEMPLATE,
            class_name=class_name,
            students=students,
            subjects=class_subjects,
            participation=participation,
            labor_data=labor_data,
            message=request.args.get('message')
        )

    @flask_app.route('/save/<class_name>', methods=['POST'])
    def save_data(class_name):
        try:
            if class_name not in school_data:
                return "Класс не найден", 404

            class_subjects = get_subjects_for_class(class_name)

            # Сохраняем выборы по предметам
            for student in school_data[class_name]:
                for subject in class_subjects:
                    radio_name = f"{student}_{subject}"
                    value = request.form.get(radio_name)
                    if value in ['home', 'school']:
                        olympiad_data[class_name][student][subject] = value
                    else:
                        olympiad_data[class_name][student][subject] = None

            # Сохраняем направления Труда
            if 'Труд' in class_subjects:
                if '_labor' not in olympiad_data[class_name]:
                    olympiad_data[class_name]['_labor'] = {}

                for student in school_data[class_name]:
                    labor_value = request.form.get(f'labor_{student}')
                    if labor_value:
                        olympiad_data[class_name]['_labor'][student] = labor_value
                    else:
                        olympiad_data[class_name]['_labor'][student] = None

            save_olympiad_data()

            labor_data = olympiad_data[class_name].get('_labor', {})

            return render_template_string(
                CLASS_PAGE_TEMPLATE,
                class_name=class_name,
                students=school_data[class_name],
                subjects=class_subjects,
                participation=olympiad_data[class_name],
                labor_data=labor_data,
                message='✅ Данные успешно сохранены!'
            )

        except Exception as e:
            return f"Ошибка: {e}", 500

    @flask_app.route('/reset/<class_name>')
    def reset_class(class_name):
        """Сброс данных для конкретного класса"""
        if class_name in olympiad_data:
            class_subjects = get_subjects_for_class(class_name)
            for student in olympiad_data[class_name]:
                if student != '_labor':  # Не сбрасываем служебный ключ
                    for subject in class_subjects:
                        olympiad_data[class_name][student][subject] = None
            # Сбрасываем направления Труда
            if '_labor' in olympiad_data[class_name]:
                for student in olympiad_data[class_name]['_labor']:
                    olympiad_data[class_name]['_labor'][student] = None
            save_olympiad_data()

        class_subjects = get_subjects_for_class(class_name)
        labor_data = olympiad_data.get(class_name, {}).get('_labor', {})

        return render_template_string(
            CLASS_PAGE_TEMPLATE,
            class_name=class_name,
            students=school_data[class_name],
            subjects=class_subjects,
            participation=olympiad_data[class_name],
            labor_data=labor_data,
            message='✅ Данные класса сброшены!'
        )

    @flask_app.route('/stats')
    def stats_page():
        """Страница статистики"""
        stats_by_grade = {}

        for class_name in school_data:
            match = re.search(r'(\d+)', class_name)
            if match:
                grade = f"{match.group(1)}"
            else:
                grade = "Другие"

            if grade not in stats_by_grade:
                stats_by_grade[grade] = []

            class_subjects = get_subjects_for_class(class_name)
            total_students = len(school_data[class_name])
            total_choices = 0
            total_home = 0
            total_school = 0

            subjects_data = []
            for subject in class_subjects:
                home_count = 0
                school_count = 0
                for student_data in olympiad_data.get(class_name, {}).values():
                    if isinstance(student_data, dict):
                        value = student_data.get(subject)
                        if value == 'home':
                            home_count += 1
                        elif value == 'school':
                            school_count += 1

                total = home_count + school_count
                total_choices += total
                total_home += home_count
                total_school += school_count

                percent = round((total / total_students * 100) if total_students > 0 else 0, 1)

                subjects_data.append({
                    'name': subject,
                    'home': home_count,
                    'school': school_count,
                    'total': total,
                    'percent': percent
                })

            subjects_data.sort(key=lambda x: x['total'], reverse=True)

            stats_by_grade[grade].append({
                'class_name': class_name,
                'total_students': total_students,
                'total_choices': total_choices,
                'total_home': total_home,
                'total_school': total_school,
                'subjects': subjects_data
            })

        for grade in stats_by_grade:
            stats_by_grade[grade].sort(key=lambda x: x['class_name'])

        sorted_stats = {}
        for grade in sorted(stats_by_grade.keys(), key=lambda x: int(x) if x.isdigit() else 999):
            sorted_stats[grade] = stats_by_grade[grade]

        return render_template_string(STATS_PAGE, stats_by_grade=sorted_stats)

    @flask_app.route('/api/stats')
    def api_stats():
        stats = {}
        for class_name in school_data:
            class_subjects = get_subjects_for_class(class_name)
            stats[class_name] = {
                'total_students': len(school_data[class_name]),
                'participation_by_subject': {}
            }
            for subject in class_subjects:
                home_count = 0
                school_count = 0
                for student_data in olympiad_data.get(class_name, {}).values():
                    if isinstance(student_data, dict):
                        value = student_data.get(subject)
                        if value == 'home':
                            home_count += 1
                        elif value == 'school':
                            school_count += 1
                stats[class_name]['participation_by_subject'][subject] = {
                    'home': home_count,
                    'school': school_count
                }
        return jsonify(stats)

    @flask_app.route('/print/<class_name>')
    def print_class(class_name):
        """Маршрут для генерации страницы печати"""
        if class_name not in school_data:
            return "Класс не найден", 404

        students = school_data[class_name]
        participation = olympiad_data.get(class_name, {})
        class_subjects = get_subjects_for_class(class_name)
        labor_data = participation.get('_labor', {})

        # Подготовка данных для удобного вывода в таблицу печати
        print_data = []
        for idx, student in enumerate(students, start=1):
            home_subjects = []
            school_subjects = []
            student_choices = participation.get(student, {})

            for subject in class_subjects:
                if student_choices.get(subject) == 'home':
                    home_subjects.append(subject)
                elif student_choices.get(subject) == 'school':
                    school_subjects.append(subject)

            print_data.append({
                'num': idx,
                'name': student,
                'home': ', '.join(home_subjects),
                'school': ', '.join(school_subjects),
                'labor': labor_data.get(student, '')
            })

        return render_template_string(
            PRINT_PAGE_TEMPLATE,
            class_name=class_name,
            print_data=print_data
        )

def load_students_from_excel(file_path):
    """Загрузка данных учеников из Excel файла"""
    global school_data
    try:
        school_data = {}
        excel_data = pd.read_excel(file_path, sheet_name=None)

        allowed_classes = []
        for grade in range(4, 12):
            for letter in ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж', 'З', 'И', 'К', 'Л', 'М', 'Н', 'О', 'П', 'Р', 'С']:
                allowed_classes.append(f"{grade}{letter}")

        for sheet_name, df in excel_data.items():
            normalized_name = ''.join(sheet_name.split())
            sheet_name = ''.join(sheet_name.split())

            match = re.match(r'(\d+)\s*([А-ЯA-Z])', normalized_name)
            if not match:
                continue

            grade = int(match.group(1))
            if grade < 4 or grade > 11:
                continue

            if df.empty:
                continue

            students = []
            for index, row in df.iterrows():
                first_cell = str(row.iloc[0]).strip().lower()
                if first_cell in ['фамилия', 'фио', 'ученик', 'фамилия имя отчество']:
                    continue

                last_name = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
                first_name = str(row.iloc[1]).strip() if len(row) > 1 and pd.notna(row.iloc[1]) else ''
                middle_name = str(row.iloc[2]).strip() if len(row) > 2 and pd.notna(row.iloc[2]) else ''

                if not last_name or last_name == 'nan':
                    continue

                fio_parts = [last_name]
                if first_name and first_name != 'nan':
                    fio_parts.append(first_name)
                if middle_name and middle_name != 'nan':
                    fio_parts.append(middle_name)

                full_name = ' '.join(fio_parts)
                students.append(full_name)

            if students:
                school_data[sheet_name] = students
                print(f"Загружен класс {sheet_name}: {len(students)} учеников")

        print(f"Всего загружено классов: {len(school_data)}")
        print(f"Загруженные классы: {list(school_data.keys())}")
        return True

    except Exception as e:
        print(f"Ошибка загрузки Excel: {e}")
        return False


def reset_olympiad_data():
    """Сброс данных об участии (все радио не выбраны)"""
    global olympiad_data
    olympiad_data = {}
    for class_name in school_data:
        olympiad_data[class_name] = {}
        class_subjects = get_subjects_for_class(class_name)
        for student in school_data[class_name]:
            olympiad_data[class_name][student] = {subject: None for subject in class_subjects}
        # Инициализируем данные по Труду
        if 'Труд' in class_subjects:
            olympiad_data[class_name]['_labor'] = {student: None for student in school_data[class_name]}

    save_olympiad_data()
    print("✅ Данные олимпиад сброшены (все выборы не отмечены)")


def load_olympiad_data():
    """Загрузка данных об участии"""
    global olympiad_data
    if os.path.exists(data_file):
        print("📁 Загружаем существующие данные из файла...")
        with open(data_file, 'r', encoding='utf-8') as f:
            olympiad_data = json.load(f)

        data_updated = False
        for class_name in school_data:
            if class_name not in olympiad_data:
                olympiad_data[class_name] = {}
                data_updated = True

            class_subjects = get_subjects_for_class(class_name)

            # Инициализируем данные для учеников
            for student in school_data[class_name]:
                if student not in olympiad_data[class_name]:
                    olympiad_data[class_name][student] = {subject: None for subject in class_subjects}
                    data_updated = True
                else:
                    for subject in class_subjects:
                        if subject not in olympiad_data[class_name][student]:
                            olympiad_data[class_name][student][subject] = None
                            data_updated = True
                    current_subjects = list(olympiad_data[class_name][student].keys())
                    for existing_subject in current_subjects:
                        if existing_subject not in class_subjects and existing_subject != '_labor':
                            del olympiad_data[class_name][student][existing_subject]
                            data_updated = True

            # Инициализируем данные по Труду
            if 'Труд' in class_subjects:
                if '_labor' not in olympiad_data[class_name]:
                    olympiad_data[class_name]['_labor'] = {}
                    data_updated = True
                for student in school_data[class_name]:
                    if student not in olympiad_data[class_name]['_labor']:
                        olympiad_data[class_name]['_labor'][student] = None
                        data_updated = True

        if data_updated:
            save_olympiad_data()
            print("✅ Структура данных обновлена")
    else:
        print("📝 Файл данных не найден, создаем новый...")
        reset_olympiad_data()


def save_olympiad_data():
    """Сохранение данных об участии"""
    with open(data_file, 'w', encoding='utf-8') as f:
        json.dump(olympiad_data, f, ensure_ascii=False, indent=2)
    print("💾 Данные сохранены в файл")


def get_statistics():
    """Получение статистики для GUI"""
    stats = {}
    for class_name in school_data:
        stats[class_name] = {}
        class_subjects = get_subjects_for_class(class_name)
        for subject in class_subjects:
            home_count = 0
            school_count = 0
            for student_data in olympiad_data.get(class_name, {}).values():
                if isinstance(student_data, dict):
                    value = student_data.get(subject)
                    if value == 'home':
                        home_count += 1
                    elif value == 'school':
                        school_count += 1
            total = home_count + school_count
            stats[class_name][subject] = f"🏠{home_count}/🏫{school_count} (всего {total})"
    return stats


def run_server():
    """Запуск сервера"""
    global server_running, first_run
    server_running = True

    if first_run and os.path.exists(data_file):
        print("📋 Файл данных существует - загружаем предыдущие выборы")
    elif first_run:
        print("🆕 Первый запуск - создаем чистые данные")

    first_run = False
    flask_app.run(host='0.0.0.0', port=5000, debug=False, use_reloader=False)


def stop_server():
    """Остановка сервера"""
    global server_running
    server_running = False


# GUI функции
def load_excel_file_gui(file_label, stats_tree):
    """Загрузка Excel файла с данными учеников"""
    global excel_file
    file_path = filedialog.askopenfilename(
        title="Выберите Excel файл с данными учеников",
        filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
    )

    if file_path:
        try:
            if load_students_from_excel(file_path):
                excel_file = file_path

                if not school_data:
                    messagebox.showwarning("Предупреждение",
                                           "Не найдено классов 4-11 в файле.\n\n"
                                           "Убедитесь, что:\n"
                                           "1. В файле есть листы для классов 4-11\n"
                                           "2. Данные содержат столбцы с Фамилией, Именем и Отчеством")
                    return

                if os.path.exists(data_file):
                    response = messagebox.askyesno(
                        "Загрузка данных",
                        "Обнаружен файл с предыдущими выборами.\n\nЗагрузить существующие данные? (Да)\nИли сбросить все выборы? (Нет)"
                    )
                    if response:
                        print("📁 Загружаем существующие данные...")
                        load_olympiad_data()
                    else:
                        print("🔄 Сбрасываем все данные...")
                        reset_olympiad_data()
                else:
                    print("🆕 Создаем новые данные...")
                    load_olympiad_data()

                file_label.config(text=f"Загружен: {os.path.basename(file_path)}")
                messagebox.showinfo("Успех", f"Данные успешно загружены!\nКлассы: {', '.join(school_data.keys())}")
                update_statistics_gui(stats_tree)
            else:
                messagebox.showerror("Ошибка", "Не удалось загрузить данные из файла")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при загрузке файла: {e}")


def reset_all_data_gui(stats_tree):
    """Сброс всех данных олимпиад"""
    if messagebox.askyesno("Сброс данных",
                           "Вы уверены, что хотите сбросить ВСЕ данные олимпиад?\nВсе выборы будут удалены!"):
        reset_olympiad_data()
        update_statistics_gui(stats_tree)
        messagebox.showinfo("Сброс", "Все данные олимпиад сброшены!")


def start_server_gui(status_label):
    """Запуск сервера"""
    global server_thread
    if not school_data:
        messagebox.showwarning("Предупреждение", "Сначала загрузите Excel файл с данными")
        return

    if not server_running:
        server_thread = threading.Thread(target=run_server)
        server_thread.daemon = True
        server_thread.start()
        hostname = socket.gethostname()
        local_ip = socket.gethostbyname(hostname)
        status_label.config(text=f"Статус: Сервер запущен на http://{local_ip}:5000",
                            foreground="green")
        messagebox.showinfo("Сервер запущен", "Сервер успешно запущен!\nВы можете открыть его в браузере.")
        create_internet_shortcut()


def create_internet_shortcut():
    """
    Создает ярлык интернет-ресурса (.url)
    """

    hostname = socket.gethostname()
    local_ip = socket.gethostbyname(hostname)

    with open("Выбор предметов ВсОШ.url", 'w', encoding='utf-8') as f:
        f.write(f'[InternetShortcut]\n')
        f.write(f'URL=http://{local_ip}:5000\n')
        f.write(f'IDList=\n')
        f.write(f'IconIndex=0\n')
        f.write(f'HotKey=0\n')

    print(f"Ярлык создан: Выбор предметов ВсОШ.url")


def stop_server_gui(status_label):
    """Остановка сервера"""
    if server_running:
        stop_server()
        status_label.config(text="Статус: Сервер остановлен", foreground="red")
        messagebox.showinfo("Сервер остановлен", "Сервер был остановлен.")
    else:
        messagebox.showinfo("Информация", "Сервер уже остановлен")


def open_browser_gui():
    """Открытие браузера"""
    if server_running:
        hostname = socket.gethostname()
        local_ip = socket.gethostbyname(hostname)
        webbrowser.open(f"http://{local_ip}:5000")
    else:
        messagebox.showwarning("Предупреждение", "Сервер не запущен")


def update_statistics_gui(stats_tree):
    """Обновление статистики в таблице"""
    for item in stats_tree.get_children():
        stats_tree.delete(item)

    if not school_data or not olympiad_data:
        return

    for class_name, students in school_data.items():
        participating_students = 0

        if class_name in olympiad_data:
            for student, subjects_data in olympiad_data[class_name].items():
                if student == '_labor':
                    continue
                if any(v is not None for v in subjects_data.values()):
                    participating_students += 1

        total_students = len(students)
        if total_students > 0:
            participation_percentage = (participating_students / total_students) * 100
        else:
            participation_percentage = 0

        percentage_text = f"{participation_percentage:.1f}%"

        stats_tree.insert("", "end", values=(class_name, percentage_text))


def save_diploma_to_xlsx():
    global data_diploma
    protocol = filedialog.askopenfilename(
        title="Выберите протокол с результатами ВсОШ ШЭ",
        filetypes=[("Excel documents", "*.xlsx"), ("All files", "*.*")])

    if not protocol:
        messagebox.showwarning("Предупреждение", "Файл не выбран")
        return

    sheets = pd.read_excel(protocol, sheet_name=None)
    res = pd.concat(sheets.values(), ignore_index=True)
    res = res[res["Статус"].isin(['Победитель', 'Призёр'])]
    res['Результат'] = res['Предмет'] + ' (' + res['Статус'] + ')'
    res['Класс участника'] = res['Класс участника'].str.upper().str.replace(":", "", regex=False)
    res[["Код", "ОО"]] = res["Школа"].str.split(" - ", n=1, expand=True)
    res = res[["Код", "ОО", 'Класс участника', 'Участник', 'Результат']]
    res = res.groupby(["Код", "ОО", "Класс участника", "Участник"])["Результат"].agg(", ".join).reset_index()

    aux_data_diploma = zip(res['Класс участника'], res['Участник'], res['Результат'], res['ОО'])
    data_diploma = dict()
    for x in aux_data_diploma:
        if not x[0] in data_diploma:
            data_diploma[x[0]] = [x[1:]]
        else:
            data_diploma[x[0]].append(x[1:])

    directory = os.path.dirname(protocol)
    filename = os.path.basename(protocol)
    new_filename = f"Обработан {filename}"
    new_protocol = os.path.join(directory, new_filename)

    res.to_excel(new_protocol, index=False)
    messagebox.showinfo("Успех", f"Данные сохранены в файл:\n{new_protocol}")


def save_to_xlsx():
    global data_file

    with open('olympiad_data.json', 'r', encoding='utf-8') as f:
        data = json.load(f)

    subject_data = defaultdict(list)
    students_data = 0

    # Собираем данные по Труду для отдельной страницы
    labor_data = []  # (класс, ученик, направление)

    for class_name, students in data.items():
        # Собираем данные по Труду
        if '_labor' in students:
            for student_name, direction in students['_labor'].items():
                if direction:  # если направление выбрано
                    labor_data.append((class_name, student_name, direction))

        # Остальные предметы
        for student_name, subjects in students.items():
            if student_name == '_labor':
                continue
            students_data += 1
            for subject, location in subjects.items():
                if location in ['home', 'school']:
                    location_label = "на дому" if location == 'home' else "в школе"
                    subject_data[subject].append({
                        'Класс': class_name,
                        'Ученик': student_name,
                        'Место': location_label
                    })

    wb = Workbook()

    default_sheet = wb.active
    wb.remove(default_sheet)

    # Создаем листы для каждого предмета
    for subject, records in sorted(subject_data.items()):
        if not records:
            continue

        sheet_name = str(subject)[:31]

        sheets = wb.sheetnames
        if not sheet_name in sheets:
            sheet = wb.create_sheet(title=sheet_name)
        else:
            continue

        headers = list(records[0].keys())

        for col_num, header in enumerate(headers, 1):
            sheet.cell(row=1, column=col_num).value = header
            sheet.cell(row=1, column=col_num).font = Font(bold=True)
            sheet.cell(row=1, column=col_num).alignment = Alignment(horizontal='center')
        sheet.column_dimensions['B'].width = 50
        sheet.column_dimensions['C'].width = 20

    sheets = wb.sheetnames
    class_data = list(data.keys())

    for sheet in sheets:
        dataset = []
        for aux in subject_data[sheet]:
            dataset.append((aux['Класс'], aux['Ученик'], aux['Место']))
        dataset.sort(key=lambda x: int(x[0][0]))

        wb.active = wb[sheet]
        activ_sheet = wb.active
        for data_row in dataset:
            activ_sheet.append(data_row)

        activ_sheet['D1'].value = 'Класс'
        activ_sheet['D1'].font = Font(bold=True)
        activ_sheet['D1'].alignment = Alignment(horizontal='center')
        activ_sheet['E1'].value = 'Кол-во'
        activ_sheet['E1'].font = Font(bold=True)
        activ_sheet['E1'].alignment = Alignment(horizontal='center')
        activ_sheet['F1'].value = 'Из них: на дому'
        activ_sheet['F1'].font = Font(bold=True)
        activ_sheet['F1'].alignment = Alignment(horizontal='center')
        activ_sheet['G1'].value = 'Из них: в школе'
        activ_sheet['G1'].font = Font(bold=True)
        activ_sheet['G1'].alignment = Alignment(horizontal='center')

        for row in range(len(class_data)):
            activ_sheet.cell(row=row + 2, column=4).value = class_data[row]
            activ_sheet.cell(row=row + 2, column=5).value = f'=COUNTIF(A:A, "{class_data[row]}")'
            activ_sheet.cell(row=row + 2, column=6).value = f'=COUNTIFS(A:A, "{class_data[row]}", C:C, "на дому")'
            activ_sheet.cell(row=row + 2, column=7).value = f'=COUNTIFS(A:A, "{class_data[row]}", C:C, "в школе")'

        activ_sheet.cell(row=row + 3, column=4).value = 'Всего: '
        activ_sheet.cell(row=row + 3, column=4).font = Font(bold=True)
        activ_sheet.cell(row=row + 3, column=4).alignment = Alignment(horizontal='center')
        end_cell = activ_sheet.cell(row=row + 2, column=5).coordinate
        activ_sheet.cell(row=row + 3, column=5).value = f'=SUM(E2:{end_cell})'
        activ_sheet.cell(row=row + 3, column=6).value = f'=SUM(F2:{get_column_letter(6)}{row + 2})'
        activ_sheet.cell(row=row + 3, column=7).value = f'=SUM(G2:{get_column_letter(7)}{row + 2})'

    # Страница статистики
    statistic_sheet = wb.create_sheet(title='Статистика')
    col = 2

    statistic_sheet['A1'].value = 'Класс'
    statistic_sheet['A1'].font = Font(bold=True)
    statistic_sheet['A1'].alignment = Alignment(horizontal='center')
    statistic_sheet.row_dimensions[1].height = 105
    statistic_sheet.column_dimensions['A'].width = 11
    for row in range(len(class_data)):
        statistic_sheet.cell(row=row + 2, column=col - 1).value = class_data[row]
        statistic_sheet.cell(row=row + 2, column=col - 1).alignment = Alignment(horizontal='center')

    statistic_sheet.cell(row=row + 3, column=col - 1).value = 'Всего:'
    statistic_sheet.cell(row=row + 3, column=col - 1).font = Font(bold=True)
    statistic_sheet.cell(row=row + 3, column=col - 1).alignment = Alignment(horizontal='center')

    statistic_sheet.cell(row=row + 4, column=col - 1).value = '% участия:'
    statistic_sheet.cell(row=row + 4, column=col - 1).font = Font(bold=True)
    statistic_sheet.cell(row=row + 4, column=col - 1).alignment = Alignment(horizontal='center')

    for sheet in sheets:
        start = statistic_sheet.cell(row=1, column=col).coordinate

        statistic_sheet[start] = sheet
        statistic_sheet[start].font = Font(bold=True)
        statistic_sheet[start].alignment = Alignment(horizontal='center', text_rotation=90)
        col_letter = get_column_letter(col)
        statistic_sheet.column_dimensions[col_letter].width = 7

        for row in range(len(class_data)):
            statistic_sheet.cell(row=row + 2, column=col).value = f"='{sheet}'!E{row + 2}"

        stop = statistic_sheet.cell(row=row + 2, column=col).coordinate
        statistic_sheet.cell(row=row + 3, column=col).value = f'=SUM({start}:{stop})'
        statistic_sheet.cell(row=row + 3, column=col).font = Font(bold=True)
        statistic_sheet.cell(row=row + 3, column=col).alignment = Alignment(horizontal='center')

        statistic_sheet.cell(row=row + 4, column=col).value = f'=SUM({start}:{stop})/{students_data}'
        statistic_sheet.cell(row=row + 4, column=col).alignment = Alignment(horizontal='center')
        statistic_sheet.cell(row=row + 4, column=col).number_format = '0.0%'

        col_home = col + 1
        statistic_sheet.cell(row=1, column=col_home).value = '🏠'
        statistic_sheet.cell(row=1, column=col_home).font = Font(bold=True)
        statistic_sheet.cell(row=1, column=col_home).alignment = Alignment(horizontal='center', text_rotation=90)
        statistic_sheet.column_dimensions[get_column_letter(col_home)].width = 4

        for row in range(len(class_data)):
            statistic_sheet.cell(row=row + 2, column=col_home).value = f"='{sheet}'!F{row + 2}"

        statistic_sheet.cell(row=row + 3,
                             column=col_home).value = f'=SUM({get_column_letter(col_home)}2:{get_column_letter(col_home)}{row + 2})'
        statistic_sheet.cell(row=row + 3, column=col_home).font = Font(bold=True)
        statistic_sheet.cell(row=row + 3, column=col_home).alignment = Alignment(horizontal='center')

        col_school = col + 2
        statistic_sheet.cell(row=1, column=col_school).value = '🏫'
        statistic_sheet.cell(row=1, column=col_school).font = Font(bold=True)
        statistic_sheet.cell(row=1, column=col_school).alignment = Alignment(horizontal='center', text_rotation=90)
        statistic_sheet.column_dimensions[get_column_letter(col_school)].width = 4

        for row in range(len(class_data)):
            statistic_sheet.cell(row=row + 2, column=col_school).value = f"='{sheet}'!G{row + 2}"

        statistic_sheet.cell(row=row + 3,
                             column=col_school).value = f'=SUM({get_column_letter(col_school)}2:{get_column_letter(col_school)}{row + 2})'
        statistic_sheet.cell(row=row + 3, column=col_school).font = Font(bold=True)
        statistic_sheet.cell(row=row + 3, column=col_school).alignment = Alignment(horizontal='center')

        row_end = row + 4
        gray_fill = PatternFill(start_color="F3F3F3", end_color="F3F3F3", fill_type="solid")
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                             top=Side(style='thin'), bottom=Side(style='thin'))
        if col % 2 == 0:
            for r in statistic_sheet.iter_rows(min_row=1, max_row=row_end, min_col=col, max_col=col_school):
                for cell in r:
                    cell.fill = gray_fill

        for r in statistic_sheet.iter_rows(min_row=1, max_row=row_end, min_col=1, max_col=col_school):
            for cell in r:
                cell.border = thin_border

        col += 3

    # ========== Новая страница: ТРУД ==========
    if labor_data:
        labor_sheet = wb.create_sheet(title='Труд (по блокам)')

        # Заголовки
        labor_sheet['A1'].value = 'Класс'
        labor_sheet['A1'].font = Font(bold=True)
        labor_sheet['A1'].alignment = Alignment(horizontal='center')
        labor_sheet.column_dimensions['A'].width = 12

        labor_sheet['B1'].value = 'ФИО ученика'
        labor_sheet['B1'].font = Font(bold=True)
        labor_sheet['B1'].alignment = Alignment(horizontal='center')
        labor_sheet.column_dimensions['B'].width = 35

        labor_sheet['C1'].value = 'Направление'
        labor_sheet['C1'].font = Font(bold=True)
        labor_sheet['C1'].alignment = Alignment(horizontal='center')
        labor_sheet.column_dimensions['C'].width = 40

        # Данные
        labor_sorted = sorted(labor_data, key=lambda x: int(x[0][0]) if x[0][0].isdigit() else 0)

        for row_idx, (class_name, student_name, direction) in enumerate(labor_sorted, start=2):
            labor_sheet.cell(row=row_idx, column=1).value = class_name
            labor_sheet.cell(row=row_idx, column=1).alignment = Alignment(horizontal='center')
            labor_sheet.cell(row=row_idx, column=2).value = student_name
            labor_sheet.cell(row=row_idx, column=3).value = direction

        # Статистика по направлениям
        start_stat_row = len(labor_sorted) + 4
        labor_sheet.cell(row=start_stat_row, column=1).value = 'СТАТИСТИКА ПО НАПРАВЛЕНИЯМ'
        labor_sheet.cell(row=start_stat_row, column=1).font = Font(bold=True, size=12)

        # Собираем статистику
        direction_stats = defaultdict(int)
        class_stats = defaultdict(lambda: defaultdict(int))

        for class_name, student_name, direction in labor_data:
            direction_stats[direction] += 1
            class_stats[class_name][direction] += 1

        # Выводим статистику по направлениям
        stat_row = start_stat_row + 2
        labor_sheet.cell(row=stat_row, column=1).value = 'Направление'
        labor_sheet.cell(row=stat_row, column=1).font = Font(bold=True)
        labor_sheet.cell(row=stat_row, column=2).value = 'Всего учеников'
        labor_sheet.cell(row=stat_row, column=2).font = Font(bold=True)

        for idx, (direction, count) in enumerate(sorted(direction_stats.items()), start=1):
            labor_sheet.cell(row=stat_row + idx, column=1).value = direction
            labor_sheet.cell(row=stat_row + idx, column=2).value = count
            labor_sheet.cell(row=stat_row + idx, column=2).alignment = Alignment(horizontal='center')

        # Статистика по классам
        stat_row_classes = stat_row + len(direction_stats) + 3
        labor_sheet.cell(row=stat_row_classes, column=1).value = 'СТАТИСТИКА ПО КЛАССАМ'
        labor_sheet.cell(row=stat_row_classes, column=1).font = Font(bold=True, size=12)

        stat_row_classes += 2
        labor_sheet.cell(row=stat_row_classes, column=1).value = 'Класс'
        labor_sheet.cell(row=stat_row_classes, column=1).font = Font(bold=True)

        # Определяем все уникальные направления
        all_directions = sorted(direction_stats.keys())
        for col_idx, direction in enumerate(all_directions, start=2):
            labor_sheet.cell(row=stat_row_classes, column=col_idx).value = direction
            labor_sheet.cell(row=stat_row_classes, column=col_idx).font = Font(bold=True)
            labor_sheet.cell(row=stat_row_classes, column=col_idx).alignment = Alignment(horizontal='center',
                                                                                         wrap_text=True)

        labor_sheet.cell(row=stat_row_classes, column=len(all_directions) + 2).value = 'Всего'
        labor_sheet.cell(row=stat_row_classes, column=len(all_directions) + 2).font = Font(bold=True)

        for class_idx, class_name in enumerate(
                sorted(class_stats.keys(), key=lambda x: int(x[0]) if x[0].isdigit() else 0)):
            row = stat_row_classes + class_idx + 1
            labor_sheet.cell(row=row, column=1).value = class_name
            labor_sheet.cell(row=row, column=1).alignment = Alignment(horizontal='center')

            total = 0
            for col_idx, direction in enumerate(all_directions, start=2):
                count = class_stats[class_name].get(direction, 0)
                labor_sheet.cell(row=row, column=col_idx).value = count
                labor_sheet.cell(row=row, column=col_idx).alignment = Alignment(horizontal='center')
                total += count

            labor_sheet.cell(row=row, column=len(all_directions) + 2).value = total
            labor_sheet.cell(row=row, column=len(all_directions) + 2).alignment = Alignment(horizontal='center')
            labor_sheet.cell(row=row, column=len(all_directions) + 2).font = Font(bold=True)

        # Применяем границы к статистике
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                             top=Side(style='thin'), bottom=Side(style='thin'))
        for r in labor_sheet.iter_rows(min_row=stat_row, max_row=stat_row + len(direction_stats),
                                       min_col=1, max_col=2):
            for cell in r:
                cell.border = thin_border

        for r in labor_sheet.iter_rows(min_row=stat_row_classes,
                                       max_row=stat_row_classes + len(class_stats),
                                       min_col=1, max_col=len(all_directions) + 2):
            for cell in r:
                cell.border = thin_border

    wb.save('Отчёт о выборе предметов ВсОШ ШЭ.xlsx')
    print('Выгрузка в Excel завершена')


def create_diploma():
    global data_diploma
    x = 0

    if not template_diploma_path:
        messagebox.showerror("Ошибка", "Сначала выберите шаблон грамоты")
        return

    current_dir = os.getcwd()
    diplomas_folder = os.path.join(current_dir, "Грамоты")

    if not os.path.exists(diplomas_folder):
        os.makedirs(diplomas_folder)
        print(f"Создана папка: {diplomas_folder}")

    for class_name, data in data_diploma.items():
        current_doc = Document(template_diploma_path)

        last_paragraph = current_doc.paragraphs[-1]
        run = last_paragraph.add_run()
        run.add_break(WD_BREAK.PAGE)

        for student_idx, (student_name, aux_subject, school) in enumerate(data):
            template_doc = Document(template_diploma_path)

            status = ''
            if 'Победитель' in aux_subject and 'Призёр' in aux_subject:
                status = 'победитель и призёр'
            elif 'Призёр' in aux_subject:
                status = 'призёр'
            elif 'Победитель' in aux_subject:
                status = 'победитель'

            count_subject = 'у'
            if ',' in aux_subject:
                count_subject = 'ам'

            name_parts = student_name.split()
            replacements = {
                '%ОО%': school,
                '%Класс%': class_name,
                '%Предметы%': aux_subject,
                '%Фамилия%': name_parts[0] if len(name_parts) > 0 else '',
                '%Имя%': name_parts[1] if len(name_parts) > 1 else '',
                '%призёр и победитель%': status,
                '%окон%': count_subject
            }

            replace_placeholders_in_document(template_doc, replacements)

            for element in template_doc.element.body:
                current_doc.element.body.append(element)

            if student_idx < len(data) - 1:
                if current_doc.paragraphs:
                    last_paragraph = current_doc.paragraphs[-1]
                    run = last_paragraph.add_run()
                    run.add_break(WD_BREAK.PAGE)

        x += 100 / len(data_diploma)
        progress_bar_diploma['value'] = x
        progress_bar_diploma.update()

        filename = f'{class_name}_Грамоты_ВСоШ_ШЭ.docx'
        filename = os.path.join(diplomas_folder, filename)
        current_doc.save(filename)


def replace_placeholders_in_document(doc, replacements):
    """
    Заменяет плейсхолдеры во всём документе, сохраняя форматирование
    replacements: словарь {плейсхолдер: значение}
    """
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            text = run.text
            new_text = text

            for placeholder, value in replacements.items():
                if placeholder in new_text:
                    new_text = new_text.replace(placeholder, str(value))

            run.text = new_text

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        text = run.text
                        new_text = text
                        for placeholder, value in replacements.items():
                            if placeholder in new_text:
                                new_text = new_text.replace(placeholder, str(value))
                        run.text = new_text


def create_statement():
    global data_file
    x = 0

    current_dir = os.getcwd()
    statements_folder = os.path.join(current_dir, "Заявления")

    if not os.path.exists(statements_folder):
        os.makedirs(statements_folder)
        print(f"Создана папка: {statements_folder}")

    with open('olympiad_data.json', 'r', encoding='utf-8') as f:
        data = json.load(f)

    for class_name, students in data.items():
        current_doc = Document()
        flag_statement = False

        for student_name, subjects in students.items():
            if student_name == '_labor':
                continue

            home_subjects = []
            school_subjects = []

            for subject, location in subjects.items():
                if location == 'home':
                    home_subjects.append(subject)
                elif location == 'school':
                    school_subjects.append(subject)

            if not home_subjects and not school_subjects:
                continue

            flag_statement = True

            home_str = ', '.join(home_subjects) if home_subjects else ''
            school_str = ', '.join(school_subjects) if school_subjects else ''

            replacements = (
                ('Англ. язык', 'Английский язык'),
                ('Инф.без.', 'Информатика (информационная безопасность)'),
                ('ИИ', 'Информатика (искусственный интеллект)'),
                ('Искусство', 'Искусство МХК'),
                ('Исп.язык', 'Испанский язык'),
                ('Итал.язык', 'Итальянский язык'),
                ('Кит.язык', 'Китайский язык'),
                ('Нем.язык', 'Немецкий язык'),
                ('ОБЗР', 'Основы безопасности и защиты Родины'),
                ('Программирование', 'Информатика (программирование)'),
                ('Робототехника', 'Информатика (робототехника)'),
                ('Физ.культура', 'Физическая культура'),
                ('Фран.язык', 'Французский язык')
            )

            for a, b in replacements:
                home_str = home_str.replace(a, b)
                school_str = school_str.replace(a, b)

            full_subjects = []
            if home_str:
                full_subjects.append(f"на дому: {home_str}")
            if school_str:
                full_subjects.append(f"в школе: {school_str}")

            subjects_text = '; '.join(full_subjects)

            if current_doc.paragraphs:
                last_paragraph = current_doc.paragraphs[-1]
                run = last_paragraph.add_run()
                run.add_break(WD_BREAK.PAGE)

            new_doc = Document(template_path)

            for element in new_doc.element.body:
                current_doc.element.body.append(element)

            paragraphs = current_doc.paragraphs
            start_index = 0

            for idx, para in enumerate(paragraphs):
                if "ЗАЯВЛЕНИЕ" in para.text and idx > start_index:
                    start_index = idx
                    break

            for para in paragraphs[start_index:]:
                for run in para.runs:
                    text = run.text
                    if '%ФИО%' in text:
                        run.text = text.replace('%ФИО%', student_name)
                    elif '%Класс%' in text:
                        run.text = text.replace('%Класс%', class_name)
                    elif '%олимпиады%' in text:
                        run.text = text.replace('%олимпиады%', subjects_text)

        x += 100 / len(data)
        progress_bar['value'] = x
        progress_bar.update()

        filename = f'Заявление ВсОШ ШЭ {class_name}.docx'
        filename = os.path.join(statements_folder, filename)
        current_doc.save(filename)


def select_template():
    """Выбор шаблона заявления в формате *.docx"""
    global template_path

    file_path = filedialog.askopenfilename(
        title="Выберите шаблон заявления",
        filetypes=[("Word documents", "*.docx"), ("All files", "*.*")]
    )

    if file_path:
        template_path = file_path
        return True
    return False


def select_template_diploma():
    """Выбор шаблона грамот в формате *.docx"""
    global template_diploma_path

    file_path = filedialog.askopenfilename(
        title="Выберите шаблон грамоты",
        filetypes=[("Word documents", "*.docx"), ("All files", "*.*")]
    )

    if file_path:
        template_diploma_path = file_path
        return True
    return False


def update_template_status(template_status_label):
    """Обновление статуса шаблона"""
    global template_path

    if template_path:
        template_status_label.config(text=f"Шаблон: {template_path}", foreground="green")
    else:
        template_status_label.config(text="Шаблон не выбран", foreground="gray")


def update_template_diploma_status(template_status_label):
    """Обновление статуса шаблона"""
    global template_diploma_path

    if template_diploma_path:
        template_status_label.config(text=f"Шаблон: {template_diploma_path}", foreground="green")
    else:
        template_status_label.config(text="Шаблон не выбран", foreground="gray")


def select_template_with_status(template_status_label):
    """Выбор шаблона заявлений с обновлением статуса"""
    if select_template():
        update_template_status(template_status_label)


def select_template_diploma_with_status(template_status_label):
    """Выбор шаблона для грамот с обновлением статуса"""
    if select_template_diploma():
        update_template_diploma_status(template_status_label)


def setup_gui():
    """Создание графического интерфейса"""
    global progress_bar
    global progress_bar_diploma
    root = tk.Tk()
    root.title("Филиал ЕДУ ИМЦ Ленинского района. Сбор сведений ВсОШ Школьный этап v1.1b©")
    root.geometry("650x700")

    main_frame = ttk.Frame(root, padding="10")
    main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

    title_label = ttk.Label(main_frame, text="🏆 Управление сервером олимпиад", font=("Arial", 16, "bold"))
    title_label.grid(row=0, column=0, columnspan=4, pady=(0, 20))

    # Загрузка данных
    file_frame = ttk.LabelFrame(main_frame, text="Загрузка данных", padding="10")
    file_frame.grid(row=1, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=(0, 10))

    file_label = ttk.Label(file_frame, text="Файл не выбран")
    file_label.grid(row=0, column=0, sticky=(tk.W, tk.E))

    ttk.Button(file_frame, text="Выбрать Excel файл",
               command=lambda: load_excel_file_gui(file_label, stats_tree)).grid(row=0, column=1, padx=(10, 0))

    ttk.Button(file_frame, text="Сбросить все данные",
               command=lambda: reset_all_data_gui(stats_tree)).grid(row=0, column=2, padx=(10, 0))

    # Управление сервером
    server_frame = ttk.LabelFrame(main_frame, text="Управление сервером", padding="10")
    server_frame.grid(row=2, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=(0, 10))

    status_label = ttk.Label(server_frame, text="Статус: Сервер остановлен", foreground="red",
                             font=("Arial", 10, "bold"))
    status_label.grid(row=1, column=0, columnspan=4, pady=(0, 0), sticky="w")

    ttk.Button(server_frame, text="▶ Запустить сервер", command=lambda: start_server_gui(status_label), width=30,
               style="Server.TButton").grid(row=0, column=0, padx=(0, 10), sticky="w")

    ttk.Button(server_frame, text="⏹ Остановить сервер", command=lambda: stop_server_gui(status_label), width=30,
               style="Server.TButton").grid(row=0, column=1, padx=(0, 10), sticky="w")

    ttk.Button(server_frame, text="🌐 Открыть в браузере", command=open_browser_gui, width=30,
               style="Server.TButton").grid(row=0, column=2, padx=(0, 10), sticky="w")

    # Заявления
    export_frame = ttk.LabelFrame(main_frame, text="Заявления", padding="10")
    export_frame.grid(row=3, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=(0, 10))

    template_button = ttk.Button(export_frame, text="📄 Шаблон заявления", width=30)
    template_button.grid(row=0, column=0, padx=(0, 10))

    ttk.Button(export_frame, text="📝 Сформировать заявления", command=create_statement,
               width=30).grid(row=0, column=1, padx=(0, 10))

    ttk.Button(export_frame, text="📊 Выгрузить в Excel", command=save_to_xlsx,
               width=30).grid(row=0, column=2, padx=(0, 10))

    template_status_label = ttk.Label(export_frame, text="Шаблон не выбран", foreground="gray")
    template_status_label.grid(row=1, column=0, columnspan=3, pady=(5, 0), sticky="w")

    template_button.config(command=lambda: select_template_with_status(template_status_label))

    progress_bar = ttk.Progressbar(export_frame, mode='determinate', length=100)
    progress_bar.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 0))

    # Грамоты
    diploma_frame = ttk.LabelFrame(main_frame, text="Грамоты ВсОШ ШЭ", padding="10")
    diploma_frame.grid(row=4, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=(0, 10))

    template_diploma_button = ttk.Button(diploma_frame, text="📄 Шаблон грамоты", width=30)
    template_diploma_button.grid(row=0, column=0, padx=(0, 10))

    ttk.Button(diploma_frame, text="📊 Протокол ВсОШ ШЭ", command=save_diploma_to_xlsx,
               width=30).grid(row=0, column=1, padx=(0, 10))

    ttk.Button(diploma_frame, text="📝 Сформировать грамоты", command=create_diploma,
               width=30).grid(row=0, column=2, padx=(0, 10))

    template_status_diploma_label = ttk.Label(diploma_frame, text="Шаблон не выбран", foreground="gray")
    template_status_diploma_label.grid(row=1, column=0, columnspan=3, pady=(5, 0), sticky="w")

    template_diploma_button.config(command=lambda: select_template_diploma_with_status(template_status_diploma_label))

    progress_bar_diploma = ttk.Progressbar(diploma_frame, mode='determinate', length=80)
    progress_bar_diploma.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 0))

    # Статистика
    stats_frame = ttk.LabelFrame(main_frame, text="Статистика участия", padding="10")
    stats_frame.grid(row=5, column=0, columnspan=4, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))

    columns = ("Класс", "Процент участия")
    stats_tree = ttk.Treeview(stats_frame, columns=columns, show="headings", height=6)

    stats_tree.heading("Класс", text="Класс")
    stats_tree.column("Класс", width=200, anchor="center")

    stats_tree.heading("Процент участия", text="Процент участия")
    stats_tree.column("Процент участия", width=200, anchor="center")

    stats_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

    scrollbar = ttk.Scrollbar(stats_frame, orient="vertical", command=stats_tree.yview)
    scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
    stats_tree.configure(yscrollcommand=scrollbar.set)

    ttk.Button(stats_frame, text="Обновить статистику",
               command=lambda: update_statistics_gui(stats_tree)).grid(row=1, column=0, pady=(10, 0))

    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)
    main_frame.columnconfigure(0, weight=1)
    main_frame.rowconfigure(4, weight=1)
    file_frame.columnconfigure(0, weight=1)
    server_frame.columnconfigure(0, weight=1)

    export_frame.columnconfigure(0, weight=0)
    export_frame.columnconfigure(1, weight=0)
    export_frame.columnconfigure(2, weight=0)

    stats_frame.columnconfigure(0, weight=1)
    stats_frame.rowconfigure(0, weight=1)

    server_frame.columnconfigure(0, weight=0)
    server_frame.columnconfigure(1, weight=0)
    server_frame.columnconfigure(2, weight=0)
    server_frame.columnconfigure(3, weight=1)

    root.after(100, show_instruction_window)

    return root


def show_instruction_window():
    """Показывает окно с краткой инструкцией при запуске программы"""
    instruction_window = tk.Toplevel()
    instruction_window.title("Краткая инструкция")
    instruction_window.geometry("700x650")
    instruction_window.resizable(False, False)

    instruction_window.transient()
    instruction_window.grab_set()

    main_frame = ttk.Frame(instruction_window, padding="15")
    main_frame.pack(fill=tk.BOTH, expand=True)

    title_label = ttk.Label(main_frame, text="📋 КРАТКАЯ ИНСТРУКЦИЯ", font=("Arial", 16, "bold"))
    title_label.pack(pady=(0, 15))

    text_frame = ttk.Frame(main_frame)
    text_frame.pack(fill=tk.BOTH, expand=True)

    text_widget = tk.Text(text_frame, wrap=tk.WORD, font=("Courier New", 11), padx=10, pady=10, bg="#f8f9fa")
    scrollbar = ttk.Scrollbar(text_frame, command=text_widget.yview)
    text_widget.configure(yscrollcommand=scrollbar.set)

    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    instruction_text = """🚀 ЗАГРУЗИТЕ ДАННЫЕ
- Нажмите "Выбрать Excel файл" — загрузите список учеников (листы с названиями классов: 4А, 5Б, 8В...). 
  *Используется выгрузка из ГИС СО "ЕЦП"  Электронный журнал. Файл формируется в разделе: Учет учащихся и педагогов → Классы → Сформировать списки: Для всех классов, Поле в отчёте - любое (обычно выбирают Пол)
- Нажмите "Шаблон заявления" — выберите файл шаблона (.docx)

🌐 СОБЕРИТЕ ДАННЫЕ
- Нажмите "Запустить сервер"
- Нажмите "Открыть в браузере" или раздайте ярлык на сервер "Выбор предметов ВсОШ"
- В браузере для каждого ученика отмечайте:
  🏠 Дома — если олимпиаду будет писать дома
  🏫 В школе — если олимпиаду будет писать в школе
  *Данные сохраняются на компьютере, где запущен сервер. Программу можно закрывать и открывать - данные о выборе сохраняются (при загрузке списков класса будет диалог для подтверждения использования уже собранных данных).
- Нажимайте "Сохранить изменения"

📄 ПОЛУЧИТЕ РЕЗУЛЬТАТ
- Нажмите "Сформировать заявления" — формируются готовые заявления в папке "Заявления" по классам.
  *В заявлении предметы будут разделены на "на дому:" и "в школе:"
- Нажмите "Выгрузить в Excel" — получите статистику выбора предметов со списком учащихся, разделением по месту проведения, общая статистика выбора предметов и расчёт процента участия в классе и школе.

🏆 ФОРМИРОВАНИЕ ГРАМОТ ШКОЛЬНОГО ЭТАПА ВсОШ:
- Выберите "Шаблон грамоты"
- Загрузите "Протокол ВсОШ ШЭ" из РБДО:  Школьный этап → Результаты диагностик → Экспорт → (ШЭ)
- Нажмите "Сформировать грамоты" — формируются грамоты в папке "Грамоты" по классам
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Шаблоны заявления и грамот содержат поля подстановки для их автоматического заполнения из полученных данных.

Поле заявления         Описание
%ФИО%                  ФИО ученика
%Класс%                Класс ученика
%олимпиады%            Список предметов с указанием места (на дому: / в школе:)

Поле грамоты           Описание
%ОО%                   Образовательная организация (школа)
%Класс%                Класс ученика
%Предметы%             Предметы с результатами
%Фамилия%              Фамилия ученика
%Имя%                  Имя ученика
%призёр и победитель%  Статус (победитель, призёр или оба сразу)
%окон%                 Окончание слова "предмет"
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"""

    text_widget.insert("1.0", instruction_text)
    text_widget.configure(state="disabled")

    button_frame = ttk.Frame(main_frame)
    button_frame.pack(fill=tk.X, pady=(15, 0))

    close_button = ttk.Button(button_frame, text="Понятно, закрыть", command=instruction_window.destroy, width=20)
    close_button.pack()

    instruction_window.update_idletasks()
    x = (instruction_window.winfo_screenwidth() // 2) - (instruction_window.winfo_width() // 2)
    y = (instruction_window.winfo_screenheight() // 2) - (instruction_window.winfo_height() // 2)
    instruction_window.geometry(f'+{x}+{y}')

    instruction_window.grab_release()


def main():
    """Главная функция"""
    print("🚀 Запуск приложения управления олимпиадами...")

    setup_flask_app()
    root = setup_gui()
    root.mainloop()


if __name__ == '__main__':
    main()