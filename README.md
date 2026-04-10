# Анализ пролонгаций договоров — 2023 год

Расчёт коэффициентов пролонгации (К1, К2) по аккаунт-менеджерам отдела сопровождения клиентов за 2023 год.

## Запуск в Google Colab

[![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/RezPoint/prolongation-analysis/blob/main/prolongation_analysis.ipynb)

1. Открыть ноутбук по кнопке выше
2. `Среда выполнения → Выполнить всё`
3. Excel-отчёт скачается автоматически

## Запуск локально

```bash
pip install pandas xlsxwriter openpyxl
python generate_excel.py
```

## Структура проекта

```
prolongation_analysis.ipynb   # ноутбук с пояснениями и расчётами
config.py                     # константы (месяцы, цвета, исключения)
data.py                       # загрузка и очистка данных
calculations.py               # расчёт K1, K2, сводные таблицы
excel_report.py               # генерация листов Excel
generate_excel.py             # точка входа
prolongations.csv             # данные о пролонгациях
financial_data.csv            # финансовые данные по проектам
```

## Коэффициенты

**К1** — доля отгрузки проектов, пролонгированных в первый месяц после завершения.

**К2** — доля отгрузки проектов, пролонгированных во второй месяц (среди тех, что не пролонгировались в первый).

Оба коэффициента считаются как взвешенное среднее: сумма отгрузки пролонгированных проектов / сумма отгрузки всех проектов-кандидатов.

## Стек

Python 3.10+, pandas, xlsxwriter
