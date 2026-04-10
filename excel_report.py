import numpy as np
import pandas as pd
from config import (
    MONTHS_2023,
    DARK_BLUE, MED_BLUE, LIGHT_BLUE, ACCENT_BLUE, DEPT_BG, WHITE,
)

COLOR_SCALE = {
    'type': '3_color_scale',
    'min_color': '#FF6B6B',
    'mid_color': '#FFE066',
    'max_color': '#70B85F',
}


def create_formats(wb):
    def fmt(**kw):
        defaults = dict(border=1, valign='vcenter', align='center',
                        font_size=10, text_wrap=True)
        defaults.update(kw)
        return wb.add_format(defaults)

    return {
        'title':      fmt(bold=True, bg_color=DARK_BLUE,   font_color=WHITE, font_size=14, border=0),
        'header':     fmt(bold=True, bg_color=MED_BLUE,    font_color=WHITE),
        'ann_header': fmt(bold=True, bg_color=ACCENT_BLUE, font_color=WHITE),
        'subheader':  fmt(bold=True, bg_color=LIGHT_BLUE),
        'dept':       fmt(bold=True, bg_color=DEPT_BG),
        'dept_pct':   fmt(bold=True, bg_color=DEPT_BG,  num_format='0.0%'),
        'dept_num':   fmt(bold=True, bg_color=DEPT_BG,  num_format='# ##0'),
        'dept_rub':   fmt(bold=True, bg_color=DEPT_BG,  num_format='# ##0 ₽', align='right'),
        'normal':     fmt(),
        'pct':        fmt(num_format='0.0%'),
        'num':        fmt(num_format='# ##0'),
        'rub':        fmt(num_format='# ##0 ₽', align='right'),
        'na':         fmt(font_color='#AAAAAA', italic=True),
        'note':       fmt(font_size=9, align='left', border=0, italic=True, font_color='#555555'),
        'info':       fmt(bg_color='#EBF3FB', align='left', font_size=9, border=0),
        'best':       fmt(bold=True, bg_color='#C6EFCE', font_color='#375623'),
        'best_pct':   fmt(bold=True, bg_color='#C6EFCE', font_color='#375623', num_format='0.0%'),
        'worst':      fmt(bold=True, bg_color='#FFC7CE', font_color='#9C0006'),
        'worst_pct':  fmt(bold=True, bg_color='#FFC7CE', font_color='#9C0006', num_format='0.0%'),
        'kpi':        fmt(bold=True, font_size=22, border=0, bg_color=LIGHT_BLUE),
        'kpi_label':  fmt(font_size=9, border=0, bg_color=LIGHT_BLUE, font_color='#555555'),
        'dash_sub':   fmt(bold=True, bg_color=ACCENT_BLUE, font_color=WHITE, font_size=11, border=0),
        'dash_title': fmt(bold=True, bg_color=DARK_BLUE,   font_color=WHITE, font_size=16, border=0),
        'dash_info':  fmt(bg_color='#D6E4F0', font_color=DARK_BLUE, font_size=11, bold=True,
                          border=0, align='left'),
    }


def _write_pct_or_dash(ws, row, col, val, fmt_val, fmt_dash):
    if pd.isna(val):
        ws.write(row, col, '—', fmt_dash)
    else:
        ws.write(row, col, val, fmt_val)


def _short_months(month_list):
    return [m.replace(' 2023', '').replace(' 2024', '') for m in month_list]


def write_sheet_coefficients(wb, F, pivot_k1, pivot_k2, month_order, row_order):
    ws = wb.add_worksheet('Коэффициенты пролонгации')
    ws.set_zoom(90)
    ws.freeze_panes(4, 2)
    ws.set_column(0, 0, 30)
    ws.set_column(1, len(month_order), 9)

    short = _short_months(month_order)

    ws.merge_range(0, 0, 0, len(month_order),
                   'Коэффициенты пролонгации договоров — 2023 год', F['title'])
    ws.set_row(0, 28)
    ws.merge_range(1, 0, 1, len(month_order),
                   'К1 — пролонгация в 1-й месяц; К2 — пролонгация во 2-й месяц '
                   '(среди не пролонгировавшихся в 1-й). Цвет: красный → жёлтый → зелёный.',
                   F['info'])
    ws.set_row(1, 22)

    ws.write(3, 0, 'Менеджер', F['header'])
    for ci, label in enumerate(short):
        f = F['ann_header'] if month_order[ci] == '2023 год' else F['header']
        ws.write(3, ci + 1, label, f)

    def write_block(title, pivot, data_row_start):
        ws.merge_range(data_row_start - 1, 0, data_row_start - 1, len(month_order),
                       title, F['subheader'])
        for ri, am in enumerate(row_order):
            r = data_row_start + ri
            is_dept = (am == 'Отдел в целом')
            ws.write(r, 0, am, F['dept'] if is_dept else F['normal'])
            for ci, mc in enumerate(month_order):
                val = pivot.loc[am, mc] if am in pivot.index else np.nan
                _write_pct_or_dash(ws, r, ci + 1, val,
                                   F['dept_pct'] if is_dept else F['pct'],
                                   F['dept']     if is_dept else F['na'])
        return data_row_start + len(row_order)

    k1_end = write_block(
        'К1 — Пролонгация в первый месяц после завершения', pivot_k1, 5
    )
    write_block(
        'К2 — Пролонгация во второй месяц (среди не пролонгировавшихся в первый)',
        pivot_k2, k1_end + 2
    )

    total_rows = k1_end + 2 + len(row_order)
    ws.conditional_format(5, 1, total_rows, len(month_order), COLOR_SCALE)


def write_sheet_annual(wb, F, annual_df, row_order):
    ws = wb.add_worksheet('Итоги за год')
    ws.set_zoom(100)
    ws.set_column(0, 0, 30)
    ws.set_column(1, 6, 16)

    ws.merge_range(0, 0, 0, 6,
                   'Итоги пролонгации за 2023 год — по менеджерам', F['title'])
    ws.set_row(0, 28)
    ws.merge_range(1, 0, 1, 6,
                   'Годовые коэффициенты: сумма числителей / сумма знаменателей '
                   'по всем месяцам (взвешенное среднее).',
                   F['info'])
    ws.set_row(1, 20)

    for ci, h in enumerate(['Менеджер', 'К1 (год)', 'К2 (год)',
                             'База К1, ₽', 'Пролонг. К1, ₽', 'База К2, ₽', 'Пролонг. К2, ₽']):
        ws.write(3, ci, h, F['header'])
    ws.set_row(3, 30)

    for ri, am in enumerate(row_order):
        sub = annual_df[annual_df['Менеджер'] == am]
        if sub.empty:
            continue
        r_data = sub.iloc[0]
        is_dept = (am == 'Отдел в целом')
        rf, pf, nf = (F['dept'], F['dept_pct'], F['dept_rub']) if is_dept \
                else (F['normal'], F['pct'], F['rub'])

        ws.write(4 + ri, 0, am, rf)
        _write_pct_or_dash(ws, 4 + ri, 1, r_data['k1'], pf, rf)
        _write_pct_or_dash(ws, 4 + ri, 2, r_data['k2'], pf, rf)
        ws.write(4 + ri, 3, r_data['k1_den'], nf)
        ws.write(4 + ri, 4, r_data['k1_num'], nf)
        ws.write(4 + ri, 5, r_data['k2_den'], nf)
        ws.write(4 + ri, 6, r_data['k2_num'], nf)

    ws.conditional_format(4, 1, 4 + len(row_order) - 1, 2, COLOR_SCALE)


def write_sheet_counts(wb, F, counts_df, row_order):
    ws = wb.add_worksheet('Детализация по проектам')
    ws.set_zoom(85)
    ws.freeze_panes(4, 2)
    ws.set_column(0, 0, 30)
    ws.set_column(1, 12, 10)

    ws.merge_range(0, 0, 0, 12,
                   'Количество проектов и суммы отгрузки — 2023 год', F['title'])
    ws.set_row(0, 28)

    short_months = _short_months(MONTHS_2023)

    metrics = [
        ('Завершилось проектов (база К1)',     'Завершилось (M-1)',               F['num'], F['dept_num']),
        ('Пролонгировано в 1-й месяц (К1)',    'Пролонг. К1',                     F['num'], F['dept_num']),
        ('Не пролонгировано в 1-й месяц',      'Не пролонг. К1',                  F['num'], F['dept_num']),
        ('База К2 (не пролонг. в M-1)',        'Не пролонг. в M-1 (из M-2)',      F['num'], F['dept_num']),
        ('Пролонгировано во 2-й месяц (К2)',   'Пролонг. К2',                     F['num'], F['dept_num']),
        ('Выручка завершившихся, ₽',           'Выручка завершившихся (M-1)',      F['rub'], F['dept_rub']),
        ('Выручка пролонг. К1, ₽',            'Выручка пролонг. К1 (M)',          F['rub'], F['dept_rub']),
        ('Выручка базы К2, ₽',                'Выручка не пролонг. в M-1 (M-2)', F['rub'], F['dept_rub']),
        ('Выручка пролонг. К2, ₽',            'Выручка пролонг. К2 (M)',          F['rub'], F['dept_rub']),
    ]

    cur_row = 2
    for metric_name, field, nfmt, dfmt in metrics:
        ws.merge_range(cur_row, 0, cur_row, 12, metric_name, F['subheader'])
        cur_row += 1
        ws.write(cur_row, 0, 'Менеджер', F['header'])
        for ci, sm in enumerate(short_months):
            ws.write(cur_row, ci + 1, sm, F['header'])
        cur_row += 1
        for am in row_order:
            is_dept = (am == 'Отдел в целом')
            ws.write(cur_row, 0, am, dfmt if is_dept else F['normal'])
            for ci, mc in enumerate(MONTHS_2023):
                sub = counts_df[(counts_df['Менеджер'] == am) & (counts_df['Месяц'] == mc)]
                val = sub[field].values[0] if not sub.empty else 0
                ws.write(cur_row, ci + 1, val, dfmt if is_dept else nfmt)
            cur_row += 1
        cur_row += 1


def write_sheet_chart(wb, F, pivot_k1, row_order):
    ws = wb.add_worksheet('График К1')
    ws.set_zoom(90)
    ws.set_column(0, 0, 20)
    ws.set_column(1, 12, 9)

    ws.merge_range(0, 0, 0, 12, 'Динамика К1 по месяцам 2023 года', F['title'])
    ws.set_row(0, 28)

    short = _short_months(MONTHS_2023)
    ws.write(2, 0, 'Менеджер / Месяц', F['header'])
    for ci, sm in enumerate(short):
        ws.write(2, ci + 1, sm, F['header'])

    data_start = 3
    for ri, am in enumerate(row_order):
        is_dept = (am == 'Отдел в целом')
        ws.write(data_start + ri, 0, am, F['dept'] if is_dept else F['normal'])
        for ci, mc in enumerate(MONTHS_2023):
            val = pivot_k1.loc[am, mc] if am in pivot_k1.index else np.nan
            if pd.isna(val):
                ws.write_blank(data_start + ri, ci + 1, None, F['na'])
            else:
                ws.write(data_start + ri, ci + 1, val,
                         F['dept_pct'] if is_dept else F['pct'])

    chart = wb.add_chart({'type': 'line'})
    chart.set_title({'name': 'Динамика К1 по месяцам 2023 года'})
    chart.set_x_axis({'name': 'Месяц'})
    chart.set_y_axis({'name': 'К1', 'num_format': '0%'})
    chart.set_size({'width': 820, 'height': 380})
    chart.set_style(10)

    colors = ['#1F3864', '#4472C4', '#ED7D31', '#A9D18E', '#FF6B6B',
              '#FFE066', '#B4C7E7', '#70AD47', '#9DC3E6']
    for ri, am in enumerate(row_order):
        chart.add_series({
            'name':       ['График К1', data_start + ri, 0],
            'categories': ['График К1', 2, 1, 2, 12],
            'values':     ['График К1', data_start + ri, 1, data_start + ri, 12],
            'line': {
                'color':     colors[ri % len(colors)],
                'width':     3.0 if am == 'Отдел в целом' else 1.5,
                'dash_type': 'solid' if am == 'Отдел в целом' else 'dash',
            },
        })
    chart.set_legend({'position': 'bottom'})
    ws.insert_chart(f'A{data_start + len(row_order) + 3}', chart)


def write_sheet_dashboard(wb, F, annual_df, row_order):
    ws = wb.add_worksheet('Дашборд')
    ws.set_zoom(100)
    ws.set_column(0, 0, 32)
    ws.set_column(1, 6, 14)
    ws.set_tab_color(ACCENT_BLUE)

    ws.merge_range(0, 0, 0, 6, 'Отчёт о пролонгациях — 2023 год', F['dash_title'])
    ws.set_row(0, 36)
    ws.merge_range(1, 0, 1, 6,
                   'Аналитический дашборд для руководителя отдела сопровождения клиентов',
                   F['dash_info'])
    ws.set_row(1, 22)

    dept = annual_df[annual_df['Менеджер'] == 'Отдел в целом'].iloc[0]
    ws.merge_range(3, 0, 3, 2, 'ПОКАЗАТЕЛИ ОТДЕЛА ЗА 2023 ГОД', F['dash_sub'])
    ws.set_row(3, 22)
    ws.merge_range(4, 0, 5, 0, f'{dept["k1"]:.1%}', F['kpi'])
    ws.merge_range(4, 1, 5, 1, f'{dept["k2"]:.1%}', F['kpi'])
    ws.set_row(4, 40)
    ws.write(6, 0, 'К1 — пролонгация в 1-й месяц', F['kpi_label'])
    ws.write(6, 1, 'К2 — пролонгация во 2-й месяц', F['kpi_label'])
    ws.set_row(6, 22)

    am_annual = annual_df[(annual_df['Менеджер'] != 'Отдел в целом') & annual_df['k1'].notna()]
    best_am  = am_annual.loc[am_annual['k1'].idxmax(), 'Менеджер'] if len(am_annual) else None
    worst_am = am_annual.loc[am_annual['k1'].idxmin(), 'Менеджер'] if len(am_annual) else None

    ws.merge_range(8, 0, 8, 6, 'ИТОГИ ПО МЕНЕДЖЕРАМ ЗА 2023 ГОД', F['dash_sub'])
    ws.set_row(8, 22)
    for ci, h in enumerate(['Менеджер', 'К1 год', 'К2 год',
                             'База К1, ₽', 'Пролонг. К1, ₽', 'База К2, ₽', 'Пролонг. К2, ₽']):
        ws.write(9, ci, h, F['header'])
    ws.set_row(9, 22)

    for ri, am in enumerate(row_order):
        sub = annual_df[annual_df['Менеджер'] == am]
        if sub.empty:
            continue
        r = sub.iloc[0]
        is_dept = (am == 'Отдел в целом')
        if is_dept:
            rf, pf, nf = F['dept'], F['dept_pct'], F['dept_rub']
        elif am == best_am:
            rf, pf, nf = F['best'], F['best_pct'], F['best']
        elif am == worst_am:
            rf, pf, nf = F['worst'], F['worst_pct'], F['worst']
        else:
            rf, pf, nf = F['normal'], F['pct'], F['rub']

        ws.write(10 + ri, 0, am, rf)
        _write_pct_or_dash(ws, 10 + ri, 1, r['k1'], pf, rf)
        _write_pct_or_dash(ws, 10 + ri, 2, r['k2'], pf, rf)
        ws.write(10 + ri, 3, r['k1_den'], nf)
        ws.write(10 + ri, 4, r['k1_num'], nf)
        ws.write(10 + ri, 5, r['k2_den'], nf)
        ws.write(10 + ri, 6, r['k2_num'], nf)

    lr = 10 + len(row_order) + 2
    ws.merge_range(lr,     0, lr,     2, 'Лучший результат по К1 среди менеджеров', F['best'])
    ws.merge_range(lr + 1, 0, lr + 1, 2, 'Худший результат по К1 среди менеджеров', F['worst'])
    ws.merge_range(lr + 3, 0, lr + 3, 6,
                   'Примечания: «—» означает отсутствие проектов-кандидатов для пролонгации. '
                   'К1 > 100% возможен, если объём проекта вырос после продления. '
                   'АМ берётся из prolongations.csv (первичные данные). '
                   '"без А/М" исключены из всех расчётов.',
                   F['note'])
