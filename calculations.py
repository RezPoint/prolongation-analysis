import numpy as np
import pandas as pd
from config import MONTHS_2023, EXCLUDE_AM


def build_col_index(month_cols):
    return {col: i for i, col in enumerate(month_cols)}


def get_col_by_offset(base_col, offset, col_to_idx, month_cols):
    if base_col not in col_to_idx:
        return None
    idx = col_to_idx[base_col] + offset
    return month_cols[idx] if 0 <= idx < len(month_cols) else None


def calc_k1_k2(df_subset, month_col, col_to_idx, month_cols):
    prev_col  = get_col_by_offset(month_col, -1, col_to_idx, month_cols)
    prev2_col = get_col_by_offset(month_col, -2, col_to_idx, month_cols)

    if prev_col:
        base_k1 = df_subset[
            (df_subset['last_month_col'] == prev_col) & (df_subset[prev_col] > 0)
        ]
        k1_den = base_k1[prev_col].sum()
        k1_num = base_k1.loc[base_k1[month_col] > 0, month_col].sum()
        k1 = k1_num / k1_den if k1_den > 0 else np.nan
    else:
        k1_num = k1_den = 0
        k1 = np.nan

    if prev_col and prev2_col:
        base_k2 = df_subset[
            (df_subset['last_month_col'] == prev2_col)
            & (df_subset[prev2_col] > 0)
            & (df_subset[prev_col] == 0)
        ]
        k2_den = base_k2[prev2_col].sum()
        k2_num = base_k2.loc[base_k2[month_col] > 0, month_col].sum()
        k2 = k2_num / k2_den if k2_den > 0 else np.nan
    else:
        k2_num = k2_den = 0
        k2 = np.nan

    return {
        'k1': k1, 'k1_num': k1_num, 'k1_den': k1_den,
        'k2': k2, 'k2_num': k2_num, 'k2_den': k2_den,
    }


def calc_counts(df_subset, month_col, entity_name, col_to_idx, month_cols):
    prev_col  = get_col_by_offset(month_col, -1, col_to_idx, month_cols)
    prev2_col = get_col_by_offset(month_col, -2, col_to_idx, month_cols)
    row = {'Менеджер': entity_name, 'Месяц': month_col}

    if prev_col:
        base = df_subset[(df_subset['last_month_col'] == prev_col) & (df_subset[prev_col] > 0)]
        row['Завершилось (M-1)']           = len(base)
        row['Пролонг. К1']                 = (base[month_col] > 0).sum()
        row['Не пролонг. К1']              = len(base) - row['Пролонг. К1']
        row['Выручка завершившихся (M-1)'] = base[prev_col].sum()
        row['Выручка пролонг. К1 (M)']    = base.loc[base[month_col] > 0, month_col].sum()
    else:
        row.update({
            'Завершилось (M-1)': 0, 'Пролонг. К1': 0, 'Не пролонг. К1': 0,
            'Выручка завершившихся (M-1)': 0, 'Выручка пролонг. К1 (M)': 0,
        })

    if prev_col and prev2_col:
        base2 = df_subset[
            (df_subset['last_month_col'] == prev2_col)
            & (df_subset[prev2_col] > 0)
            & (df_subset[prev_col] == 0)
        ]
        row['Не пролонг. в M-1 (из M-2)']      = len(base2)
        row['Пролонг. К2']                       = (base2[month_col] > 0).sum()
        row['Выручка не пролонг. в M-1 (M-2)']  = base2[prev2_col].sum()
        row['Выручка пролонг. К2 (M)']           = base2.loc[base2[month_col] > 0, month_col].sum()
    else:
        row.update({
            'Не пролонг. в M-1 (из M-2)': 0, 'Пролонг. К2': 0,
            'Выручка не пролонг. в M-1 (M-2)': 0, 'Выручка пролонг. К2 (M)': 0,
        })

    return row


def calc_all_metrics(df, managers, col_to_idx, month_cols):
    coef_rows  = []
    count_rows = []
    df_dept    = df[df['AM'] != EXCLUDE_AM]

    entities = [('Отдел в целом', df_dept)] + [(am, df[df['AM'] == am]) for am in sorted(managers)]

    for month_col in MONTHS_2023:
        for name, subset in entities:
            r = calc_k1_k2(subset, month_col, col_to_idx, month_cols)
            coef_rows.append({'Менеджер': name, 'Месяц': month_col, **r})
            count_rows.append(calc_counts(subset, month_col, name, col_to_idx, month_cols))

    return pd.DataFrame(coef_rows), pd.DataFrame(count_rows)


def calc_annual(results_df):
    annual_rows = []
    for entity in results_df['Менеджер'].unique():
        sub = results_df[results_df['Менеджер'] == entity]
        k1n, k1d = sub['k1_num'].sum(), sub['k1_den'].sum()
        k2n, k2d = sub['k2_num'].sum(), sub['k2_den'].sum()
        annual_rows.append({
            'Менеджер': entity,
            'Месяц':    '2023 год',
            'k1':       k1n / k1d if k1d > 0 else np.nan,
            'k1_num':   k1n, 'k1_den': k1d,
            'k2':       k2n / k2d if k2d > 0 else np.nan,
            'k2_num':   k2n, 'k2_den': k2d,
        })
    return pd.DataFrame(annual_rows)


def build_pivots(all_results, row_order):
    month_order = MONTHS_2023 + ['2023 год']
    pivot_k1 = (all_results.pivot(index='Менеджер', columns='Месяц', values='k1')
                            .reindex(columns=month_order).reindex(row_order))
    pivot_k2 = (all_results.pivot(index='Менеджер', columns='Месяц', values='k2')
                            .reindex(columns=month_order).reindex(row_order))
    return pivot_k1, pivot_k2, month_order
