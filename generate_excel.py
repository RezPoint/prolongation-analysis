import warnings
import pandas as pd
import xlsxwriter

warnings.filterwarnings('ignore')

from config       import EXCLUDE_AM
from data         import load_data, get_month_cols, clean_financial, deduplicate_prolongations, build_project_df
from calculations import build_col_index, calc_all_metrics, calc_annual, build_pivots
from excel_report import create_formats, write_sheet_dashboard, write_sheet_coefficients, \
                         write_sheet_annual, write_sheet_counts, write_sheet_chart


def generate_report(output_path='prolongation_report.xlsx'):
    prol, fin  = load_data()
    month_cols = get_month_cols(fin)
    fin_agg    = clean_financial(fin, month_cols)
    prol       = deduplicate_prolongations(prol)
    df         = build_project_df(prol, fin_agg, month_cols)
    col_to_idx = build_col_index(month_cols)

    managers  = sorted([am for am in df['AM'].unique() if am != EXCLUDE_AM])
    row_order = ['Отдел в целом'] + managers

    results_df, counts_df = calc_all_metrics(df, managers, col_to_idx, month_cols)
    annual_df   = calc_annual(results_df)
    all_results = pd.concat([results_df, annual_df], ignore_index=True)
    pivot_k1, pivot_k2, month_order = build_pivots(all_results, row_order)

    wb = xlsxwriter.Workbook(output_path)
    F  = create_formats(wb)

    write_sheet_dashboard(wb, F, annual_df, row_order)
    write_sheet_coefficients(wb, F, pivot_k1, pivot_k2, month_order, row_order)
    write_sheet_annual(wb, F, annual_df, row_order)
    write_sheet_counts(wb, F, counts_df, row_order)
    write_sheet_chart(wb, F, pivot_k1, row_order)

    wb.get_worksheet_by_name('Дашборд').activate()
    wb.close()
    print(f'Отчёт сохранён: {output_path}')
    return results_df, annual_df, counts_df


if __name__ == '__main__':
    generate_report()
