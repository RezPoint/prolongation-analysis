import pandas as pd
from config import NON_MONTH_COLS


def load_data(prol_path='prolongations.csv', fin_path='financial_data.csv'):
    prol = pd.read_csv(prol_path, encoding='utf-8').rename(columns={'month': 'last_month'})
    fin  = pd.read_csv(fin_path,  encoding='utf-8')
    return prol, fin


def get_month_cols(fin):
    return [c for c in fin.columns if c not in NON_MONTH_COLS]


def _clean_value(v):
    """'стоп', 'в ноль', 'end', NaN → 0."""
    if pd.isna(v):
        return 0.0
    s = str(v).strip().lower()
    if s in ('стоп', 'в ноль', 'end', ''):
        return 0.0
    s = s.replace('\xa0', '').replace(' ', '').replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return 0.0


def clean_financial(fin, month_cols):
    for col in month_cols:
        fin[col] = fin[col].apply(_clean_value)
    fin_agg = fin.groupby('id', as_index=False)[month_cols].sum()
    account_map = fin.drop_duplicates('id').set_index('id')['Account']
    fin_agg['Account'] = fin_agg['id'].map(account_map)
    return fin_agg


def deduplicate_prolongations(prol):
    before = len(prol)
    prol = prol.drop_duplicates(subset=['id', 'last_month'], keep='first')
    print(f'Дедупликация prolongations: {before} -> {len(prol)} строк')
    return prol


def build_project_df(prol, fin_agg, month_cols):
    df = prol.merge(fin_agg, on='id', how='left')
    df['last_month_col'] = df['last_month'].apply(lambda m: m.strip().capitalize())
    for col in month_cols:
        df[col] = df[col].fillna(0.0)
    return df
