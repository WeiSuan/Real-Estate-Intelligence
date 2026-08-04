import os
from typing import Optional, Dict

import pandas as pd
import numpy as np
import pyodbc
import urllib.parse
from sqlalchemy import create_engine

from datetime import datetime, timedelta

def prev_month_end(ref_date: datetime | None = None) -> datetime.date:
    today = ref_date or datetime.today()
    first_of_month = today.replace(day=1)
    last_prev = first_of_month - timedelta(days=1)
    return last_prev.date()

def to_int_then_str(series: pd.Series) -> pd.Series:
    """Convert values to integer (nullable) then to string; missing -> empty string."""
    num = pd.to_numeric(series, errors="coerce")
    int_series = num.where(num.notnull(), pd.NA).astype("Int64")
    obj = int_series.astype(object).where(int_series.notnull(), "")
    return obj.astype(str)


def _read_sql_file(path: str, encoding: str = 'big5') -> str:
    with open(path, encoding=encoding) as f:
        return f.read()


def _open_connection(conn_params: Optional[Dict[str, str]] = None):
    params = conn_params or {}
    drv = params.get('DRIVER', "ODBC Driver 17 for SQL Server")
    server = params.get('SERVER', 'Shindbbackup01')
    #database = params.get('DATABASE', 'shin_daily')
    database = params.get('DATABASE', 'shin_monthly')
    uid = params.get('UID')
    pwd = params.get('PWD')

    conn_str_parts = [f"DRIVER={{{drv}}}", f"SERVER={server}", f"DATABASE={database}"]
    if uid is not None and pwd is not None:
        conn_str_parts.append(f"UID={uid}")
        conn_str_parts.append(f"PWD={pwd}")
    conn_str = ';'.join(conn_str_parts)

    con = pyodbc.connect(conn_str)
    try:
        con.setencoding(encoding='big5')
    except Exception:
        pass
    return con


def _open_sqlalchemy_engine(conn_params: Optional[Dict[str, str]] = None):
    """Return a SQLAlchemy engine using the provided connection params.

    Params accepted: DRIVER, SERVER, DATABASE, UID, PWD
    """
    params = conn_params or {}
    drv = params.get('DRIVER', "ODBC Driver 17 for SQL Server")
    server = params.get('SERVER', 'Shindbbackup01')
    database = params.get('DATABASE', 'shin_monthly')
    uid = params.get('UID')
    pwd = params.get('PWD')

    odbc_parts = [f"DRIVER={{{drv}}}", f"SERVER={server}", f"DATABASE={database}"]
    if uid is not None and pwd is not None:
        odbc_parts.append(f"UID={uid}")
        odbc_parts.append(f"PWD={pwd}")
    odbc_str = ';'.join(odbc_parts)
    quoted = urllib.parse.quote_plus(odbc_str)
    engine = create_engine(f"mssql+pyodbc:///?odbc_connect={quoted}")
    return engine


def get_mssql_df(sql_statement_path: Optional[str] = None,
                        conn_params: Optional[Dict[str, str]] = None,
                        sql_encoding: str = 'big5') -> pd.DataFrame:
    """Load Data from SQL.

    Args:
        sql_statement_path: path to the .sql file. This argument is required;
        conn_params: optional dict of connection params (DRIVER, SERVER, DATABASE, UID, PWD).
        sql_encoding: file encoding for the SQL file (default 'big5').

    Returns:
        pandas.DataFrame with the query result.
    """        
    
    if not sql_statement_path:
        raise ValueError("sql_statement_path is required. Provide path to the .sql file.")

    sql_query = _read_sql_file(sql_statement_path, encoding=sql_encoding)

    # HINT: original pyodbc usage (kept as comment for reference)
    # con = _open_connection(conn_params)
    # try:
    #     df = pd.read_sql_query(sql_query, con)
    # finally:
    #     try:
    #         con.close()
    #     except Exception:
    #         pass

    # Use SQLAlchemy engine (recommended by pandas) to execute the query
    engine = _open_sqlalchemy_engine(conn_params)
    try:
        with engine.connect() as conn:
            df = pd.read_sql_query(sql_query, conn)
    finally:
        try:
            engine.dispose()
        except Exception:
            pass

    return df

def generate_city_gage_for_city(
    summary_table: pd.DataFrame,
    key_mapping: pd.DataFrame,
    loan_remain: pd.DataFrame,
    loan_immovable_gage_detail: pd.DataFrame,
    city_name: str,
) -> pd.DataFrame:
    """Generate city gage allocation for a given city.

    Returns a DataFrame with columns:
    ['額度關係人','申購人','初審編號(歸案)','業別第一碼(擔保品)分類',
     'city_name','district_result','value_worth','value_capital','value_remain']

    Behaviour (vectorized and robust):
    - Selects summary rows with positive `{city}_remain`.
    - Joins to loan_remain (via `key_mapping`) to get `loan_no` and then to
      `loan_immovable_gage_detail` filtered by `city_result == city_name`.
    - For each group (初審編號(歸案), 業別第一碼(擔保品)分類):
      * If the group has more than one non-empty `district_result`, drop
        rows in that group whose `district_result` is empty.
      * If the group has zero non-empty `district_result`, set
        `district_result` to '分區不明'.
    - Compute unique district counts per group and apportion the city
      columns `{city}_worth`, `{city}_capital`, `{city}_remain` by dividing
      by that unique count (leave as-is when count == 0).

    The function makes minimal assumptions about input column existence and
    will create missing columns with NaN/0 as needed.
    """

    # column names used in the pipeline
    pre_col = '初審編號(歸案)'
    code_col = '業別第一碼(擔保品)分類'
    rel_col = '額度關係人'
    cust_col = '申購人'

    city_worth = f"{city_name}_worth"
    city_capital = f"{city_name}_capital"
    city_remain = f"{city_name}_remain"

    # quick guard: if no city_remain column, return empty df
    if city_remain not in summary_table.columns:
        return pd.DataFrame(columns=[rel_col, cust_col, pre_col, code_col, 'city_name', 'district_result', 'value_worth', 'value_capital', 'value_remain'])

    # 1) seed table: summary rows with city remain > 0
    city_init = summary_table.loc[summary_table[city_remain].fillna(0) > 0].copy()

    # 2) attach original pre_examine_no (via key_mapping) and loan_no (via loan_remain)
    key_merge = key_mapping[[ 'pre_examine_no', 'pre_examine_no_return' ]].copy()
    city_init = city_init.merge(key_merge, left_on=pre_col, right_on='pre_examine_no_return', how='left')

    loan_map = loan_remain[[ 'pre_examine_no_major', 'loan_no', 'loan_code_type' ]].copy()
    city_init = city_init.merge(loan_map, left_on='pre_examine_no', right_on='pre_examine_no_major', how='left')

    # 3) filter detail rows for the city and relevant loan_no
    detail = loan_immovable_gage_detail.copy()
    if 'city_result' in detail.columns:
        detail = detail.loc[detail['city_result'] == city_name].copy()
    # ensure loan_no types align
    detail['loan_no'] = detail['loan_no'].astype(str)
    city_init['loan_no'] = city_init['loan_no'].astype(str)
    detail = detail.loc[detail['loan_no'].isin(city_init['loan_no'].dropna().unique())].copy()

    # 4) merge to get per-loan district rows
    city_table = city_init.merge(detail, on='loan_no', how='left')

    # 5) ensure required columns exist
    for c in [rel_col, cust_col, pre_col, code_col, city_worth, city_capital, city_remain, 'loan_no', 'city_result', 'district_result']:
        if c not in city_table.columns:
            city_table[c] = np.nan

    pre1 = city_table[[rel_col, cust_col, pre_col, code_col, city_worth, city_capital, city_remain, 'loan_no', 'city_result', 'district_result']].drop_duplicates().reset_index(drop=True)

    # 6) build pre2: drop empty district rows when group has >1 non-empty district
    cg = pre1.copy()
    cg['district_nonempty'] = cg['district_result'].notna() & cg['district_result'].astype(str).str.strip().ne('')
    grp = cg.groupby([pre_col, code_col], as_index=False)['district_nonempty'].sum().rename(columns={'district_nonempty': 'district_nonempty_count'})
    cg = cg.merge(grp, on=[pre_col, code_col], how='left')
    mask_remove = (cg['district_nonempty_count'] > 1) & (~cg['district_nonempty'])
    cg2 = cg.loc[~mask_remove].copy()
    # if whole group empty, mark unknown
    cg2.loc[cg2['district_nonempty_count'] == 0, 'district_result'] = '分區不明'
    city_gage_pre2 = cg2.drop(columns=['district_nonempty', 'district_nonempty_count']).reset_index(drop=True)

    # 7) compute unique district counts and apportion values
    district_counts = city_gage_pre2.groupby([pre_col, code_col], as_index=False)['district_result'].nunique().rename(columns={'district_result': 'district_result_unique_count'})
    pre3 = city_gage_pre2.merge(district_counts, on=[pre_col, code_col], how='left')
    pre3['district_result_unique_count'] = pre3['district_result_unique_count'].fillna(0).astype(int)

    def _apportion(series_name):
        s = pre3[series_name].replace({np.nan: 0})
        cnt = pre3['district_result_unique_count']
        # when cnt>0 divide, else keep original (as float)
        return np.where(cnt > 0, s.astype(float) / cnt, s.astype(float))

    pre3['value_worth'] = _apportion(city_worth)
    pre3['value_capital'] = _apportion(city_capital)
    pre3['value_remain'] = _apportion(city_remain)
    pre3['city_name'] = city_name

    out_cols = [rel_col, cust_col, pre_col, code_col, 'city_name', 'district_result', 'value_worth', 'value_capital', 'value_remain']
    result_pre3 = pre3[out_cols].drop_duplicates().reset_index(drop=True)

    # 8) handle NULL (summary entries with 初審編號(歸案)=='NULL') – keep previous behaviour
    pre4 = summary_table.loc[summary_table[pre_col].isin(["NULL"])].copy()
    if not pre4.empty:
        pre4['district_result'] = '分區不明'
        pre4['city_name'] = city_name
        for col in [city_worth, city_capital, city_remain]:
            if col not in pre4.columns:
                pre4[col] = 0
        pre4 = pre4[[pre_col, code_col, 'city_name', 'district_result', city_worth, city_capital, city_remain]].drop_duplicates().reset_index(drop=True)
        pre4 = pre4.rename(columns={city_worth: 'value_worth', city_capital: 'value_capital', city_remain: 'value_remain'})
        # ensure column names align for concat
        pre4 = pre4.rename(columns={pre_col: pre_col, code_col: code_col})
        # create placeholder columns for relation and customer then reorder
        pre4[rel_col] = np.nan
        pre4[cust_col] = np.nan
        pre4 = pre4[[rel_col, cust_col, pre_col, code_col, 'city_name', 'district_result', 'value_worth', 'value_capital', 'value_remain']]
        final = pd.concat([result_pre3, pre4], ignore_index=True, sort=False)
    else:
        final = result_pre3.copy()

    final['value_worth'] = pd.to_numeric(final['value_worth'], errors='coerce')
    final['value_capital'] = pd.to_numeric(final['value_capital'], errors='coerce')
    final['value_remain'] = pd.to_numeric(final['value_remain'], errors='coerce')

    final = final.loc[final.district_result.notna()]
    return final.drop_duplicates().reset_index(drop=True)
