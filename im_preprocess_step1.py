import pandas as pd
import numpy as np
from datetime import timedelta

table1_file_path = './Result/截至20260531/table1_all_approval_cases.xlsx'
table2_file_path = './Result/截至20260531/table2_primary_key_mapping.xlsx'
table3_file_path = './Result/截至20260531/table3_loan_remain_cases.xlsx'
table4_file_path = './Result/截至20260531/table4_loan_remain_immovable_gage.xlsx'

def to_numeric_or_NA(series):
    s_num = pd.to_numeric(series, errors='coerce')
    return s_num.where(~s_num.isnull(), 'NA')


def distribute_to_cities(df, value_col, city_flag_cols, city_count_col='city_count', city_none_col='city_none'):
    """Distribute a value column into city columns based on city flag columns.

    Behaviour:
    - If `city_none_col` == 1, set that column to the full `value_col`.
    - If `city_count_col` == 1, set each city column = flag * value_col.
    - If `city_count_col` > 1, set each city column = flag * (value_col / city_count_col).

    This function is vectorized and returns a new DataFrame.
    """
    df = df.copy()

    # ensure numeric dtypes to avoid incompatible-assignment FutureWarning
    df[city_flag_cols] = df[city_flag_cols].astype(float)
    df[city_none_col] = df[city_none_col].astype(float)
    df[value_col] = pd.to_numeric(df[value_col], errors='coerce').astype(float)

    # when no city flags (city_none==1), propagate full value into city_none
    mask_none = df[city_none_col] == 1
    if mask_none.any():
        df.loc[mask_none, city_none_col] = df.loc[mask_none, value_col]

    # when exactly one city is flagged, multiply flags by full value
    mask_count1 = df[city_count_col] == 1
    if mask_count1.any():
        df.loc[mask_count1, city_flag_cols] = df.loc[mask_count1, city_flag_cols].multiply(
            df.loc[mask_count1, value_col], axis=0
        )

    # when multiple cities flagged, divide value by city_count then multiply
    mask_count_gt1 = df[city_count_col] > 1
    if mask_count_gt1.any():
        share = df.loc[mask_count_gt1, value_col].div(df.loc[mask_count_gt1, city_count_col])
        df.loc[mask_count_gt1, city_flag_cols] = df.loc[mask_count_gt1, city_flag_cols].multiply(
            share, axis=0
        )

    # ensure other rows have zeros in city_flag_cols
    other_mask = ~(mask_none | mask_count1 | mask_count_gt1)
    if other_mask.any():
        df.loc[other_mask, city_flag_cols] = 0

    return df

# 讀取 table1_all_approval_cases.xlsx：已核准案件資料
all_approval = pd.read_excel(table1_file_path, sheet_name = "RawData")

for col in ['pre_examine_no', 'customer_id_no', 'qtgroup_id']:
    all_approval[col] = pd.to_numeric(all_approval[col], errors='coerce').astype('Int64').astype(str)

all_approval['approval_date'] = pd.to_datetime(all_approval['approval_date'], errors='coerce')
all_approval['approved_amount'] = to_numeric_or_NA(all_approval['approved_amount'])


# 讀取 table2_primary_key_mapping.xlsx 的 sheetName = "代碼對應表"
key_mapping = pd.read_excel(table2_file_path, sheet_name='代碼對應表')
for col in ['pre_examine_no', 'pre_examine_no_return', 'pre_examine_no_return_previous']:
    key_mapping[col] = pd.to_numeric(key_mapping[col], errors='coerce').astype('Int64').astype(str)

for col in ['approved_amount', 'total_approved_amount', 'total_approved_amount_previous']:
    key_mapping[col] = to_numeric_or_NA(key_mapping[col])

merged = pd.merge(all_approval, key_mapping, on='pre_examine_no', how='left', suffixes=('', '_key'))

# 讀取 table3_loan_remain_cases.xlsx 的 sheetName = "RawData"
loan_remain = pd.read_excel(table3_file_path, sheet_name='RawData')
for col in ['pre_examine_no', 'pre_examine_no_major', 'customer_id_no', 'loan_no']:
    loan_remain[col] = pd.to_numeric(loan_remain[col], errors='coerce').astype('Int64').astype(str)

for col in ['loan_capital', 'loan_remain_capital'] :
    loan_remain[col] = to_numeric_or_NA(loan_remain[col])

loan_remain['loan_capital'] = pd.to_numeric(loan_remain['loan_capital'], errors='coerce').fillna(0).astype(int)
loan_remain['loan_remain_capital'] = pd.to_numeric(loan_remain['loan_remain_capital'], errors='coerce').fillna(0).astype(int)

loan_remain_merged = loan_remain.drop(columns=[col for col in ['pre_examine_no', 'main_no'] if col in loan_remain.columns], errors='ignore').rename(columns={'pre_examine_no_major': 'pre_examine_no'})
key_mapping_merged = key_mapping[['pre_examine_no', 'pre_examine_no_return']].copy()
loan_remain_merged = pd.merge(loan_remain_merged, key_mapping_merged, on='pre_examine_no', how='left')
# 將 pre_examine_no_return NA 或空白補為 'NULL'
loan_remain_merged['pre_examine_no_return'] = loan_remain_merged['pre_examine_no_return'].fillna('NULL')
loan_remain_merged['pre_examine_no_return'] = loan_remain_merged['pre_examine_no_return'].replace('', 'NULL')

# (初審)擔保值
pre_worth = pd.read_excel(table2_file_path, sheet_name='初審單案擔保率')
pre_worth["pre_examine_no"] = pd.to_numeric(pre_worth["pre_examine_no"], errors='coerce').astype('Int64').astype(str)
pre_worth = pre_worth[["pre_examine_no", "pre_immovable_worth"]].drop_duplicates()

# (合約)擔保值
loan_worth = pd.read_excel(table2_file_path, sheet_name='合約單案擔保率(已撥)')
loan_worth["loan_no"] = pd.to_numeric(loan_worth["loan_no"], errors='coerce').astype('Int64').astype(str)
loan_worth = loan_worth[["loan_no", "immovable_worth"]]
loan_remain_merged = pd.merge(loan_remain_merged, loan_worth, on = 'loan_no', how = 'left')

# loan_remain_merged拆分三段處理：
# (a) pre_examine_no == "<NA>"，展延/大額還款合約另外處裡。 
loan_remain_null = loan_remain_merged[loan_remain_merged['pre_examine_no'].isin(["<NA>"])].copy()

# (b) pre_examine_no != "<NA>" and loan_code_type == "(2)建融"，immovable_worth強制補0。
loan_remain_notnull_construction = loan_remain_merged[~loan_remain_merged['pre_examine_no'].isin(["<NA>"]) & \
                                                      loan_remain_merged["loan_code_type"].isin(["(2)建融"])].copy()
loan_remain_notnull_construction['immovable_worth'] = 0.

# (c) pre_examine_no != "<NA>" and loan_code_type != "(2)建融"，immovable_worth若為空值則補pre_immovable_worth。
loan_remain_notnull_notconstruction = loan_remain_merged[~(loan_remain_merged['pre_examine_no'].isin(["<NA>"]) | \
                                                           loan_remain_merged["loan_code_type"].isin(["(2)建融"]))].copy()


loan_remain_notnull_notconstruction = pd.merge(loan_remain_notnull_notconstruction, pre_worth, on='pre_examine_no', how='left')

# 產生暫時性針對合約撥款金額計算出來的擔保值
# loan_remain_notnull_notconstruction["tmp_pre_immovable_worth"] = loan_remain_notnull_notconstruction.loan_capital * loan_remain_notnull_notconstruction.pre_immovable_worth_rate
# for col, pre_col in [('immovable_worth', 'tmp_pre_immovable_worth'), ('immovable_worth_rate', 'pre_immovable_worth_rate')]:
#     if col in loan_remain_notnull_notconstruction.columns and pre_col in loan_remain_notnull_notconstruction.columns:
#         loan_remain_notnull_notconstruction[col] = loan_remain_notnull_notconstruction[col].combine_first(loan_remain_notnull_notconstruction[pre_col])
# loan_remain_notnull_notconstruction = loan_remain_notnull_notconstruction.drop(columns=[col for col in ['pre_immovable_worth', 'tmp_pre_immovable_worth', 'pre_immovable_worth_rate'] if col in loan_remain_notnull_notconstruction.columns])

# 彙整
loan_remain_merged = pd.concat([loan_remain_notnull_construction,
                                loan_remain_notnull_notconstruction,
                                loan_remain_null], ignore_index=True).sort_values(['pre_examine_no', 'loan_no']).reset_index(drop=True)

# 讀取 table4_loan_remain_immovable_gage.xlsx 的 sheetName = "RawData"
loan_immovable_gage = pd.read_excel(table4_file_path, sheet_name='merged_city_expanded')
loan_immovable_gage["loan_no"] = pd.to_numeric(loan_immovable_gage["loan_no"], errors='coerce').astype('Int64').astype(str)


# 合約資訊彙整
loan_immovable_city = [col for col in loan_immovable_gage.columns if col != "loan_no"]
loan_remain_merged = pd.merge(loan_remain_merged, loan_immovable_gage, on='loan_no', how='left')

for col in loan_immovable_city:
    loan_remain_merged[col] = loan_remain_merged[col].fillna(0)

loan_remain_merged['city_count'] = loan_remain_merged[loan_immovable_city].sum(axis=1)
loan_remain_merged['city_none'] = (loan_remain_merged['city_count'] == 0).astype(int)

# === 「擔保值」城市表 ===
gage_cols = ['loan_no', 'immovable_worth'] + loan_immovable_city + ['city_count', 'city_none']
loan_city_immovable_gage = loan_remain_merged[gage_cols].copy()

loan_city_immovable_gage = distribute_to_cities(
    loan_city_immovable_gage,
    value_col='immovable_worth',
    city_flag_cols=loan_immovable_city,
    city_count_col='city_count',
    city_none_col='city_none'
)

# === 「授信本金」城市表 ===
gage_cols_capital = ['loan_no', 'loan_capital'] + loan_immovable_city + ['city_count', 'city_none']
loan_city_capital = loan_remain_merged[gage_cols_capital].copy()

loan_city_capital = distribute_to_cities(
    loan_city_capital,
    value_col='loan_capital',
    city_flag_cols=loan_immovable_city,
    city_count_col='city_count',
    city_none_col='city_none'
)

# === 「本金餘額」城市表 ===
gage_cols_remain = ['loan_no', 'loan_remain_capital'] + loan_immovable_city + ['city_count', 'city_none']
loan_city_remain = loan_remain_merged[gage_cols_remain].copy()

loan_city_remain = distribute_to_cities(
    loan_city_remain,
    value_col='loan_remain_capital',
    city_flag_cols=loan_immovable_city,
    city_count_col='city_count',
    city_none_col='city_none'
)

loan_remain_merged = pd.merge(
    loan_remain_merged, 
    loan_city_immovable_gage, 
    on='loan_no', 
    how='left', 
    suffixes=('', '_worth')).merge(
        loan_city_capital, 
        on='loan_no', 
        how='left', 
        suffixes=('', '_capital')).merge(
            loan_city_remain, 
            on='loan_no', 
            how='left',
            suffixes=('', '_remain')
        )


# 新增 approval_date_return
# 先取得所有 pre_examine_no_return 對應的 approval_date
pre_return_dates = merged.groupby('pre_examine_no_return')['approval_date'].agg(list).to_dict()

def get_approval_date_return(row):
    pre_code_types = set(row.get('pre_code_type', '').split(',')) if 'pre_code_type' in row else set()
    dates = pre_return_dates.get(row['pre_examine_no_return'], [])
    if len(dates) == 0:
        return np.nan
    if {'(1)土融', '(2)建融'}.issubset(pre_code_types):
        # 兩者皆存在，取最小日期
        valid_dates = [d for d in dates if pd.notnull(d)]
        return min(valid_dates) if valid_dates else np.nan
    else:
        return np.nan

merged['approval_date_return'] = merged.apply(get_approval_date_return, axis=1)

# 新增 approval_date_returnV2
# 若 approval_date_return 為空，則用 approval_date 補上
merged['approval_date_returnV2'] = merged['approval_date_return']
merged['approval_date_returnV2'] = merged['approval_date_returnV2'].where(
    merged['approval_date_returnV2'].notnull(), merged['approval_date']
)


# 新增 pre_code_type_both_group
# 判斷 pre_code_type 是否同時有 '(1)土融' 及 '(2)建融'。目前用於判斷同案土建融的有效額度期限
def has_both_code_type(series):
    types = set()
    for val in series:
        types.update([v.strip() for v in str(val).split(',') if v.strip()])
    return 1 if {'(1)土融', '(2)建融'}.issubset(types) else 0

pre_code_type_both_group = merged.groupby('pre_examine_no_return')['pre_code_type'].apply(has_both_code_type).reset_index().rename(columns={'pre_code_type': 'pre_code_type_both_group'})
merged = pd.merge(merged, pre_code_type_both_group, on='pre_examine_no_return', how='left')


# 新增 approval_valid_date
def calc_valid_date(row):
    base_date = row['approval_date_returnV2']
    if pd.isnull(base_date):
        return np.nan
    if row.get('bus_name', '') == '資金貸與' and row.get('pre_code_type', '') == '(2)建融':
        return base_date + pd.DateOffset(months=12)
    elif row.get('pre_code_type', '') == '(2)建融' and row.get('pre_code_type_both_group', 0) == 1:
        return base_date + pd.DateOffset(months=48)
    else:
        return base_date + pd.DateOffset(months=6)

merged['approval_valid_date'] = merged.apply(calc_valid_date, axis=1)

# 新增 quota_valid 欄位
# 指定日期
cutoff_date = pd.to_datetime('2026-05-31')
merged['quota_valid'] = np.where(
    merged['approval_valid_date'] < cutoff_date,
    0,
    1
)

# 新增剩餘額度欄位
# 型別轉換，確保為數值
merged['approved_amount'] = pd.to_numeric(merged['approved_amount'], errors='coerce').fillna(0)
merged['total_loan_capital'] = pd.to_numeric(merged['total_loan_capital'], errors='coerce').fillna(0)
merged['approved_amount_remain'] = merged['approved_amount'] - merged['total_loan_capital']
# 新增未失效額度
merged['approved_amount_remain_valid'] = merged.apply(lambda x: x['approved_amount_remain'] if x.get('quota_valid', 0) == 1 else 0, axis=1)
# 新增已失效額度
merged['approved_amount_remain_invalid'] = merged.apply(lambda x: x['approved_amount_remain'] if x.get('quota_valid', 0) == 0 else 0, axis=1)

# === 表一 ===
# 2.1 取qtgroup_name、customer_name、pre_examine_no_return、pre_code_type、total_approved_amount，distinct
table1 = merged[[
    'qtgroup_name', 'customer_name', 'pre_examine_no_return', 'pre_code_type', 'total_approved_amount'
]].drop_duplicates()
# 2.2 取pre_examine_no_return，加總approved_amount(取名為total_approved_amount_system)
table1_sum = merged.groupby('pre_examine_no_return', as_index=False)['approved_amount'].sum().rename(columns={'approved_amount': 'total_approved_amount_system'})
# 2.3 生成表一
table1_final = pd.merge(table1, table1_sum, on='pre_examine_no_return', how='left')
table1_final["total_approved_amount_diff"] = table1_final["total_approved_amount"] - table1_final["total_approved_amount_system"]

# === 表二 ===
table2 = merged.groupby(['pre_examine_no_return', 'pre_code_type'], as_index=False).agg({
    'approved_amount': 'sum',
    # 'approved_amount_remain': 'sum',
    'approved_amount_remain_valid': 'sum',
    'approved_amount_remain_invalid': 'sum',
    'total_loan_capital': 'sum' # 所有已撥款紀錄
}).rename(columns={
    'approved_amount': 'approved_amount_return',
    # 'approved_amount_remain': 'approved_amount_remain_return',
    'approved_amount_remain_valid': 'approved_amount_remain_valid_return',
    'approved_amount_remain_invalid': 'approved_amount_remain_invalid_return',
    'total_loan_capital': 'total_loan_capital_original'
})

# === 表三 ===
# 以 loan_remain_merged 統計 group by pre_examine_no_return, loan_code_type
# build aggregation dict programmatically to avoid repeating city column names
agg_dict = {
    'loan_no': lambda x: x.nunique(),
    'loan_capital': 'sum',
    'loan_remain_capital': 'sum',
    'loan_extend': 'sum' if 'loan_extend' in loan_remain_merged.columns else (lambda x: 0),
    'immovable_worth': 'sum',
}

for city in loan_immovable_city:
    agg_dict[f"{city}_worth"] = 'sum'

for city in loan_immovable_city:
    agg_dict[f"{city}_capital"] = 'sum'

for city in loan_immovable_city:
    agg_dict[f"{city}_remain"] = 'sum'


agg_dict['city_none_worth'] = 'sum'
agg_dict['city_none_capital'] = 'sum'
agg_dict['city_none_remain'] = 'sum'

table3 = loan_remain_merged.groupby(['pre_examine_no_return', 'loan_code_type'], as_index=False).agg(agg_dict)

table3 = table3.rename(columns={
    'loan_no': 'loan_no_count',
    'loan_capital': 'total_loan_capital',
    'loan_remain_capital': 'total_loan_remain',
    'loan_extend': 'total_loan_extend',
})



# === 表三延伸表(城市 / 融資類型 / 剩餘本金 / 擔保品價值) ===
city_list = loan_immovable_city + ["city_none"]
city_gage_rows = []
for city in city_list:
    for idx, row in loan_city_immovable_gage.iterrows():
        city_gage_rows.append({
            'city': city,
            'loan_no': row['loan_no'],
            'immovable_gage_worth': row[city],
            'immovable_worth': row['immovable_worth']
        })
city_gage_df = pd.DataFrame(city_gage_rows)

# 2. 從 table3 取出每個 city 的 {city}_remain 及 {city}_worth，分別命名為 loan_capital_remain 與 immovable_gage_worth
table3_v2 = pd.merge(table3, table1_final[["pre_examine_no_return", "customer_name"]].drop_duplicates(), on=['pre_examine_no_return'], how='left')
# 當 pre_examine_no_return 為 NULL 或 NA 時，customer_name 補 "資料不明"
table3_v2['customer_name'] = table3_v2.apply(
    lambda x: '資料不明' if (pd.isna(x['pre_examine_no_return']) or x['pre_examine_no_return'] == 'NULL' or pd.isna(x['customer_name'])) else x['customer_name'], axis=1)

city_stat_rows = []
for city in city_list:
    remain_col = f"{city}_remain"
    worth_col = f"{city}_worth"
    if remain_col in table3_v2.columns and worth_col in table3_v2.columns:
        for idx, row in table3_v2.iterrows():
            city_stat_rows.append({
                'pre_examine_no_return': row['pre_examine_no_return'],
                'loan_code_type': row['loan_code_type'],
                'city': city,
                'customer_name': row['customer_name'], 
                'loan_capital_remain': row[remain_col],
                'immovable_gage_worth': row[worth_col]
            })
city_stat_df = pd.DataFrame(city_stat_rows)

city_stat_df = city_stat_df.groupby(['city', 'loan_code_type', 'customer_name'], as_index=False).agg({
    'loan_capital_remain': 'sum',
    'immovable_gage_worth': 'sum'
}).sort_values(['city', 'loan_code_type', 'customer_name']).reset_index(drop=True)
# 刪除 loan_capital_remain 與 immovable_gage_worth 同時為 0 的資料
city_stat_df = city_stat_df[~((city_stat_df['loan_capital_remain'] == 0) & (city_stat_df['immovable_gage_worth'] == 0))].reset_index(drop=True)

# === 總表 ===
# 合併
summary_table = pd.merge(table1_final, table2, on=['pre_examine_no_return', 'pre_code_type'], how='left')
addition_summary_table = pd.DataFrame({
    "pre_examine_no_return": ['NULL', 'NULL', 'NULL', 'NULL', 'NULL'],
    "pre_code_type": ['(1)土融', '(2)建融', '(3)土建融', '(4)餘屋', '(5)其他'],
})
summary_table = pd.concat([summary_table, addition_summary_table], axis = 0).reset_index(drop=True)
summary_table = pd.merge(
    summary_table,
    table3,
    left_on=['pre_examine_no_return', 'pre_code_type'],
    right_on=['pre_examine_no_return', 'loan_code_type'],
    how='left'
)

# build summary column list programmatically to avoid duplication
base_cols = [
    "qtgroup_name",
    "customer_name",
    "total_approved_amount",
    "total_approved_amount_system",
    "total_approved_amount_diff",
    "pre_examine_no_return",
    "pre_code_type",
    "approved_amount_return",
    "approved_amount_remain_valid_return",
    "approved_amount_remain_invalid_return",
    "total_loan_capital_original",
    "loan_no_count",
    "total_loan_capital",
    "total_loan_remain",
    "total_loan_extend",
    "immovable_worth",
]

worth_cols = [f"{city}_worth" for city in loan_immovable_city] + ["city_none_worth"]
capital_cols = [f"{city}_capital" for city in loan_immovable_city] + ["city_none_capital"]
remain_cols = [f"{city}_remain" for city in loan_immovable_city] + ["city_none_remain"]

summary_cols = base_cols + worth_cols + capital_cols + remain_cols
summary_table = summary_table[summary_cols]

# 輸出四個表
with pd.ExcelWriter(r'./Result/截至20260531/Summary_截至20260531.xlsx') as writer:
    merged.to_excel(writer, sheet_name='RawData', index=False)
    # table1_final.to_excel(writer, sheet_name='表一', index=False)
    # table2.to_excel(writer, sheet_name='表二', index=False)
    table3.to_excel(writer, sheet_name='合約本餘_擔保品', index=False)
    city_stat_df.to_excel(writer, sheet_name='城市統計表_擔保品', index=False)
    summary_table.to_excel(writer, sheet_name='總表', index=False)