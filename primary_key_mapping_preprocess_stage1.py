# total_approved_amount.py
"""
建立案件總額度檢視：
1. 讀取"不動產_已核准案件.xlsx" sheet="已核准(已撥+未撥款)"，取所需欄位。
2. 根據 customer_name 分組，approval_date 遞增排序。
3. 依據規則產生 pre_examine_no_return。
4. 依 pre_examine_no_return 彙總 approved_amount。
5. left join 回原始表格，存成 excel。
"""
import pandas as pd
from datetime import datetime
from pandas.tseries.offsets import DateOffset

# 1. 讀取資料
input_file = "./Result/截至20260430/table1_all_approval_cases.xlsx"
sheet_name_rawdata = "RawData"

previous_file = "./Result/截至20260430/table2_primary_key_mapping_previous.xlsx"
code_mapping_sheet = "代碼對應表"
pre_examine_guarantee_rate = "初審單案擔保率"
loan_guarantee_rate = "合約單案擔保率(已撥)"

loan_file = "./Result/截至20260430/table3_loan_remain_cases.xlsx"
sheet_name_rawdata = "RawData"

output_file = "./Result/截至20260430/table2_primary_key_mapping(TEST)_V3.xlsx"

def to_int_then_str(series: pd.Series) -> pd.Series:
    """Convert values to integer (nullable) then to string; missing -> empty string."""
    num = pd.to_numeric(series, errors="coerce")
    int_series = num.where(num.notnull(), pd.NA).astype("Int64")
    obj = int_series.astype(object).where(int_series.notnull(), "")
    return obj.astype(str)

# 處理pre_examine_df(最新已核准案件)
usecols = ["pre_examine_no", "approval_date", "approved_amount", "customer_name", "pre_code_type", "total_loan_capital"]
df = pd.read_excel(input_file, sheet_name=sheet_name_rawdata, usecols=usecols, dtype={"pre_examine_no": str, "customer_name": str})
df_code_mapping_previous = pd.read_excel(previous_file, sheet_name=code_mapping_sheet)
df_pre_examine_guarantee_rate = pd.read_excel(previous_file, sheet_name=pre_examine_guarantee_rate)
df_loan_guarantee_rate = pd.read_excel(previous_file, sheet_name=loan_guarantee_rate)

# 處理loan_df(最新已撥貸案件)
loan_df = pd.read_excel(loan_file, sheet_name=sheet_name_rawdata)
for col in ["pre_examine_no", "main_no", "pre_examine_no_major", "examine_group_no", "customer_id_no", "loan_no"]:
    if col in loan_df.columns:
        loan_df[col] = to_int_then_str(loan_df[col])
loan_df["immovable_worth_system"] = pd.to_numeric(loan_df["immovable_worth_system"], errors="coerce").fillna(0).astype(int)

# 處理 df(最新已核准案件)
df["pre_examine_no"] = to_int_then_str(df["pre_examine_no"])
if not pd.api.types.is_datetime64_any_dtype(df["approval_date"]):
    df["approval_date"] = pd.to_datetime(df["approval_date"], errors="coerce")
if not pd.api.types.is_integer_dtype(df["approved_amount"]):
    df["approved_amount"] = pd.to_numeric(df["approved_amount"], errors="coerce").fillna(0).astype(int)
if not pd.api.types.is_integer_dtype(df["total_loan_capital"]):
    df["total_loan_capital"] = pd.to_numeric(df["total_loan_capital"], errors="coerce").fillna(0).astype(int)

# 處理 df_previous (前次代碼對應表)
df_code_mapping_previous["pre_examine_no"] = to_int_then_str(df_code_mapping_previous["pre_examine_no"])
df_code_mapping_previous["pre_examine_no_return"] = to_int_then_str(df_code_mapping_previous["pre_examine_no_return"])
df_code_mapping_previous.columns = [item_ + "_previous" for item_ in df_code_mapping_previous.columns]

# 處理df_pre_examine_guarantee_rate(前次初審擔保率)
df_pre_examine_guarantee_rate["pre_examine_no"] = to_int_then_str(df_pre_examine_guarantee_rate["pre_examine_no"])
df_pre_examine_guarantee_rate.columns = [item_ + "_previous" if item_ != "pre_examine_no" else item_ for item_ in df_pre_examine_guarantee_rate.columns]

# 處理df_loan_guarantee_rate(前次合約擔保率)
df_loan_guarantee_rate["loan_no"] = to_int_then_str(df_loan_guarantee_rate["loan_no"])
df_loan_guarantee_rate.columns = [item_ + "_previous" for item_ in df_loan_guarantee_rate.columns]

# df 歸戶
df = df.sort_values(["customer_name", "approval_date"]).reset_index(drop=True)
pre_examine_no_return = []
for customer, group in df.groupby("customer_name", sort=False):
    group = group.sort_values("pre_examine_no").reset_index()
    prev_date = None
    prev_no_return = None
    for idx, row in group.iterrows():
        if idx == group.index[0]:
            pre_examine_no_return.append(row["pre_examine_no"])
            prev_date = row["approval_date"]
            prev_no_return = row["pre_examine_no"]
        else:
            curr_date = row["approval_date"]
            if pd.notnull(curr_date) and pd.notnull(prev_date):
                # 若相差1個月內
                if abs((curr_date.year - prev_date.year) * 12 + (curr_date.month - prev_date.month)) <= 1:
                    pre_examine_no_return.append(prev_no_return)
                else:
                    pre_examine_no_return.append(row["pre_examine_no"])
                    prev_no_return = row["pre_examine_no"]
                prev_date = curr_date
            else:
                pre_examine_no_return.append(row["pre_examine_no"])
                prev_date = curr_date
                prev_no_return = row["pre_examine_no"]
                
df["pre_examine_no_return"] = pre_examine_no_return
agg_df = df.groupby("pre_examine_no_return", as_index=False)["approved_amount"].sum().rename(columns={"approved_amount": "total_approved_amount"})
result = pd.merge(df, agg_df, on="pre_examine_no_return", how="left")

# 合併
df_new = df.merge(
    df_code_mapping_previous, 
    left_on="pre_examine_no", 
    right_on="pre_examine_no_previous", 
    how="left").merge(
        df_pre_examine_guarantee_rate,
        on = "pre_examine_no",
        how="left"
    )

# 1. 註記前一次pre_examine_no未出現者。
# 2. 註記特定期間內有進行條件變更者。
df_new["update_tag"] = 1 * df_new["pre_examine_no_previous"].isna()
tagged_returns = (
    df_new.loc[df_new["update_tag"] == 1, "pre_examine_no_return"]
    .dropna()
    .unique()
    .tolist()
)

df_manual_revised = df_new[df_new["pre_examine_no_return"].isin(tagged_returns)].copy()
df_manual_revised = df_manual_revised.sort_values(["pre_examine_no_return", "pre_examine_no"]).reset_index(drop=True)


loan_df_new = loan_df.merge(
    df_loan_guarantee_rate,
    left_on="loan_no",
    right_on="loan_no_previous",
    how="left"
)
loan_df_new["update_tag"] = 1 * loan_df_new["loan_no_previous"].isna()
loan_manual_revised = loan_df_new[loan_df_new["loan_no_previous"].isna()].copy()

with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    df_new.to_excel(writer, sheet_name="pre_examine_merged", index=False)
    df_manual_revised.to_excel(writer, sheet_name="pre_examine_guarantee_revised", index=False)
    loan_df_new.to_excel(writer, sheet_name="loan_merged", index=False)
    loan_manual_revised.to_excel(writer, sheet_name="loan_guarantee_revised", index=False)