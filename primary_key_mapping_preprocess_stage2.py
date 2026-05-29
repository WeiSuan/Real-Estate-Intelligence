import pandas as pd
import numpy as np

INPUT_STAGE1 = "./Result/截至20260430/table2_primary_key_mapping_preprocess1.xlsx"
INPUT_MANUAL = "./Result/截至20260430/table2_primary_key_mapping_manual_preprocess.xlsx"
OUTPUT_FILE = "./Result/截至20260430/table2_primary_key_mapping.xlsx"

def to_int_then_str(series: pd.Series) -> pd.Series:
    num = pd.to_numeric(series, errors="coerce")
    int_series = num.where(num.notnull(), pd.NA).astype("Int64")
    obj = int_series.astype(object).where(int_series.notnull(), "")
    return obj.astype(str)


def pick_col(df: pd.DataFrame, candidates):
    for c in candidates:
        if c in df.columns:
            return c
    return None


def main():
    # 1) 讀取 stage1 檔案
    pre_examine_merged = pd.read_excel(INPUT_STAGE1, sheet_name="pre_examine_merged")
    loan_merged = pd.read_excel(INPUT_STAGE1, sheet_name="loan_merged")

    # 2) 讀取 manual 檔案
    pre_manual = pd.read_excel(INPUT_MANUAL, sheet_name="pre_examine_guarantee_revised")
    loan_manual = pd.read_excel(INPUT_MANUAL, sheet_name="loan_guarantee_revised")

    # 對可能的 key 欄位先 int->str
    for df in (pre_examine_merged, pre_manual):
        for col in ("pre_examine_no", "pre_examine_no_major"):
            if col in df.columns:
                df[col] = to_int_then_str(df[col])

    for df in (loan_merged, loan_manual):
        if "loan_no" in df.columns:
            df["loan_no"] = to_int_then_str(df["loan_no"])

    # pre_manual 只取需要欄位（容錯不同拼字）
    pre_manual_col = pick_col(pre_manual, ["pre_examine_no", "pre_examine_no_previous"])
    pre_manual_worth_col = pick_col(pre_manual, ["pre_immovable_worth_mannual", "pre_immovable_worth_manual"])
    pre_manual_sel = pre_manual[[pre_manual_col, pre_manual_worth_col]].rename(columns={pre_manual_col: "pre_examine_no", pre_manual_worth_col: "pre_immovable_worth_mannual"})

    # loan_manual 只取需要欄位
    loan_manual_col = pick_col(loan_manual, ["loan_no", "loan_no_previous"])
    loan_manual_worth_col = pick_col(loan_manual, ["immovable_worth_manual", "immovable_worth_mannual"])
    loan_manual_sel = loan_manual[[loan_manual_col, loan_manual_worth_col]].rename(columns={loan_manual_col: "loan_no", loan_manual_worth_col: "immovable_worth_manual"})

    # 合併 pre_examine
    pre_merged = pre_examine_merged.merge(pre_manual_sel, on="pre_examine_no", how="left")
    # 新欄位 pre_immovable_worth：若 manual 存在則用之，否則用 pre_immovable_worth_previous
    prev_col = pick_col(pre_merged, ["pre_immovable_worth_previous", "pre_immovable_worth_previous_previous"])
    pre_merged["pre_immovable_worth"] = np.where(
        pre_merged["pre_immovable_worth_mannual"].notna() & (pre_merged["pre_immovable_worth_mannual"] != ""),
        pre_merged["pre_immovable_worth_mannual"],
        pre_merged[prev_col] if prev_col in pre_merged.columns else np.nan,
    )

    # 合併 loan
    loan_merged2 = loan_merged.merge(loan_manual_sel, on="loan_no", how="left")
    prev_loan_col = pick_col(loan_merged2, ["immovable_worth_previous", "immovable_worth_previous_previous"])
    loan_merged2["immovable_worth"] = np.where(
        loan_merged2["immovable_worth_manual"].notna() & (loan_merged2["immovable_worth_manual"] != ""),
        loan_merged2["immovable_worth_manual"],
        loan_merged2[prev_loan_col] if prev_loan_col in loan_merged2.columns else np.nan,
    )

    primary_mapping_code_stage2 = pre_merged.copy()
    primary_mapping_code_stage2 = primary_mapping_code_stage2[["pre_examine_no", "pre_examine_no_return", "approved_amount", "pre_examine_no_return_previous", "total_approved_amount_previous"]]
    tmp_ = primary_mapping_code_stage2.groupby("pre_examine_no_return", as_index=False)["approved_amount"].sum().rename(columns={"approved_amount": "total_approved_amount"})
    primary_mapping_code_stage2 = primary_mapping_code_stage2.merge(tmp_, on="pre_examine_no_return", how="left")
    primary_mapping_code_stage2 = primary_mapping_code_stage2[["pre_examine_no", "pre_examine_no_return", "approved_amount", "total_approved_amount", "pre_examine_no_return_previous", "total_approved_amount_previous"]]
    primary_mapping_code_stage2.sort_values(["pre_examine_no_return", "pre_examine_no"], inplace=True)

    pre_examine_worth_rate_stage2 = pre_merged[["pre_examine_no", "approved_amount", "pre_immovable_worth"]].copy()
    loan_worth_rate_stage2 = loan_merged2[["loan_no", "loan_capital", "immovable_worth", "immovable_worth_system"]].copy()

    # 儲存為 xlsx，兩個 sheet
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        primary_mapping_code_stage2.to_excel(writer, sheet_name="代碼對應表", index=False)
        pre_examine_worth_rate_stage2.to_excel(writer, sheet_name="初審單案擔保率", index=False)
        loan_worth_rate_stage2.to_excel(writer, sheet_name="合約單案擔保率(已撥)", index=False)

if __name__ == "__main__":
    main()
