import re
import pandas as pd
from typing import Optional, Dict

def parse_address(address: Optional[str]) -> Dict[str, Optional[str]]:
    """
    拆解地址，回傳縣市、行政區、無法辨識標記。
    """
    if address is None or str(address).strip() == "" or pd.isna(address):
        return {"無法辨識": "Y", "縣市": None, "行政區": None}

    address = str(address).replace("臺", "台")

    # 個別認定特殊 case
    if "新竹市關新路" in address:
        return {"無法辨識": "N", "縣市": "新竹市", "行政區": "東區"}
    if "東勢寮171" in address:
        return {"無法辨識": "N", "縣市": "台南市", "行政區": "善化區"}
    if "台南善化區中山路500巷4號" in address:
        return {"無法辨識": "N", "縣市": "台南市", "行政區": "善化區"}
    if "彰化市進德路9巷2號" in address:
        return {"無法辨識": "N", "縣市": "彰化縣", "行政區": "彰化市"}

    # 明確符合「縣市 + 行政區」格式
    m = re.match(r"^(.*?[市縣])(.*?[區鎮鄉市])", address)
    if m:
        city = m.group(1)
        district = m.group(2)
        if city and district:
            return {"無法辨識": "N", "縣市": city, "行政區": district}

    return {"無法辨識": "Y", "縣市": None, "行政區": None}


def preprocess_building_address(input_xlsx: str, output_xlsx: str) -> None:
    """
    處理建物地址資料，輸出結構化結果為 xlsx。
    input_xlsx: 原始XLSX檔案路徑
    output_xlsx: 處理後輸出檔案路徑（xlsx）
    """
    df = pd.read_excel(input_xlsx, sheet_name="building_address", dtype=str)
    # 保留原始資料內容
    results = df['building_address'].apply(parse_address)
    results_df = pd.DataFrame(list(results))
    df = pd.concat([df, results_df], axis=1)
    df.to_excel(output_xlsx, index=False)


if __name__ == "__main__":
    # 範例用法
    preprocess_building_address(
        input_xlsx="不動產_擔保品.xlsx",
        output_xlsx="不動產_擔保品_building_address_拆解結果.xlsx"
    )
