import re
import pandas as pd
from typing import Optional, Dict, Union

from data_function import get_mssql_df, prev_month_end

def parse_address(address: Optional[str]) -> Dict[str, Optional[str]]:
    """Return dict with keys: clear_or_not, city, district, section (section=None for building)."""
    if address is None or pd.isna(address) or str(address).strip() == "":
        return {"clear_or_not": "Y", "city": None, "district": None, "section": None}
    text = str(address).replace("臺", "台")

    # 個別認定特殊 case (restore cases from building_address_preprocess.py)
    if "新竹市關新路" in text:
        return {"clear_or_not": "N", "city": "新竹市", "district": "東區", "section": None}
    if "東勢寮171" in text:
        return {"clear_or_not": "N", "city": "台南市", "district": "善化區", "section": None}
    if "台南善化區中山路500巷4號" in text:
        return {"clear_or_not": "N", "city": "台南市", "district": "善化區", "section": None}
    if "彰化市進德路9巷2號" in text:
        return {"clear_or_not": "N", "city": "彰化縣", "district": "彰化市", "section": None}
    if "嘉義市大雅路一段368巷216號" in text:
        return {"clear_or_not": "N", "city": "嘉義市", "district": "東區", "section": None}

    m = re.match(r"^(.*?[市縣])(.*?[區鎮鄉市])", text)
    if m:
        city = m.group(1)
        district = m.group(2)
        if city and district:
            return {"clear_or_not": "N", "city": city, "district": district, "section": None}
    return {"clear_or_not": "Y", "city": None, "district": None, "section": None}


def parse_land_section(land_section: Optional[str]) -> Dict[str, Optional[str]]:
    """Return dict with keys: clear_or_not, city, district, section."""
    if land_section is None or pd.isna(land_section) or str(land_section).strip() == "" or str(land_section).strip().upper() == "NULL":
        return {"clear_or_not": "Y", "city": None, "district": None, "section": None}
    text = str(land_section).replace("臺", "台")
    m = re.match(r"^([\u4e00-\u9fa5]{2,3}[縣市])([\u4e00-\u9fa5]{1,4}[區鎮鄉市])([\u4e00-\u9fa5\d]+段)$", text)
    if m:
        m_return = {"clear_or_not": "N", "city": m.group(1), "district": m.group(2), "section": m.group(3)}

    # 個案hard code對應（來源: land_section_preprocess.py）
    hard_code_cases = {
        "北投區崇仰段": ("N", "台北市", "北投區", "崇仰段"),
        "臺北市大同區市府段": ("N", "台北市", "大同區", "市府段"),
        "三重段": ("N", "新北市", "三重區", "三重段"),
        "三重區中民段": ("N", "新北市", "三重區", "中民段"),
        "三重區集美段": ("N", "新北市", "三重區", "集美段"),
        "三重區成功段": ("N", "新北市", "三重區", "成功段"),
        "新莊區中原段": ("N", "新北市", "新莊區", "中原段"),
        "土城區永和段": ("N", "新北市", "土城區", "永和段"),
        "港子嘴段": ("N", "新北市", "板橋區", "港子嘴段"),
        "金美段": ("N", "新北市", "金山區", "金美段"),
        "濛濛谷段": ("N", "新北市", "新店區", "濛濛谷"),
        "蘆洲區復興段": ("N", "新北市", "蘆洲區", "復興段"),
        "秀岡一段": ("N", "新北市", "新店區", "秀岡一段"),
        "聖德段": ("N", "桃園市", "中壢區", "聖德段"),
        "青芝段": ("N", "桃園市", "中壢區", "青芝段"),
        "中路三段": ("N", "桃園市", "桃園區", "中路三段"),
        "下田心子段": ("N", "桃園市", "新屋區", "下田心子段"),
        "楊梅區和平段": ("N", "桃園市", "楊梅區", "和平段"),
        "新鼻段": ("N", "桃園市", "蘆竹區", "新鼻段"),
        "宏華段": ("N", "桃園市", "蘆竹區", "宏華段"),
        "草富段": ("N", "桃園市", "觀音區", "草富段"),
        "龍潭區石門": ("N", "桃園市", "龍潭區", "石門段"),
        "龍潭區金龍段": ("N", "桃園市", "龍潭區", "金龍段"),
        "桃園市 桃園區 中山段": ("N", "桃園市", "桃園區", "中山段"),
        "大觀段" : ("N", "桃園市", "觀音區", "大觀段"), # 有分「觀音區」跟「大園區」
        "二重溪" : ("N", "桃園市", "楊梅區", "二重溪段"),
        "大忠段" : ("N", "桃園市", "八德區", "大忠段"),
        "三層段" : ("N", "桃園市", "大溪區", "三層段"),
        "新竹市中寮段": ("N", "新竹市", "北區", "中寮段"),
        "新竹市隆恩段": ("N", "新竹市", "東區", "隆恩段"),
        "新竹市古賢段": ("N", "新竹市", "東區", "古賢段"),
        "新竹市仙水段": ("N", "新竹市", "東區", "仙水段"),
        "新竹市師院段": ("N", "新竹市", "香山區", "師院段"),
        "新竹市西濱段": ("N", "新竹市", "香山區", "西濱段"),
        "新竹市忠孝段": ("N", "新竹市", "香山區", "忠孝段"),
        "新竹市和平段": ("N", "新竹市", "香山區", "和平段"),
        "新竹市港南段": ("N", "新竹市", "香山區", "港南段"),
        "新竹縣北埔鄉鄉所段": ("N", "新竹縣", "北埔鄉", "鄉所段"),
        "新港段": ("N", "新竹市", "北區", "新港段"),
        "長嶺段": ("N", "新竹縣", "湖口鄉", "長嶺段"),
        "苑裡鎮苑西段": ("N", "苗栗縣", "苑裡鎮", "苑西段"),
        "銅鑼鄉老雞隆段": ("N", "苗栗縣", "銅鑼鄉", "老雞隆段"),
        "銅鑼鄉福興段": ("N", "苗栗縣", "銅鑼鄉", "福興段"),
        "新博愛段": ("N", "苗栗縣", "竹南鎮", "新博愛段"),
        "福興段": ("N", "苗栗縣", "銅鑼鄉", "福興段"),
        "文發段": ("N", "苗栗縣", "苗栗市", "文發段"),
        "文峰段": ("N", "苗栗縣", "苗栗市", "文峰段"),
        "芒埔段": ("N", "苗栗縣", "苗栗市", "芒埔段"),
        "育賢段": ("N", "台中市", "太平區", "育賢段"),
        "台中市南屯區鎮福段": ("N", "台中市", "南屯區", "鎮福段"),        
        "古夷段": ("N", "彰化縣", "彰化市", "古夷段"),
        "鹿鳴段": ("N", "彰化縣", "鹿港鎮", "鹿鳴段"),
        "名間鄉吳厝段": ("N", "南投縣", "名間鄉", "吳厝段"),
        "斗六市大竹圍段": ("N", "雲林縣", "斗六市", "大竹圍段"),
        "虎尾鎮龍安段": ("N", "雲林縣", "虎尾鎮", "龍安段"),
        "虎尾鎮芳草段": ("N", "雲林縣", "虎尾鎮", "芳草段"),
        "虎尾鎮竹安段": ("N", "雲林縣", "虎尾鎮", "竹安段"),
        "虎尾鎮大屯段": ("N", "雲林縣", "虎尾鎮", "大屯段"),
        "虎尾鎮大學段": ("N", "雲林縣", "虎尾鎮", "大學段"),
        "嘉義市新太平段": ("N", "嘉義市", "東區", "新太平段"),
        "短竹段":("N", "嘉義市", "東區", "短竹段"),
        "嘉義市頂庄段": ("N", "嘉義市", "東區", "頂庄段"),
        "嘉義市竹圍子段": ("N", "嘉義市", "西區", "竹圍子段"),
        "嘉義市育人段": ("N", "嘉義市", "西區", "育人段"),
        "嘉義市盧東段": ("N", "嘉義市", "東區", "盧東段"),
        "嘉義市盧厝段": ("N", "嘉義市", "東區", "盧厝段"),
        "嘉義市崇文段": ("N", "嘉義市", "東區", "崇文段"),
        "六甲區忠孝段": ("N", "台南市", "六甲區", "忠孝段"),
        "白河區國泰段": ("N", "台南市", "白河區", "國泰段"),
        "東勢角段": ("N", "台南市", "佳里區", "東勢角段"),
        "南區鹽埕段": ("N", "台南市", "南區", "鹽埕段"),
        "後壁區侯伯段": ("N", "台南市", "後壁區", "侯伯段"),
        "將軍區漚汪段": ("N", "台南市", "將軍區", "漚汪段"),
        "西港區成功段": ("N", "台南市", "西港區", "成功段"),
        "歸仁北段": ("N", "台南市", "歸仁區", "歸仁北段"),
        "歸仁區大苓段": ("N", "台南市", "歸仁區", "大苓段"),
        "大苓段": ("N", "台南市", "歸仁區", "大苓段"),
        "永康區大灣段": ("N", "台南市", "永康區", "大灣段"),
        "台南關廟區南雄段": ("N", "台南市", "關廟區", "南雄段"),
        "台南市佳里區鎮山段": ("N", "台南市", "佳里區", "鎮山段"),        
        "台南市善化區東勢寮": ("N", "台南市", "善化區", "東勢寮"),        
        "三民區大港段": ("N", "高雄市", "三民區", "大港段"),
        "三民區三塊厝段": ("N", "高雄市", "三民區", "三塊厝段"),
        "大樹區大樹保安段": ("N", "高雄市", "大樹區", "大樹保安段"),
        "岡山區友情段": ("N", "高雄市", "岡山區", "友情段"),
        "岡山區石螺段": ("N", "高雄市", "岡山區", "石螺段"),
        "楠梓區高昌段": ("N", "高雄市", "楠梓區", "高昌段"),
        "高昌段": ("N", "高雄市", "楠梓區", "高昌段"),
        "楠梓區楠都段": ("N", "高雄市", "楠梓區", "楠都段"),
        "楠梓區藍田中段": ("N", "高雄市", "楠梓區", "藍田中段"),
        "鳳山區鳳翔段": ("N", "高雄市", "鳳山區", "鳳翔段"),
        "鳳翔段": ("N", "高雄市", "鳳山區", "鳳翔段"),
        "內埔鄉豐田段": ("N", "屏東縣", "內埔鄉", "豐田段"),
        "屏東縣林邊鄉東林段": ("N", "屏東縣", "林邊鄉", "東林段"),
        "萬丹鄉新安段": ("N", "屏東縣", "萬丹鄉", "新安段"),
        "冬山鄉八仙段": ("N", "宜蘭縣", "冬山鄉", "八仙段"),
        "同結段": ("N", "宜蘭縣", "員山鄉", "同結段"),
        "金門縣金湖鎮料羅村測" : ("N", "金門縣", "金湖鎮", "料羅村測段")
    }
    key = str(land_section).strip()
    if key in hard_code_cases:
        n, city, admin, section = hard_code_cases[key]
        return {"clear_or_not": n, "city": city, "district": admin, "section": section}
    elif 'm_return' in locals():
        return m_return
    else:
        return {"clear_or_not": "Y", "city": None, "district": None, "section": None}


def _ensure_cols(df: pd.DataFrame, cols):
    for c in cols:
        if c not in df.columns:
            df[c] = None


def preprocess_building_land(input_df: pd.DataFrame, output_xlsx: str) -> None:
    """
    Read Excel, add primary_key, parse building and land, then left-join back with prefixes.

    - building join: key = primary_key, prefix = 'building_'
    - land join: if original df has column 'land', attempt join using that as key to land parsed primary_key; otherwise join on primary_key. Prefix = 'land_'
    """
    df = input_df.copy()

    # 新增 primary_key 作為後續表的串接key_id
    df['primary_key'] = (df.index + 1).astype(int)

    # parse building_address
    building_parsed = pd.DataFrame(list(df['building_address'].apply(parse_address)))
    _ensure_cols(building_parsed, ['clear_or_not', 'city', 'district', 'section'])
    building_parsed = building_parsed[['clear_or_not', 'city', 'district', 'section']]
    building_parsed.insert(0, 'primary_key', df['primary_key'].values)
    building_parsed = building_parsed.rename(columns={
        'clear_or_not': 'building_clear_or_not',
        'city': 'building_city',
        'district': 'building_district',
        'section': 'building_section'
    })

    # parse land_section
    land_parsed = pd.DataFrame(list(df['land_section'].apply(parse_land_section)))
    _ensure_cols(land_parsed, ['clear_or_not', 'city', 'district', 'section'])
    land_parsed = land_parsed[['clear_or_not', 'city', 'district', 'section']]
    land_parsed.insert(0, 'primary_key', df['primary_key'].values)
    land_parsed = land_parsed.rename(columns={
        'clear_or_not': 'land_clear_or_not',
        'city': 'land_city',
        'district': 'land_district',
        'section': 'land_section'
    })

    merged = df.merge(building_parsed, on='primary_key', how='left')

    if 'land' in df.columns:
        merged = merged.merge(land_parsed, left_on='land', right_on='primary_key', how='left', suffixes=(False, False))
        if 'primary_key_y' in merged.columns:
            merged = merged.drop(columns=['primary_key_y'])
        if 'primary_key_x' in merged.columns:
            merged = merged.rename(columns={'primary_key_x': 'primary_key'})
    else:
        merged = merged.merge(land_parsed, on='primary_key', how='left')

    merged.rename(columns={'land_section_x': 'land_section', 'land_section_y': 'land_location'}, inplace=True)

    # city_result：如果 building 明確 (N) 則取 building_city，否則取 land_city
    merged['city_result'] = merged.apply(
        lambda r: r.get('building_city') if str(r.get('building_clear_or_not')) == 'N' else r.get('land_city'),
        axis=1
    )
    # district_result：如果 building 明確 (N) 則取 building_district，否則取 land_district
    merged['district_result'] = merged.apply(
        lambda r: r.get('building_district') if str(r.get('building_clear_or_not')) == 'N' else r.get('land_district'),
        axis=1
    )
    # location_result：若 building_clear_or_not == 'N' 則空字串，否則取 land_location
    merged['location_result'] = merged.apply(
        lambda r: '' if str(r.get('building_clear_or_not')) == 'N' else (r.get('land_location') if r.get('land_location') is not None else ''),
        axis=1
    )
    # both_status：當兩者皆為 'N' -> '兩項並存'；任一為 'N' -> '至少存在一項'；兩者皆為 'Y' -> '無資料'
    def _both_status(r):
        b = str(r.get('building_clear_or_not'))
        l = str(r.get('land_clear_or_not'))
        if b == 'N' and l == 'N':
            return '兩項並存'
        if b == 'N' or l == 'N':
            return '至少存在一項'
        return '無資料'

    merged['both_status'] = merged.apply(_both_status, axis=1)

    merged_city_only = merged[["loan_no", "city_result"]].copy().drop_duplicates().sort_values(by="loan_no")

    counties = [
        # 北部
        '基隆市',
        '台北市',
        '新北市',
        '桃園市',
        '新竹縣',
        '新竹市',
        '苗栗縣',

        # 中部
        '南投縣',
        '台中市',
        '彰化縣',
        '雲林縣',

        # 南部
        '嘉義縣',
        '嘉義市',
        '台南市',        
        '高雄市',
        '屏東縣',
        
        # 東部
        '宜蘭縣',
        '花蓮縣',
        '台東縣',

        # 外島        
        '澎湖縣',
        '金門縣',
        '連江縣',
    ]

    all_loan_nos = merged['loan_no'].drop_duplicates()

    if not merged_city_only.empty:
        city_pivot = pd.crosstab(merged_city_only['loan_no'], merged_city_only['city_result']).astype(int)
        city_pivot = city_pivot.reindex(all_loan_nos, fill_value=0)
    else:
        city_pivot = pd.DataFrame(0, index=all_loan_nos, columns=counties)

    for c in counties:
        if c not in city_pivot.columns:
            city_pivot[c] = 0

    city_pivot = city_pivot[counties].reset_index()
    city_pivot = city_pivot.rename(columns={city_pivot.columns[0]: 'loan_no'})
    # city_pivot['city_count'] = (city_pivot[counties].sum(axis=1) >= 1).astype(int)
    merged_city_expanded = city_pivot

    with pd.ExcelWriter(output_xlsx, engine='openpyxl') as writer:
        merged.to_excel(writer, sheet_name='merged', index=False)
        merged_city_only.to_excel(writer, sheet_name='merged_city_only', index=False)
        merged_city_expanded.to_excel(writer, sheet_name='merged_city_expanded', index=False)

if __name__ == '__main__':
    
    end_of_last_month = prev_month_end().strftime('%Y%m%d')    
    OUTPUT_XLSX = f'./Result/截至{end_of_last_month}/table4_loan_remain_immovable_gage.xlsx'

    immovable_gage_df = get_mssql_df(
            sql_statement_path=r"./SQL/不動產_擔保品資訊.sql",
            conn_params={
                'SERVER': 'Shindbbackup01',
                'DATABASE': 'shin_monthly',
                'UID': 'scoreap',
                'PWD': 'Nii2Sc@re'
            })

    preprocess_building_land(immovable_gage_df, OUTPUT_XLSX)
