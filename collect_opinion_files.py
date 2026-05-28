# 讀取審查意見(ALL)RawData.xlsx中的"初審編號"
import os
import pandas as pd
import shutil
rawdata_path = '不動產_尚有本餘初審.xlsx'
df_raw = pd.read_excel(rawdata_path, sheet_name="TEST")
pre_examine_nos = df_raw['pre_examine_no'].dropna().astype(int).astype(str).unique().tolist()
# 網路磁碟路徑
base_path = r'\\shinsvr1\shinMountFiles\publicReport\preexamine'

# 準備下載目錄
download_dir = './審查報告'
os.makedirs(download_dir, exist_ok=True)

results = []

for pre_no in pre_examine_nos:
    folder = os.path.join(base_path, pre_no, 'uploadFiles')
    found_files = []
    if os.path.exists(folder):
        for fname in os.listdir(folder):
            if '審查意見' in fname:
                file_path = os.path.join(folder, fname)
                last_modified = os.path.getmtime(file_path)
                found_files.append({
                    'pre_examine_no': pre_no,
                    'file_name': fname,
                    'last_modified': last_modified,
                    'file_path': file_path
                })
        if found_files:
            # 取最新一筆
            latest = max(found_files, key=lambda x: x['last_modified'])
            # 下載最新檔案
            dst_path = os.path.join(download_dir, f"{pre_no}_{latest['file_name']}")
            shutil.copy2(latest['file_path'], dst_path)
            # 記錄
            results.append({
                'pre_examine_no': pre_no,
                'file_name': latest['file_name'],
                'last_modified': pd.to_datetime(latest['last_modified'], unit='s')
            })
        else:
            # 有資料夾但沒找到相關檔案
            results.append({
                'pre_examine_no': pre_no,
                'file_name': '無',
                'last_modified': None
            })
    else:
        # 若資料夾不存在也記錄
        results.append({
            'pre_examine_no': pre_no,
            'file_name': '無',
            'last_modified': None
        })

# 轉成DataFrame
result_df = pd.DataFrame(results)

# 輸出成Excel
result_df.to_excel('審查意見檔案彙整.xlsx', index=False)
print('彙整完成，已輸出Excel檔案。')
