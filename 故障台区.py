import requests
import json
import time
import pandas as pd
from datetime import datetime
import os
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# ==================== 1. 配置区 ====================

EXCEL_FILE = "power_failure_history.xlsx"
SHEET_LATEST = "Latest_Data"  # 存储最新全量数据（用于下次对比）
SHEET_DIFF = "Differences_Log"  # 存储差异记录

# 定义列映射: (Excel表头, JSON字段名)
COLUMNS_MAP = [
    ("序号", None),
    ("地市局", "cityOrgName"),
    ("区县局", "disOrgName"),
    ("供电所", "townOrgName"),
    ("变电站编号", "subsNo"),
    ("变电站名称", "subsName"),
    ("线路编号", "lineNo"),
    ("线路名称", "lineName"),
    ("停电事件编号", "pfEventNo"),
    ("停电类型", "pfTypeCodeDesc"),
    ("是否转供", "pfTrFlagDesc"),
    ("停电开始时间", "pfStartTime"),
    ("停电结束时间", "pfEndTime"),
    ("停电时长(h)", "pfDuration"),
    ("影响用户数", "powerFailNum"),
    ("预计电量损失", "prospectPowerLoss"),
]

# 用于标识唯一行的关键字段 (防止数据顺序变化导致误判)
KEY_FIELDS = ["pfEventNo"]

TARGET_URL = 'http://10.40.229.190/ams-collect1-pompowoutmotana-service/member/V1/tqPowerFailure/getTqPowerFailureList'
CANDIDATE_URLS = [
    TARGET_URL,
    'http://10.40.229.191/ams-collect1-pompowoutmotana-service/member/V1/linePowerFailure/getLinePowerFailureList'
]

headers = {
    'Accept': 'application/json, text/plain, */*',
    'Content-Type': 'application/json',
    'Cookie': 'Authorization=eyJhbGciOiJSUzI1NiJ9.eyJqdGkiOiJ6MkotUUg5ckdDWjJwTWRObF9INDI0WlowbUFDR3pRdE10aG8wYjF0dHRFIiwiand0TWF4VmFsaWRpdHlUaW1lIjozMH0.Uxydl5rg5U4QNJYPabYQllZSi7jKKTSGFgOTUMxu-CTguoixqRBc7NCd_rT3G4YoULdQQxaGEy0XzL0mN8Gq5mhT1_3lnIKmsiG7YL8kezFPAasgNOB_qXX-8nC5SbTAm42zxdHjXqoFACPyxN739fz1fqryUL98CGMzKhMQb58pF421bwG8uGhvrmirUVX-NOAMULc3XZ5YVIvYexLabIEGFNf4R92H4d8TeogPIxUOQJdDNQ0fPiCnGjqHFApZdnsqCZCnWoo4yGxfhaAnZSbYiJUCkAqKK1Ocjk3nDbu6BsL7tnAv7FRAbQvNDsoUPawlAu4eCiALgTSSxzFSlA',  # 认证信息（Token），实际使用时需替换为真实值
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.6261.95 Safari/537.36',
    'X-Requested-With': 'XMLHttpRequest'
}

payload = {
    "supplyOrgNo": "050111",
    "orgTypeCode": "4",
    "pfTypeCode": "",
    "pfTrFlag": "",
    "pfDuration": "",
    "subsName": "",
    "lineNo": "",
    "lineName": "",
    "subsNo": "",
    "pfStartTime": "2026-11-13 00:00:00",
    "pfEndTime": "2026-11-16 23:59:59",
    "pageNo": 1,
    "pageSize": 50,
    "total": 0,
    "type": "1"
}


# ==================== 2. 工具函数 ====================

def init_session():
    """初始化带重试机制的 Session"""
    session = requests.Session()
    retries = Retry(total=3, backoff_factor=1, status_forcelist=[502, 503, 504])
    session.mount('http://', HTTPAdapter(max_retries=retries))
    return session


session = init_session()


def fetch_data_with_fallback(urls, headers, payload):
    """容错获取数据"""
    last_error = None
    for url in urls:
        try:
            print(f"正在尝试请求: {url}")
            response = session.post(url, headers=headers, json=payload, timeout=30)
            if response.status_code == 200:
                return response.json()
            else:
                last_error = f"HTTP {response.status_code}: {response.text[:100]}"
                print(f"请求失败 ({last_error})，尝试下一个...")
        except Exception as e:
            last_error = str(e)
            print(f"异常发生 ({last_error})，尝试下一个...")
            continue
    raise Exception(f"所有已知URL均不可用，最后错误: {last_error}")


def convert_to_dataframe(records):
    """将 JSON 列表转换为 DataFrame"""
    if not records:
        return pd.DataFrame(columns=[c[0] for c in COLUMNS_MAP])

    df_data = []
    for i, row in enumerate(records, 1):
        new_row = {}
        for display_name, field_key in COLUMNS_MAP:
            if field_key is None:
                new_row[display_name] = i
            else:
                val = row.get(field_key, "")
                new_row[display_name] = "" if val is None else str(val)
        df_data.append(new_row)

    return pd.DataFrame(df_data, columns=[c[0] for c in COLUMNS_MAP])


def load_previous_data():
    """从 Excel 加载上一次的全量数据"""
    if os.path.exists(EXCEL_FILE):
        try:
            xl = pd.ExcelFile(EXCEL_FILE)
            if SHEET_LATEST in xl.sheet_names:
                prev_df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_LATEST)
                prev_df = prev_df.reindex(columns=[c[0] for c in COLUMNS_MAP]).fillna("")
                for col in prev_df.columns:
                    prev_df[col] = prev_df[col].astype(str)
                return prev_df
        except Exception as e:
            print(f"[WARN] Failed to load previous data: {e}")
    return None


def save_to_excel(latest_df, diff_df=None):
    """保存最新数据和差异数据到 Excel"""
    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        latest_df.to_excel(writer, index=False, sheet_name=SHEET_LATEST)
        if diff_df is not None and not diff_df.empty:
            diff_df.to_excel(writer, index=False, sheet_name=SHEET_DIFF)


def compare_data(prev_df, curr_df):
    """
    对比两个 DataFrame，返回差异 DataFrame
    """
    if prev_df is None or prev_df.empty:
        if curr_df.empty:
            return pd.DataFrame()
        diff = curr_df.copy()
        diff['Change_Type'] = 'NEW'
        return diff

    if curr_df.empty:
        diff = prev_df.copy()
        diff['Change_Type'] = 'DELETED'
        return diff

    key_cols = [col for name, col in COLUMNS_MAP if col in KEY_FIELDS]
    if not key_cols:
        key_cols = ['序号']

    curr_keys = curr_df[key_cols].drop_duplicates()
    prev_keys = prev_df[key_cols].drop_duplicates()

    new_keys = curr_keys[~curr_keys.set_index(key_cols).index.isin(prev_keys.set_index(key_cols).index)]
    deleted_keys = prev_keys[~prev_keys.set_index(key_cols).index.isin(curr_keys.set_index(key_cols).index)]

    diffs = []

    if not new_keys.empty:
        added_rows = curr_df[curr_df.set_index(key_cols).index.isin(new_keys.set_index(key_cols).index)].copy()
        added_rows['Change_Type'] = 'ADDED'
        diffs.append(added_rows)

    if not deleted_keys.empty:
        removed_rows = prev_df[prev_df.set_index(key_cols).index.isin(deleted_keys.set_index(key_cols).index)].copy()
        removed_rows['Change_Type'] = 'REMOVED'
        diffs.append(removed_rows)

    common_keys = curr_keys[curr_keys.set_index(key_cols).index.isin(prev_keys.set_index(key_cols).index)]
    if not common_keys.empty:
        merged = curr_df.merge(prev_df, on=key_cols, suffixes=('_New', '_Old'))
        value_cols = [c for c in curr_df.columns if c not in key_cols and c != '序号']
        change_mask = pd.Series(False, index=merged.index)

        for col in value_cols:
            new_col = f"{col}_New"
            old_col = f"{col}_Old"
            change_mask |= (merged[new_col].astype(str) != merged[old_col].astype(str))

        changed_rows = merged[change_mask]
        if not changed_rows.empty:
            change_details = []
            for idx, row in changed_rows.iterrows():
                for col in value_cols:
                    new_val = row[f"{col}_New"]
                    old_val = row[f"{col}_Old"]
                    if str(new_val) != str(old_val):
                        detail_row = row[key_cols].to_dict()
                        detail_row['Field_Name'] = col
                        detail_row['Old_Value'] = old_val
                        detail_row['New_Value'] = new_val
                        detail_row['Change_Type'] = 'MODIFIED'
                        change_details.append(detail_row)

            if change_details:
                mod_df = pd.DataFrame(change_details)
                cols_order = key_cols + ['Field_Name', 'Old_Value', 'New_Value', 'Change_Type']
                diffs.append(mod_df)

    if not diffs:
        return pd.DataFrame()

    return pd.concat(diffs, ignore_index=True)


# ==================== 3. 主流程 =================

def main():
    print("开始执行数据监控任务...")

    # 1. 获取最新数据
    try:
        raw_data = fetch_data_with_fallback(CANDIDATE_URLS, headers, payload)
    except Exception as e:
        print(f"最终请求失败: {e}")
        return

    # 解析数据
    if isinstance(raw_data, dict):
        records = raw_data.get('data') or raw_data.get('list') or raw_data.get('rows') or []
    elif isinstance(raw_data, list):
        records = raw_data
    else:
        records = []

    if not records:
        print("没有数据，完整响应：")
        print(json.dumps(raw_data, ensure_ascii=False, indent=2))
        return

    print(f"共获取到 {len(records)} 条记录\n")

    # 2. 转换为 DataFrame
    curr_df = convert_to_dataframe(records)

    # 3. 加载上次数据
    prev_df = load_previous_data()

    # 4. 对比数据
    diff_df = compare_data(prev_df, curr_df)

    # 5. 处理结果
    magss = []  # 初始化变量 magss，用于存储差异数据

    if diff_df.empty:
        print("✅ 无差异，数据保持一致。")
    else:
        print(f"⚠️  检测到差异！共 {len(diff_df)} 条变更。")
        print("-" * 30)
        # 打印差异详情
        print(diff_df.to_string(index=False))
        print("-" * 30)

        # 将差异数据存入变量 magss (转为字典列表，方便后续使用)
        magss = diff_df.to_dict(orient='records')

    # 6. 保存到 Excel
    try:
        save_to_excel(curr_df, diff_df if not diff_df.empty else None)
        print(f"💾 数据已保存至 {EXCEL_FILE}")
    except PermissionError:
        print("[ERROR] 无法写入 Excel 文件，请检查文件是否被占用。")
    except Exception as e:
        print(f"[ERROR] 保存 Excel 失败: {e}")

    # 7. 输出最终结果
    print("\n--- 变量 magss 内容 ---")
    if magss:
        print(json.dumps(magss, ensure_ascii=False, indent=4))
    else:
        print("[] (空列表)")


if __name__ == "__main__":
    main()