"""
停电事件数据监控脚本
功能：从接口拉取最新数据 -> 与上次全量对比 -> 记录差异 -> 保存到 Excel
作者：AI助手优化版
"""

# ==================== 导入依赖库 ====================

import os  # 用于操作文件路径、判断文件是否存在
import json  # 用于把 Python 字典/列表转成 JSON 字符串（打印差异时用）
import time  # 用于计算脚本运行耗时
import logging  # 用于写日志（比 print 更专业，能记录时间、级别）
from datetime import datetime, timedelta  # 用于获取当前时间，给差异记录打时间戳
from typing import Optional, Dict, Any, List  # 类型注解，让代码更规范、IDE 提示更友好

import requests  # 用于发送 HTTP 请求（POST 请求接口）
import pandas as pd  # 用于数据处理，把数据变成表格（DataFrame），方便对比
from requests.adapters import HTTPAdapter  # 用于给 requests 配置连接池和重试策略
from urllib3.util.retry import Retry  # 用于定义重试规则（比如失败重试3次）

# ==================== 1. 配置区（改这里就行，不用动下面的逻辑） ====================

# Excel 文件名，数据会存在这个文件里
EXCEL_FILE = "power_failure_history.xlsx"

# Sheet 页名称：存最新全量数据（下次运行拿来对比用）
SHEET_LATEST = "Latest_Data"

# Sheet 页名称：存本次检测到的差异
SHEET_DIFF = "Differences_Log"

# Sheet 页名称：累积所有历史差异（每次的差异都追加，不覆盖）
SHEET_ALL_DIFF = "All_Diff_History"

# 列映射：左边是 Excel 表头显示的名字，右边是接口返回的 JSON 字段名
# 如果右边是 None，说明这列不需要从接口取（比如"序号"我们自己生成）
COLUMNS_MAP = [
    ("序号", None),  # 序号，自己数
    # ("地市局", "cityOrgName"),  # 地市局名称
    ("区县局", "disOrgName"),  # 区县局名称
    ("供电所", "townOrgName"),  # 供电所名称
    # ("变电站编号", "subsNo"),  # 变电站编号
    # ("变电站名称", "subsName"),  # 变电站名称
    # ("线路编号", "lineNo"),  # 线路编号
    ("线路名称", "lineName"),  # 线路名称
    ("停电事件编号", "pfEventNo"),  # 停电事件编号（关键字段，用来唯一标识一条记录）
    ("停电类型", "pfTypeCodeDesc"),  # 停电类型描述
    # ("是否转供", "pfTrFlagDesc"),  # 是否转供描述
    ("停电开始时间", "pfStartTime"),  # 停电开始时间
    ("停电结束时间", "pfEndTime"),  # 停电结束时间
    ("停电时长(h)", "pfDuration"),  # 停电时长（小时）
    # ("影响用户数", "powerFailNum"),  # 影响用户数
    # ("预计电量损失", "prospectPowerLoss"),  # 预计电量损失
]

# 关键字段：用哪些字段来判断"两条记录是不是同一条"
# 这里用停电事件编号，因为它是唯一的
KEY_FIELDS = ["pfEventNo"]   #pfEventNo=停电事件编号

# 请求头：模拟浏览器访问，告诉服务器我们发送的是 JSON 数据
HEADERS = {
    'Accept': 'application/json, text/plain, */*',  # 告诉服务器我们接受什么格式的响应
    'Content-Type': 'application/json',  # 告诉服务器我们发送的是 JSON
    'Cookie': 'Authorization=eyJhbGciOiJSUzI1NiJ9.eyJqdGkiOiJ6MkotUUg5ckdDWjJwTWRObF9INDI0WlowbUFDR3pRdE10aG8wYjF0dHRFIiwiand0TWF4VmFsaWRpdHlUaW1lIjozMH0.Uxydl5rg5U4QNJYPabYQllZSi7jKKTSGFgOTUMxu-CTguoixqRBc7NCd_rT3G4YoULdQQxaGEy0XzL0mN8Gq5mhT1_3lnIKmsiG7YL8kezFPAasgNOB_qXX-8nC5SbTAm42zxdHjXqoFACPyxN739fz1fqryUL98CGMzKhMQb58pF421bwG8uGhvrmirUVX-NOAMULc3XZ5YVIvYexLabIEGFNf4R92H4d8TeogPIxUOQJdDNQ0fPiCnGjqHFApZdnsqCZCnWoo4yGxfhaAnZSbYiJUCkAqKK1Ocjk3nDbu6BsL7tnAv7FRAbQvNDsoUPawlAu4eCiALgTSSxzFSlA',  # 认证信息（Token），实际使用时需替换为真实值
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                  '(KHTML, like Gecko) Chrome/122.0.6261.95 Safari/537.36',  # 模拟浏览器标识
    'X-Requested-With': 'XMLHttpRequest',  # 告诉服务器这是 AJAX 请求
}

# 请求参数：发给接口的数据，相当于在网页上填的查询条件
PAYLOAD = {
    "supplyOrgNo": "050111",  # 供电单位编号
    "orgTypeCode": "4",  # 单位类型
    "pfTypeCode": "2",  # 停电类型
    "pfTrFlag": "",  # 是否转供（空=全部）
    "pfDuration": "",  # 停电时长筛选（空=全部）
    "subsName": "",  # 变电站名称筛选
    "lineNo": "",  # 线路编号筛选
    "lineName": "",  # 线路名称筛选
    "subsNo": "",  # 变电站编号筛选
    # "pfStartTime": "2026-11-15 00:00:00",  # 查询开始时间
    # "pfEndTime": "2026-11-18 23:59:59",  # 查询结束时间
    "pfStartTime": (datetime.now() - timedelta(days=3)).strftime("%Y-%m-%d 00:00:00"),             # 3天前零点
    "pfEndTime": datetime.now().strftime("%Y-%m-%d 23:59:59"),                 # 今天23:59:59
    "pageNo": 1,  # 第几页
    "pageSize": 1000,  # 每页多少条
    "total": 0,  # 总数（接口可能需要）
    "type": "1",  # 类型参数
}

# 接口地址列表：主地址 + 备用地址，主地址挂了自动试备用
URL_LIST = [
    'http://10.40.229.190/ams-collect1-pompowoutmotana-service/member/V1/tqPowerFailure/getTqPowerFailureList',
    'http://10.40.229.191/ams-collect1-pompowoutmotana-service/member/V1/linePowerFailure/getLinePowerFailureList',
]

# 请求超时时间（秒）：超过这个时间没响应就放弃
REQUEST_TIMEOUT = 30

# 最大重试次数：请求失败最多重试几次
MAX_RETRIES = 3

# 重试间隔因子：每次重试等待的时间会递增（1秒、2秒、4秒...）
BACKOFF_FACTOR = 1

# 哪些 HTTP 状态码触发重试（502=网关错误, 503=服务不可用, 504=网关超时）
STATUS_FORCELIST = (502, 503, 504)

# ==================== 2. 日志配置 ====================

# 日志格式：时间 [级别] 消息
LOG_FORMAT = "%(asctime)s [%(levelname)s] %(message)s"

# 时间格式
DATE_FORMAT = "%Y-%m-%d %H:%M:%S"

# 配置日志：同时输出到屏幕和文件
logging.basicConfig(
    level=logging.INFO,  # 日志级别：INFO 及以上才记录（DEBUG < INFO < WARNING < ERROR）
    format=LOG_FORMAT,  # 使用上面定义的格式
    datefmt=DATE_FORMAT,  # 时间格式
    handlers=[
        logging.StreamHandler(),  # 输出到屏幕（控制台）
        logging.FileHandler("power_failure_monitor.log", encoding="utf-8"),  # 同时写入日志文件
    ],
)

# 创建一个 logger 对象，后面用 logger.info() / logger.error() 来记录
logger = logging.getLogger(__name__)


# ==================== 3. 工具函数 ====================

def init_session() -> requests.Session:
    """
    初始化一个带重试机制的 requests Session 对象
    Session 可以复用 TCP 连接，比每次请求都新建连接更快
    """
    # 创建一个 Session 对象（相当于一个浏览器会话）
    session = requests.Session()

    # 创建重试策略对象
    retry = Retry(
        total=MAX_RETRIES,  # 最多重试3次
        backoff_factor=BACKOFF_FACTOR,  # 重试间隔：1秒、2秒、4秒递增
        status_forcelist=STATUS_FORCELIST,  # 遇到这些HTTP状态码就重试
        allowed_methods=("POST",),  # 只对 POST 请求重试（POST 一般不是幂等的，但这里接口是查询所以可以）
        raise_on_status=False,  # 重试完还是失败的话，不抛异常，返回最后一次的响应
    )

    # 创建一个 HTTP 适配器，绑定重试策略
    adapter = HTTPAdapter(max_retries=retry)

    # 把适配器挂载到 http 和 https 两种协议上
    session.mount("http://", adapter)
    session.mount("https://", adapter)

    # 返回配置好的 session
    return session


# 创建一个全局 session 对象，整个脚本共用这一个
session = init_session()


def fetch_data(urls: List[str], headers: Dict, payload: Dict) -> Optional[Dict]:
    """
    容错获取数据：依次尝试 URL 列表中的每个地址
    只要有一个成功就返回结果，全部失败就抛异常
    """
    # 记录最后一个错误信息（用于全部失败时展示）
    last_error = None

    # 遍历所有 URL
    for url in urls:
        try:
            # 记录日志：正在请求哪个地址
            logger.info("正在请求: %s", url)

            # 发送 POST 请求，传入 headers 和 JSON 格式的 payload
            resp = session.post(url, headers=headers, json=payload, timeout=REQUEST_TIMEOUT)

            # 如果 HTTP 状态码是 200（成功）
            if resp.status_code == 200:
                logger.info("请求成功: %s", url)
                # 把响应内容解析成 JSON（Python 字典），返回
                return resp.json()

            # 如果不是 200，记录错误信息
            last_error = f"HTTP {resp.status_code}: {resp.text[:200]}"
            logger.warning("请求失败 (%s)，尝试下一个...", last_error)

        except requests.RequestException as e:
            # 如果连接失败、超时等异常
            last_error = str(e)
            logger.warning("请求异常 (%s)，尝试下一个...", last_error)

    # 如果所有 URL 都试过了还是失败，抛出一个连接错误
    raise ConnectionError(f"所有 URL 均不可用，最后错误: {last_error}")


def extract_records(raw: Any) -> List[Dict]:
    """
    从接口返回的 JSON 中提取出记录列表
    因为不同接口返回格式可能不一样，这里做兼容处理
    """
    # 如果返回的直接是个列表，就直接返回
    if isinstance(raw, list):
        return raw

    # 如果返回的是个字典
    if isinstance(raw, dict):
        # 按常见字段名依次尝试提取
        for key in ("data", "list", "rows", "records", "result", "items"):
            val = raw.get(key)
            if isinstance(val, list):
                return val
            # 如果 data 字段是个对象（比如 {list: [...], total: 100}）
            if isinstance(val, dict):
                for inner_key in ("list", "records", "rows", "items"):
                    inner = val.get(inner_key)
                    if isinstance(inner, list):
                        return inner
        # 如果都没找到，返回空列表
        return []

    # 其他情况也返回空列表
    return []


def normalize_value(val: Any) -> str:
    """
    统一字段值格式：
    - None -> 空字符串
    - NaN（pandas 的空值）-> 空字符串
    - 其他 -> 转成字符串并去掉首尾空格
    这样对比时不会因为格式不一致导致误判
    """
    if val is None:
        return ""
    if isinstance(val, float) and pd.isna(val):
        return ""
    return str(val).strip()


def convert_to_dataframe(records: List[Dict]) -> pd.DataFrame:
    """
    把接口返回的 JSON 记录列表，转换成 pandas 的 DataFrame（表格）
    这样后面可以用 pandas 方便地进行数据对比
    """
    # 提取所有列名（Excel 表头）
    columns = [c[0] for c in COLUMNS_MAP]

    # 如果记录为空，返回一个只有表头的空 DataFrame
    if not records:
        return pd.DataFrame(columns=columns)

    # 用来存放每一行数据的列表
    df_data = []

    # 遍历每条记录，i 是序号（从1开始）
    for i, row in enumerate(records, 1):
        new_row = {}  # 一行数据，用字典存
        # 遍历列映射配置
        for display_name, field_key in COLUMNS_MAP:
            if field_key is None:
                # 如果字段是 None（比如"序号"），用我们自己数的 i
                new_row[display_name] = i
            else:
                # 从接口返回的数据中取对应字段的值，用 normalize_value 统一格式
                new_row[display_name] = normalize_value(row.get(field_key))
        # 把这一行加入列表
        df_data.append(new_row)

    # 用列表创建 DataFrame，指定列顺序
    df = pd.DataFrame(df_data, columns=columns)

    # 把关键字段统一转成字符串并去空格（防止数字和字符串比较时不一致）
    for name, field in COLUMNS_MAP:
        if field in KEY_FIELDS:
            df[name] = df[name].astype(str).str.strip()

    return df


def load_previous_data() -> Optional[pd.DataFrame]:
    """
    从 Excel 文件中加载上一次保存的全量数据
    用于和本次新数据做对比
    """
    # 如果 Excel 文件不存在，返回 None
    if not os.path.exists(EXCEL_FILE):
        return None

    try:
        # 打开 Excel 文件，读取所有 Sheet 名
        xl = pd.ExcelFile(EXCEL_FILE)

        # 如果 Latest_Data 这个 Sheet 不存在，返回 None
        if SHEET_LATEST not in xl.sheet_names:
            return None

        # 读取 Latest_Data Sheet 的数据
        prev_df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_LATEST)

        # 确保列的顺序和当前配置一致（防止列顺序变化导致对比出错）
        prev_df = prev_df.reindex(columns=[c[0] for c in COLUMNS_MAP]).fillna("")

        # 把所有列都转成字符串并去空格（统一格式）
        for col in prev_df.columns:
            prev_df[col] = prev_df[col].astype(str).str.strip()

        return prev_df

    except Exception as e:
        # 如果读取失败（比如文件被占用），记录警告并返回 None
        logger.warning("加载历史数据失败: %s", e)
        return None


def compare_data(prev_df: Optional[pd.DataFrame],
                 curr_df: pd.DataFrame) -> pd.DataFrame:
    """
    核心函数：对比新旧数据，找出差异
    返回差异 DataFrame，包含变更类型（ADDED / REMOVED / MODIFIED）
    """
    # 找出关键字段对应的列名（用于匹配新旧记录）
    key_col_names = [name for name, field in COLUMNS_MAP if field in KEY_FIELDS]
    if not key_col_names:
        key_col_names = ["序号"]  # 兜底：如果没有关键字段，用序号

    # 情况1：没有历史数据（首次运行）
    if prev_df is None or prev_df.empty:
        if curr_df.empty:
            return pd.DataFrame()  # 新数据也是空的，没差异
        diff = curr_df.copy()  # 复制当前数据
        diff["Change_Type"] = "ADDED"  # 全部标记为新增
        return diff

    # 情况2：有历史数据，但新数据是空的（说明所有记录都被删了）
    if curr_df.empty:
        diff = prev_df.copy()
        diff["Change_Type"] = "REMOVED"  # 全部标记为删除
        return diff

    # 提取新旧数据的关键字段，去重后转成索引（方便快速查找）
    prev_keys = prev_df[key_col_names].drop_duplicates().set_index(key_col_names).index
    curr_keys = curr_df[key_col_names].drop_duplicates().set_index(key_col_names).index

    # 计算新增的键：在新数据里有，但旧数据里没有
    new_keys = curr_keys.difference(prev_keys)

    # 计算删除的键：在旧数据里有，但新数据里没有
    removed_keys = prev_keys.difference(curr_keys)

    # 计算共同的键：新旧数据里都有（需要检查是否有字段修改）
    common_keys = curr_keys.intersection(prev_keys)

    # 用来存放所有差异的列表
    diffs = []

    # ---- 处理新增的记录 ----
    if not new_keys.empty:
        # 找出新数据中关键字段在 new_keys 里的行
        mask = curr_df.set_index(key_col_names).index.isin(new_keys)
        added = curr_df[mask].copy()
        added["Change_Type"] = "ADDED"  # 标记为新增
        diffs.append(added)

    # ---- 处理删除的记录 ----
    if not removed_keys.empty:
        mask = prev_df.set_index(key_col_names).index.isin(removed_keys)
        removed = prev_df[mask].copy()
        removed["Change_Type"] = "REMOVED"  # 标记为删除
        diffs.append(removed)

    # ---- 处理修改的记录 ----
    if not common_keys.empty:
        # 找出需要对比的值列（排除关键字段和序号）
        value_cols = [c for c in curr_df.columns if c not in key_col_names and c != "序号"]

        # 把新旧数据按关键字段合并（类似 SQL 的 JOIN）
        merged = curr_df.merge(prev_df, on=key_col_names, suffixes=("_New", "_Old"))

        # 逐行逐字段对比
        change_details = []
        for _, row in merged.iterrows():
            # 取出关键字段的值（用于标识是哪条记录）
            key_vals = {k: row[k] for k in key_col_names}

            # 遍历每个值字段
            for col in value_cols:
                new_val = row[f"{col}_New"]  # 新值
                old_val = row[f"{col}_Old"]  # 旧值

                # 如果新值不等于旧值，说明这个字段被修改了
                if str(new_val) != str(old_val):
                    change_details.append({
                        **key_vals,  # 关键字段
                        "Field_Name": col,  # 哪个字段被改了
                        "Old_Value": old_val,  # 改之前的值
                        "New_Value": new_val,  # 改之后的值
                        "Change_Type": "MODIFIED",  # 标记为修改
                    })

        # 如果有修改记录，转成 DataFrame 加入差异列表
        if change_details:
            diffs.append(pd.DataFrame(change_details))

    # 如果没有任何差异，返回空 DataFrame
    if not diffs:
        return pd.DataFrame()

    # 把所有差异拼接成一个完整的 DataFrame 返回
    return pd.concat(diffs, ignore_index=True)


def save_to_excel(latest_df: pd.DataFrame,
                  diff_df: Optional[pd.DataFrame] = None) -> None:
    """
    保存数据到 Excel：
    - Latest_Data Sheet：覆盖写入最新全量数据
    - Differences_Log Sheet：覆盖写入本次差异
    - All_Diff_History Sheet：追加写入本次差异（保留历史）
    """
    # 确保文件所在目录存在（如果路径里有目录的话）
    os.makedirs(os.path.dirname(os.path.abspath(EXCEL_FILE)) or ".", exist_ok=True)

    # 读取已有的 Excel 中的所有 Sheet（为了保留历史差异）
    existing_sheets = {}
    if os.path.exists(EXCEL_FILE):
        try:
            with pd.ExcelFile(EXCEL_FILE) as xl:
                for sheet in xl.sheet_names:
                    existing_sheets[sheet] = pd.read_excel(xl, sheet_name=sheet)
        except Exception as e:
            logger.warning("读取已有 Excel 失败，将覆盖: %s", e)

    # 处理累积差异历史
    all_diff = existing_sheets.get(SHEET_ALL_DIFF, pd.DataFrame())
    if diff_df is not None and not diff_df.empty:
        diff_df = diff_df.copy()
        # 在差异数据最前面插入一列：检测时间
        diff_df.insert(0, "Detect_Time", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        # 把本次差异追加到历史差异后面
        all_diff = pd.concat([all_diff, diff_df], ignore_index=True)

    # 写入 Excel（mode="w" 表示覆盖整个文件）
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="w") as writer:
        # 写入最新全量数据
        latest_df.to_excel(writer, index=False, sheet_name=SHEET_LATEST)

        # 如果有差异，写入本次差异 Sheet
        if diff_df is not None and not diff_df.empty:
            diff_df.to_excel(writer, index=False, sheet_name=SHEET_DIFF)

        # 如果有累积历史差异，写入历史 Sheet
        if not all_diff.empty:
            all_diff.to_excel(writer, index=False, sheet_name=SHEET_ALL_DIFF)


# ==================== 4. 主流程 ====================

def main():
    """
    主函数：整个脚本的执行入口
    """
    # 记录开始时间（用于计算总耗时）
    start_ts = time.time()

    # 打印分隔线和开始时间
    logger.info("=" * 50)
    logger.info("开始执行数据监控任务: %s",
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    logger.info("=" * 50)

    # ---- 第1步：从接口获取最新数据 ----
    try:
        raw_data = fetch_data(URL_LIST, HEADERS, PAYLOAD)
    except Exception as e:
        logger.error("最终请求失败: %s", e)
        return  # 获取失败就直接结束

    # ---- 第2步：从返回数据中提取记录列表 ----
    records = extract_records(raw_data)
    if not records:
        # 如果没拿到数据，打印原始响应（方便排查接口问题）
        logger.warning("接口返回为空，原始响应: %s",
                       json.dumps(raw_data, ensure_ascii=False, indent=2)[:500])
        return

    logger.info("共获取到 %d 条记录", len(records))

    # ---- 第3步：把记录转成 DataFrame 表格 ----
    curr_df = convert_to_dataframe(records)

    # ---- 第4步：加载上次保存的全量数据 ----
    prev_df = load_previous_data()
    if prev_df is not None:
        logger.info("已加载历史数据: %d 条", len(prev_df))
    else:
        logger.info("无历史数据，首次运行，将全量记录为新增")

    # ---- 第5步：对比新旧数据，找出差异 ----
    diff_df = compare_data(prev_df, curr_df)

    # ---- 第6步：处理结果 ----
    # magss 变量：用来存储差异数据（你的原始代码里有这个变量，保留下来）
    magss = []

    if diff_df.empty:
        # 没有差异
        logger.info("✅ 无差异，数据保持一致。")
    else:
        # 有差异，打印出来
        logger.info("⚠️  检测到差异！共 %d 条变更。", len(diff_df))
        logger.info("-" * 30)
        logger.info("\n%s", diff_df.to_string(index=False))
        logger.info("-" * 30)
        # 把差异数据转成字典列表，存入 magss
        magss = diff_df.to_dict(orient="records")

    # ---- 第7步：保存到 Excel ----
    try:
        save_to_excel(curr_df, diff_df if not diff_df.empty else None)
        logger.info("💾 数据已保存至 %s", os.path.abspath(EXCEL_FILE))
    except PermissionError:
        logger.error("❌ 无法写入 Excel，请检查文件是否被其他程序占用。")
        return
    except Exception as e:
        logger.error("❌ 保存 Excel 失败: %s", e)
        return

    # ---- 第8步：输出 magss 变量内容 ----
    # ---- 第8步：输出 magss 变量内容并保存为指定格式 ----
    logger.info("\n--- 变量 magss 内容 ---")
    magss_formatted = []
    if magss:
        for item in magss:
            line_name = item.get("线路名称", "")
            pf_type = item.get("停电类型", "")
            start_time = item.get("停电开始时间", "")
            end_time = item.get("停电结束时间", "")

            # 去掉秒，只保留到分钟
            if len(start_time) >= 19:
                start_time = start_time[:16]
            if len(end_time) >= 19:
                end_time = end_time[:16]

            formatted_str = f"[{line_name}\\{pf_type}\\停电开始时间: {start_time}\\停电结束时间: {end_time}]"
            magss_formatted.append(formatted_str)

        logger.info("%s", "\n".join(magss_formatted))
    else:
        logger.info("[] (空列表)")

    # ---- 第9步：保存 magss 到 JSON 文件，供 send_msg.py 读取 ----
    save_magss(magss_formatted)

    # 计算并打印总耗时
    elapsed = time.time() - start_ts
    logger.info("任务完成，耗时 %.2f 秒", elapsed)


# ==================== magss 持久化相关 ====================

# magss 数据文件路径（与当前脚本在同一目录下）
MAGSS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "magss_data.json")


def save_magss(magss_list):
    """将 magss 格式化后的字符串列表保存到 JSON 文件"""
    with open(MAGSS_FILE, "w", encoding="utf-8") as f:
        json.dump(magss_list, f, ensure_ascii=False, indent=2)
# 当这个文件被直接运行时（不是被其他文件 import），执行 main() 函数
if __name__ == "__main__":
    main()
