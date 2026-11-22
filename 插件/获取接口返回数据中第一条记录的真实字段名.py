# 导入 requests 库，用于发送 HTTP 请求（如 POST/GET）
import requests
# 导入 json 库，用于处理 JSON 格式的数据（如解析、格式化输出）
import json

# 定义目标接口的 URL 地址
url = 'http://10.40.229.190/ams-collect1-pompowoutmotana-service/member/V1/tqPowerFailure/getTqPowerFailureList'

# 定义请求头 headers，包含认证信息、内容类型等
headers = {
    # 接受的内容类型，表示支持 JSON、纯文本等
    'Accept': 'application/json, text/plain, */*',
    # 接受的语言环境，优先中文
    'Accept-Language': 'zh-CN,zh;q=0.9',
    # 认证令牌，用于身份验证（需替换为真实有效的 Token）
    'Cookie': 'Authorization=eyJhbGciOiJSUzI1NiJ9.eyJqdGkiOiJ6MkotUUg5ckdDWjJwTWRObF9INDI0WlowbUFDR3pRdE10aG8wYjF0dHRFIiwiand0TWF4VmFsaWRpdHlUaW1lIjozMH0.Uxydl5rg5U4QNJYPabYQllZSi7jKKTSGFgOTUMxu-CTguoixqRBc7NCd_rT3G4YoULdQQxaGEy0XzL0mN8Gq5mhT1_3lnIKmsiG7YL8kezFPAasgNOB_qXX-8nC5SbTAm42zxdHjXqoFACPyxN739fz1fqryUL98CGMzKhMQb58pF421bwG8uGhvrmirUVX-NOAMULc3XZ5YVIvYexLabIEGFNf4R92H4d8TeogPIxUOQJdDNQ0fPiCnGjqHFApZdnsqCZCnWoo4yGxfhaAnZSbYiJUCkAqKK1Ocjk3nDbu6BsL7tnAv7FRAbQvNDsoUPawlAu4eCiALgTSSxzFSlA',  # 认证信息（Token），实际使用时需替换为真实值
    # 发送的内容类型为 JSON
    'Content-Type': 'application/json',
    # 模拟浏览器 User-Agent，避免被服务器拦截
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.6261.95 Safari/537.36',
    # 标识这是一个 XMLHttpRequest（AJAX）请求
    'X-Requested-With': 'XMLHttpRequest'
}

# 定义请求体 payload，即 POST 请求要发送的 JSON 数据
payload = {
    "supplyOrgNo": "050111",  # 供电单位编号
    "orgTypeCode": "4",  # 机构类型代码
    "pfTypeCode": "2",  # 停电类型代码
    # "pfTrFlag": "",  # 是否转供标志
    # "pfDuration": "",  # 停电时长
    # "subsName": "",  # 变电站名称
    # "lineNo": "",  # 线路编号
    # "lineName": "",  # 线路名称
    # "subsNo": "",  # 变电站编号
    "pfStartTime": "2026-11-13 00:00:00",  # 停电开始时间
    "pfEndTime": "2026-11-16 23:59:59",  # 停电结束时间
    "pageNo": 1,  # 页码
    "pageSize": 50,  # 每页条数
    "total": 0,  # 总条数
    "type": "1"  # 类型
}

# 发送 POST 请求，将 url、headers、payload 和超时时间（30秒）传入
response = requests.post(url, headers=headers, json=payload, timeout=30)

# 判断 HTTP 响应状态码是否为 200（成功）
if response.status_code == 200:
    # 将响应体解析为 Python 字典（JSON 转 dict）
    data = response.json()

    # 尝试从常见的键名中获取记录列表，防止接口返回结构不同
    # 依次尝试 'data'、'list'、'rows'，都找不到则默认为空列表 []
    records = data.get('data') or data.get('list') or data.get('rows') or []

    # 如果记录列表为空，说明没有获取到任何数据
    if not records:
        print("未获取到任何数据")
    else:
        # 初始化变量：用于存储最完整的数据样本
        best_item = None
        # 初始化变量：记录当前找到的最大非空字段数量，初始值为 -1
        max_non_null_count = -1

        # 遍历所有记录，逐条检查
        for item in records:
            # 统计该对象中非 null 的字段数量
            # sum(1 for v in item.values() if v is not None) 是生成器表达式，
            # 遍历每个值 v，如果 v 不是 None 就计 1，最后求和
            non_null_count = sum(1 for v in item.values() if v is not None)

            # 如果当前对象的非空字段数大于之前记录的最大值
            if non_null_count > max_non_null_count:
                # 更新最大非空字段数
                max_non_null_count = non_null_count
                # 更新最完整的数据样本为当前对象
                best_item = item

        # 如果找到了有效样本（best_item 不为 None 且非空字段数大于 0）
        if best_item and max_non_null_count > 0:
            # 打印提示信息，显示找到的样本包含多少个非空字段
            print(f"✅ 找到最完整的数据样本（包含 {max_non_null_count} 个非空字段）：")

            # 打印换行和标题
            print("\n【真实字段名列表】:")
            # 获取并打印该样本的所有字段名（键），以列表形式展示
            print(list(best_item.keys()))

            # 打印换行和标题
            print("\n【该条数据的完整内容】:")
            # 将该样本格式化为 JSON 字符串并打印
            # ensure_ascii=False 确保中文正常显示（不转义为 \uXXXX）
            # indent=4 设置缩进为 4 个空格，使输出更易读
            print(json.dumps(best_item, ensure_ascii=False, indent=4))
        else:
            # 如果所有数据的非空字段数都为 0，说明数据无效
            print("⚠️ 所有数据均为空值，请检查 payload 参数是否足够精确，或接口是否需要其他必填字段。")
else:
    # 如果状态码不是 200，说明请求失败
    # 打印状态码和错误详情，方便排查问题
    print(f"请求失败: HTTP {response.status_code}")
    print(response.text)