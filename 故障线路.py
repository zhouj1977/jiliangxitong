import requests
import json

# url = 'http://10.40.229.190/ams-collect1-pompowoutmotana-service/member/V1/linePowerFailure/getLinePowerFailureList'
url = 'http://10.40.229.190/ams-collect1-smageneraltools-service/member/V1/notiCenter/getNotiCenterTreeList'
headers = {
    'Accept': 'application/json, text/plain, */*',
    'Accept-Language': 'zh-CN,zh;q=0.9',
    'Authorization': 'Bearer eyJhbGciOiJSUzI1NiJ9.eyJqdGkiOiJsTTdrdHVrLVlHd3JoMDRDTXJNck5TcEZoeWhHeEVkeDBzNXJEd0h6aHgwIiwiand0TWF4VmFsaWRpdHlUaW1lIjozMH0.OjrxwBHxIX4GwGS7APY6th5akZzt0CJCRRmBW5WoYeq4JbsBYq4VRRqQsO_9ZBOXu-2f5TOmsnHu8KTvAB7RrBgwpKKJyS2QaeR21sh6cVbAtSWFQg5-SVNaefbTAY93j24gzfy-_qD0myNljonoRaFgodjZ3J19LFwHJIIf7re1C79S6WZ9xU0gAN5NlMW_vyI8GrBlepqLIZzJ-MvEVqJLPf3ylclXac7auHm3XZxCRdq3zVlrD6Y9BcEm4PERTeBvt-7m3WJv8XJzP4jbYa_-KlbI5tzBTf0fBt0R-MZhHFQ9Nz6XCHQgUe4y2yrhangJi4t9GEByEwcUEGsfnA',
    'Content-Type': 'application/json',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.6261.95 Safari/537.36',
    'X-Requested-With': 'XMLHttpRequest'
}

payload = {
    "supplyOrgNo": "050111",
    "orgTypeCode": "4",
    "pfTypeCode": "2",
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

response = requests.post(url, headers=headers, json=payload, timeout=10)
data = response.json()

# 提取数据列表
if isinstance(data, dict):
    records = data.get('data') or data.get('list') or data.get('rows') or []
elif isinstance(data, list):
    records = data
else:
    records = []

if not records:
    print("没有数据，完整响应：")
    print(json.dumps(data, ensure_ascii=False, indent=2))
else:
    # 定义表头和对应的字段映射（根据常见命名习惯推测）
    headers_map = [
        ("序号", "index"),
        ("地市局", "cityBureau"),
        ("区县局", "countyBureau"),
        ("供电所", "supplyOrgName"),
        ("变电站编号", "subsNo"),
        ("变电站名称", "subsName"),
        ("线路编号", "lineNo"),
        ("线路名称", "lineName"),
        ("停电事件编号", "eventNo"),
        ("停电类型", "pfTypeName")
    ]

    # 打印表头
    header_str = " | ".join([h[0] for h in headers_map])
    print(header_str)
    print("-" * len(header_str) * 2)

    # 打印数据行
    for i, row in enumerate(records, 1):
        row_data = []
        for display_name, field_key in headers_map:
            if field_key == "index":
                value = str(i)
            else:
                # 尝试获取映射字段，如果没有则尝试获取原始字段名（如lineName）
                value = row.get(field_key, "")
                # 如果值太长（如事件编号），做截断处理
                if isinstance(value, str) and len(value) > 15:
                    value = value[:15] + "..."
            row_data.append(str(value))

        print(" | ".join(row_data))

    print("\n" + " = " * 60)
    print("提示：如果上述字段显示为空，说明接口返回的真实字段名不同。")
    print("第一条数据的真实字段名如下：")
    print(list(records[0].keys()))
