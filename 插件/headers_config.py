# headers_config.py
import os
import time

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TOKEN_FILE = os.path.join(BASE_DIR, "auth_token.txt")

# TOKEN_FILE = "auth_token.txt"

def get_headers():
    """读取本地 Token 生成请求头"""
    if not os.path.exists(TOKEN_FILE):
        raise FileNotFoundError(f"❌ 错误: 未找到 Token 文件 '{TOKEN_FILE}'，请先执行登录获取 Token。")

    with open(TOKEN_FILE, 'r', encoding='utf-8') as f:
        token = f.read().strip()

    if not token:
        raise ValueError(f"❌ 错误: Token 文件内容为空。")

    return {
        'Accept': 'application/json, text/plain, */*',
        'Content-Type': 'application/json',
        'cookie': token,
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                      '(KHTML, like Gecko) Chrome/122.0.6261.95 Safari/537.36',
        'X-Requested-With': 'XMLHttpRequest',
    }


def refresh_and_save_token():
    """
    【核心】模拟从网页获取新 Token 并覆盖写入 TXT 文件
    在实际项目中，这里可以替换为调用登录接口或 Selenium 自动登录逻辑
    """
    print("⚠️ 检测到 Token 失效，正在自动重新获取...")

    # === 模拟获取新 Token 的过程 ===
    # 实际开发中，这里放你的 requests.post("登录接口") 或 Selenium 操作
    time.sleep(1)
    new_token = f"Bearer new_token_{int(time.time())}"
    # ================================

    # 覆盖写入 TXT 文件
    with open(TOKEN_FILE, 'w', encoding='utf-8') as f:
        f.write(new_token)

    print(f"✅ 成功: 新 Token 已获取并保存。")
    return new_token