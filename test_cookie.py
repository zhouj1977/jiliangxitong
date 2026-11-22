# test_cookie.py
import requests
from 插件.headers_config import get_headers

def test_website_cookie():
    print("🔍 正在读取本地 Cookie 并发起请求...")

    try:
        headers = get_headers()

        # ==========================================
        # ⚠️ 【核心配置区】填入 URL 和对应的登录特征
        # ==========================================
        url_list = [
            ("https://account.sogou.com/home/login"),  # 1. 首选
            # ("https://www.bilibili.com/video/BV1tcsV6jEfx/", "退出"),  # 2. 备用
        ]
        # ==========================================

        data_found = False

        for url, login_feature in url_list:
            try:
                print(f"🌐 正在尝试: {url}")
                response = requests.get(url, headers=headers, timeout=10)

                # 【新增 1】先检查 HTTP 状态码
                if response.status_code != 200:
                    print(f"⚠️ 请求返回异常状态码: {response.status_code}，尝试下一个...")
                    continue

                # 【新增 2】检查返回内容是否为空
                if not response.text:
                    print("⚠️ 服务器返回了空内容，尝试下一个...")
                    continue

                # 【新增 3】安全地解析 JSON（如果目标网站返回的是 JSON）
                # 如果目标网站返回的是普通 HTML 网页，直接用 response.text 即可，无需 json()
                # 这里假设你是用页面特征检测，所以直接搜文本：
                html_content = response.text

                if login_feature in html_content:
                    print(f"🎉 验证成功！Cookie 有效！")
                    print(f"✅ 在页面中成功找到了特征: '{login_feature}'")
                    data_found = True
                    break  # 成功，跳出循环
                else:
                    print(f"⚠️ 页面打开成功，但未找到特征 '{login_feature}'，尝试下一个...")
                    # 调试用：打印前 100 个字符，看看服务器到底返回了什么
                    print(f"📄 网页内容预览: {html_content[:100]}...")

            except requests.exceptions.RequestException as e:
                # 【新增 4】捕获网络层面的异常（超时、断网等）
                print(f"❌ 该 URL 网络请求失败 ({type(e).__name__})，尝试下一个...")
                continue

        if not data_found:
            print("❌ 验证失败：所有 URL 均无法确认登录状态，Cookie 可能已失效。")

    except FileNotFoundError as e:
        print(f"❌ {e}")
    except Exception as e:
        print(f"❌ 发生未知错误: {e}")


if __name__ == "__main__":
    test_website_cookie()