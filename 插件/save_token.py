# save_cookie.py
import time
from selenium import webdriver
from selenium.common.exceptions import NoSuchWindowException  # 【新增】导入窗口异常类
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.support.ui import WebDriverWait

LOGIN_URL = "https://www.bilibili.com/video/BV1tcsV6jEfx/"
TOKEN_FILE = "auth_token.txt"

edge_options = EdgeOptions()
# ⚠️ 确保这里的 msedge.exe 路径是你电脑真实的
edge_options.binary_location = r"C:\Users\Administrator\AppData\Local\Microsoft\Edge\Bin\edgex.exe"


def get_and_save_cookie():
    driver = None
    try:
        driver = webdriver.Edge(
            service=EdgeService(EdgeChromiumDriverManager().install()),
            options=edge_options
        )
        driver.get(LOGIN_URL)

        print("🌐 Edge 浏览器已打开，请在浏览器中手动完成登录...")
        print("⚠️ 注意：在登录成功前，请【不要】手动关闭浏览器窗口！")

        # 【核心优化】使用显式等待，并增加窗口存活检查
        # 最多等待 300 秒，每隔 1 秒检查一次 SESSDATA 是否存在
        WebDriverWait(driver, 300, poll_frequency=1.0).until(
            lambda d: any(c['name'] == 'SESSDATA' and c['value'] for c in d.get_cookies())
        )

        print("✅ 检测到登录成功，正在提取 Cookie...")
        time.sleep(1)

        cookies = driver.get_cookies()
        if not cookies:
            print("❌ 警告: 未获取到任何 Cookie，请检查是否真的登录成功。")
            return

        cookie_str = "; ".join([f"{c['name']}={c['value']}" for c in cookies])
        with open(TOKEN_FILE, "w", encoding="utf-8") as f:
            f.write(cookie_str)
        print(f"🎉 成功: Cookie 已保存至 '{TOKEN_FILE}'")
        print(f"📄 预览: {cookie_str[:50]}...")

    # 【核心优化】专门捕获“窗口被关闭”的异常
    except NoSuchWindowException:
        print("❌ 错误: 浏览器窗口在登录过程中被意外关闭了！")
        print("💡 提示: 请重新运行脚本，并在登录完成前保持浏览器开启。")

    except Exception as e:
        print(f"❌ 获取 Cookie 失败: {e}")

    finally:
        if driver:
            try:
                driver.quit()
                print("🔒 浏览器已正常关闭。")
            except Exception:
                pass  # 如果浏览器已经被手动关了，忽略关闭时的报错


if __name__ == "__main__":
    get_and_save_cookie()