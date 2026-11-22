"""⚠️ 重要提示：此脚本仅适用于 Windows 系统，且需要南网eLink客户端已登录并保持运行状态。"""

# ==================== 导入模块 ====================

# 导入 ctypes：Python 调用 C 语言动态链接库（DLL）的桥梁
# 这里主要用于调用 Windows API（如查找窗口、模拟按键等）
import ctypes
# 导入 time：用于控制程序执行节奏（sleep 暂停），等待界面响应
import time
# 导入 subprocess：用于在 Python 中启动外部命令（这里是 PowerShell）
# 因为 SendKeys 和剪贴板操作需要通过 PowerShell 完成
import subprocess



# ==================== 初始化 Windows API ====================

# 加载 Windows 系统的 user32.dll 动态链接库
# user32.dll 包含了大量与用户界面相关的 API 函数：
#   - FindWindowW: 根据标题查找窗口句柄
#   - SetForegroundWindow: 将指定窗口置于前台（激活窗口）
#   - ShowWindow: 显示/还原最小化的窗口
#   - IsIconic: 判断窗口是否处于最小化状态
#   - keybd_event: 模拟键盘按键事件（按下/释放）
user32 = ctypes.windll.user32

# ==================== 配置常量 ====================

# 目标群聊名称（必须与 eLink 中显示的群名完全一致，否则搜索不到）
GROUP_NAME = "测试项目工作群"
# 要发送的消息内容（纯文本，支持中文）
# MESSAGE = magss
# magss 数据文件路径（与 monitor.py 在同一目录下）
MAGSS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "magss_data.json")
# 南网eLink 主窗口的标题（用于 FindWindowW 查找窗口）
# 注意：这个值必须与任务栏/窗口标题栏上显示的文字完全匹配
ELINK_TITLE = "南网eLink"

# ==================== 虚拟键码常量 ====================

# VK_RETURN (0x0D)：回车键的虚拟键码
# 用于确认搜索、发送消息等操作
VK_RETURN = 0x0D
# VK_DOWN (0x28)：方向键【下】的虚拟键码
# 用于在 @提及 的下拉菜单中选中第一个匹配项
VK_DOWN = 0x28


# ==================== 函数定义 ====================

def activate(hwnd):
    """
    激活指定窗口：将其从最小化状态恢复，并置为前台焦点窗口

    参数:
        hwnd: 窗口句柄（由 FindWindowW 返回）
    """
    # 如果没找到窗口（hwnd 为 0 或 None），直接退出程序并报错
    if not hwnd:
        raise SystemExit("没找到 南网eLink 窗口")

    # 检查窗口是否处于最小化状态
    # IsIconic 返回非零值表示窗口被最小化了
    if user32.IsIconic(hwnd):
        # ShowWindow(hwnd, 9)：参数 9 对应 SW_RESTORE，表示还原窗口到正常大小
        # 如果窗口最小化在任务栏，需要先还原才能操作
        user32.ShowWindow(hwnd, 9)

    # 将该窗口设置为前台活动窗口（获得输入焦点）
    # 这样后续的键盘操作才会发送到 eLink 而不是其他程序
    user32.SetForegroundWindow(hwnd)

    # 等待 0.5 秒，给窗口切换留出时间
    # 太短可能导致后续按键发到了错误的窗口
    time.sleep(0.5)


def send_keys(keys):
    """
    通过 PowerShell 的 WScript.Shell.SendKeys 模拟按键

    参数:
        keys: 按键字符串，遵循 SendKeys 语法
              例如: '^F' 表示 Ctrl+F，'^v' 表示 Ctrl+V，'{ENTER}' 表示回车
    """
    # 构造 PowerShell 命令：
    # New-Object -ComObject WScript.Shell 创建 COM 对象
    # .SendKeys('...') 执行按键序列
    ps = "(New-Object -ComObject WScript.Shell).SendKeys('" + keys + "')"
    # 通过 subprocess 调用 powershell 执行该命令
    # -command 参数后跟要执行的 PowerShell 语句
    subprocess.run(['powershell', '-command', ps])


def paste_text(text):
    """
    将文本粘贴到当前焦点位置（通过剪贴板实现）

    为什么用剪贴板而不是直接打字？
    因为 SendKeys 对中文支持不好，容易乱码或失败
    而剪贴板可以完美处理任意 Unicode 字符（包括中文、特殊符号）

    参数:
        text: 要粘贴的文本内容
    """
    # 第一步：将文本写入系统剪贴板
    # Set-Clipboard -Value "..." 是 PowerShell 原生命令
    # text.replace('"', '\\"') 是为了转义文本中的双引号，防止破坏 PowerShell 命令结构
    ps = 'Set-Clipboard -Value "' + text.replace('"', '\\"') + '"'

    # 第二步：追加 Ctrl+V 按键命令，触发粘贴动作
    # ^v 在 SendKeys 中表示 Ctrl+V
    ps += "; (New-Object -ComObject WScript.Shell).SendKeys('^v')"

    # 执行组合命令：先设剪贴板，再按 Ctrl+V
    subprocess.run(['powershell', '-command', ps])


def press_key(vk):
    """
    使用 Windows API keybd_event 模拟单个物理按键

    为什么不用 SendKeys？
    对于简单的单键（如回车、方向键），keybd_event 更稳定、更快
    而且不依赖 PowerShell，减少子进程开销

    参数:
        vk: 虚拟键码（Virtual Key Code），如 VK_RETURN=0x0D, VK_DOWN=0x28
    """
    # KEYEVENTF_KEYUP (0x0002)：标志位，表示"释放按键"
    # 模拟完整按键需要两步：先按下（无标志位），再释放（带 KEYUP 标志）
    KEYEVENTF_KEYUP = 0x0002

    # 第一步：按下按键
    # keybd_event(bVk, bScan, dwFlags, dwExtraInfo)
    #   bVk: 虚拟键码
    #   bScan: 扫描码（一般传 0）
    #   dwFlags: 标志位（0 表示按下）
    #   dwExtraInfo: 额外信息（一般传 0）
    user32.keybd_event(vk, 0, 0, 0)

    # 第二步：释放按键
    # 带上 KEYEVENTF_KEYUP 标志，表示松开
    user32.keybd_event(vk, 0, KEYEVENTF_KEYUP, 0)


def at_mention(name):
    """
    在聊天框中 @指定人员

    流程：敲 @ → 粘贴名字过滤 → 方向键下选中 → 回车确认

    参数:
        name: 要 @ 的人员姓名（需与 eLink 通讯录中的名字匹配）
    """
    # ① 先敲 @ 符号唤起提及（mention）下拉菜单
    # @ 是 ASCII 字符，SendKeys 处理很稳定
    send_keys('@')
    time.sleep(1)  # 等待 1 秒，让下拉菜单完全弹出

    # ② 粘贴中文名字进行过滤
    # 不能直接用 SendKeys 打中文（会乱码），所以走剪贴板
    # 粘贴后 eLink 会自动过滤出匹配的成员列表
    paste_text(name)
    time.sleep(1.5)  # 等待 1.5 秒，等过滤结果渲染完成（中文搜索可能稍慢）

    # ③ 按方向键【下】选中第一个匹配项
    # 虽然默认高亮就是第一个，但显式按一次 Down 确保选中状态正确
    press_key(VK_DOWN)
    time.sleep(0.3)  # 短暂等待，确保选中生效

    # ④ 按回车确认选择该人员
    # 确认后，输入框中会出现 "@黄仕伟" 这样的提及标记
    press_key(VK_RETURN)
    time.sleep(0.5)  # 等待确认完成


# ==================== 主流程 ====================

def main():
    """
    主函数：串联所有步骤，完成消息发送
    """
    # 第一步：查找南网eLink 主窗口
    # FindWindowW(lpClassName, lpWindowName)
    #   第一个参数 None 表示不限定类名
    #   第二个参数 ELINK_TITLE 是窗口标题
    # 返回值是窗口句柄（HWND），找不到则返回 0
    hwnd = user32.FindWindowW(None, ELINK_TITLE)

    # 第二步：激活窗口（恢复最小化 + 置前）
    activate(hwnd)

    # ---------- 第三步：搜索并进入目标群 ----------

    # 按 Ctrl+F 打开 eLink 的搜索框
    # ^F 在 SendKeys 中表示 Ctrl+F
    send_keys('^F')
    time.sleep(0.8)  # 等待搜索框出现并获得焦点

    # 粘贴群名称进行搜索
    # 使用剪贴板方式，避免中文输入问题
    paste_text(GROUP_NAME)
    time.sleep(1.5)  # 等待搜索结果加载出来

    # 按回车进入第一个匹配的群聊
    press_key(VK_RETURN)
    time.sleep(1)  # 等待群聊界面加载完成

    # ---------- 第四步：@指定人员 ----------

    # 调用 at_mention 函数，@ 指定的人
    # 如果需要 @ 多人，可以在这里循环调用，例如：
    # for person in ["黄仕伟", "张三", "李四"]:
    #     at_mention(person)
    # 如果不需要 @ 人，直接注释掉这一行即可
    at_mention("黄仕伟")  # ← 传入具体人名

    # ---------- 第五步：粘贴消息内容 ----------

    # 将预设的消息文本粘贴到输入框
    paste_text(MESSAGE)
    time.sleep(0.5)  # 等待粘贴完成

    # ---------- 第六步：回车发送 ----------

    # 按回车键发送消息
    press_key(VK_RETURN)

    # 打印成功提示
    print(f"已发送到 {GROUP_NAME}")


# ==================== 程序入口 ====================

# Python 标准写法：只有当这个文件被直接运行时才执行 main()
# 如果被其他模块 import，则不会自动执行
if __name__ == "__main__":
    main()