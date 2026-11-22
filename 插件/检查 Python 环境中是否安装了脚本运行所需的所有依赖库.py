# 尝试执行下面的导入操作。
# try 块用于捕获可能发生的错误（这里是导入模块时可能出现的错误）。
try:
    # 导入 requests 库（第三方库，需手动安装）
    import requests

    # 导入 json 库（Python 内置标准库，无需安装，用于处理 JSON 数据）
    import json

    # 导入 time 库（Python 内置标准库，无需安装，用于时间控制如 sleep）
    import time

    # 导入 pandas 库并简写为 pd（第三方库，需手动安装，用于数据处理）
    import pandas as pd

    # 从 datetime 模块中导入 datetime 类（Python 内置标准库，无需安装，用于日期时间处理）
    from datetime import datetime

    # 导入 os 库（Python 内置标准库，无需安装，用于操作系统相关功能如文件路径判断）
    import os

    # 如果上面所有 import 都成功执行到这里，说明所有依赖库都已正确安装
    print("✅ 所有依赖库均已安装")

# 如果 try 块中任何一行 import 失败，会抛出 ImportError 异常
# as e 表示将捕获到的异常对象赋值给变量 e，方便后续获取错误信息
except ImportError as e:
    # e.name 是异常对象的属性，值为“缺失的模块名称”
    # 例如：缺少 pandas 时，e.name 就是 "pandas"
    print(f"❌ 缺少依赖库: {e.name}")

    # 根据缺失的模块名，生成对应的 pip 安装命令，提示用户执行
    # 例如：缺少 pandas 时，输出 "请运行: pip install pandas"
    print(f"请运行: pip install {e.name}")

"""补充说明
ImportError 的触发场景：
当 import 的模块未安装（如 requests、pandas、openpyxl）；
或模块安装但损坏/版本不兼容导致无法导入时，才会抛出该异常。
而 json、time、os、datetime 是 Python 内置库，正常环境下不会触发 ImportError。
e.name 的作用： ImportError 异常的 name 属性会直接返回「缺失的模块名」，因此可以精准定位是哪个库没装，并自动生成对应的 pip install 命令，无需手动排查。
扩展性： 若脚本新增了其他第三方依赖（如 openpyxl），只需在 try 块中添加对应的 import 语句即可，异常捕获逻辑会自动适配新依赖的检查。"""