# license_manager.py - 永久免授权版

def check_license():
    # 直接告诉主程序：验证通过！
    # 返回 True 和一个提示语
    return True, "永久免授权版"

def get_machine_code():
    return "无需机器码"

# 下面这行是为了防止单独运行时报错，留着没事
if __name__ == "__main__":
    print("当前已是免授权模式")