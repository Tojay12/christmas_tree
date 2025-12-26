import time
import pyautogui
import pygetwindow as gw
import subprocess
import os
import winreg
import traceback

# -------------------
# 配置
# -------------------
CLICK_PER_VIEW = 10       # 每屏点击会话数量（可适当大一点）
ROW_HEIGHT = 60           # 每行会话高度
LOOP_INTERVAL = 600*10     # 循环间隔 10 分钟
SCROLL_STEP = -500        # 每次滚动距离（负数向下）

# -------------------
# 查找/启动企业微信
# -------------------
def find_wxwork_window():
    wins = [w for w in gw.getWindowsWithTitle("企业微信") if w.isVisible]
    return wins[0] if wins else None

def find_wxwork_path():
    default_paths = [
        r"C:\Program Files (x86)\Tencent\WeChatWork\WXWork.exe",
        r"C:\Program Files\Tencent\WeChatWork\WXWork.exe"
    ]
    for path in default_paths:
        if os.path.exists(path):
            return path
    try:
        reg_paths = [
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\WeWork",
            r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\WeWork"
        ]
        for reg_path in reg_paths:
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path)
            exe_path, _ = winreg.QueryValueEx(key, "DisplayIcon")
            if os.path.exists(exe_path):
                return exe_path
    except:
        pass
    return None

def activate_or_start_wxwork():
    win = find_wxwork_window()
    if win:
        win.activate()
        time.sleep(1)
    else:
        exe_path = find_wxwork_path()
        if exe_path:
            subprocess.Popen(exe_path)
            time.sleep(5)
            win = find_wxwork_window()
            if win:
                win.activate()
                time.sleep(1)
            else:
                raise Exception("启动企业微信失败")
        else:
            raise Exception("未找到企业微信路径")
    return win

# -------------------
# 智能滚动 + 点击
# -------------------
def scroll_and_click_smart(win):
    chat_list_x = win.left + 100
    chat_y_start = win.top + 150
    chat_y_end = win.bottom - 50

    prev_pos = None
    reached_bottom = False

    while not reached_bottom:
        # 点击当前可见会话
        y = chat_y_start
        while y < chat_y_end:
            pyautogui.click(chat_list_x, y)
            y += ROW_HEIGHT
            time.sleep(0.2)

        # 滚动
        pyautogui.moveTo(chat_list_x, chat_y_start)
        pyautogui.scroll(SCROLL_STEP)
        time.sleep(0.5)

        # 判断是否到底
        current_pos = pyautogui.position()
        if prev_pos == current_pos:
            reached_bottom = True
        prev_pos = current_pos

# -------------------
# 主循环
# -------------------
def main_loop():
    while True:
        try:
            win = activate_or_start_wxwork()
            scroll_and_click_smart(win)
            print("✅ 已尝试清空全部会话未读")
        except Exception as e:
            print("❌ 异常:", e)
            traceback.print_exc()
        print(f"⏳ 等待 {LOOP_INTERVAL//60} 分钟后再次检查")
        time.sleep(LOOP_INTERVAL)

if __name__ == "__main__":
    main_loop()
