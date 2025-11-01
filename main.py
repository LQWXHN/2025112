import cv2
import threading
import time
from PIL import Image, ImageDraw
import pyautogui
import win32gui
import win32con
import pystray
from pathlib import Path
import sys
import os
import win32com.client  # 用于读取系统设备信息

# 全局变量：控制程序运行、摄像头状态、内置摄像头索引列表
running = True
camera_active = False
hwnd = None
built_in_camera_indexes = []  # 存储内置摄像头的索引

def resource_path(relative_path):
    """处理打包后资源路径问题"""
    if hasattr(sys, '_MEIPASS'):
        return Path(sys._MEIPASS) / relative_path
    return Path(__file__).parent / relative_path

def get_built_in_camera_indexes():
    """获取电脑内置摄像头的索引（核心：区分内置/外置）"""
    built_in_indexes = []
    # 读取系统视频设备（WMI查询）
    try:
        wmi = win32com.client.GetObject("winmgmts:")
        # 查询所有视频输入设备
        devices = wmi.InstancesOf("Win32_PnPEntity")
        for device in devices:
            device_name = str(device.Name).lower() if device.Name else ""
            # 筛选内置摄像头关键词（可根据实际设备名称补充）
            built_in_keywords = ["integrated", "内置", "集成", "built-in", "laptop", "notebook"]
            external_keywords = ["usb", "external", "外接", "usb camera", "webcam usb"]
            
            # 判定为内置摄像头：包含内置关键词 + 不包含外置关键词
            is_built_in = any(keyword in device_name for keyword in built_in_keywords) and not any(keyword in device_name for keyword in external_keywords)
            if is_built_in and "camera" in device_name:
                # 遍历可能的索引，匹配设备对应的摄像头索引
                for idx in range(10):  # 最多检测10个索引
                    cap = None
                    try:
                        cap = cv2.VideoCapture(idx, cv2.CAP_DSHOW)
                        if cap.isOpened():
                            # 简单验证：内置摄像头通常分辨率较低（如720P），可辅助判断
                            width = cap.get(cv2.CAP_PROP_FRAME_WIDTH)
                            height = cap.get(cv2.CAP_PROP_FRAME_HEIGHT)
                            # 内置摄像头常见分辨率（可根据实际调整）
                            common_built_in_res = [(640, 480), (1280, 720), (1920, 1080)]
                            if (width, height) in common_built_in_res and idx not in built_in_indexes:
                                built_in_indexes.append(idx)
                    except Exception:
                        pass
                    finally:
                        if cap:
                            cap.release()
    except Exception:
        # 异常时降级：默认索引0（多数电脑内置摄像头为索引0）
        built_in_indexes = [0]
    return list(set(built_in_indexes))  # 去重

def check_built_in_camera():
    """仅检测内置摄像头是否打开（每2秒检测一次）"""
    global camera_active
    global built_in_camera_indexes
    # 初始化内置摄像头索引（程序启动时获取一次）
    if not built_in_camera_indexes:
        built_in_camera_indexes = get_built_in_camera_indexes()
        # 若未识别到，默认检测索引0（兜底）
        if not built_in_camera_indexes:
            built_in_camera_indexes = [0]
    
    while running:
        temp_active = False
        # 仅遍历内置摄像头的索引
        for idx in built_in_camera_indexes:
            cap = None
            try:
                cap = cv2.VideoCapture(idx, cv2.CAP_DSHOW)
                if cap.isOpened():
                    temp_active = True
                    break  # 任一内置摄像头打开即判定为活跃
            except Exception:
                pass
            finally:
                if cap:
                    cap.release()
        camera_active = temp_active
        time.sleep(2)

def create_red_dot_image():
    """创建红色圆点图像（10x10像素，半透明）"""
    image = Image.new("RGBA", (10, 10), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)
    draw.ellipse((0, 0, 9, 9), fill=(255, 0, 0, 200))
    return image

def create_tray_icon():
    """创建系统托盘图标（用于退出程序）"""
    def on_quit(icon, item):
        global running
        running = False
        icon.stop()
        if hwnd:
            win32gui.DestroyWindow(hwnd)
        sys.exit()

    image = create_red_dot_image()
    icon = pystray.Icon("BuiltInCameraDetector", image, "内置摄像头检测工具")
    icon.menu = pystray.Menu(pystray.MenuItem("退出", on_quit))
    return icon

def create_floating_window():
    """创建无边框悬浮窗（右上角红点）"""
    global hwnd
    screen_width, screen_height = pyautogui.size()
    win_x = screen_width - 30
    win_y = 20
    win_width, win_height = 10, 10

    wc = win32gui.WNDCLASS()
    wc.lpszClassName = "FloatingRedDot_BuiltIn"
    wc.lpfnWndProc = lambda h, msg, wparam, lparam: 0
    wc.hInstance = win32gui.GetModuleHandle(None)
    win32gui.RegisterClass(wc)

    hwnd = win32gui.CreateWindowEx(
        win32con.WS_EX_LAYERED | win32con.WS_EX_TRANSPARENT | win32con.WS_EX_TOPMOST,
        wc.lpszClassName,
        "RedDotWindow_BuiltIn",
        win32con.WS_POPUP,
        win_x, win_y, win_width, win_height,
        None, None, wc.hInstance, None
    )

    win32gui.SetLayeredWindowAttributes(hwnd, 0, 200, win32con.LWA_ALPHA)
    win32gui.ShowWindow(hwnd, win32con.SW_SHOW)

def update_floating_window():
    """更新悬浮窗显示（内置摄像头打开则显示红点）"""
    while running:
        if camera_active:
            if hwnd:
                win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
                hdc = win32gui.GetDC(hwnd)
                image = create_red_dot_image()
                win32gui.BitBlt(hdc, 0, 0, 10, 10, image.getdata(), 0, 0, win32con.SRCCOPY)
                win32gui.ReleaseDC(hwnd, hdc)
        else:
            if hwnd:
                win32gui.ShowWindow(hwnd, win32con.SW_HIDE)
        time.sleep(0.5)

def set_startup(auto_start=True):
    """设置开机自启（Windows系统）"""
    app_name = "BuiltInCameraDetector"
    app_path = sys.executable
    startup_path = Path(os.getenv("APPDATA")) / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Startup"
    shortcut_path = startup_path / f"{app_name}.lnk"

    if auto_start:
        try:
            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortcut(str(shortcut_path))
            shortcut.TargetPath = app_path
            shortcut.WorkingDirectory = str(Path(app_path).parent)
            shortcut.Save()
        except Exception:
            pass
    else:
        if shortcut_path.exists():
            shortcut_path.unlink(missing_ok=True)

def main():
    global running
    # 1. 设置开机自启
    set_startup(auto_start=True)

    # 2. 创建悬浮窗
    create_floating_window()

    # 3. 启动内置摄像头检测线程
    camera_thread = threading.Thread(target=check_built_in_camera, daemon=True)
    camera_thread.start()

    # 4. 启动悬浮窗更新线程
    update_thread = threading.Thread(target=update_floating_window, daemon=True)
    update_thread.start()

    # 5. 启动系统托盘
    tray_icon = create_tray_icon()
    tray_icon.run()

    # 6. 退出清理
    running = False
    camera_thread.join()
    update_thread.join()
    if hwnd:
        win32gui.DestroyWindow(hwnd)

if __name__ == "__main__":
    main()