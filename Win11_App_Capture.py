import numpy as np
import torch
import cv2
from ctypes import windll, byref, c_ubyte, Structure, POINTER, c_int, sizeof
from ctypes.wintypes import RECT, HWND, HDC, HBITMAP, HGDIOBJ
import win32gui
import win32con

# --- 1. Windows API 定义 (核心截图逻辑) ---
class BITMAPINFOHEADER(Structure):
    _fields_ = [("biSize", c_int), ("biWidth", c_int), ("biHeight", c_int),
                ("biPlanes", c_int), ("biBitCount", c_int), ("biCompression", c_int),
                ("biSizeImage", c_int), ("biXPelsPerMeter", c_int),
                ("biYPelsPerMeter", c_int), ("biClrUsed", c_int),
                ("biClrImportant", c_int)]

class BITMAPINFO(Structure):
    _fields_ = [("bmiHeader", BITMAPINFOHEADER), ("bmiColors", c_int * 3)]

# 全局变量存储 ROI (感兴趣区域)
# 格式: {'窗口标题': (x, y, w, h)}
ROI_STORAGE = {}

class Win11CaptureNode:
    def __init__(self):
        pass

    @classmethod
    def INPUT_TYPES(s):
        return {
            "required": {
                "window_title": ("STRING", {"default": "Notepad"}), # 目标窗口标题
                "reset_roi": ("BOOLEAN", {"default": False}),       # 是否重置选区
                # 用 seed 变化来强制 ComfyUI 刷新，实现"实时"
                "seed": ("INT", {"default": 0, "min": 0, "max": 0xffffffffffffffff}),
            }
        }

    RETURN_TYPES = ("IMAGE",)
    RETURN_NAMES = ("image",)
    FUNCTION = "capture"
    CATEGORY = "🖥️ Desktop Capture"

    def capture_window(self, hwnd):
        # 1. 获取窗口尺寸
        rect = RECT()
        windll.user32.GetWindowRect(hwnd, byref(rect))
        w = rect.right - rect.left
        h = rect.bottom - rect.top

        if w == 0 or h == 0:
            return None

        # 2. 创建设备上下文
        hwndDC = windll.user32.GetWindowDC(hwnd)
        mfcDC = windll.gdi32.CreateCompatibleDC(hwndDC)
        saveBitMap = windll.gdi32.CreateCompatibleBitmap(hwndDC, w, h)
        windll.gdi32.SelectObject(mfcDC, saveBitMap)

        # 3. 核心：使用 PrintWindow (支持后台/遮挡)
        # PW_RENDERFULLCONTENT = 0x00000002 (Win11 关键参数)
        result = windll.user32.PrintWindow(hwnd, mfcDC, 2) 

        # 4. 提取位图数据
        bmpinfo = BITMAPINFO()
        bmpinfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER)
        bmpinfo.bmiHeader.biWidth = w
        bmpinfo.bmiHeader.biHeight = -h  # 负数表示自上而下
        bmpinfo.bmiHeader.biPlanes = 1
        bmpinfo.bmiHeader.biBitCount = 32
        bmpinfo.bmiHeader.biCompression = 0 # BI_RGB

        buffer_len = h * w * 4
        buffer = (c_ubyte * buffer_len)()
        
        windll.gdi32.GetDIBits(mfcDC, saveBitMap, 0, h, buffer, byref(bmpinfo), 0)

        # 5. 转换为 Numpy 数组
        image = np.frombuffer(buffer, dtype=np.uint8).reshape((h, w, 4))

        # 清理内存
        windll.gdi32.DeleteObject(saveBitMap)
        windll.gdi32.DeleteDC(mfcDC)
        windll.user32.ReleaseDC(hwnd, hwndDC)

        # 剔除 Alpha 通道 (BGRA -> BGR)
        return image[:, :, :3] 

    def capture(self, window_title, reset_roi, seed):
        # 1. 查找窗口句柄
        hwnd = win32gui.FindWindow(None, window_title)
        if not hwnd:
            # 模糊搜索尝试
            def callback(h, params):
                txt = win32gui.GetWindowText(h)
                if window_title.lower() in txt.lower() and win32gui.IsWindowVisible(h):
                    params.append(h)
            hwnds = []
            win32gui.EnumWindows(callback, hwnds)
            if hwnds:
                hwnd = hwnds[0]
            else:
                print(f"Window '{window_title}' not found, returning black image.")
                return (torch.zeros((1, 512, 512, 3)),)

        # 2. 执行截图
        img_np = self.capture_window(hwnd)
        if img_np is None:
            return (torch.zeros((1, 512, 512, 3)),)

        # 3. 处理 ROI (鼠标框选)
        global ROI_STORAGE
        
        # 如果请求重置 ROI 或者该窗口还没选过区
        if reset_roi or window_title not in ROI_STORAGE:
            print(f"请在弹出的 '{window_title}' 截图中框选区域，并按 Enter 确认...")
            # 弹出一个 OpenCV 窗口让用户画框
            # 为了防止图太大，可以缩放一下，这里直接显示原图
            roi = cv2.selectROI("Select Region (Press Enter to Confirm)", img_np, showCrosshair=True, fromCenter=False)
            cv2.destroyAllWindows()
            
            # roi 格式是 (x, y, w, h)
            # 如果用户没选直接关掉，roi会全是0
            if roi[2] > 0 and roi[3] > 0:
                ROI_STORAGE[window_title] = roi
            else:
                # 默认全选
                h, w, _ = img_np.shape
                ROI_STORAGE[window_title] = (0, 0, w, h)

        # 4. 裁剪图像
        x, y, w, h = ROI_STORAGE[window_title]
        
        # 安全检查，防止窗口变小导致越界
        img_h, img_w, _ = img_np.shape
        x = min(x, img_w - 1)
        y = min(y, img_h - 1)
        w = min(w, img_w - x)
        h = min(h, img_h - y)

        crop_img = img_np[y:y+h, x:x+w]

        # 5. 格式转换 OpenCV (BGR) -> ComfyUI (RGB Tensor)
        # OpenCV 是 BGR, ComfyUI 需要 RGB
        img_rgb = cv2.cvtColor(crop_img, cv2.COLOR_BGR2RGB)
        
        # 归一化到 0-1 浮点数
        img_float = img_rgb.astype(np.float32) / 255.0
        
        # 转换为 Torch Tensor (Batch, Height, Width, Channel)
        img_tensor = torch.from_numpy(img_float)[None,]

        return (img_tensor,)

# 节点映射
NODE_CLASS_MAPPINGS = {
    "Win11AppCapture": Win11CaptureNode
}

NODE_DISPLAY_NAME_MAPPINGS = {
    "Win11AppCapture": "🪟 Win11 App Capture (ROI)"
}
