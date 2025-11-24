import torch
import numpy as np
from PIL import Image, ImageTk
import tkinter as tk
import ctypes
from ctypes import wintypes
import platform

# --- Windows API 定义 ---
user32 = ctypes.windll.user32
gdi32 = ctypes.windll.gdi32
shcore = ctypes.windll.shcore

# 开启高DPI感知
try:
    shcore.SetProcessDpiAwareness(2) # PROCESS_PER_MONITOR_DPI_AWARE
except Exception:
    user32.SetProcessDPIAware()

# PrintWindow 标志
PW_CLIENTONLY = 1
PW_RENDERFULLCONTENT = 2 

# --- 关键修复：手动定义 BITMAPINFOHEADER ---
class BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [
        ('biSize', wintypes.DWORD),
        ('biWidth', wintypes.LONG),
        ('biHeight', wintypes.LONG),
        ('biPlanes', wintypes.WORD),
        ('biBitCount', wintypes.WORD),
        ('biCompression', wintypes.DWORD),
        ('biSizeImage', wintypes.DWORD),
        ('biXPelsPerMeter', wintypes.LONG),
        ('biYPelsPerMeter', wintypes.LONG),
        ('biClrUsed', wintypes.DWORD),
        ('biClrImportant', wintypes.DWORD),
    ]

class BackgroundCapture:
    """
    处理后台截图的核心类，使用 PrintWindow API
    """
    def capture(self, hwnd):
        # 获取窗口尺寸
        r = wintypes.RECT()
        user32.GetWindowRect(hwnd, ctypes.byref(r))
        width = r.right - r.left
        height = r.bottom - r.top
        
        if width <= 0 or height <= 0:
            return None

        # 创建设备上下文 (DC)
        hwndDC = user32.GetWindowDC(hwnd)
        mfcDC = gdi32.CreateCompatibleDC(hwndDC)
        saveDC = gdi32.CreateCompatibleBitmap(hwndDC, width, height)
        
        gdi32.SelectObject(mfcDC, saveDC)

        # 核心：PrintWindow 后台截图
        # 尝试使用 PW_RENDERFULLCONTENT (Win8.1+)
        result = user32.PrintWindow(hwnd, mfcDC, PW_RENDERFULLCONTENT)
        if result == 0:
            # 回退旧版方法
            result = user32.PrintWindow(hwnd, mfcDC, 0)

        if result == 0:
            # 如果依然失败，清理并返回
            gdi32.DeleteObject(saveDC)
            gdi32.DeleteDC(mfcDC)
            user32.ReleaseDC(hwnd, hwndDC)
            print("PrintWindow failed.")
            return None

        # 初始化结构体
        bmpinfo = BITMAPINFOHEADER()
        bmpinfo.biSize = ctypes.sizeof(BITMAPINFOHEADER)
        bmpinfo.biWidth = width
        bmpinfo.biHeight = -height # 负数表示自上而下 (Top-down)
        bmpinfo.biPlanes = 1
        bmpinfo.biBitCount = 32
        bmpinfo.biCompression = 0 # BI_RGB
        
        buffer_len = width * height * 4
        buffer = ctypes.create_string_buffer(buffer_len)
        
        # 将位图数据复制到 buffer
        # 0 = DIB_RGB_COLORS
        lines_copied = gdi32.GetDIBits(mfcDC, saveDC, 0, height, buffer, ctypes.byref(bmpinfo), 0)
        
        # 清理资源
        gdi32.DeleteObject(saveDC)
        gdi32.DeleteDC(mfcDC)
        user32.ReleaseDC(hwnd, hwndDC)

        if lines_copied == 0:
            return None

        # 转为 PIL Image (注意 Windows 位图通常是 BGR 顺序)
        try:
            image = Image.frombuffer("RGBX", (width, height), buffer, "raw", "BGRX", 0, 1)
            return image.convert("RGB")
        except Exception as e:
            print(f"Buffer conversion error: {e}")
            return None

class ROISelector:
    """
    弹窗让用户在静态截图上框选区域
    """
    def __init__(self, pil_image, window_name):
        self.selection = None
        self.root = tk.Tk()
        self.root.title(f"Select Region for: {window_name}")
        self.root.attributes('-topmost', True)
        
        self.tk_image = ImageTk.PhotoImage(pil_image)
        self.w, self.h = pil_image.size
        
        # 限制窗口最大尺寸，防止图片过大超出屏幕
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        
        win_w = min(self.w, screen_w - 100)
        win_h = min(self.h, screen_h - 100)
        
        self.root.geometry(f"{win_w}x{win_h}")
        
        # 使用 Scrollbar 防止大图无法显示完全
        vbar = tk.Scrollbar(self.root, orient=tk.VERTICAL)
        hbar = tk.Scrollbar(self.root, orient=tk.HORIZONTAL)
        vbar.pack(side=tk.RIGHT, fill=tk.Y)
        hbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.canvas = tk.Canvas(self.root, width=win_w, height=win_h, scrollregion=(0,0,self.w,self.h), xscrollcommand=hbar.set, yscrollcommand=vbar.set, cursor="cross")
        vbar.config(command=self.canvas.yview)
        hbar.config(command=self.canvas.xview)
        
        self.canvas.pack(expand=True, fill=tk.BOTH)
        self.canvas.create_image(0, 0, image=self.tk_image, anchor=tk.NW)
        
        self.canvas.bind("<ButtonPress-1>", self.on_press)
        self.canvas.bind("<B1-Motion>", self.on_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)
        
        self.rect = None
        self.start_x = 0
        self.start_y = 0

    def on_press(self, event):
        # 需要加上滚动条偏移量
        self.start_x = self.canvas.canvasx(event.x)
        self.start_y = self.canvas.canvasy(event.y)
        self.rect = self.canvas.create_rectangle(self.start_x, self.start_y, self.start_x, self.start_y, outline='red', width=2)

    def on_drag(self, event):
        cur_x = self.canvas.canvasx(event.x)
        cur_y = self.canvas.canvasy(event.y)
        self.canvas.coords(self.rect, self.start_x, self.start_y, cur_x, cur_y)

    def on_release(self, event):
        cur_x = self.canvas.canvasx(event.x)
        cur_y = self.canvas.canvasy(event.y)
        
        x1, y1 = self.start_x, self.start_y
        x2, y2 = cur_x, cur_y
        
        self.selection = (
            int(min(x1, x2)), int(min(y1, y2)),
            int(max(x1, x2)), int(max(y1, y2))
        )
        self.root.destroy()

    def get_roi(self):
        self.root.mainloop()
        return self.selection

# 全局缓存
_roi_cache = {}

class BackgroundCaptureNode:
    def __init__(self):
        self.capturer = BackgroundCapture()
    
    @classmethod
    def INPUT_TYPES(s):
        return {
            "required": {
                "window_title": ("STRING", {"default": "Untitled", "multiline": False}),
                "reset_roi": ("BOOLEAN", {"default": False, "label_on": "Reselect Region", "label_off": "Use Cached Region"}),
                "seed": ("INT", {"default": 0, "min": 0, "max": 0xffffffffffffffff}),
            },
        }

    RETURN_TYPES = ("IMAGE",)
    RETURN_NAMES = ("image",)
    FUNCTION = "capture_background"
    CATEGORY = "Image/Capture"

    def find_window_handle(self, partial_title):
        target_hwnd = None
        def callback(hwnd, lParam):
            nonlocal target_hwnd
            if not user32.IsWindowVisible(hwnd):
                return True
            length = user32.GetWindowTextLengthW(hwnd)
            if length == 0: return True
            
            buff = ctypes.create_unicode_buffer(length + 1)
            user32.GetWindowTextW(hwnd, buff, length + 1)
            
            if partial_title.lower() in buff.value.lower():
                target_hwnd = hwnd
                return False
            return True
            
        user32.EnumWindows(ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)(callback), 0)
        return target_hwnd

    def capture_background(self, window_title, reset_roi, seed):
        global _roi_cache
        
        if platform.system() != "Windows":
            print("Background capture only works on Windows.")
            return (torch.zeros((1, 64, 64, 3)),)

        hwnd = self.find_window_handle(window_title)
        if not hwnd:
            print(f"Window '{window_title}' not found.")
            return (torch.zeros((1, 64, 64, 3)),)

        full_img = self.capturer.capture(hwnd)
        if full_img is None:
            print("Failed to capture window (Minimzed or Invalid).")
            return (torch.zeros((1, 64, 64, 3)),)

        roi = _roi_cache.get(window_title)
        
        if reset_roi or roi is None:
            print("Opening selection window...")
            try:
                selector = ROISelector(full_img, window_title)
                roi = selector.get_roi()
            except Exception as e:
                print(f"GUI Error: {e}")
                roi = None
            
            if roi and (roi[2] - roi[0] > 0) and (roi[3] - roi[1] > 0):
                _roi_cache[window_title] = roi
                print(f"ROI Cached: {roi}")
            else:
                roi = (0, 0, full_img.width, full_img.height)
                _roi_cache[window_title] = roi

        safe_x1 = max(0, roi[0])
        safe_y1 = max(0, roi[1])
        safe_x2 = min(full_img.width, roi[2])
        safe_y2 = min(full_img.height, roi[3])
        
        if safe_x2 <= safe_x1 or safe_y2 <= safe_y1:
            final_img = full_img
        else:
            final_img = full_img.crop((safe_x1, safe_y1, safe_x2, safe_y2))

        np_img = np.array(final_img).astype(np.float32) / 255.0
        return (torch.from_numpy(np_img)[None,],)

NODE_CLASS_MAPPINGS = {
    "Win11BackgroundCapture": BackgroundCaptureNode
}

NODE_DISPLAY_NAME_MAPPINGS = {
    "Win11BackgroundCapture": "👻 Win11 Background Capture (Occluded)"
}
