import win32clipboard
import win32gui
import win32con
import ctypes

class ClipboardWatcher(object):
    def __init__(self):
        self.hwnd = win32gui.CreateWindowEx(0, "STATIC", "", 0, 0, 0, 0, 0, 0, 0, 0, None)
        win32gui.AddClipboardFormatListener(self.hwnd)
        self.callback = None

    def register_callback(self, callback):
        self.callback = callback

    def _clipboard_changed(self, hwnd, msg, wparam, lparam):
        if self.callback:
            self.callback()
        return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)

    def start_watcher(self):
        msg = ctypes.wintypes.MSG()
        while ctypes.windll.user32.GetMessageW(ctypes.byref(msg), self.hwnd, 0, 0) != 0:
            if msg.message == win32con.WM_CLIPBOARDUPDATE:
                self._clipboard_changed(msg.hwnd, msg.message, msg.wParam, msg.lParam)
            ctypes.windll.user32.TranslateMessage(ctypes.byref(msg))
            ctypes.windll.user32.DispatchMessageW(ctypes.byref(msg))

def on_clipboard_change():
    try:
        win32clipboard.OpenClipboard()
        data = win32clipboard.GetClipboardData()
        win32clipboard.CloseClipboard()
        print(f"剪贴板内容发生变化: {data}")
    except Exception as e:
        print(f"剪贴板读取失败: {e}")

if __name__ == "__main__":
    watcher = ClipboardWatcher()
    watcher.register_callback(on_clipboard_change)
    watcher.start_watcher()
