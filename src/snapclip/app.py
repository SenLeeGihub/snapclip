from __future__ import annotations

import contextlib
import ctypes
import io
import logging
import os
import tempfile
import threading
import time
from ctypes import wintypes
from dataclasses import dataclass
from typing import Optional, Tuple

import mss
import pystray
import pywintypes
import win32api
import win32clipboard
import win32con
import win32gui
from PIL import Image, ImageDraw

from . import __version__

# Win32 constants that are not provided by win32con
MOD_NOREPEAT = 0x4000
ERROR_CLASS_ALREADY_EXISTS = getattr(win32con, "ERROR_CLASS_ALREADY_EXISTS", 1410)
LRESULT = ctypes.c_longlong if ctypes.sizeof(ctypes.c_void_p) == 8 else ctypes.c_long


def _setup_logging() -> None:
    """Configure silent file logging to help diagnose issues without UI."""
    log_path = os.path.join(tempfile.gettempdir(), "snapclip.log")
    logging.basicConfig(
        filename=log_path,
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
    )
    logging.info("SnapClip starting up (pid=%s)", os.getpid())


@dataclass
class CaptureSession:
    start: Optional[Tuple[int, int]] = None
    current: Optional[Tuple[int, int]] = None
    final: Optional[Tuple[int, int]] = None

    def bounds(self) -> Optional[Tuple[int, int, int, int]]:
        if not self.start or not self.final:
            return None
        left = min(self.start[0], self.final[0])
        top = min(self.start[1], self.final[1])
        right = max(self.start[0], self.final[0])
        bottom = max(self.start[1], self.final[1])
        width = right - left
        height = bottom - top
        if width < 3 or height < 3:
            return None
        return left, top, width, height


class SnapClipApp:
    HOTKEY_ID = 1
    ESC_HOTKEY_ID = 2
    WM_APP_EXIT = win32con.WM_APP + 1
    WM_APP_CAPTURE_COMPLETE = win32con.WM_APP + 2

    def __init__(self) -> None:
        if os.name != "nt":
            raise SystemExit("SnapClip only supports Windows.")
        self._instance = win32api.GetModuleHandle(None)
        self._class_name = "SnapClipHiddenWindow"
        self._class_atom = None
        self._hwnd = None
        self._icon: Optional[pystray.Icon] = None
        self._mss = mss.mss()
        self._capture_session: Optional[CaptureSession] = None
        self._user32 = ctypes.windll.user32
        self._user32.CallNextHookEx.argtypes = [
            ctypes.c_void_p,
            ctypes.c_int,
            wintypes.WPARAM,
            wintypes.LPARAM,
        ]
        self._user32.CallNextHookEx.restype = LRESULT
        self._mouse_proc_ref = None
        self._mouse_hook = None
        self._autotest = os.environ.get("SNAPCLIP_AUTOTEST") == "1"
        self._autotest_timer: Optional[threading.Timer] = None

    def run(self) -> None:
        logging.info("Initializing SnapClip")
        self._create_message_window()
        self._register_primary_hotkey()
        self._start_tray_icon()
        if self._autotest:
            self._autotest_timer = threading.Timer(2.0, self.request_exit)
            self._autotest_timer.daemon = True
            self._autotest_timer.start()
        try:
            win32gui.PumpMessages()
        finally:
            self._teardown()

    def request_exit(self) -> None:
        if self._hwnd:
            logging.info("Exit requested")
            win32gui.PostMessage(self._hwnd, self.WM_APP_EXIT, 0, 0)

    def _create_message_window(self) -> None:
        wndclass = win32gui.WNDCLASS()
        wndclass.hInstance = self._instance
        wndclass.lpszClassName = self._class_name
        wndclass.lpfnWndProc = self._handle_win32_event
        try:
            self._class_atom = win32gui.RegisterClass(wndclass)
        except win32gui.error as exc:
            if exc.winerror != ERROR_CLASS_ALREADY_EXISTS:
                raise
        self._hwnd = win32gui.CreateWindowEx(
            win32con.WS_EX_LEFT,
            self._class_name,
            "SnapClipMessageWindow",
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            self._instance,
            None,
        )
        if not self._hwnd:
            raise RuntimeError("Failed to create hidden window")

    def _handle_win32_event(self, hwnd, msg, wparam, lparam):
        if msg == win32con.WM_HOTKEY:
            if wparam == self.HOTKEY_ID:
                self._arm_capture()
            elif wparam == self.ESC_HOTKEY_ID:
                self._cancel_capture()
            return 0
        if msg == self.WM_APP_CAPTURE_COMPLETE:
            self._complete_capture()
            return 0
        if msg == self.WM_APP_EXIT:
            win32gui.DestroyWindow(hwnd)
            return 0
        if msg == win32con.WM_DESTROY:
            win32gui.PostQuitMessage(0)
            return 0
        return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)

    def _register_primary_hotkey(self) -> None:
        if not self._user32.RegisterHotKey(
            self._hwnd,
            self.HOTKEY_ID,
            win32con.MOD_ALT | win32con.MOD_SHIFT | MOD_NOREPEAT,
            ord("A"),
        ):
            raise RuntimeError("Failed to register hotkey Alt+Shift+A")
        logging.info("Registered global hotkey Alt+Shift+A")

    def _register_escape_hotkey(self) -> None:
        self._user32.RegisterHotKey(
            self._hwnd,
            self.ESC_HOTKEY_ID,
            MOD_NOREPEAT,
            win32con.VK_ESCAPE,
        )

    def _unregister_hotkeys(self) -> None:
        with contextlib.suppress(Exception):
            self._user32.UnregisterHotKey(self._hwnd, self.HOTKEY_ID)
        with contextlib.suppress(Exception):
            self._user32.UnregisterHotKey(self._hwnd, self.ESC_HOTKEY_ID)

    def _start_tray_icon(self) -> None:
        image = self._build_tray_image()
        menu = pystray.Menu(
            pystray.MenuItem(f"SnapClip {__version__}", lambda *args, **kwargs: None, enabled=False),
            pystray.MenuItem("Exit", lambda icon, item: self.request_exit()),
        )
        self._icon = pystray.Icon("snapclip", image, "SnapClip", menu=menu)
        self._icon.run_detached()
        logging.info("Tray icon started")

    def _stop_tray_icon(self) -> None:
        if self._icon:
            logging.info("Stopping tray icon")
            self._icon.stop()
            self._icon = None

    def _build_tray_image(self) -> Image.Image:
        size = 64
        img = Image.new("RGBA", (size, size), (18, 18, 18, 255))
        draw = ImageDraw.Draw(img)
        draw.rectangle((10, 10, size - 11, size - 11), outline=(255, 255, 255, 255), width=3)
        draw.rectangle((22, 22, size - 23, size - 23), fill=(255, 64, 64, 255))
        return img

    def _arm_capture(self) -> None:
        if self._capture_session:
            logging.debug("Capture already armed")
            return
        self._capture_session = CaptureSession()
        if not self._install_mouse_hook():
            logging.error("Failed to install mouse hook")
            self._capture_session = None
            return
        self._register_escape_hotkey()
        logging.info("Capture armed")

    def _cancel_capture(self) -> None:
        if not self._capture_session:
            return
        logging.info("Capture canceled")
        self._end_capture_session()

    def _complete_capture(self) -> None:
        if not self._capture_session:
            return
        bounds = self._capture_session.bounds()
        if not bounds:
            logging.info("Selection too small or invalid; ignoring")
            self._end_capture_session()
            return
        left, top, width, height = bounds
        logging.info("Capturing region left=%s top=%s width=%s height=%s", left, top, width, height)
        success = False
        try:
            shot = self._mss.grab(
                {
                    "left": left,
                    "top": top,
                    "width": width,
                    "height": height,
                }
            )
            image = Image.frombytes("RGB", shot.size, shot.rgb)
            success = self._copy_to_clipboard(image)
        except Exception:
            logging.exception("Failed to capture selection")
        finally:
            self._end_capture_session()
        if success:
            logging.info("Capture copied to clipboard")
        else:
            logging.error("Failed to write capture to clipboard")

    def _install_mouse_hook(self) -> bool:
        if self._mouse_hook:
            return True
        pointer_size = ctypes.sizeof(ctypes.c_void_p)
        ulong_ptr = ctypes.c_ulonglong if pointer_size == 8 else ctypes.c_ulong

        class POINT(ctypes.Structure):
            _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]

        class MSLLHOOKSTRUCT(ctypes.Structure):
            _fields_ = [
                ("pt", POINT),
                ("mouseData", ctypes.c_ulong),
                ("flags", ctypes.c_ulong),
                ("time", ctypes.c_ulong),
                ("dwExtraInfo", ulong_ptr),
            ]

        def low_level_mouse_proc(nCode, wParam, lParam):
            if nCode == win32con.HC_ACTION and self._capture_session and lParam:
                info = ctypes.cast(lParam, ctypes.POINTER(MSLLHOOKSTRUCT)).contents
                pos = (info.pt.x, info.pt.y)
                if wParam == win32con.WM_LBUTTONDOWN:
                    self._capture_session.start = pos
                    self._capture_session.current = pos
                elif wParam == win32con.WM_MOUSEMOVE and self._capture_session.start:
                    self._capture_session.current = pos
                elif wParam == win32con.WM_LBUTTONUP and self._capture_session.start:
                    self._capture_session.final = pos
                    win32gui.PostMessage(self._hwnd, self.WM_APP_CAPTURE_COMPLETE, 0, 0)
            return self._user32.CallNextHookEx(self._mouse_hook, nCode, wParam, lParam)

        self._mouse_proc_ref = ctypes.WINFUNCTYPE(
            LRESULT, ctypes.c_int, wintypes.WPARAM, wintypes.LPARAM
        )(low_level_mouse_proc)
        hook = self._user32.SetWindowsHookExW(
            win32con.WH_MOUSE_LL,
            self._mouse_proc_ref,
            0,
            0,
        )
        if not hook:
            self._mouse_proc_ref = None
            logging.error("SetWindowsHookExW failed: %s", ctypes.GetLastError())
            return False
        self._mouse_hook = hook
        return True

    def _end_capture_session(self) -> None:
        self._remove_mouse_hook()
        self._capture_session = None
        with contextlib.suppress(Exception):
            self._user32.UnregisterHotKey(self._hwnd, self.ESC_HOTKEY_ID)

    def _remove_mouse_hook(self) -> None:
        if self._mouse_hook:
            self._user32.UnhookWindowsHookEx(self._mouse_hook)
            self._mouse_hook = None
        self._mouse_proc_ref = None

    def _copy_to_clipboard(self, image: Image.Image) -> bool:
        bmp = io.BytesIO()
        image.convert("RGB").save(bmp, "BMP")
        data = bmp.getvalue()[14:]
        bmp.close()
        for _ in range(5):
            try:
                win32clipboard.OpenClipboard()
                win32clipboard.EmptyClipboard()
                win32clipboard.SetClipboardData(win32con.CF_DIB, data)
                return True
            except pywintypes.error:
                time.sleep(0.05)
            finally:
                with contextlib.suppress(Exception):
                    win32clipboard.CloseClipboard()
        return False

    def _teardown(self) -> None:
        if self._autotest_timer:
            self._autotest_timer.cancel()
        self._cancel_capture()
        self._stop_tray_icon()
        self._unregister_hotkeys()
        if self._hwnd:
            with contextlib.suppress(Exception):
                win32gui.DestroyWindow(self._hwnd)
            self._hwnd = None
        if self._class_atom:
            with contextlib.suppress(Exception):
                win32gui.UnregisterClass(self._class_name, self._instance)
            self._class_atom = None
        if self._mss:
            with contextlib.suppress(Exception):
                self._mss.close()
            self._mss = None
        logging.info("SnapClip shutdown complete")


def main() -> None:
    _setup_logging()
    app = SnapClipApp()
    try:
        app.run()
    except KeyboardInterrupt:
        app.request_exit()


__all__ = ["main", "SnapClipApp"]
