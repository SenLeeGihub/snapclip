# SnapClip - Windows stealth screenshot to clipboard (no UI)

## Purpose
A minimal, stealth screenshot tool for Windows.
Runs in background (no window) with a tray icon.
User presses Alt+Shift+A, then selects a region by mouse drag (blind selection).
On mouse release, capture the region and put image into Windows clipboard so Ctrl+V pastes it anywhere.

## Hard Requirements
1. No main window. Only system tray icon.
2. Global hotkey Alt+Shift+A arms capture mode.
3. In capture mode: track left-mouse drag (down -> move -> up) to define rectangle.
4. During capture: NO overlay, NO rectangle, NO tooltip, NO dialog, NO toast. Absolutely no on-screen UI.
5. On left button up: capture region immediately and write to Clipboard as an image (Ctrl+V works in Word/WeChat/PPT/Paint).
6. ESC cancels capture mode and does NOT modify clipboard.
7. Very small selections (<3x3) are treated as cancel.
8. Tray menu: Exit/Quit (and optional version item that does not open dialogs).

## Platform
- Windows 10/11
- Python 3.11+ (tested on 3.12)

## Allowed Libraries
- pywin32 (RegisterHotKey, message loop, clipboard)
- mss (screen capture)
- Pillow (image processing)
- pystray (tray icon)

## Architecture Notes
- Use Win32 message loop for hotkey and for receiving hook callbacks cleanly.
- Capture mode uses WH_MOUSE_LL low-level mouse hook to record:
  - start point on WM_LBUTTONDOWN
  - end point on WM_LBUTTONUP
  - optionally update current point on WM_MOUSEMOVE (not required)
- Ensure proper unhook on cancel/complete/exit.
- Multi-monitor: handle virtual screen coordinates correctly (may be negative).
- Logging: optional file logging to %TEMP%\snapclip.log; must stay silent (no UI).

## Entrypoint
- Must run via: `python -m snapclip`
