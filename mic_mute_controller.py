import win32api
import win32gui
import keyboard
import win32con
from win32gui import EnumWindows, GetWindowText, GetClassName, GetForegroundWindow, SendMessage
import ctypes

target_application_hwnd = None

def on_keyevent(event):
    print(f'{event.name} {event.event_type} {event.scan_code:X} {event.scan_code}')

def toggle_mic_mute():
    global target_application_hwnd
    print('toggle mic mute')
    WM_APPCOMMAND = 0x319
    APPCOMMAND_MICROPHONE_VOLUME_MUTE = 0x180000
    hwnd_active = GetForegroundWindow() # VOLUME_MUTEを送付する先はForeground Windowで無くても良いかも。
    win32api.SendMessage(hwnd_active, WM_APPCOMMAND, None, APPCOMMAND_MICROPHONE_VOLUME_MUTE)
    if target_application_hwnd is not None:
        activate_application(target_application_hwnd)
    else:
        print('Window is not registered')

def activate_application(hwnd):
    print(f'Activate : {GetWindowText(hwnd)}')
    if win32gui.IsIconic(hwnd):
        win32gui.ShowWindow(hwnd, 1)
    win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE + win32con.SWP_NOSIZE + win32con.SWP_SHOWWINDOW)
    win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE + win32con.SWP_NOSIZE + win32con.SWP_SHOWWINDOW)
    ctypes.windll.user32.SetForegroundWindow(hwnd) # win32guiのSetForegroundWindowが動かなかったので暫定処置

def register_meeting_window():
    global target_application_hwnd
    hwnd_active = GetForegroundWindow()
    target_application_hwnd = hwnd_active
    print(f'Registered : {GetWindowText(hwnd_active)} {GetClassName(hwnd_active)}')

if __name__ == '__main__':
    keyboard.add_hotkey('insert', toggle_mic_mute, suppress=True)
    keyboard.add_hotkey('shift+insert', register_meeting_window, suppress=True)
    # keyboard.hook(on_keyevent) # for keycode check

    while True:
        keyboard.wait()
