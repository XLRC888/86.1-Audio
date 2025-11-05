print("Smart 86.1 Audio by XLRC888")
print("WARNING: Make sure to close the code via CTRL+C otherwise audio may stay altered.")
print("Mode: Automatic by default. Press Ctrl+Left to force LEFT, Ctrl+Right to force RIGHT, Ctrl+Up to return to AUTO.")

import win32gui
import win32api
import time
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
import atexit

TARGET_APPS = ["chrome", "brave", "zen", "msedge", "opera", "opera_gx", "firefox"]


def find_target_window():
    """Find any visible top-level window whose title contains one of TARGET_APPS.
    Returns (hwnd, title) or (None, "") if not found.
    This allows the script to act on background windows (not only the focused one).
    """
    matches = []

    def _enum(hwnd, _):
        try:
            if not win32gui.IsWindowVisible(hwnd):
                return True
            title = win32gui.GetWindowText(hwnd).lower()
            if not title:
                return True
            for app in TARGET_APPS:
                if app in title:
                    matches.append((hwnd, title))
                    break
        except Exception:
            pass
        return True

    win32gui.EnumWindows(_enum, None)
    if matches:
        return matches[0]
    return (None, "")

def reset_audio():
    set_audio_pan(1.0, 1.0)

def set_audio_pan(left_vol, right_vol):
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    volume.SetChannelVolumeLevelScalar(0, left_vol, None)
    volume.SetChannelVolumeLevelScalar(1, right_vol, None)

def get_window_zone(hwnd):
    rect = win32gui.GetWindowRect(hwnd)
    screen_w = win32api.GetSystemMetrics(0)

    center_x = ((rect[0] + rect[2]) / 2) / screen_w

    return "left" if center_x < 0.5 else "right"

def adjust_audio_by_zone(zone):
    if zone == "left":
        set_audio_pan(1.0, 0.0)
    elif zone == "right":
        set_audio_pan(0.0, 1.0)
    else:
        set_audio_pan(1.0, 1.0)

def main():
    atexit.register(reset_audio)
    print("Smart 86.1 Audio active (Safely exit: Ctrl+C)")
    
    try:
        manual_mode = None

        prev_left = prev_right = prev_up = 0

        while True:
            left_state = win32api.GetAsyncKeyState(0x25)
            right_state = win32api.GetAsyncKeyState(0x27)
            up_state = win32api.GetAsyncKeyState(0x26)
            ctrl_down = (win32api.GetAsyncKeyState(0xA2) & 0x8000) or (win32api.GetAsyncKeyState(0xA3) & 0x8000) or (win32api.GetAsyncKeyState(0x11) & 0x8000)

            if ctrl_down and (left_state & 0x8000) and not (prev_left & 0x8000):
                manual_mode = 'left'
                print("Manual mode: LEFT (Ctrl+Left detected)" + ' ' * 20)
            if ctrl_down and (right_state & 0x8000) and not (prev_right & 0x8000):
                manual_mode = 'right'
                print("Manual mode: RIGHT (Ctrl+Right detected)" + ' ' * 20)
            if ctrl_down and (up_state & 0x8000) and not (prev_up & 0x8000):
                manual_mode = None
                print("Automatic mode: using window position" + ' ' * 20)

            prev_left, prev_right, prev_up = left_state, right_state, up_state

            hwnd, title = find_target_window()

            if any and manual_mode is not None:
                zone = manual_mode
                print(f"Manual: {zone.upper()}" + ('  |  App: ' + title if title else ''), end='\r')
                adjust_audio_by_zone(zone)
            else:
                if hwnd:
                    zone = get_window_zone(hwnd)
                    print(f"Location: {zone.upper()}  |  App: {title}", end='\r')
                    adjust_audio_by_zone(zone)
                else:
                    set_audio_pan(1.0, 1.0)

            time.sleep(0.15)
            
    except KeyboardInterrupt:
        print("\nProgram terminated")
    finally:
        reset_audio()

if __name__ == "__main__":
    main()
