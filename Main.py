import win32gui
import win32process
import win32con
import psutil
import sys
import ctypes


def find_windows_by_process_name(process_name):
    """Find all window handles for a given process name."""
    result = []

    # Callback function for EnumWindows
    def enum_windows_callback(hwnd, results):
        if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd):
            # Get process ID for this window
            _, process_id = win32process.GetWindowThreadProcessId(hwnd)
            try:
                # Get process name by ID
                process = psutil.Process(process_id)
                if process.name().lower() == process_name.lower() or \
                        process.name().lower() == f"{process_name.lower()}.exe":
                    results.append((hwnd, win32gui.GetWindowText(hwnd)))
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass
        return True

    windows = []
    win32gui.EnumWindows(enum_windows_callback, windows)
    return windows


def resize_window(hwnd, width, height):
    """Resize a window to the specified width and height."""
    # Get current window position
    x, y, right, bottom = win32gui.GetWindowRect(hwnd)
    current_width = right - x
    current_height = bottom - y

    print(f"Current window size: {current_width}x{current_height}")

    # Resize window, keeping the same position
    return win32gui.SetWindowPos(
        hwnd,
        win32con.HWND_TOP,  # z-order
        x,  # x position
        y,  # y position
        width,  # new width
        height,  # new height
        win32con.SWP_NOZORDER  # flags (don't change Z-order)
    )


def main():
    # Check if running as admin
    if not ctypes.windll.shell32.IsUserAnAdmin():
        print("Note: Some windows may require administrator privileges to resize.")
        print("Consider running this script as administrator if it doesn't work.")
        print()

    # Get process name from user
    process_name = input("Enter process name (e.g., notepad): ")

    # Find windows for the process
    windows = find_windows_by_process_name(process_name)

    if not windows:
        print(f"No windows found for process: {process_name}")
        return

    # Display found windows
    print(f"Found {len(windows)} window(s):")
    for i, (_, title) in enumerate(windows):
        print(f"{i}: {title}")

    # Select window if multiple found
    selected_index = 0
    if len(windows) > 1:
        selected_index = int(input("Enter window index to resize: "))

        if selected_index < 0 or selected_index >= len(windows):
            print("Invalid index")
            return

    target_window = windows[selected_index][0]

    # Get new dimensions
    width = int(input("Enter new width: "))
    height = int(input("Enter new height: "))

    # Resize window
    try:
        resize_window(target_window, width, height)
        print(f"Window resized successfully to {width}x{height}")
    except Exception as e:
        print(f"Failed to resize window: {e}")


if __name__ == "__main__":
    main()
