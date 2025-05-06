import win32gui
import win32process
import win32con
import psutil
import sys
import ctypes
import tkinter as tk
from tkinter import ttk, messagebox


class WindowResizerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Window Resizer")

        # Add icon to the window
        try:
            self.root.iconbitmap("icon.ico")
        except tk.TclError:
            print("Warning: Icon file not found")

        self.root.geometry("500x500")
        self.root.resizable(True, False)

        # Check admin privileges
        self.is_admin = ctypes.windll.shell32.IsUserAnAdmin()
        if not self.is_admin:
            self.show_admin_warning()

        # Create main frame
        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Process name entry
        ttk.Label(main_frame, text="Process Name:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.process_entry = ttk.Entry(main_frame, width=30)
        self.process_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5)
        ttk.Button(main_frame, text="Find Windows", command=self.find_windows).grid(row=0, column=2, padx=5, pady=5)

        # Windows list
        ttk.Label(main_frame, text="Available Windows:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.window_list_frame = ttk.Frame(main_frame)
        self.window_list_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)

        # Create a treeview for windows list
        self.windows_tree = ttk.Treeview(self.window_list_frame, columns=('title',), show='headings')
        self.windows_tree.heading('title', text='Window Title')
        self.windows_tree.column('title', width=350)
        self.windows_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Add a scrollbar
        scrollbar = ttk.Scrollbar(self.window_list_frame, orient=tk.VERTICAL, command=self.windows_tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.windows_tree.configure(yscrollcommand=scrollbar.set)

        # Resize options
        resize_frame = ttk.LabelFrame(main_frame, text="Resize Options", padding="10")
        resize_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)

        ttk.Label(resize_frame, text="Width:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.width_entry = ttk.Entry(resize_frame, width=10)
        self.width_entry.grid(row=0, column=1, sticky=tk.W, pady=5, padx=5)

        ttk.Label(resize_frame, text="Height:").grid(row=0, column=2, sticky=tk.W, pady=5)
        self.height_entry = ttk.Entry(resize_frame, width=10)
        self.height_entry.grid(row=0, column=3, sticky=tk.W, pady=5, padx=5)

        ttk.Button(resize_frame, text="Get Current Size", command=self.get_current_size).grid(row=0, column=4,
                                                                                              sticky=tk.W, pady=5,
                                                                                              padx=5)

        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)

        ttk.Button(button_frame, text="Resize Window", command=self.resize_selected_window).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Exit", command=root.destroy).pack(side=tk.RIGHT, padx=5)

        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        status_bar = ttk.Label(root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # Store windows data
        self.windows = []

    def show_admin_warning(self):
        messagebox.showwarning(
            "Administrator Privileges",
            "Note: Some windows may require administrator privileges to resize.\n"
            "Consider running this application as administrator if it doesn't work."
        )

    def find_windows_by_process_name(self, process_name):
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

    def find_windows(self):
        """Find windows based on the entered process name."""
        process_name = self.process_entry.get().strip()
        if not process_name:
            messagebox.showerror("Error", "Please enter a process name")
            return

        # Clear existing items in the treeview
        for item in self.windows_tree.get_children():
            self.windows_tree.delete(item)

        # Find windows for the process
        self.windows = self.find_windows_by_process_name(process_name)

        if not self.windows:
            self.status_var.set(f"No windows found for process: {process_name}")
            return

        # Add windows to the treeview
        for i, (hwnd, title) in enumerate(self.windows):
            self.windows_tree.insert('', tk.END, values=(title,), iid=str(i))

        self.status_var.set(f"Found {len(self.windows)} window(s) for process: {process_name}")

    def get_current_size(self):
        """Get the current size of the selected window."""
        selected_items = self.windows_tree.selection()
        if not selected_items:
            messagebox.showerror("Error", "Please select a window from the list")
            return

        selected_idx = int(selected_items[0])
        hwnd = self.windows[selected_idx][0]

        # Get current window position
        x, y, right, bottom = win32gui.GetWindowRect(hwnd)
        current_width = right - x
        current_height = bottom - y

        # Set the values in the entry fields
        self.width_entry.delete(0, tk.END)
        self.width_entry.insert(0, str(current_width))

        self.height_entry.delete(0, tk.END)
        self.height_entry.insert(0, str(current_height))

        self.status_var.set(f"Current window size: {current_width}x{current_height}")

    def resize_window(self, hwnd, width, height):
        """Resize a window to the specified width and height."""
        # Get current window position
        x, y, right, bottom = win32gui.GetWindowRect(hwnd)

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

    def resize_selected_window(self):
        """Resize the selected window with the specified dimensions."""
        selected_items = self.windows_tree.selection()
        if not selected_items:
            messagebox.showerror("Error", "Please select a window from the list")
            return

        try:
            width = int(self.width_entry.get().strip())
            height = int(self.height_entry.get().strip())

            if width <= 0 or height <= 0:
                messagebox.showerror("Error", "Width and height must be positive integers")
                return

            selected_idx = int(selected_items[0])
            hwnd = self.windows[selected_idx][0]

            # Resize window
            self.resize_window(hwnd, width, height)

            self.status_var.set(f"Window resized successfully to {width}x{height}")

        except ValueError:
            messagebox.showerror("Error", "Width and height must be valid integers")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to resize window: {str(e)}")


def main():
    root = tk.Tk()
    app = WindowResizerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()