import win32gui
import win32process
import win32con
import psutil
import sys
import ctypes
import tkinter as tk
from tkinter import ttk, messagebox
import json
import os


class WindowResizerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Window Resizer")

        # Configuration file path
        self.config_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")

        # Default settings
        self.settings = {
            "dark_mode": False,
            "window_width": 700,
            "window_height": 600
        }

        # Load settings
        self.load_settings()

        # Add icon to the window
        try:
            self.root.iconbitmap("icon.ico")
        except tk.TclError:
            print("Warning: Icon file not found")

        self.root.geometry(f"{self.settings['window_width']}x{self.settings['window_height']}")
        self.root.resizable(True, False)

        # Check admin privileges
        self.is_admin = ctypes.windll.shell32.IsUserAnAdmin()
        if not self.is_admin:
            self.show_admin_warning()

        # Create style object
        self.style = ttk.Style()

        # Create main frame
        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Theme toggle frame at the top
        theme_frame = ttk.Frame(main_frame)
        theme_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 5))

        # Theme toggle
        self.theme_var = tk.BooleanVar(value=self.settings["dark_mode"])
        self.theme_check = ttk.Checkbutton(
            theme_frame,
            text="Dark Mode",
            variable=self.theme_var,
            command=self.toggle_theme
        )
        self.theme_check.pack(side=tk.RIGHT)

        # Process filter frame
        filter_frame = ttk.Frame(main_frame)
        filter_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)

        # Process name filter
        ttk.Label(filter_frame, text="Filter by Process Name:").pack(side=tk.LEFT, padx=(0, 5))
        self.process_entry = ttk.Entry(filter_frame, width=30)
        self.process_entry.pack(side=tk.LEFT, padx=(0, 5))

        ttk.Button(filter_frame, text="Apply Filter", command=self.apply_filter).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(filter_frame, text="Show All Windows", command=self.find_all_windows).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(filter_frame, text="Refresh", command=self.refresh_windows).pack(side=tk.LEFT, padx=(0, 5))

        # Windows list
        ttk.Label(main_frame, text="Available Windows:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.window_list_frame = ttk.Frame(main_frame)
        self.window_list_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)

        # Create a treeview for windows list with process information
        self.windows_tree = ttk.Treeview(self.window_list_frame,
                                         columns=('title', 'process_name', 'pid'),
                                         show='headings')
        self.windows_tree.heading('title', text='Window Title')
        self.windows_tree.heading('process_name', text='Process Name')
        self.windows_tree.heading('pid', text='PID')
        self.windows_tree.column('title', width=350)
        self.windows_tree.column('process_name', width=150)
        self.windows_tree.column('pid', width=80)
        self.windows_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Add a scrollbar
        scrollbar = ttk.Scrollbar(self.window_list_frame, orient=tk.VERTICAL, command=self.windows_tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.windows_tree.configure(yscrollcommand=scrollbar.set)

        # Resize options
        resize_frame = ttk.LabelFrame(main_frame, text="Window Size and Position", padding="10")
        resize_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)

        # Size inputs (width and height)
        ttk.Label(resize_frame, text="Width:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.width_entry = ttk.Entry(resize_frame, width=10)
        self.width_entry.grid(row=0, column=1, sticky=tk.W, pady=5, padx=5)

        ttk.Label(resize_frame, text="Height:").grid(row=0, column=2, sticky=tk.W, pady=5)
        self.height_entry = ttk.Entry(resize_frame, width=10)
        self.height_entry.grid(row=0, column=3, sticky=tk.W, pady=5, padx=5)

        # Position inputs (X and Y coordinates)
        ttk.Label(resize_frame, text="X Position:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.x_entry = ttk.Entry(resize_frame, width=10)
        self.x_entry.grid(row=1, column=1, sticky=tk.W, pady=5, padx=5)

        ttk.Label(resize_frame, text="Y Position:").grid(row=1, column=2, sticky=tk.W, pady=5)
        self.y_entry = ttk.Entry(resize_frame, width=10)
        self.y_entry.grid(row=1, column=3, sticky=tk.W, pady=5, padx=5)

        ttk.Button(resize_frame, text="Get Current Properties", command=self.get_current_properties).grid(
            row=1, column=4, sticky=tk.W, pady=5, padx=5)

        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)

        ttk.Button(button_frame, text="Apply Changes", command=self.modify_selected_window).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Exit", command=self.on_exit).pack(side=tk.RIGHT, padx=5)

        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        self.status_bar = ttk.Label(root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # Store windows data
        self.windows = []
        self.all_windows = []

        # apply theme
        self.apply_theme()

        # Load all windows on startup
        self.find_all_windows()

        # Register closing event
        self.root.protocol("WM_DELETE_WINDOW", self.on_exit)

    def load_settings(self):
        """Load settings from the configuration file."""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r') as f:
                    loaded_settings = json.load(f)
                    # Update default settings with loaded values
                    self.settings.update(loaded_settings)
        except Exception as e:
            print(f"Error loading settings: {e}")

    def save_settings(self):
        """Save current settings to the configuration file."""
        try:
            # Update window size in settings
            if self.root.winfo_width() > 0 and self.root.winfo_height() > 0:
                self.settings["window_width"] = self.root.winfo_width()
                self.settings["window_height"] = self.root.winfo_height()

            # Update dark mode setting
            self.settings["dark_mode"] = self.theme_var.get()

            # Save settings to file
            with open(self.config_file, 'w') as f:
                json.dump(self.settings, f, indent=4)
        except Exception as e:
            print(f"Error saving settings: {e}")

    def apply_theme(self):
        """Apply light or dark theme to the application."""
        if self.settings["dark_mode"]:
            # Dark theme
            self.style.theme_use('clam')  # base theme

            # Configure colors for dark mode
            self.style.configure('TFrame', background='#2E2E2E')
            self.style.configure('TLabel', background='#2E2E2E', foreground='#FFFFFF')
            self.style.configure('TButton', background='#555555', foreground='#FFFFFF')
            self.style.configure('TCheckbutton', background='#2E2E2E', foreground='#FFFFFF')
            self.style.configure('TLabelframe', background='#2E2E2E', foreground='#FFFFFF')
            self.style.configure('TLabelframe.Label', background='#2E2E2E', foreground='#FFFFFF')
            self.style.configure('TEntry', fieldbackground='#3E3E3E', foreground='#FFFFFF')

            # Configure Treeview colors
            self.style.configure('Treeview',
                                 background='#3E3E3E',
                                 fieldbackground='#3E3E3E',
                                 foreground='#FFFFFF')
            self.style.map('Treeview',
                           background=[('selected', '#4A6984')],
                           foreground=[('selected', '#FFFFFF')])

            # Root window and menu
            self.root.configure(background='#2E2E2E')

            # Configure status bar if it exists
            if hasattr(self, 'status_bar'):
                self.status_bar.configure(background='#333333', foreground='#FFFFFF')

        else:
            # Light theme (default)
            self.style.theme_use('clam')  # Reset to default

            # Configure colors for light mode
            self.style.configure('TFrame', background='#F0F0F0')
            self.style.configure('TLabel', background='#F0F0F0', foreground='#000000')
            self.style.configure('TButton', background='#E0E0E0', foreground='#000000')
            self.style.configure('TCheckbutton', background='#F0F0F0', foreground='#000000')
            self.style.configure('TLabelframe', background='#F0F0F0', foreground='#000000')
            self.style.configure('TLabelframe.Label', background='#F0F0F0', foreground='#000000')
            self.style.configure('TEntry', fieldbackground='#FFFFFF', foreground='#000000')

            # Configure Treeview colors
            self.style.configure('Treeview',
                                 background='#FFFFFF',
                                 fieldbackground='#FFFFFF',
                                 foreground='#000000')
            self.style.map('Treeview',
                           background=[('selected', '#0078D7')],
                           foreground=[('selected', '#FFFFFF')])

            # Root window and menu
            self.root.configure(background='#F0F0F0')

            # Configure status bar if it exists
            if hasattr(self, 'status_bar'):
                self.status_bar.configure(background='#F0F0F0', foreground='#000000')

    def toggle_theme(self):
        """Toggle between light and dark theme."""
        self.settings["dark_mode"] = self.theme_var.get()
        self.apply_theme()
        self.save_settings()

        theme_name = "Dark" if self.settings["dark_mode"] else "Light"
        self.status_var.set(f"{theme_name} theme applied")

    def show_admin_warning(self):
        messagebox.showwarning(
            "Administrator Privileges",
            "Note: Some windows may require administrator privileges to resize.\n"
            "Consider running this application as administrator if it doesn't work."
        )

    def find_all_windows(self):
        """Find all visible windows with their process information."""
        self.status_var.set("Finding all visible windows...")

        # Clear existing items in the treeview
        for item in self.windows_tree.get_children():
            self.windows_tree.delete(item)

        self.all_windows = []

        # Callback function for EnumWindows
        def enum_windows_callback(hwnd, results):
            if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd):
                # Get process ID for this window
                _, process_id = win32process.GetWindowThreadProcessId(hwnd)
                window_title = win32gui.GetWindowText(hwnd)

                try:
                    # Get process name by ID
                    process = psutil.Process(process_id)
                    process_name = process.name()
                    results.append((hwnd, window_title, process_name, process_id))
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    # If we can't get process information, still add the window
                    results.append((hwnd, window_title, "Unknown", process_id))

            return True

        win32gui.EnumWindows(enum_windows_callback, self.all_windows)

        # Update the windows list with all windows
        self.windows = self.all_windows
        self.update_windows_display()

        self.status_var.set(f"Found {len(self.windows)} visible windows")

    def apply_filter(self):
        """Apply filter based on the entered process name."""
        process_name = self.process_entry.get().strip().lower()
        if not process_name:
            # If no process name is provided show all windows
            self.find_all_windows()
            return

        # Filter windows by process name
        self.windows = [
            window for window in self.all_windows
            if process_name in window[2].lower()
        ]

        # Update the display
        self.update_windows_display()

        self.status_var.set(f"Found {len(self.windows)} windows matching '{process_name}'")

    def refresh_windows(self):
        """Refresh the windows list."""
        self.find_all_windows()

        # Re-apply filter if one exists
        process_name = self.process_entry.get().strip()
        if process_name:
            self.apply_filter()

    def update_windows_display(self):
        """Update the windows display in the treeview."""
        # Clear existing items in the treeview
        for item in self.windows_tree.get_children():
            self.windows_tree.delete(item)

        # Add windows to the treeview
        for i, (hwnd, title, process_name, pid) in enumerate(self.windows):
            self.windows_tree.insert('', tk.END, values=(title, process_name, pid), iid=str(i))

    def get_current_properties(self):
        """Get the current properties (size and position) of the selected window."""
        selected_items = self.windows_tree.selection()
        if not selected_items:
            messagebox.showerror("Error", "Please select a window from the list")
            return

        selected_idx = int(selected_items[0])
        hwnd = self.windows[selected_idx][0]

        # Get current window position and size
        x, y, right, bottom = win32gui.GetWindowRect(hwnd)
        current_width = right - x
        current_height = bottom - y

        # Set the values in the entry fields
        self.width_entry.delete(0, tk.END)
        self.width_entry.insert(0, str(current_width))

        self.height_entry.delete(0, tk.END)
        self.height_entry.insert(0, str(current_height))

        # Set the X and Y coordinate values
        self.x_entry.delete(0, tk.END)
        self.x_entry.insert(0, str(x))

        self.y_entry.delete(0, tk.END)
        self.y_entry.insert(0, str(y))

        self.status_var.set(f"Current window size: {current_width}x{current_height}, Position: ({x}, {y})")

    def modify_window(self, hwnd, width=None, height=None, x=None, y=None):
        """Modify a window's size and/or position."""
        # Get current window position and size
        curr_x, curr_y, right, bottom = win32gui.GetWindowRect(hwnd)
        curr_width = right - curr_x
        curr_height = bottom - curr_y

        # Use current values if new ones are not provided
        width = width if width is not None else curr_width
        height = height if height is not None else curr_height
        x = x if x is not None else curr_x
        y = y if y is not None else curr_y

        # Apply changes to the window
        return win32gui.SetWindowPos(
            hwnd,
            win32con.HWND_TOP,  # z-order
            x,  # x position
            y,  # y position
            width,  # width
            height,  # height
            win32con.SWP_NOZORDER  # flags (don't change Z-order)
        )

    def modify_selected_window(self):
        """Modify the selected window with the specified dimensions and position."""
        selected_items = self.windows_tree.selection()
        if not selected_items:
            messagebox.showerror("Error", "Please select a window from the list")
            return

        try:
            # Get size values (width and height)
            width = int(self.width_entry.get().strip()) if self.width_entry.get().strip() else None
            height = int(self.height_entry.get().strip()) if self.height_entry.get().strip() else None

            # Get position values (x and y)
            x = int(self.x_entry.get().strip()) if self.x_entry.get().strip() else None
            y = int(self.y_entry.get().strip()) if self.y_entry.get().strip() else None

            # Validate values if provided
            if width is not None and width <= 0:
                messagebox.showerror("Error", "Width must be a positive integer")
                return

            if height is not None and height <= 0:
                messagebox.showerror("Error", "Height must be a positive integer")
                return

            selected_idx = int(selected_items[0])
            hwnd = self.windows[selected_idx][0]

            # Modify window size and position
            self.modify_window(hwnd, width, height, x, y)

            # Construct status message
            status_parts = []
            if width is not None and height is not None:
                status_parts.append(f"size: {width}x{height}")
            if x is not None and y is not None:
                status_parts.append(f"position: ({x}, {y})")

            status_message = f"Window updated - " + ", ".join(status_parts)
            self.status_var.set(status_message)

        except ValueError:
            messagebox.showerror("Error", "Width, height, X and Y values must be valid integers")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to modify window: {str(e)}")

    def on_exit(self):
        """Handle application exit."""
        self.save_settings()
        self.root.destroy()

    # Legacy method for compatibility
    def resize_selected_window(self):
        """Legacy method redirecting to modify_selected_window."""
        self.modify_selected_window()


def main():
    root = tk.Tk()
    app = WindowResizerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()