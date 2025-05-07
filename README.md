# Window Resizer

A Python application that allows you to view all visible windows and resize them. This tool is useful for managing and manipulating application windows on Windows operating systems.

## Features

- Display all visible windows with their process names and PIDs
- Filter windows by process name
- Get and modify window properties (size and position)
- Resize and reposition windows directly from the application
- Refresh window list on demand

## Screenshots

(Add screenshots of your application here)

## Requirements

- Windows operating system
- Python 3.6+
- Required Python packages:
  - psutil
  - pywin32

## Installation

1. Clone this repository:
```bash
git clone https://github.com/allenmonkey970/window-resizer.git
cd window-resizer
```

2. Install required packages:
```bash
pip install -r requirements.txt
```

## Usage

Run the application:
```bash
python window_resizer.py
```

### Basic Operations:
1. When launched, the application will automatically display all visible windows
2. To filter windows, enter a process name in the filter box and click "Apply Filter"
3. To resize a window:
   - Select it from the list
   - Click "Get Current Properties"
   - Modify the width, height, X, or Y values
   - Click "Apply Changes"

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.
