# InternetMonitor

This script monitors your internet connection and logs disconnection and reconnection events into an Excel file with timestamps.

## Features
- Logs the exact date and time when the internet connection is lost and restored.
- Saves logs into an Excel file (`internet_disconnection_log.xlsx`).

## Requirements
- Python 3.x
- `openpyxl` library
- `requests` library
- `pyinstaller` (for creating a standalone executable)

## Installation

### 1. Clone the Repository
```bash
git clone https://github.com/yourusername/internet-monitor.git
cd internet-monitor

### 2. Install Dependencies
```bash
pip install openpyxl requests
Usage

### Running the Script
To run the script directly:

```bash
python internet_monitor.py
Creating a Standalone Executable
1. Install pyinstaller
```bash
pip install pyinstaller
2. Create the Executable
```bash
pyinstaller --onefile --noconsole internet_monitor.py
3. Move the Executable
After running the above command, you will find the executable file in the dist directory inside your project folder. Move this file to your desired location.

Running the Executable in the Background
Create a Shortcut:

Right-click on the executable and select "Create shortcut".
Move the Shortcut to the Startup Folder:

Press Win + R, type shell:startup, and press Enter. This opens the Startup folder.
Move the shortcut you created into this folder.
This ensures that the executable will run automatically in the background whenever your computer starts.

Contributing
Feel free to submit issues or pull requests if you have any suggestions or improvements.

License
This project is licensed under the MIT License.