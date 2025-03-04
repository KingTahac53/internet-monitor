import openpyxl
from openpyxl import Workbook
from datetime import datetime
import requests
import time


def check_internet_connection(url='http://www.google.com/', timeout=5):
    try:
        _ = requests.get(url, timeout=timeout)
        return True
    except requests.ConnectionError:
        return False


def insert_row_with_timestamp(file_path, status_message):
    if not os.path.exists(file_path):
        # Create a new Excel file with headers if it doesn't exist
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Logs"
        sheet.append(["Timestamp", "Status Message"])  # Add headers
        workbook.save(file_path)

    try:
        # Load the workbook and add the log entry
        workbook = load_workbook(file_path)
        sheet = workbook.active
        from datetime import datetime
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        sheet.append([timestamp, status_message])
        workbook.save(file_path)
    except Exception as e:
        print(f"Error while accessing the Excel file: {e}")


def monitor_internet_connection(file_name, check_interval=10):
    was_connected = None  # Start with unknown state

    while True:
        current_state = check_internet_connection()

        if current_state and was_connected is False:
            insert_row_with_timestamp(file_name, "Internet Connected")

        if not current_state and was_connected is not False:
            insert_row_with_timestamp(file_name, "Internet Disconnected")

        was_connected = current_state
        time.sleep(check_interval)


if __name__ == "__main__":
    file_name = 'internet_disconnection_log.xlsx'
    monitor_internet_connection(file_name)
