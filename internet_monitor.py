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

def insert_row_with_timestamp(file_name, message):
    try:
        workbook = openpyxl.load_workbook(file_name)
    except FileNotFoundError:
        workbook = Workbook()
        workbook.active.append(["Timestamp", "Message"])  # Adding header row
    
    sheet = workbook.active
    
    current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    new_row = [current_time, message]
    
    sheet.append(new_row)
    
    workbook.save(file_name)

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
