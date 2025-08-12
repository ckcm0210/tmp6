
import os
import datetime

def get_timestamp():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def format_message(message):
    return f"[{get_timestamp()}] {message}"

def open_external_file(file_path):
    try:
        if os.path.exists(file_path):
            os.startfile(file_path)
    except Exception:
        pass
