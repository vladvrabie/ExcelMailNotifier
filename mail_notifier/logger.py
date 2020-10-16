import os
import pathlib
from datetime import datetime

INFO = 'INFO'
ERROR = 'ERROR'

FILE_NAME = "log_file.log"
MAX_FILE_SIZE = 5000000  # bytes; roughly 5 MB


def _half_if_necessary(file_path):
    if os.path.exists(file_path):
        byte_size = os.path.getsize(file_path)

        if byte_size > MAX_FILE_SIZE:
            with open(file_path, 'r') as logging_file:
                lines = logging_file.readlines()
                number_of_lines = len(lines)
                middle_line = number_of_lines // 2

            with open(file_path, 'w+') as logging_file:
                logging_file.writelines(lines[middle_line:])


def log(message, message_type=INFO):
    home = str(pathlib.Path.home())
    log_path = os.path.join(home, FILE_NAME)
    _half_if_necessary(log_path)
    with open(log_path, 'a') as logging_file:
        logging_file.write(
            f'[{str(datetime.today())}] {message_type}: {message}\n'
        )
