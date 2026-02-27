# -*- coding: utf-8 -*-
# General.py

import traceback
import sys

def log_error():
    try:
        exc_type, exc_value, exc_traceback = sys.exc_info()
        filename = exc_traceback.tb_frame.f_code.co_filename
        lineno = exc_traceback.tb_lineno
        error_massage = """
        [ERROR] Exception Occurred:
        - File: {}
        - Line: {}
        - Error Type: {}
        - Message: {}
        """.format(filename, lineno, exc_type.__name__, str(exc_value))
        print(error_massage)
        traceback.print_exc()
    except Exception as e:
        print("log_error failed with exception {}".format(e))


