import time
import datetime
import pywintypes
import base64

NO_DATE = 4501

def check_type(value):
    if value == None: return True
    if isinstance(value, (bool, int, float, str, unicode, buffer)): return True
    if type(value).__name__ == "time": return True
    return False

def fix_type(value):
    if type(value).__name__ == "time":
        if value.year == NO_DATE: return None
        value = datetime.datetime(year=value.year, month=value.month, day=value.day, hour=value.hour,
            minute=value.minute,second=value.second)
    return value
