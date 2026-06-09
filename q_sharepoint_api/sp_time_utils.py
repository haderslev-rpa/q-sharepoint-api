# q_sharepoint_api/sp_time_utils.py

from datetime import datetime
from zoneinfo import ZoneInfo  # modul (tidszoner)


def dk_timestamp():
    """
    Returnerer dansk timestamp: DD-MM-YYYY HH:MM
    """
    now = datetime.now(ZoneInfo("Europe/Copenhagen"))
    return now.strftime("%d-%m-%Y %H:%M")