import os
import pandas as pd
import pynetbox
import requests
import re
import urllib3
import warnings
from urllib3.exceptions import InsecureRequestWarning
from openpyxl import load_workbook
import config
from datetime import datetime


filepath = config.filepath
NetBox_URL = config.NetBox_URL
NetBox_Token = config.NetBox_Token
sitename = config.sitename
sheetname = config.sheetname


def main():
    try:
        device_year_of_investment = "2022"
        # convert string to data time format YYYY-MM-DD HH:MM:SS
        date_object = datetime.strptime(device_year_of_investment, "%m/%d/%Y")
        formatted_date = date_object.strftime("%Y-%m-%d %H:%M:%S")
        device_year_of_investment = formatted_date

    except Exception as e:
        print(f"Error during execution: {e}")

if __name__ == "__main__":
    main()

