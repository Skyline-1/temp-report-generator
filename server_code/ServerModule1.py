import anvil.server
import pandas as pd
import os
import time
import statistics
#import win32com.client
import math
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import date
from docx.shared import Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import RGBColor
from anvil import *
import random
# This is a server module. It runs on the Anvil server,
# rather than in the user's browser.
#
# To allow anvil.server.call() to call functions here, we mark
# them with @anvil.server.callable.
# Here is an example - you can replace it with your own:
#
@anvil.server.callable
def save_user_choice(start_digit, author_name, start_date, start_time, end_date, end_time, temp_value, application_name, company_name):
    start_date_str = start_date.strftime('%Y-%m-%d')
    start_str = start_date_str + " " + start_time
    start_datetime = pd.to_datetime(start_str)
    end_date_str = end_date.strftime('%Y-%m-%d')
    end_str = end_date_str + " " + end_time
    end_datetime = pd.to_datetime(end_str)
    print("Hello")

