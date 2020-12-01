#!/usr/bin/env python3.7

from openpyxl import Workbook
import time

book = Workbook()
sheet = book.active

sheet['A1'] = 56
sheet['A2'] = 43

now = time.strftime("%x")
sheet['A3'] = now
sheet['A5'] = now

book.save("sample.xlsx")
