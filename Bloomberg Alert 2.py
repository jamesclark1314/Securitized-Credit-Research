# -*- coding: utf-8 -*-
"""
Created on Thu Dec 22 12:31:07 2022

@author: trpjs86
"""

import pandas as pd

repython = pd.read_excel('RePython.xlsx', header = None)
repython.columns = ['']

# Drop na
repython = repython.dropna()

# Write to excel
with pd.ExcelWriter('BBG Alert Template.xlsx', engine = 'openpyxl',
                    mode = 'a', if_sheet_exists = 'replace') as writer:
    repython.to_excel(writer, 'Alerts2', index = False)
