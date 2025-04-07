# -*- coding: utf-8 -*-
"""
Created on Wed Apr 10 12:24:25 2024

@author: Usama Ahsan
"""

import time
import pandas as pd
import pyautogui
import pyperclip
import urllib

df = pd.read_excel('contacts.xlsx')
msg = urllib.parse.quote_plus(
'''
MESSAGE IN MARKDOWN.
''')

time.sleep(5)
for i in df.iterrows():
    # NUM should be in the following format. +COUNTRYCODE|NUMBER. For example, +921234567890, +61123456789, etc.
    # Each number should be present in each consecutive rows, there should be no spaces in between. You can add the names as well just to keep the track.
    num = i[1][1]
    num = num.replace(' ','')
    num = num.replace('(','')
    num = num.replace(')','')
    num = num.replace('-','')
    pyperclip.copy(f'https://web.whatsapp.com/send?phone={num}&text={msg}')
    pyautogui.moveTo(300,77)
    time.sleep(0.1)
    pyautogui.click()
    time.sleep(0.2)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.2)
    pyautogui.press('enter')   
    time.sleep(14)
    pyautogui.press('enter')
    print(f'{i[0]}: {i[1]["Name"]}')
    time.sleep(3)
    
