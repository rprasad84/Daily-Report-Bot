import pyautogui
import pyperclip
import time
import re
import os
from datetime import datetime

pyautogui.FAILSAFE = True   # move mouse to a screen corner to stop the bot
pyautogui.PAUSE = 0.5

# Date and time, generated when the bot runs
now = datetime.now()
date_time = now.strftime("%Y-%m-%d %H:%M:%S")
date_only = now.strftime("%Y-%m-%d")

# File names with today's date
os.makedirs("output", exist_ok=True)
excel_file = os.path.abspath("output/daily_report_" + date_only + ".xlsx")
screenshot_file = os.path.abspath("output/daily_report_" + date_only + ".png")

print("Starting in 5 seconds. Do not touch the mouse or keyboard.")
time.sleep(5)

# 1. Open Edge and go to the website
pyautogui.hotkey("win", "r")
time.sleep(1)
pyautogui.write("msedge https://www.moneycontrol.com/markets/indian-indices/")
pyautogui.press("enter")
time.sleep(10)

# 2. Copy the page and find the Nifty 50 value
pyperclip.copy("")
pyautogui.hotkey("ctrl", "a")
time.sleep(1)
pyautogui.hotkey("ctrl", "c")
time.sleep(1)
page_text = pyperclip.paste()

stock_value = "N/A"
position = page_text.lower().find("nifty 50")
if position != -1:
    part = page_text[position: position + 300]
    found = re.search(r"\d{1,3}(?:,\d{3})+(?:\.\d+)?", part)
    if found:
        stock_value = found.group(0)

print("Nifty 50 value:", stock_value)
comment = "Market data captured, review the trend"

# 3. Open Excel with a new blank workbook
pyautogui.hotkey("win", "r")
time.sleep(1)
pyautogui.write("excel")
pyautogui.press("enter")
#time.sleep(8)
#pyautogui.hotkey("ctrl", "n")
time.sleep(3)

# 4. Type the header row and the data row
pyautogui.write("Date & Time")
pyautogui.press("tab")
pyautogui.write("Nifty 50 Value")
pyautogui.press("tab")
pyautogui.write("Comment")
pyautogui.press("enter")

pyautogui.write(date_time)
pyautogui.press("tab")
pyautogui.write(stock_value)
pyautogui.press("tab")
pyautogui.write(comment)
pyautogui.press("enter")

# Make columns wide enough to read
pyautogui.hotkey("ctrl", "a")
pyautogui.press("alt")
pyautogui.press("h")
pyautogui.press("o")
pyautogui.press("i")
pyautogui.hotkey("ctrl", "home")

# 5. Save the Excel file
if os.path.exists(excel_file):
    os.remove(excel_file)
pyautogui.press("f12")
time.sleep(3)
pyperclip.copy(excel_file)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(3)

# 6. Take a screenshot of the final sheet
pyautogui.screenshot(screenshot_file)

print("Done!")
print("Excel file:", excel_file)
print("Screenshot:", screenshot_file)
print("Task completed.")
