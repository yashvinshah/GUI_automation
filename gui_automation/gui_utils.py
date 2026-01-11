import pyautogui
import time

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.5  # small pause between actions to avoid timing issues

def wait(seconds=1.0):
    time.sleep(seconds)

def press(*keys):
    pyautogui.hotkey(*keys)

def type_text(text):
    pyautogui.write(str(text), interval=0.05)
