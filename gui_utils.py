# This file has utility functions for GUI automation using PyAutoGUI.

import pyautogui
import time
from pathlib import Path

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.4  # global pause between actions


def wait(seconds=1.0):
    time.sleep(seconds)


def click_image(image_name, confidence=0.9, timeout=10):
    """
    Clicks an image on screen using PyAutoGUI.
    Image must be inside images/ directory.
    """
    image_path = Path("images") / image_name
    start = time.time()

    while time.time() - start < timeout:
        location = pyautogui.locateCenterOnScreen(
            str(image_path), confidence=confidence
        )
        if location:
            pyautogui.click(location)
            return True
        wait(0.5)

    return False


def type_text(text):
    pyautogui.write(str(text), interval=0.05)


def press(*keys):
    pyautogui.hotkey(*keys)
