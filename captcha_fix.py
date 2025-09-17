import time
import json
import random
import subprocess
import calendar
import string
import sys

from appium import webdriver
from appium.options.android import UiAutomator2Options
from appium.webdriver.common.appiumby import AppiumBy
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
import pytesseract
from PIL import Image
import anthropic

def solve_captcha_advanced(driver, device_name, screen_size):
    """EXACT advanced captcha method from comp.py but with 10 seconds as requested"""
    print("🔧 Using advanced CAPTCHA solving from comp.py (10 seconds)")
    time.sleep(3)
    
    # Find CAPTCHA button using comp.py's bulletproof method
    def find_element_bulletproof(by, value, timeout=10, retry_attempts=3):
        """Bulletproof element finding from comp.py"""
        for attempt in range(retry_attempts):
            try:
                elements = WebDriverWait(driver, timeout).until(
                    lambda d: d.find_elements(by, value)
                )
                
                if elements:
                    # Filter for displayed elements
                    visible_elements = []
                    for elem in elements:
                        try:
                            if elem.is_displayed():
                                visible_elements.append(elem)
                        except:
                            visible_elements.append(elem)
                    
                    if visible_elements:
                        return visible_elements[0]  # Return first visible
                
                time.sleep(0.5)
            except:
                if attempt < retry_attempts - 1:
                    time.sleep(0.5)
        
        return None
    
    # EXACT selectors from comp.py
    selectors = [
        (AppiumBy.ANDROID_UIAUTOMATOR, 'new UiSelector().className("android.widget.Button").textContains("Press").clickable(true).enabled(true)'),
        (AppiumBy.XPATH, "//android.widget.Button[contains(@text,'Press')]")
    ]
    
    button = None
    for by, selector in selectors:
        button = find_element_bulletproof(by, selector, timeout=8)
        if button:
            print(f"✅ Found CAPTCHA button with selector: {by}")
            break
    
    if button:
        try:
            # Method 1: Get button location and use ADB with exact coordinates
            location = button.location
            size = button.size
            x = location['x'] + size['width'] // 2
            y = location['y'] + size['height'] // 2
            
            print(f"📍 Button location: x={x}, y={y}")
            
            # Method 1A: Try native long press first (but with 10 seconds)
            try:
                driver.execute_script("mobile: longClickGesture", {
                    "elementId": button.id,
                    "duration": 10000  # 10 seconds as requested
                })
                print("✅ Native long press (10s)")
                time.sleep(4)
                return True
            except Exception as e:
                print(f"⚠️ Native long press failed: {e}, trying ADB...")
            
            # Method 1B: ADB with exact button coordinates (10 seconds)
            result = subprocess.run([
                "adb", "-s", device_name, "shell", "input", "touchscreen", "swipe",
                str(x), str(y), str(x), str(y), "10000"  # 10 seconds as requested
            ], check=False, capture_output=True, text=True)
            
            if result.returncode == 0:
                print("✅ ADB long press with button coordinates (10s)")
                time.sleep(4)
                return True
            else:
                print(f"⚠️ ADB with coordinates failed: {result.stderr}")
                
        except Exception as e:
            print(f"❌ Button coordinate method failed: {e}")
    
    # Method 2: Coordinate fallback using screen center (EXACT from comp.py)
    print("🔧 Using coordinate fallback...")
    x = screen_size['width'] // 2
    y = int(screen_size['height'] * 0.6)
    
    print(f"📍 Fallback coordinates: x={x}, y={y}")
    
    result = subprocess.run([
        "adb", "-s", device_name, "shell", "input", "touchscreen", "swipe",
        str(x), str(y), str(x), str(y), "10000"  # 10 seconds as requested
    ], check=False, capture_output=True, text=True)
    
    if result.returncode == 0:
        print("✅ Coordinate fallback long press (10s)")
        time.sleep(4)
        return True
    else:
        print(f"❌ All CAPTCHA methods failed. ADB error: {result.stderr}")
        return False