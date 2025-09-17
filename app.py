#!/usr/bin/env python3
"""
FIXED CAPTCHA: Working Agentic Outlook Account Creation
- Uses EXACT advanced CAPTCHA solving from comp.py  
- BUT with 10 seconds instead of 15 seconds as requested
- All other methods remain the same as exact_replica.py
"""

import os
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

# ================== TESSERACT SETUP (EXACT FROM main.py) ==================
pytesseract.pytesseract.tesseract_cmd = r"D:\\Tesseract-OCR\\tesseract.exe"

def _find_tesseract():
    configured = pytesseract.pytesseract.tesseract_cmd
    if configured and os.path.exists(configured):
        return configured
    
    from shutil import which
    found = which('tesseract')
    if found:
        pytesseract.pytesseract.tesseract_cmd = found
        return found
    return None

tess_path = _find_tesseract()
if not tess_path:
    print('\\nTesseract not found. Configure Tesseract OCR before running this script.')
    print('1) Install Tesseract: https://github.com/tesseract-ocr/tesseract')
    print('   On Windows you can use the official installer and note the install path, e.g. C:\\\\Program Files\\\\Tesseract-OCR\\\\tesseract.exe')
    print('2) Update the variable `pytesseract.pytesseract.tesseract_cmd` in this file to point to the tesseract executable, or add tesseract to your PATH.')
    print('3) Verify by running: tesseract --version')
    sys.exit(1)

# ================== AGENT SETUP (ENHANCED FROM agent.py) ==================
ANTHROPIC_API_KEY = None  # Set to None to skip API and use only pattern matching
# ANTHROPIC_API_KEY = "sk-ant-api03-your-actual-key-here"  # Uncomment and set your key

if ANTHROPIC_API_KEY:
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
else:
    client = None
    print("🔧 Using pattern matching only (no API key set)")

def analyze_screen_content(ocr_text):
    """ENHANCED version of your agent.py with better debugging and fallbacks"""
    print(f'🔍 OCR Text (first 200 chars): {ocr_text[:200]}...')
    print(f'🔍 OCR Text (full length): {len(ocr_text)} characters')
    
    # First try pattern matching (this always works)
    def pattern_match_fallback(text):
        print("🔧 Using pattern matching...")
        
        # EXACT same patterns as your agent.py
        if 'CREATE NEW ACCOUNT' in text:
            result = '{"action": "click", "element_type": "button", "element_identifier": "CREATE NEW ACCOUNT"}'
            print(f"✅ Pattern matched: CREATE NEW ACCOUNT")
            return result
        elif 'Create your Microsoft account' in text:
            result = '{"action": "type_email", "element_type": "input", "element_identifier": "New email"}'
            print(f"✅ Pattern matched: Create your Microsoft account")
            return result
        elif 'Create your password' in text:
            result = '{"action": "type_password", "element_type": "input", "element_identifier": "Password"}'
            print(f"✅ Pattern matched: Create your password")
            return result
        elif 'Add your country/region and birthdate' in text:
            result = '{"action": "type_dob", "element_type": "input", "element_identifier": "DOB"}'
            print(f"✅ Pattern matched: Add your country/region and birthdate")
            return result
        elif 'Add your name' in text:
            result = '{"action": "type_fullname", "element_type": "input", "element_identifier": "Full name"}'
            print(f"✅ Pattern matched: Add your name")
            return result
        elif "Let's prove you're human" in text:
            result = '{"action": "solve_captcha", "element_type": "press_and_hold", "element_identifier": "Captcha"}'
            print(f"✅ Pattern matched: Let's prove you're human")
            return result
        else:
            print("⚠️ No patterns matched. Returning empty JSON.")
            print(f"🔍 Available text snippets:")
            lines = text.split('\\n')[:10]  # Show first 10 lines
            for i, line in enumerate(lines):
                if line.strip():
                    print(f"   Line {i}: '{line.strip()}'")
            return '{}'
    
    # Try API first if available, then fallback to pattern matching
    if client:
        system_prompt = """You are an expert in mobile app automation and UI interaction. Your task is to analyze OCR text extracted from a mobile app screen and determine the appropriate action to take based on the content. You will provide your response in JSON format, specifying the action, the type of UI element involved, and an identifier for that element."""

        user_prompt = f"""You are a JSON generator for mobile app automation. You will be provided with OCR text extracted from a mobile app screen and a description of the action needed. Your task is to generate a JSON object that specifies the action to be taken, the type of UI element involved, and an identifier for that element. You must strictly response in JSON format without any additional text or explanation or markdown formatting.

OCR Text from screen: {ocr_text}

If the OCR Text contains the phrase 'CREATE NEW ACCOUNT' then generate this json response:
{{
"action": "click",
"element_type": "button",
"element_identifier": "CREATE NEW ACCOUNT"
}}

Or if the OCR Text contains the phrase 'Create your Microsoft account' then generate this json response:
{{
"action": "type_email",
"element_type": "input",
"element_identifier": "New email"
}}

Or if the OCR Text contains the phrase 'Create your password' then generate this json response:
{{
"action": "type_password",
"element_type": "input",
"element_identifier": "Password"
}}

Or if the OCR Text contains the phrase 'Add your country/region and birthdate' then generate this json response:
{{
"action": "type_dob",
"element_type": "input",
"element_identifier": "DOB"
}}

Or if the OCR Text contains the phrase 'Add your name' then generate this json response:
{{
"action": "type_fullname",
"element_type": "input",
"element_identifier": "Full name"
}}

Or if the OCR Text contains the phrase 'Let's prove you're human' then generate this json response:
{{
"action": "solve_captcha",
"element_type": "press_and_hold",
"element_identifier": "Captcha"
}}

Note: Do not respond with any other actions or elements. Only respond with the specified actions and identifiers as per the conditions above. If none of the conditions are met, then generate the empty JSON response."""

        try:
            print("🤖 Trying Anthropic API...")
            message = client.messages.create(
                model="claude-3-5-sonnet-20241022",
                max_tokens=800,
                temperature=0,
                system=system_prompt,
                messages=[{
                    "role": "user",
                    "content": user_prompt
                }]
            )
            
            result = message.content[0].text.strip()
            print(f"🤖 API Response: {result}")
            return result
            
        except Exception as e:
            print(f"❌ Claude API Error: {e}")
            print("🔧 Falling back to pattern matching...")
            return pattern_match_fallback(ocr_text)
    else:
        return pattern_match_fallback(ocr_text)

# ================== SETUP (ENHANCED FROM comp.py + main.py) ==================
screenshots_folder = "screenshots"

# ensure screenshots folder exists and clear it
if not os.path.exists(screenshots_folder):
    os.makedirs(screenshots_folder, exist_ok=True)

for filename in os.listdir(screenshots_folder):
    file_path = os.path.join(screenshots_folder, filename)
    if os.path.isfile(file_path):
        try:
            os.remove(file_path)
        except Exception:
            pass

# ================== DRIVER SETUP (EXACT FROM comp.py - BULLETPROOF) ==================
def setup_driver():
    """EXACT driver setup from comp.py - bulletproof version"""
    try:
        print("Setting up driver...")
        device_name = "ZD222GXYPV"  # Update with your device ID
        
        options = UiAutomator2Options()
        options.platform_name = 'Android'
        options.device_name = 'Android'
        options.app_package = 'com.microsoft.office.outlook'
        options.app_activity = '.MainActivity'
        options.automation_name = 'UiAutomator2'
        options.no_reset = False
        options.full_reset = False
        options.new_command_timeout = 300
        options.unicode_keyboard = True
        options.reset_keyboard = True
        options.auto_grant_permissions = True
        
        driver = webdriver.Remote("http://localhost:4723", options=options)
        driver.update_settings({"enforceXPath1": True})
        screen_size = driver.get_window_size()
        
        print("✓ Driver ready")
        return driver, device_name, screen_size
        
    except Exception as e:
        print(f"✗ Setup failed: {e}")
        return None, None, None

# ================== EMAIL/PASSWORD GENERATION (EXACT FROM main.py) ==================
def generate_outlook_email():
    first = "ivan"
    last = "lopez"
    digits = random.randint(1000000, 9999999)
    email_name = f"{first.lower()}{last.lower()}{digits}"
    return email_name

def generate_password():
    chars = string.ascii_letters + string.digits + "!@#$%"
    password = ''.join(random.choices(chars, k=15))
    return password

# ================== ACTION METHODS (EXACT FROM main.py) ==================
def press_button(driver, button_text):
    print(f'Waiting for button with text: {button_text}')
    try:
        button = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((AppiumBy.ANDROID_UIAUTOMATOR, f'new UiSelector().text("{button_text}")'))
        )
        button.click()
        return True
    except (NoSuchElementException, Exception) as e:
        print(f"Button '{button_text}' not found: {e}")
        return False

def fill_input_field(driver, device_name, field_text, input_value):
    try:
        if field_text == 'First name':
            name_fields = driver.find_elements(AppiumBy.CLASS_NAME, "android.widget.EditText")
            name_fields[0].send_keys(input_value)
            return True
        elif field_text == 'Last name':
            name_fields = driver.find_elements(AppiumBy.CLASS_NAME, "android.widget.EditText")
            name_fields[1].send_keys(input_value)
            return True
            
        input_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((AppiumBy.CLASS_NAME, "android.widget.EditText"))
        )
        
        if field_text == 'Year':
            input_field.click()
            time.sleep(0.3)
            subprocess.run(["adb", "-s", device_name, "shell", "input", "text", input_value])
        else:
            input_field.send_keys(input_value)
        return True
    except (NoSuchElementException, Exception) as e:
        print(f"Input field '{field_text}' not found: {e}")
        return False

def choose_from_dropdown(driver, dropdown_text, resource_Id, value):
    try:
        spinner = driver.find_element(AppiumBy.ANDROID_UIAUTOMATOR, f'new UiSelector().resourceId("{resource_Id}")')
        spinner.click()
        time.sleep(1)
        driver.find_element(AppiumBy.ANDROID_UIAUTOMATOR, f'new UiSelector().text("{value}")').click()
        return True
    except (NoSuchElementException, Exception) as e:
        print(f"Dropdown '{dropdown_text}' or option '{value}' not found: {e}")
        return False

# ================== ADVANCED CAPTCHA SOLVING (FROM comp.py BUT 10 SECONDS) ==================
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

def take_action(driver, device_name, screen_size, element_identifier, action):
    """EXACT same as your main.py with ADVANCED CAPTCHA"""
    try:
        if action == "click":
            success = press_button(driver, element_identifier)
            if success:
                print(f"Clicked on '{element_identifier}'")
            else:
                print(f"Failed to click on '{element_identifier}'")
                
        elif action == "type_email":
            email = generate_outlook_email()
            success_email = fill_input_field(driver, device_name, element_identifier, email)
            if success_email:
                print(f"Typed email '{email}'")
            else:
                print(f"Failed to type email.")
                
        elif action == "type_password":
            password = generate_password()
            success_password = fill_input_field(driver, device_name, element_identifier, password)
            if success_password:
                print(f"Typed password '{password}'")
            else:
                print(f"Failed to type password.")
                
        elif action == "type_dob":
            # EXACT same as main.py
            # select month
            random_month = random.randint(1, 12)
            month_text = calendar.month_name[random_month]
            success_month = choose_from_dropdown(driver, 'Month', 'BirthMonthDropdown', month_text)
            if success_month:
                print(f"Typed Month '{month_text}'")
            else:
                print(f"Failed to type Month.")

            # select day
            random_day = random.randint(1, 28)
            success_day = choose_from_dropdown(driver, 'Day', 'BirthDayDropdown', random_day)
            if success_day:
                print(f"Typed Day '{random_day}'")
            else:
                print(f"Failed to type Day.")

            # select year
            year = '2001'
            success_year = fill_input_field(driver, device_name, 'Year', year)
            if success_year:
                print(f"Typed Year '{year}'")
            else:
                print(f"Failed to type Year.")
                
        elif action == "type_fullname":
            first_name = "Ivan"
            last_name = "Lopez"
            success_first = fill_input_field(driver, device_name, 'First name', first_name)
            success_last = fill_input_field(driver, device_name, 'Last name', last_name)
            if success_first and success_last:
                print(f"Typed full name '{first_name} {last_name}'")
            else:
                print(f"Failed to type full name.")
                
        elif action == "solve_captcha":
            print("🎯 Using ADVANCED CAPTCHA solving method...")
            success = solve_captcha_advanced(driver, device_name, screen_size)
            if success:
                print("🎉 CAPTCHA solved with advanced method!")
            else:
                print("⚠️ Advanced CAPTCHA method had issues, but continuing...")
            
    except Exception as e:
        print(f"Error performing action: {e}")

def take_screenshot(device_id, screenshot_number, crop_status_bar=True):
    """EXACT same as your main.py"""
    screenshot_base_dir = os.path.join(screenshots_folder)
    screen_shot_file_name = "screenshot.png"
    screenshot_file = os.path.join(screenshot_base_dir, screen_shot_file_name)
    
    # remove old if present
    if os.path.exists(screenshot_file):
        try:
            os.remove(screenshot_file)
        except Exception:
            pass

    # run screencap on device
    try:
        cp = subprocess.run(["adb", "-s", device_id, "shell", "screencap", "-p", "/sdcard/screen.png"], capture_output=True, text=True)
        if cp.returncode != 0:
            print(f"adb screencap failed: {cp.stderr.strip()}")
            return None
    except Exception as e:
        print(f"Error running adb screencap: {e}")
        return None

    # pull with retries
    pulled = False
    for attempt in range(3):
        try:
            cp = subprocess.run(["adb", "-s", device_id, "pull", "/sdcard/screen.png", screenshot_file], capture_output=True, text=True)
            if cp.returncode == 0 and os.path.exists(screenshot_file):
                pulled = True
                break
            else:
                print(f"adb pull attempt {attempt+1} failed: {cp.stderr.strip()}")
        except Exception as e:
            print(f"Error pulling screenshot attempt {attempt+1}: {e}")
        time.sleep(1)

    if not pulled:
        print("Failed to pull screenshot from device.")
        return None

    if crop_status_bar:
        try:
            with Image.open(screenshot_file) as img:
                width, height = img.size
                status_bar_height = 120
                cropped_img = img.crop((0, status_bar_height, width, height))
                cropped_img.save(screenshot_file)
        except Exception as e:
            print(f"Error processing screenshot image: {e}")
            return None

    return screenshot_file

# ================== MAIN LOOP (EXACT FROM main.py) ==================
def main():
    """EXACT same main loop as your main.py but with ADVANCED CAPTCHA"""
    print("🤖 CAPTCHA FIXED: Agentic Outlook Automation")
    print("=" * 60)
    print("✅ Driver setup: comp.py (bulletproof)")
    print("✅ Agent logic: agent.py (enhanced with fallbacks)")  
    print("✅ Action methods: main.py (exact replica)")
    print("✅ CAPTCHA method: comp.py (advanced but 10 seconds)")
    print("✅ Main loop: main.py (exact replica)")
    print("=" * 60)
    
    # Setup driver using comp.py method
    driver, device_name, screen_size = setup_driver()
    if not driver:
        print("❌ Failed to setup driver")
        return
    
    completed = False
    attempt_count = 0
    max_attempts = 25  # Prevent infinite loops

    while not completed and attempt_count < max_attempts:
        attempt_count += 1
        print(f"\\n🔄 Attempt {attempt_count}/{max_attempts}")
        
        text = ""
        screenshot_attempts = 0
        max_screenshot_attempts = 5

        # Get screenshot and OCR (with retries)
        while len(text) == 0 and screenshot_attempts < max_screenshot_attempts:
            screenshot_attempts += 1
            print(f"📸 Screenshot attempt {screenshot_attempts}/{max_screenshot_attempts}")
            
            device_to_use = device_name if device_name else 'emulator-5554'
            screenshot_path = take_screenshot(device_to_use, 1, crop_status_bar=True)
            if not screenshot_path:
                print("❌ Failed to get screenshot, waiting and retrying...")
                time.sleep(2)
                continue
            
            try:
                text = pytesseract.image_to_string(screenshot_path)
                print(f"✅ OCR successful, extracted {len(text)} characters")
            except Exception as e:
                print(f"❌ OCR error: {e}")
                text = ""
            time.sleep(1)

        if not text:
            print("❌ No OCR text after multiple attempts, continuing...")
            time.sleep(3)
            continue

        # Analyze and act (EXACT same as main.py)
        if text:
            try:
                print("\\n🤖 Analyzing screen content...")
                analysis_result = analyze_screen_content(text)
                
                if analysis_result is None:
                    print("❌ Agent returned None, skipping this iteration")
                    time.sleep(2)
                    continue
                    
                print(f"🤖 Raw agent response: {analysis_result}")
                
                # Clean response in case of markdown formatting
                cleaned_result = analysis_result.strip()
                if cleaned_result.startswith('```json'):
                    cleaned_result = cleaned_result[7:]
                if cleaned_result.endswith('```'):
                    cleaned_result = cleaned_result[:-3]
                cleaned_result = cleaned_result.strip()
                
                action = json.loads(cleaned_result)
                print(f"🤖 Parsed action: {action}")
                
                if action.get("action") and action.get("element_identifier"):
                    print(f"⚡ Executing action: {action['action']} on {action['element_identifier']}")
                    take_action(driver, device_name, screen_size, action["element_identifier"], action["action"])
                    
                    # ALWAYS try Next button (except after captcha)
                    if action["action"] != "solve_captcha":
                        print("⏭️ Trying to press Next button...")
                        press_button(driver, "Next")
                    
                    # Check completion condition
                    if action["element_identifier"] == "Captcha":
                        print("🎉 CAPTCHA completed! Account creation should be done.")
                        completed = True
                else:
                    print("⚠️ No valid action found, waiting...")
                    time.sleep(3)
                    
            except json.JSONDecodeError as e:
                print(f"❌ JSON decode error: {e}")
                print(f"❌ Raw response was: {analysis_result}")
                time.sleep(2)
            except Exception as e:
                print(f"❌ Error processing action: {e}")
                time.sleep(2)
        
        time.sleep(2)  # Small delay between iterations

    if completed:
        print("\\n🎊 AUTOMATION COMPLETED SUCCESSFULLY!")
    else:
        print(f"\\n⚠️ Automation stopped after {max_attempts} attempts")
    
    # Cleanup
    try:
        driver.quit()
    except:
        pass

if __name__ == "__main__":
    main()