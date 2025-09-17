#!/usr/bin/env python3
"""
FINAL CORRECTED FLOW: Agentic Outlook Account Creation to Inbox
- Uses the EXACT flow sequence you specified
- Looks for "INBOX" instead of "Search" for success detection
- Authentication → Your Data, Your Way → Add another account → Getting Better Together → Powering Your Experiences → INBOX
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
    """CORRECTED agent with EXACT flow sequence"""
    print(f'🔍 OCR Text (first 200 chars): {ocr_text[:200]}...')
    print(f'🔍 OCR Text (full length): {len(ocr_text)} characters')
    
    def pattern_match_fallback(text):
        print("🔧 Using pattern matching...")
        
        # ORIGINAL PATTERNS (account creation)
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
            
        # CORRECTED POST-CAPTCHA FLOW (exact sequence you provided)
        elif 'Authentication' in text or 'authenticating' in text or 'Please wait' in text:
            result = '{"action": "wait_auth", "element_type": "loading", "element_identifier": "Authentication"}'
            print(f"✅ Pattern matched: Authentication/Please wait")
            return result
        elif 'Your Data, Your Way' in text or 'Your data, your way' in text:
            result = '{"action": "click_next", "element_type": "button", "element_identifier": "NEXT"}'
            print(f"✅ Pattern matched: Your Data, Your Way")
            return result
        elif 'Add another account' in text or 'Would you like to add another account' in text:
            result = '{"action": "click_maybe_later", "element_type": "button", "element_identifier": "MAYBE LATER"}'
            print(f"✅ Pattern matched: Add another account")
            return result
        elif 'Getting Better Together' in text or 'Getting better together' in text:
            result = '{"action": "click_accept", "element_type": "button", "element_identifier": "ACCEPT"}'
            print(f"✅ Pattern matched: Getting Better Together")
            return result
        elif 'Powering Your Experiences' in text or 'Powering your experiences' in text:
            result = '{"action": "click_continue", "element_type": "button", "element_identifier": "CONTINUE TO OUTLOOK"}'
            print(f"✅ Pattern matched: Powering Your Experiences")
            return result
            
        # CORRECTED SUCCESS DETECTION (look for INBOX instead of Search)
        elif 'INBOX' in text.upper() or 'Inbox' in text:
            result = '{"action": "inbox_reached", "element_type": "success", "element_identifier": "INBOX"}'
            print(f"🎉 Pattern matched: INBOX REACHED!")
            return result
            
        else:
            print("⚠️ No patterns matched. Checking for common UI elements...")
            
            # Look for common button texts
            lines = text.split('\\n')
            for line in lines:
                line_clean = line.strip().upper()
                if 'NEXT' in line_clean:
                    result = '{"action": "click_next", "element_type": "button", "element_identifier": "NEXT"}'
                    print(f"✅ Found NEXT button in line: {line.strip()}")
                    return result
                elif 'MAYBE LATER' in line_clean or 'NOT NOW' in line_clean:
                    result = '{"action": "click_maybe_later", "element_type": "button", "element_identifier": "MAYBE LATER"}'
                    print(f"✅ Found MAYBE LATER button in line: {line.strip()}")
                    return result
                elif 'CONTINUE' in line_clean and 'OUTLOOK' in line_clean:
                    result = '{"action": "click_continue", "element_type": "button", "element_identifier": "CONTINUE TO OUTLOOK"}'
                    print(f"✅ Found CONTINUE TO OUTLOOK in line: {line.strip()}")
                    return result
                elif 'ACCEPT' in line_clean:
                    result = '{"action": "click_accept", "element_type": "button", "element_identifier": "ACCEPT"}'
                    print(f"✅ Found ACCEPT button in line: {line.strip()}")
                    return result
                elif 'INBOX' in line_clean:
                    result = '{"action": "inbox_reached", "element_type": "success", "element_identifier": "INBOX"}'
                    print(f"🎉 Found INBOX - SUCCESS!")
                    return result
            
            print("⚠️ No specific patterns found. Will wait and retry...")
            print(f"🔍 Available text snippets:")
            for i, line in enumerate(lines[:15]):  # Show first 15 lines
                if line.strip():
                    print(f"   Line {i}: '{line.strip()}'")
            return '{"action": "wait", "element_type": "unknown", "element_identifier": "Unknown"}'
    
    # Try API first if available, then fallback to pattern matching
    if client:
        # Enhanced system prompt for corrected flow
        system_prompt = """You are an expert in mobile app automation and UI interaction for Microsoft Outlook account creation. Your task is to analyze OCR text from mobile screens and determine actions. You handle both account creation screens AND post-CAPTCHA setup screens that lead to the inbox."""

        user_prompt = f"""Analyze this OCR text and generate the appropriate JSON action:

OCR Text: {ocr_text}

ACCOUNT CREATION ACTIONS:
- 'CREATE NEW ACCOUNT' → {{"action": "click", "element_identifier": "CREATE NEW ACCOUNT"}}
- 'Create your Microsoft account' → {{"action": "type_email", "element_identifier": "New email"}}  
- 'Create your password' → {{"action": "type_password", "element_identifier": "Password"}}
- 'Add your country/region and birthdate' → {{"action": "type_dob", "element_identifier": "DOB"}}
- 'Add your name' → {{"action": "type_fullname", "element_identifier": "Full name"}}
- 'Let's prove you're human' → {{"action": "solve_captcha", "element_identifier": "Captcha"}}

POST-CAPTCHA SETUP ACTIONS (EXACT FLOW):
- 'Authentication'/'Please wait' → {{"action": "wait_auth", "element_identifier": "Authentication"}}
- 'Your Data, Your Way' → {{"action": "click_next", "element_identifier": "NEXT"}}
- 'Add another account' → {{"action": "click_maybe_later", "element_identifier": "MAYBE LATER"}}  
- 'Getting Better Together' → {{"action": "click_accept", "element_identifier": "ACCEPT"}}
- 'Powering Your Experiences' → {{"action": "click_continue", "element_identifier": "CONTINUE TO OUTLOOK"}}

COMPLETION:
- Text contains 'INBOX' or 'Inbox' → {{"action": "inbox_reached", "element_identifier": "INBOX"}}

Return ONLY valid JSON. If no patterns match, return {{"action": "wait", "element_identifier": "Unknown"}}"""

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

if not os.path.exists(screenshots_folder):
    os.makedirs(screenshots_folder, exist_ok=True)

for filename in os.listdir(screenshots_folder):
    file_path = os.path.join(screenshots_folder, filename)
    if os.path.isfile(file_path):
        try:
            os.remove(file_path)
        except Exception:
            pass

# ================== DRIVER SETUP (EXACT FROM comp.py) ==================
def setup_driver():
    try:
        print("Setting up driver...")
        device_name = "ZD222GXYPV"  
        
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

# ================== ACTION METHODS (EXACT FROM main.py + ENHANCED) ==================
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

# ================== ADVANCED CAPTCHA (FROM comp.py BUT 10 SECONDS) ==================
def solve_captcha_advanced(driver, device_name, screen_size):
    """EXACT advanced captcha method from comp.py but with 10 seconds"""
    print("🔧 Using advanced CAPTCHA solving from comp.py (10 seconds)")
    time.sleep(3)
    
    def find_element_bulletproof(by, value, timeout=10, retry_attempts=3):
        for attempt in range(retry_attempts):
            try:
                elements = WebDriverWait(driver, timeout).until(
                    lambda d: d.find_elements(by, value)
                )
                
                if elements:
                    visible_elements = []
                    for elem in elements:
                        try:
                            if elem.is_displayed():
                                visible_elements.append(elem)
                        except:
                            visible_elements.append(elem)
                    
                    if visible_elements:
                        return visible_elements[0]
                
                time.sleep(0.5)
            except:
                if attempt < retry_attempts - 1:
                    time.sleep(0.5)
        
        return None
    
    selectors = [
        (AppiumBy.ANDROID_UIAUTOMATOR, 'new UiSelector().className("android.widget.Button").textContains("Press").clickable(true).enabled(true)'),
        (AppiumBy.XPATH, "//android.widget.Button[contains(@text,'Press')]")
    ]
    
    button = None
    for by, selector in selectors:
        button = find_element_bulletproof(by, selector, timeout=8)
        if button:
            print(f"✅ Found CAPTCHA button")
            break
    
    if button:
        try:
            location = button.location
            size = button.size
            x = location['x'] + size['width'] // 2
            y = location['y'] + size['height'] // 2
            
            print(f"📍 Button location: x={x}, y={y}")
            
            try:
                driver.execute_script("mobile: longClickGesture", {
                    "elementId": button.id,
                    "duration": 10000
                })
                print("✅ Native long press (10s)")
                time.sleep(4)
                return True
            except Exception as e:
                print(f"⚠️ Native long press failed: {e}, trying ADB...")
            
            result = subprocess.run([
                "adb", "-s", device_name, "shell", "input", "touchscreen", "swipe",
                str(x), str(y), str(x), str(y), "10000"
            ], check=False, capture_output=True, text=True)
            
            if result.returncode == 0:
                print("✅ ADB long press with button coordinates (10s)")
                time.sleep(4)
                return True
            else:
                print(f"⚠️ ADB with coordinates failed: {result.stderr}")
                
        except Exception as e:
            print(f"❌ Button coordinate method failed: {e}")
    
    print("🔧 Using coordinate fallback...")
    x = screen_size['width'] // 2
    y = int(screen_size['height'] * 0.6)
    
    print(f"📍 Fallback coordinates: x={x}, y={y}")
    
    result = subprocess.run([
        "adb", "-s", device_name, "shell", "input", "touchscreen", "swipe",
        str(x), str(y), str(x), str(y), "10000"
    ], check=False, capture_output=True, text=True)
    
    if result.returncode == 0:
        print("✅ Coordinate fallback long press (10s)")
        time.sleep(4)
        return True
    else:
        print(f"❌ All CAPTCHA methods failed. ADB error: {result.stderr}")
        return False

# ================== ENHANCED ACTION HANDLER ==================
def take_action(driver, device_name, screen_size, element_identifier, action):
    """Enhanced action handler for ALL screen types"""
    try:
        if action == "click":
            success = press_button(driver, element_identifier)
            if success:
                print(f"✅ Clicked on '{element_identifier}'")
            else:
                print(f"❌ Failed to click on '{element_identifier}'")
                
        elif action == "type_email":
            email = generate_outlook_email()
            success_email = fill_input_field(driver, device_name, element_identifier, email)
            if success_email:
                print(f"✅ Typed email '{email}'")
            else:
                print(f"❌ Failed to type email.")
                
        elif action == "type_password":
            password = generate_password()
            success_password = fill_input_field(driver, device_name, element_identifier, password)
            if success_password:
                print(f"✅ Typed password '{password}'")
            else:
                print(f"❌ Failed to type password.")
                
        elif action == "type_dob":
            # DOB handling (exact same as main.py)
            random_month = random.randint(1, 12)
            month_text = calendar.month_name[random_month]
            success_month = choose_from_dropdown(driver, 'Month', 'BirthMonthDropdown', month_text)
            if success_month:
                print(f"✅ Selected Month '{month_text}'")
            
            random_day = random.randint(1, 28)
            success_day = choose_from_dropdown(driver, 'Day', 'BirthDayDropdown', random_day)
            if success_day:
                print(f"✅ Selected Day '{random_day}'")
            
            year = '2001'
            success_year = fill_input_field(driver, device_name, 'Year', year)
            if success_year:
                print(f"✅ Entered Year '{year}'")
                
        elif action == "type_fullname":
            first_name = "Ivan"
            last_name = "Lopez"
            success_first = fill_input_field(driver, device_name, 'First name', first_name)
            success_last = fill_input_field(driver, device_name, 'Last name', last_name)
            if success_first and success_last:
                print(f"✅ Typed full name '{first_name} {last_name}'")
                
        elif action == "solve_captcha":
            print("🎯 Using ADVANCED CAPTCHA solving...")
            success = solve_captcha_advanced(driver, device_name, screen_size)
            if success:
                print("🎉 CAPTCHA solved!")
            
        # CORRECTED POST-CAPTCHA ACTIONS
        elif action == "wait_auth":
            print("⏳ Waiting for authentication to complete...")
            # Use comp.py's wait_authentication logic
            for _ in range(45):  # 90 seconds max
                try:
                    progress_bars = driver.find_elements(AppiumBy.CLASS_NAME, "android.widget.ProgressBar")
                    visible = []
                    for bar in progress_bars:
                        try:
                            if bar.is_displayed():
                                visible.append(bar)
                        except:
                            continue
                    
                    if not visible:
                        print("✅ Authentication complete")
                        time.sleep(3)
                        break
                        
                except:
                    pass
                time.sleep(2)
                
        elif action == "click_next":
            print("⏭️ Clicking NEXT button...")
            success = press_button(driver, "NEXT") or press_button(driver, "Next")
            if success:
                print("✅ Clicked NEXT")
            else:
                print("⚠️ NEXT button not found, continuing...")
                
        elif action == "click_maybe_later":
            print("⏭️ Clicking MAYBE LATER button...")  
            success = press_button(driver, "MAYBE LATER") or press_button(driver, "Maybe later") or press_button(driver, "Not now")
            if success:
                print("✅ Clicked MAYBE LATER")
            else:
                print("⚠️ MAYBE LATER button not found, continuing...")
                
        elif action == "click_accept":
            print("⏭️ Clicking ACCEPT button...")
            success = press_button(driver, "ACCEPT") or press_button(driver, "Accept")
            if success:
                print("✅ Clicked ACCEPT")
            else:
                print("⚠️ ACCEPT button not found, continuing...")
                
        elif action == "click_continue":
            print("⏭️ Clicking CONTINUE TO OUTLOOK button...")
            success = press_button(driver, "CONTINUE TO OUTLOOK") or press_button(driver, "Continue to Outlook") 
            if success:
                print("✅ Clicked CONTINUE TO OUTLOOK")
            else:
                print("⚠️ CONTINUE TO OUTLOOK button not found, continuing...")
            
        elif action == "inbox_reached":
            print("🎊 INBOX REACHED! Outlook account creation completed successfully!")
            return "COMPLETED"
            
        elif action == "wait":
            print("⏳ Waiting...")
            time.sleep(3)
            
    except Exception as e:
        print(f"❌ Error performing action: {e}")
        
    return "CONTINUE"

# ================== SCREENSHOT (EXACT FROM main.py) ==================
def take_screenshot(device_id, crop_status_bar=True):
    screenshot_base_dir = os.path.join(screenshots_folder)
    screen_shot_file_name = "screenshot.png"
    screenshot_file = os.path.join(screenshot_base_dir, screen_shot_file_name)
    
    if os.path.exists(screenshot_file):
        try:
            os.remove(screenshot_file)
        except Exception:
            pass

    try:
        cp = subprocess.run(["adb", "-s", device_id, "shell", "screencap", "-p", "/sdcard/screen.png"], capture_output=True, text=True)
        if cp.returncode != 0:
            print(f"adb screencap failed: {cp.stderr.strip()}")
            return None
    except Exception as e:
        print(f"Error running adb screencap: {e}")
        return None

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

# ================== MAIN LOOP (CORRECTED FLOW) ==================
def main():
    """CORRECTED complete flow with exact sequence"""
    print("🤖 CORRECTED FLOW: Agentic Outlook Automation to Inbox")
    print("=" * 70)
    print("✅ EXACT FLOW: Authentication → Your Data, Your Way → Add another account → Getting Better Together → Powering Your Experiences → INBOX")
    print("✅ SUCCESS: Looks for 'INBOX' text (not 'Search')")
    print("✅ METHOD: Screenshot + Agent analysis throughout")
    print("=" * 70)
    
    # Setup driver using comp.py method
    driver, device_name, screen_size = setup_driver()
    if not driver:
        print("❌ Failed to setup driver")
        return
    
    completed = False
    attempt_count = 0
    max_attempts = 50  # More attempts for complete flow
    captcha_completed = False

    while not completed and attempt_count < max_attempts:
        attempt_count += 1
        print(f"\\n🔄 Attempt {attempt_count}/{max_attempts}")
        
        # Take screenshot and OCR
        text = ""
        screenshot_attempts = 0
        max_screenshot_attempts = 3

        while len(text) == 0 and screenshot_attempts < max_screenshot_attempts:
            screenshot_attempts += 1
            print(f"📸 Screenshot attempt {screenshot_attempts}/{max_screenshot_attempts}")
            
            screenshot_path = take_screenshot(device_name, crop_status_bar=True)
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

        # Analyze with agent and act
        try:
            print("\\n🤖 Analyzing screen content...")
            analysis_result = analyze_screen_content(text)
            
            if analysis_result is None:
                print("❌ Agent returned None, skipping this iteration")
                time.sleep(2)
                continue
                
            print(f"🤖 Raw agent response: {analysis_result}")
            
            # Clean response
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
                
                result = take_action(driver, device_name, screen_size, action["element_identifier"], action["action"])
                
                # Check completion
                if result == "COMPLETED":
                    print("🎊 COMPLETE SUCCESS! Outlook account created and INBOX opened!")
                    completed = True
                    break
                    
                # Track CAPTCHA completion
                if action["action"] == "solve_captcha":
                    captcha_completed = True
                    print("🎯 CAPTCHA phase completed, now handling post-CAPTCHA flow...")
                    
                # For account creation phase, always try Next (except after CAPTCHA)
                if not captcha_completed and action["action"] not in ["solve_captcha", "wait_auth", "inbox_reached"]:
                    print("⏭️ Trying to press Next button (account creation phase)...")
                    press_button(driver, "Next")
                    
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
        print("\\n🎊🎊🎊 COMPLETE SUCCESS! 🎊🎊🎊")
        print("✅ Outlook account created successfully")  
        print("✅ All post-CAPTCHA screens handled")
        print("✅ INBOX opened and ready to use")
        print("🎉 AUTOMATION COMPLETED!")
    else:
        print(f"\\n⚠️ Automation stopped after {max_attempts} attempts")
        if captcha_completed:
            print("✅ Account creation completed, but inbox not fully reached")
        else:
            print("⚠️ Did not complete account creation")
    
    # Cleanup
    try:
        driver.quit()
    except:
        pass

if __name__ == "__main__":
    main()