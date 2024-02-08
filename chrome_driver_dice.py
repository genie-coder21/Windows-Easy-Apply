from typing import List
import pyautogui, time, subprocess, autoit, pyperclip, os, configparser, pathlib
import regex as re
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException

# Read the settings values value from the settings file
config = configparser.ConfigParser()
config.read('C:\\Users\\Gee\\Downloads\\dice-easy-apply\\chrome_driver\\settings.ini')
directory = config.get('home', 'working_directory')
keyword_folder = config.get('files', 'keywords_directory')
documents = config.get('files', 'coverletter_location')
resume_folder = config.get('files', 'resume_location')
resume_name = config.get('files', 'resume_name')
coverletter_name = config.get('files', 'coverletter_name')
coverletter_folder = config.get('files', 'coverletter_location')


def save_ids(_ids):
    #with open(f"{directory}ids.txt", "w"), open("ids.log", "w"):
        #pass

    with open(file_path, 'r') as file:
        backup = file.read()
        #backup_ids = set(backup)
        #print("backup_ids:", backup_ids)
        print("_ids:", _ids)

    with open (f"{directory}ids.txt", "w") as file:
        #_id_list = [line for line in _ids if line not in backup]
        #for item in _id_list:
        #    print("item:", item)
        if _ids not in backup:    
            file.write(f"{_ids}\n")
    
    with open(f"{directory}ids.txt", 'r') as file:
        new_id = file.readline()


        return new_id

def remove_id(_id):
    print("_id:", _id)

    with open(f"{directory}ids.txt", "r") as file:
        lines = file.readlines()
    print("lines:", lines)

    # Find the lines that match the string
    lines_without_match = [line for line in lines if _id not in line]
    print("lines_without_match:", lines_without_match)

    with open(f"{directory}ids.txt", "w") as file:
        file.writelines(lines_without_match)

def keyword():
    with open(f"{keyword_folder}keywords.txt", "r") as keywords:
        keyword = keywords.readline()
        if  not keyword:
            print("No Keywords in List... Exiting")
            exit()

        lines = keywords.readlines()
        keywords.close()                    


    with open(f"{keyword_folder}keywords.txt", "w") as keywords:
        keywords.writelines(lines[0:])
        keywords.close()

    return keyword

def link_ids():

    time.sleep(5)
    # Execute JavaScript code
    wait = WebDriverWait(driver, 30)

    elements = wait.until(EC.presence_of_all_elements_located((By.XPATH, '//*[contains(@id, "-")]')))
    #print(elements)
    # Initialize an empty list to store the filtered IDs
    filtered_ids = []

    # Extract the 'id' and its contents
    for element in elements:
        element_id = element.get_attribute('id')
        
        # Remove "ID: " from the element ID
        element_id = element_id.replace("ID: ", "")
        
        # Count the number of hyphens in the element ID
        if element_id.count('-') == 4:
            filtered_ids.append(element_id)

    #print(filtered_ids)
    #print(element_id)

    return filtered_ids

def id_converter(element, link2):

    # Splitting the link into base URL and query parameters
    base_url, query_params = link2.split('?')
    params_dict = dict(param.split('=') for param in query_params.split('&'))
    base_url = "/".join(base_url.split('/')[:3])
    # Constructing the new path with the provided element
    new_path = f"/job-detail/{element}"

    # Updating the necessary query parameter
    params_dict['searchlink'] = 'search%2F%3Fq%3Dlinux%26countryCode%3DUS%26radius%3D30%26radiusUnit%3Dmi%26page%3D1%26pageSize%3D20%26filters.postedDate%3DONE%26filters.easyApply%3Dtrue%26filters.isRemote%3Dtrue%26language%3Den'
    #params_dict['searchId'] = element

    # Joining the base URL, new path and updated query parameters
    link1 = f"{base_url}{new_path}?{'&'.join([f'{key}={value}' for key, value in params_dict.items()])}"
    with open(f"{directory}url.txt", "w") as file:
        file.write(link1)
    with open(f"{directory}url.txt", "r") as file:
        link1 = file.read().replace("\n", "")

    with open(f"{directory}url.txt", "w") as file:
        file.write(link1)

    return link1

def open_new_tab(url):
    # Assuming you already have an existing ChromeDriver instance
    driver.switch_to.new_window('tab')

    # Switches to the newly opened tab
    driver.switch_to.window(driver.window_handles[-1])

    # Navigates to the given URL
    driver.get(url)

def close_current_tab(driver):
    driver.close()
    driver.switch_to.window(driver.window_handles[-1])  # Switch to the new active tab

def info():
    from cryptography.fernet import Fernet

    key = f"{directory}info\\encryption.key"  # Path to the encryption key file

    # Load the encryption key
    with open(key, 'rb') as key_file:
        key = key_file.read()

    # Load the encrypted password
    with open('C:\\Users\\Gee\\AppData\\Local\\Chromium\\Application\\encrypted_password.dat', 'rb') as password_file:
        encrypted_password = password_file.read()

    # Create a Fernet instance
    f = Fernet(key)

    # Decrypt the password
    decrypted_password = f.decrypt(encrypted_password).decode()

    # Use the decrypted password
    return decrypted_password

def apply():
    pyautogui.doubleClick(1786, 509)
    pyautogui.hotkey('ctrl', 'c')
    time.sleep(.1)

    double_check = pyperclip.paste()

    if "Application" in double_check:
        double_check = "yes"
    else:
        double_check = "no"

    pyautogui.click(1786, 509)
    time.sleep(.1)

    start_x, start_y = 1786, 509
    end_x, end_y = 1457, 456 

    pyautogui.moveTo(start_x, start_y)
    pyautogui.dragTo(end_x, end_y, duration=.3)
    pyautogui.hotkey('ctrl', 'c')
    time.sleep(.1)

    submitted = pyperclip.paste()

        
    if "Submitted" in submitted or "this Company" in submitted or "for this Company" in submitted or "View all" in submitted or "yes" in double_check: #or "ago" in content:
        print("Application Submitted Exit")
        applied = "yes"
    else:
        applied = "no"

    return applied
    # Wait for the page to load
#    wait = WebDriverWait(driver, 10)
#    wait.until(EC.presence_of_element_located((By.TAG_NAME, 'body')))
#
#    parent_shadow_root = driver.find_element(By.CSS_SELECTOR, "#applyButton > apply-button-wc")
    
    # Wait for the shadow root to be available
#    wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#applyButton > apply-button-wc")))
    
#    child_element = driver.execute_script("return arguments[0].shadowRoot.querySelector('apply-button > div')", parent_shadow_root)

#    try:
#        if child_element is not None:
#            apply_button = child_element.find_element(By.CSS_SELECTOR, ".btn.btn-primary")
            
#            if apply_button.is_displayed():
#                apply_button.click()
#                print("Button available")
#                return "yes"
#            else:
#                return "no"
#        else:
            # Handle when the child element is None
#            print("Child element not found")
#            return "no"
    
#    except NoSuchElementException:
        # Button not found, handle it as needed
#        print("Button not available")
#        return "no"

def apply_button():

    pyautogui.click(492, 355)
    time.sleep(.2)

    for _ in range(5):
        pyautogui.hotkey('tab')
        time.sleep(0.2)

    pyautogui.hotkey('enter')
    #time.sleep(7)
    
def click_replace_button(driver):

    try:
        replace_button = driver.find_element(By.CSS_SELECTOR, "button.file-remove")
        if replace_button.is_displayed():
            replace_button.click()
        else:
            print("Replace button not visible")
    except NoSuchElementException:
        print("Replace button not found")

def click_drop_box(upload, driver):
    try:
        drop_box = driver.find_element(By.CLASS_NAME, "fsp-drop-area")
        if drop_box.is_displayed():
            drop_box.click()
            autoit.win_wait_active("Open")
            time.sleep(4.5)
            #autoit.control_send("Open", "Edit1", "C:\\Users\\Gee\\Downloads\\"f"{upload}")
            pyautogui.typewrite("C:\\Users\\Gee\\Downloads\\"f"{upload}", interval=0.0375)
            time.sleep(.2)
            autoit.control_send("Open", "Edit1", "{ENTER}")
        else:
            print("Drop box area not visible")
    except NoSuchElementException:
        print("Drop box area not visible")

def click_upload_button(driver):
 
    try:
        span_element = driver.find_element(By.CSS_SELECTOR, '[data-e2e="upload"]')
        span_text = span_element.get_attribute("innerText").strip()
        if span_text == "Upload":
            span_element.click()
            return False
        else:
            print("Span element found, but text does not match 'Upload'")
            return True
    except NoSuchElementException:
        print("Span element not found, restarting")
        return True       

def check_label_presence(driver):
    try:
        progress_bar = driver.find_element(By.CSS_SELECTOR, 'progress-bar[label="Step 1 of 3"]')
        if progress_bar.is_displayed():
            print("Step 1 of 3 found")
            found = "yes"
        else:
            print("Step 1 of 3 not visible")
            found = "no"
    except NoSuchElementException:
        print("Step 1 of 3 not found")
        found = "no"
    return found

def next_button(text, driver):
    try:
        button_element = driver.find_element(By.XPATH, "//button[contains(.,'" + text + "')]")
        if button_element.is_displayed():
            button_element.click()
        else:
            print("Button element with text '{}' not visible".format(text))
    except NoSuchElementException:
        print("Button element with text '{}' not found".format(text))

def click_apply_button(driver):
    try:
        button_element = driver.find_element(By.CSS_SELECTOR, "button.btn-next.btn-split")
        if button_element.is_displayed():
            button_element.click()
        else:
            print("Apply button element not visible")
    except NoSuchElementException:
        print("Apply button element not found")

def compare_lists(list1_file, list2_file):
    # Read the contents of list1 file
    with open(list1_file, 'r') as file:
        list1 = file.read().split(', ')

    # Read the contents of list2 file
    with open(list2_file, 'r') as file:
        list2 = file.read().splitlines()

    # Check for the matching words
    matching_words = [word for word in list1 if word in list2]

    # Set the value of do_coverletter based on matching words
    if len(matching_words) >= 3:
        do_coverletter = "yes"
    else:
        do_coverletter = "no"

    print("numbe of matches:", len(matching_words))
    return do_coverletter
    
def click_upload_cover_letter_button(driver):

    try:
        upload_button = driver.find_element(By.CSS_SELECTOR, 'button[data-v-746be088=""]')
        if upload_button.is_displayed():
            upload_button.click()
        else:
            print("Upload button not visible")
    except NoSuchElementException:
        print("Upload button not found")

# Provide the paths of the list files
list1_file = f"{directory}myskills.txt"
list2_file = f"{directory}skillslist.txt"

with open(f"{keyword_folder}keywords.txt", "r") as file:
    keywords = file.readline()
    if  not keyword:
        print("No Keywords in List... Exiting")
        exit()


if os.path.exists(f"{resume_folder}{resume_name}.docx"):
    del_path = pathlib.Path(f'{resume_folder}{resume_name}.docx')
    del_path.unlink()
else:
    print("")

if os.path.exists(f"{coverletter_folder}{coverletter_name}.docx"):
    del_path = pathlib.Path(f'{coverletter_folder}{coverletter_name}.docx')
    del_path.unlink()
else:
    print("")

network_path = r"\\vmware-host\Shared Folders\Downloads\dice-easy-apply\chrome_driver"
file_path = os.path.join(network_path, "ids.log")

url1 = "https://www.dice.com/jobs?q="
url2 = "&countryCode=US&radius=30&radiusUnit=mi&page="
url3 = "&pageSize=100&filters.postedDate=ONE&filters.easyApply=true&filters.isRemote=true&language=en"
base_url = (f"{url1}{keywords}{url2}1{url3}")
start_url = "https://www.dice.com/dashboard/login"
#print("keyword:", keywords)
chromium_path = 'C:\\Users\\Gee\\AppData\\Local\\Chromium\\Application\\chromium.exe'

# Create ChromeOptions and set the remote debugging endpoint and binary location
options = webdriver.ChromeOptions()
options.add_argument('--remote-debugging-address=127.0.0.1')
options.add_argument('--remote-debugging-port=9222')
options.add_argument('--start-maximized')
#options.add_extension('/path/to/extension.crx')
options.binary_location = chromium_path

# Create the Chrome driver instance
driver = webdriver.Chrome(options=options) # You need to have the Chrome driver installed and accessible in your PATH
driver.get(start_url)

username_field = driver.find_element(By.CSS_SELECTOR, "#email")
password_field = driver.find_element(By.CSS_SELECTOR, "#password")

information = info()
username_field.send_keys("Marylandky21@gmail.com")
password_field.send_keys(f"{information}")
password_field.send_keys(Keys.RETURN)

time.sleep(2.5)
driver.get(base_url)

with open(f"{directory}ids.txt", 'r') as file:
    List = file.readlines()


#@time.sleep(5)
#keywords = keyword()
#print("keyword:", keywords)
counter = 0
for i in range(1, 500, 1):
    counter += 1
    _ids = link_ids()
    print("_ids:", _ids)
    num_ids = len(_ids)
    #print(num_ids)
    if not keywords and not List:
        print("No keywords or ID's Available Exiting")
        exit()
    if num_ids < 100:
        i = 1
    url = f"{url1}{keywords}{url2}{i}{url3}"
    if counter == 1 and num_ids == 100:
        i = 2
        url = f"{url1}{keywords}{url2}{i}{url3}"
    if num_ids < 100:
        keywords = keyword()
    #print(url)
    #print("keyword:", keywords)
    list_ids = [f"{_id}" for _id in _ids]
    #print("list_ids:", list_ids)
    #id_logs = save_ids(_ids)
    #print("_id_logs:", id_logs)
    for identifiers in list_ids:
        print("_id:", identifiers)
        print()
        code = identifiers
        print("code:", code)
        use_id = save_ids(code)
        print("use_id:", use_id)
        if len(use_id) > 0:
            start_time = time.time()
            #print("use_id:", use_id)
            app_url = id_converter(use_id, url)
            print(app_url)
            open_new_tab(app_url)
            time.sleep(18)
            applied = apply()
            if "no" in applied:
                apply_button()
                resume = f"{directory}search and replace resume.py"
                resume_file = "Resume.pdf"
                # Call the script runs in backgroud
                mod_resume = subprocess.Popen(["python", resume])
                #print("applying")
                time.sleep(5)
                result = check_label_presence(driver)
                #print(result)
                if "yes" in result:
                    mod_resume.wait()
                    close_current_tab(driver)
                    print("1 of 3 found")
                    current_time = datetime.now().strftime("%m/%d %I:%M:%S.%f %p %Z") # [:-3] to remove the last 3 digits for milliseconds
                    with open(file_path, 'a') as file:
                        file.write(f"{code}, {current_time} VM\n")
                    remove_id(code)
                    break
                click_replace_button(driver)
                time.sleep(2)
                mod_resume.wait()
                time.sleep(2)
                click_drop_box(resume_file, driver)
                time.sleep(2)
                upload_button = click_upload_button(driver)
                if upload_button:
                    print("retry 1")
                    time.sleep(1)
                    upload_button = click_upload_button(driver)
                    if upload_button:
                        print("retry 2")
                        time.sleep(1)
                        upload_button = click_upload_button(driver)
                        if upload_button:
                            print("restarting cant find element")
                            close_current_tab(driver)
                            continue
                time.sleep(3)
                # Call the function and store the result in a variable
                do_coverletter_result = compare_lists(list1_file, list2_file)
                if "yes" in do_coverletter_result:
                    # Path to the script you want to call
                    search_coverletter = f"{directory}search and replace coverletter.py"
                    coverletter_file = "CoverLetter.pdf"
                    # Call the script and wait for it to finish
                    mod_coverletter = subprocess.Popen(["python", search_coverletter])
                    click_upload_cover_letter_button(driver)
                    time.sleep(2)
                    mod_coverletter.wait()
                    click_drop_box(coverletter_file, driver)
                    time.sleep(2)
                    upload_button = click_upload_button(driver)
                    if upload_button:
                        print("retry 1")
                        time.sleep(1)
                        continue
                    time.sleep(4)
                next_button("Next", driver)
                time.sleep(3)
                click_apply_button(driver)
                time.sleep(3)
                current_time = datetime.now().strftime("%m/%d %I:%M:%S.%f %p %Z") # [:-3] to remove the last 3 digits for milliseconds
                with open(file_path, 'a') as file:
                    file.write(f"{code}, {current_time} VM\n")
                remove_id(code)
                #id_logs = save_ids(list_ids)
                close_current_tab(driver)

                elapsed_time = time.time() - start_time
                with open(f"{directory}EasyApplyPythonElapsed_time.log", "a") as file:
                    file.write(f"Elapsed Time: {elapsed_time} seconds, Coverletter: {do_coverletter_result}\n")
            if "yes" in applied:
                print("Already Applied")
                current_time = datetime.now().strftime("%m/%d %I:%M:%S.%f %p %Z") # [:-3] to remove the last 3 digits for milliseconds
                with open(file_path, 'a') as file:
                    file.write(f"{code}, {current_time} VM\n")
                remove_id(code)
                #id_logs = save_ids(list_ids)
                close_current_tab(driver)
    else:
        print("no ids")
        driver.get(url)



