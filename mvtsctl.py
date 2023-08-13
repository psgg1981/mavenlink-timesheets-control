import pdb
import csv
from datetime import datetime 
from dotenv import load_dotenv
import logging
import os
import pandas
import re
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support.ui import Select

#setup logger
log = logging.getLogger('logger')
log.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(message)s')
ch = logging.StreamHandler()
ch.setFormatter(formatter)
log.addHandler(ch)

# Load dotenv file ( .env file needs to be placed at the same level as the python script)
load_dotenv()
EMAIL = os.getenv('EMAIL', '')
PASS = os.getenv('PASS', '')
MAVENLINK_LOGIN_URL = os.getenv('MAVENLINK_LOGIN_URL', 'https://outsystems.mavenlink.com/login')

# pdb.set_trace()
projects_df = pandas.read_csv('projects.csv', header=None)
MAVENLINK_PROJECT_LIST = projects_df.iloc[:, 0].tolist()

SETTINGS = (EMAIL, PASS, MAVENLINK_LOGIN_URL)
PARAMS = ('EMAIL', 'PASS', 'MAVENLINK_LOGIN_URL')

try:
    pos = SETTINGS.index(None)
    print('Please make sure you have a .env file with respective value(s) for the property ' + PARAMS[pos] + ' on the same folder as the python script')
    input('Press return to exit.')
    exit()
except ValueError:
    print(datetime.now().strftime('%b %d, %Y'))
    # pdb.set_trace()
    print('Starting Mavenlink Timesheets Control script...')

# uncomment to check if variables were correctly loaded
# for prop in SETTINGS:
#     log.debug(prop)

if EMAIL=='' or PASS=='':
    print("Invalid username or  password.")
    input('Press return to exit.')
    exit()
elif len(MAVENLINK_PROJECT_LIST)==0:
    print("Empty project list.")
    input('Press return to exit.')
    exit()
else:
    log.info("Environment variables loaded successfully...")

# Login Selectors
login_sso_link = (By.CSS_SELECTOR, '.components_common__anchor-that-looks-like-a-button--A0Ug4')
login_Username_Inp = (By.CSS_SELECTOR, 'Input[name="loginfmt"]')
login_Btn = (By.CSS_SELECTOR, '.button_primary.button')
login_Pass_Inp = (By.NAME, "passwd")
login_BtnNext = (By.CSS_SELECTOR, 'Input[type="submit"]')
mavenlink_dashboard = (By.CSS_SELECTOR, '.app_bar_index__heading--emxeZ')

# Dashboard
search_container = (By.CSS_SELECTOR, ".button_button__subtle--h3aQl")
search_input = (By.CSS_SELECTOR, '.app_bar_Search__input--fz17o')
search_result_list = (By.CSS_SELECTOR, 'a.app_bar_Search__listItemAnchor--Yx3BG.no-custom :first-child')

# Timesheets tab Selector
timesheets_tab = (By.CSS_SELECTOR, "li[tab='time-tracking'")

# Timesheets Export Excel section
# timesheets_max_records_select = (By.CSS_SELECTOR, '.dataTables_length label select')
timesheets_max_records_select = (By.CSS_SELECTOR, '.dataTables_length > label > select')
# timesheets_export_expand_link = (By.CSS_SELECTOR, 'a.export-report-section-toggler')
# timesheets_export_start_date_input = (By.CSS_SELECTOR, 'input.start_date.hasDatepicker')
# timesheets_export_end_date_input = (By.CSS_SELECTOR, 'input.start_date.hasDatepicker')
# timesheets_export_button = (By.CSS_SELECTOR, 'button.export-report-xlsx')
timesheets_table = (By.CSS_SELECTOR, '.data-table')

# Supress webdriver experimental errors
options = webdriver.ChromeOptions()
#options.binary_location = get_chrome()
options.add_experimental_option('excludeSwitches', ['enable-logging'])
options.add_argument("--remote-debugging-port=9222") #fix for WSL Ubuntu issue WebDriverException: Message: unknown error: DevToolsActivePort file doesn't exist
options.add_argument('incognito')

# Load chromium webdriver. It will automatically be updated if necessary
chrome_service = Service()
log.info("Loading Chrome...")
driver = webdriver.Chrome(service=chrome_service, options=options)

log.info("Loading Mavenlink timeline URL...")
# Load the Mavenlink timeline URL
driver.get(MAVENLINK_LOGIN_URL)

log.info("Waiting for form and click for SSO...")
# Wait for form and click for SSO
wait = WebDriverWait(driver, 10, 1)
wait.until(EC.element_to_be_clickable(login_sso_link)).click()

wait = WebDriverWait(driver, 30, 1)
log.info("Waiting for form and entering username...")
# Wait for form and then enter the username
wait.until(EC.element_to_be_clickable(login_Username_Inp)).send_keys(EMAIL, Keys.RETURN)

log.info("Waiting for form and entering password...")
# Wait for form and then enter the password
if (len(PASS) > 0):
    wait.until(EC.element_to_be_clickable(login_Pass_Inp)).send_keys(PASS, Keys.RETURN)

print("Waiting for 2FA confirmation...")
longwait = WebDriverWait(driver, 120, 1)

# Wait for the dashboard page to load
print("Dashboard loading...")
longwait.until(EC.presence_of_element_located(mavenlink_dashboard))

# log.debug(MAVENLINK_PROJECT_LIST)

# Function returning a tuple (hours, minutes)
def calculate_h_m(data):

    total_hours = 0
    total_minutes = 0

    for entry in data:
        match = re.search(r'(\d+)h (\d+)m', entry)
        if match:
            hours = int(match.group(1))
            minutes = int(match.group(2))
            total_hours += hours
            total_minutes += minutes

    # Convert extra minutes to hours if needed
    total_hours += total_minutes // 60
    total_minutes %= 60

    return (total_hours, total_minutes)

wait = WebDriverWait(driver, 5, 1)
log.info('Searching projects...')
# click search bar
wait.until(EC.element_to_be_clickable(search_container)).click()
is_past_1st_element = False
for item in MAVENLINK_PROJECT_LIST:          
    try:
        # repeat dashboard loading once past 1st iteration
        if(is_past_1st_element):
            # HARD CODED!!
            driver.get('https://outsystems.mavenlink.com/users/16605785/dashboard')
            longwait.until(EC.presence_of_element_located(mavenlink_dashboard))
            wait.until(EC.element_to_be_clickable(search_container)).click()

        # clear text and type project        
        wait.until(EC.element_to_be_clickable(search_input)).clear()
        wait.until(EC.element_to_be_clickable(search_input)).send_keys(item, Keys.RETURN)
    except TimeoutException:
        print('Automation anomaly: skipping project ' + item + '...')
        is_past_1st_element = True
        continue

    try:
        wait.until(EC.element_to_be_clickable(search_result_list)).click()
    except TimeoutException:
        print('Project ' + item + ' not found. Skipping retrieval.')
        is_past_1st_element = True
        continue
    
    # Click timesheets tab
    wait.until(EC.element_to_be_clickable(timesheets_tab)).click()

    # Find select dropdown to list 100 records
    select = Select(wait.until(EC.presence_of_element_located(timesheets_max_records_select)))
    select.select_by_value('100')

    # Find timesheets HTML
    timesheets_html_table = wait.until(EC.presence_of_element_located(timesheets_table)).get_attribute('outerHTML')

    # Read timesheets HTML table
    df = pandas.read_html(timesheets_html_table)
    ts_tuple = tuple(df[0]["Time"])
    total_hm = calculate_h_m(ts_tuple)
    print(  str(total_hm[0]) + 'h ' + str(total_hm[1]) + 'm, ' + item)

    is_past_1st_element = True

print('Program finished.')