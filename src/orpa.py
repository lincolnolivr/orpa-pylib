# COMMAND ---------- [markdown]
# # orpa (RPA Functions)

# COMMAND ----------
import os
import sys
import time
import signal
import atexit
import random
import zipfile
import requests
import win32api
import win32con
import pyautogui
import pyperclip
import http.client
import pandas as pd
import win32com.client
from io import BytesIO
from PIL import ImageGrab
from threading import local
import win32clipboard as clip
from datetime import datetime
from selenium import webdriver
from mouseinfo import mouseInfo
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
from win32gui import GetWindowText, GetForegroundWindow

import mylibs.osystempy as osystempy

def open_mouse_info():
    mouseInfo()

def sleep(seconds):
    time.sleep(seconds)

def wait(mintime=3,maxtime=5):
    mintime = mintime * 60
    maxtime = maxtime * 60
    x = random.uniform(mintime,maxtime)
    print('Iniciando próxima ação em ' + str(datetime.timedelta(seconds=x)) + ' min \n')
    time.sleep(x)

def get_screen_size():
    image = ImageGrab.grab()
    return image.height, image.width

def get_screen_height():
    image = ImageGrab.grab()
    return image.height

def get_screen_width():
    image = ImageGrab.grab()
    return image.width

def screenshot_to_clipboard(x_top,y_top,x_bottom,y_bottom):
    image = ImageGrab.grab(bbox=(x_top,y_top,x_bottom,y_bottom))
    output = BytesIO()
    image.convert('RGB').save(output, 'BMP')
    data = output.getvalue()[14:]
    output.close()
    clip.OpenClipboard()
    clip.EmptyClipboard()
    clip.SetClipboardData(win32con.CF_DIB, data)
    clip.CloseClipboard()

def get_downloads_folder():
    downloads_folder = os.path.join(os.path.expanduser('~'), 'Downloads')
    return downloads_folder

def set_logs_folder(folder_path=get_downloads_folder()):
    global logs_folder

    default_logs_folder = os.path.join(folder_path, 'Logs')
    if folder_path == get_downloads_folder():
        folder_path = default_logs_folder

    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

    logs_folder = folder_path
    
    return logs_folder

def start_logging(logs_name=os.path.basename(sys.argv[0]),file_path=None,include_messages=False):
    global logs_folder

    if file_path == None and logs_folder == None:
        logs_folder = set_logs_folder()
        file_path = logs_folder
    elif file_path == None and logs_folder != None:
        file_path = logs_folder

    logs_filename = os.path.join(file_path, logs_name + '.txt')

    with open(logs_filename, 'a') as f:
        f.write('\n')
        f.write('Iniciando o log em ' + format(datetime.now())[:19] + '\n')
        f.write('\n')

    # Redireciona a saída padrão para o arquivo de log
    if include_messages:
        sys.stdout = open(logs_filename, 'a')
    sys.stderr = open(logs_filename, 'a')

def start_saving_logs():
    global logs_df
    logs_df = pd.DataFrame()

# create a function to append the dataframe
def job(job_name='Job'):
    global logs_df
    df = pd.DataFrame({'Date': format(datetime.now())[:19], 'Job': job_name}, index=[0])
    try:
        logs_df = pd.concat([logs_df, df], ignore_index=True)
    except Exception as e:
        start_saving_logs()
        logs_df = pd.concat([logs_df, df], ignore_index=True)

def save_logs(file_path=os.path.join(os.path.expanduser('~'), 'Downloads'), file_prefix='', mode=None):
    global logs_df
    file_prefix = ' ' + file_prefix if file_prefix != '' else file_prefix
    logs_file_name = os.path.join(file_path, format(datetime.now())[:10].replace(' ', '-').replace(':', '-') + file_prefix + ' Logs' + '.txt')
    logs_df.to_csv(logs_file_name, mode='a', index=False,header=not(os.path.isfile(logs_file_name)))
    if mode == 'skipline':
        with open(logs_file_name, 'a') as f:
            f.write('\n')
    logs_df = pd.DataFrame()

def get_active_window():
    return print(GetWindowText(GetForegroundWindow()))

def power_automates_notify(notify_url,notify_path):
    conn = http.client.HTTPSConnection(notify_url)
    conn.request('POST', notify_path)
    response = conn.getresponse()
    if response.status == 202:
        print('Requisição bem sucedida!')
    else:
        print(f'Erro na requisição. Código de status: {response.status}')

    conn.close()

def setup_outlook(active_mapi=False):
    global outlook
    global outlook_fail
    global mapi

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        if active_mapi:
            mapi = outlook.GetNamespace("MAPI")
        outlook_fail = False

    except Exception as e:
        outlook_fail = True

def check_outlook_status():
    global outlook_fail

    if not ('outlook' in locals() or 'outlook' in globals()):
        setup_outlook()
    if outlook_fail:
        return False
    else:
        return True

def get_main_account():
    global outlook
    global outlook_fail
    global main_account

    if not check_outlook_status():
        main_account = None
        print('Fail to connect with Outlook')
    else:
        main_account = outlook.Session.Accounts[0]
    return main_account

def send_email_notification(subject, body,to=None, importance='normal'):
    global outlook
    global outlook_fail
    global main_account

    if to == None:
        main_account = get_main_account()
        to = main_account

    if not check_outlook_status():
        print('Fail to connect with Outlook')
        return False
        
    else:
        mail = outlook.CreateItem(0)
        mail.To = to
        mail.Subject = subject
        if importance == 'high': mail.Importance = 2
        elif importance == 'normal': mail.Importance = 1
        elif importance == 'low': mail.Importance = 0
        mail.Body = body
        mail.Send()

# COMMAND ---------- [markdown]
# Pyautogui

# COMMAND ----------
def press(key):
    pyautogui.press(key)

def hotkey(key1,key2):
    pyautogui.hotkey(key1,key2)

def write(text):
    pyautogui.write(text)

def typewrite(text):
    pyautogui.typewrite(text)

def press_tab(presses=1, interval=0.0, mode='normal'):
    for i in range(presses):
        if mode == 'normal':
            press('tab')
        elif mode == 'shift':
            hotkey('shift', 'tab')
        sleep(interval)

def copy_clipboard():
    pyperclip.copy("") # <- This prevents last copy replacing current copy of null.
    pyautogui.hotkey('ctrl', 'c')
    pyautogui.sleep(.01)  # ctrl-c is usually very fast but your program may execute faster
    
    return pyperclip.paste()

def pyautogui_open_microsoft_edge(link=''):
    os.startfile('C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe')
    pyautogui.sleep(2)
    pyautogui.getActiveWindow().maximize()

    if link !='':
        pyautogui.typewrite(link)
        pyautogui.press('enter')

def pyautogui_wait_until_download_edge(download_button):
    r = None
    while r == None:
        r = pyautogui.locateOnScreen(download_button, grayscale=False)

def found_all_buttons_and_click(buttons):
    for button in buttons.split(' ', 0) if type(buttons)==str else buttons:
        r = None
        while r == None:
            r = pyautogui.locateOnScreen(button)
        pyautogui.click(r)

def press_shift_pgdn(interval=2):
    pyautogui.keyDown('shiftleft')
    pyautogui.keyDown('shiftright')
    pyautogui.hotkey('pgdn')
    pyautogui.keyUp('shiftleft')
    pyautogui.keyUp('shiftright')
    pyautogui.sleep(interval)

def press_shift_pgup(interval=2):
    pyautogui.keyDown('shiftleft')
    pyautogui.keyDown('shiftright')
    pyautogui.hotkey('pgup')
    pyautogui.keyUp('shiftleft')
    pyautogui.keyUp('shiftright')
    pyautogui.sleep(interval)

def select_right_text(interval=2):
    pyautogui.keyDown('shiftleft')
    pyautogui.keyDown('shiftright')
    pyautogui.keyDown('ctrl')
    pyautogui.press('end')
    pyautogui.keyUp('shiftleft')
    pyautogui.keyUp('shiftright')
    pyautogui.keyUp('ctrl')
    pyautogui.sleep(interval)

def alternate_tabs(presses=1):
    pyautogui.keyDown('alt')
    pyautogui.sleep(.2)

    for i in range(presses):
        pyautogui.press('tab')
        pyautogui.sleep(.2)

    pyautogui.keyUp('alt')

# COMMAND ---------- [markdown]
# Selenium

# COMMAND ----------
def get_python_folder():
    return os.path.dirname(sys.executable)

def get_edge_webdriver_folder():
    python_folder = get_python_folder()
    return os.path.join(python_folder,'msedgedriver.exe')

def download_edge_webdriver(version: str,python_folder: str,download_folder: str=None):
    print('Downloading new Edge Webdriver')
    if download_folder is None: download_folder = os.path.join(os.path.expanduser('~'),'Downloads')

    try:
        url = 'https://msedgedriver.azureedge.net/' + version + '/edgedriver_win64.zip'
        response = requests.get(url, allow_redirects=True)
        download_file = os.path.join(download_folder,f'edgedriver_win64.zip')

        # Check if the download was successful
        if response.status_code == 200:
            # Save the zip file
            with open(download_file, 'wb') as f:
                f.write(response.content)

            with zipfile.ZipFile(download_file, 'r') as zip_ref:
                zip_ref.extractall(python_folder)
            os.remove(download_file)
            print(f'WebDriver do Microsoft Edge (versão {version}) baixado com sucesso!')

        else:
            print('Falha ao baixar o arquivo ZIP do WebDriver.')
    except Exception as e:
        print(f"Erro ao baixar o WebDriver do Microsoft Edge: {e}")

def is_edge_webdriver_update():
    current_edge_version = osystempy.get_edge_version()
    current_webdriver_version = osystempy.get_installed_edge_webdriver_version(get_edge_webdriver_folder())

    if current_edge_version != current_webdriver_version:
        print('Current Edge Version:', current_edge_version)
        print('Current Webdriver Version', current_webdriver_version)
        return False
    else:
        return True

def update_edge_webdriver():
    print('Checking Webdriver Version...')
    if is_edge_webdriver_update():
        print('Edge Webdriver is up to date!')
    else:
        download_edge_webdriver(osystempy.get_edge_version(),get_python_folder())

def selenium_hide_edge():
    global edge_options
    # Open Microsoft Edge Browser
    edge_options = webdriver.EdgeOptions()
    edge_options.add_experimental_option("prefs", {"download.prompt_for_download": False, 'profile.default_content_settings.popups': False})     
    edge_options.add_argument("--window-size=1920,1080")
    edge_options.add_argument("--start-maximized")
    edge_options.add_argument("--headless=new")

def selenium_open_edge(link='https://www.google.com/',hidden=False):
    global driver
    global edge_options

    if hidden:
        selenium_hide_edge()
        driver = webdriver.Edge(options=edge_options)
    else:
        driver = webdriver.Edge()
        driver.maximize_window()
    driver.get(link)

def selenium_wait_page_load(current_driver=None):
    global driver

    if current_driver == None:
        current_driver = driver

    while True:
        page_state = current_driver.execute_script('return document.readyState;')
        if page_state == 'complete':
            break
        else:
            sleep(1)

def selenium_action(xpath, action='click', keys=None,tries=10, wait=1, current_driver=None):
    global driver

    if isinstance(xpath,webdriver.Edge):
        print('Remova o orpa.driver da função ou atualize para current_driver=orpa.driver')
        return False

    if current_driver == None:
        current_driver = driver

    tries_count = 0
    while tries_count <= tries:
        try:
            if action == 'click':
                current_driver.find_element(By.XPATH, xpath).click()
            elif action == 'send_keys':
                current_driver.find_element(By.XPATH, xpath).send_keys(keys)
            elif action == 'get':
                current_driver.get(keys)
            tries_count = tries + 1
        except NoSuchElementException:
            tries_count += 1
            sleep(wait)
            if tries_count >= tries:
                print('Element not found')
                return False
            continue

def selenium_click(xpath, tries=10, wait=1,current_driver=None):
    global driver

    if current_driver == None:
        current_driver = driver

    tries_count = 0
    while tries_count <= tries:
        try:
            current_driver.find_element(By.XPATH, xpath).click()
            tries_count = tries + 1
        except NoSuchElementException:
            tries_count += 1
            sleep(wait)
            if tries_count >= tries:
                print('Element not found')
                return False
            continue

def selenium_get(url,current_driver=None):
    global driver

    if current_driver == None:
        current_driver = driver
        
    current_driver.get(url)

def selenium_quit(current_driver=None):
    global driver

    if current_driver == None:
        current_driver = driver

    current_driver.quit()

def selenium_write(xpath, keys=None,tries=10, wait=1, current_driver=None):
    global driver

    if current_driver == None:
        current_driver = driver

    tries_count = 0
    while tries_count <= tries:
        try:
            current_driver.find_element(By.XPATH, xpath).send_keys(keys)
            tries_count = tries + 1
        except NoSuchElementException:
            tries_count += 1
            sleep(wait)
            if tries_count >= tries:
                print('Element not found')
                return False
            continue
    
def selenium_press_key(current_driver=None,keys=None,qty=1):
    global driver

    if current_driver == None:
        current_driver = driver

    if keys == None:
        keys = []
        print('No keys to perform action')
    
    actions = ActionChains(current_driver) 

    for key in [keys]:
        key = key.lower()
        if key == 'up': actions.send_keys(Keys.UP * qty)
        elif key == 'down': actions.send_keys(Keys.DOWN * qty)
        elif key == 'left': actions.send_keys(Keys.LEFT * qty)
        elif key == 'right': actions.send_keys(Keys.RIGHT * qty)
        elif key == 'enter': actions.send_keys(Keys.ENTER * qty)
        elif key == 'tab': actions.send_keys(Keys.TAB * qty)
        elif key == 'esc': actions.send_keys(Keys.ESCAPE * qty)
        elif key == 'ctrl': actions.send_keys(Keys.CONTROL * qty)
        actions.perform()

def selenium_clear_dropdown_list(current_driver, xpath, tries=10, wait=1):
    global driver

    if current_driver == None:
        current_driver = driver

    tries_count = 0
    while tries_count <= tries:
        try:
            dropdown = Select(current_driver.find_element(By.XPATH, xpath))
            dropdown.deselect_all()
        except NoSuchElementException:
            tries_count += 1
            sleep(wait)
            continue
    if tries_count >= tries:
        print('Element not found')
        return False
    
def selenium_select(xpath, value,current_driver=None):
    global driver

    if current_driver == None:
        current_driver = driver

    button = Select(current_driver.find_element(By.XPATH, xpath))
    button.select_by_value(value)

# COMMAND ---------- [markdown]
# Credentials

# COMMAND ----------
def set_credentials_file_sheet_name(sheet_name):
    global credentials_file_sheet_name
    credentials_file_sheet_name = sheet_name

    return credentials_file_sheet_name

def set_credentials_folder(folder_path=get_downloads_folder(), sheet_name=None):
    global credentials_file
    global credentials_file_sheet_name

    if sheet_name is None:
        sheet_name = credentials_file_sheet_name

    default_credentials_folder = os.path.join(folder_path, '00. Credentials')
    if folder_path == get_downloads_folder():
        folder_path = default_credentials_folder

    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
    
    credentials_file = os.path.join(folder_path, 'Credentials.xlsx')

    if not os.path.exists(credentials_file):
        credentials = pd.DataFrame(columns=['App', 'Login' , 'Senha'])
        credentials.to_excel(credentials_file, sheet_name=sheet_name, index=False)

    return credentials_file

def set_credentials(app_name, login, password, sheet_name=None, file_path=None):
    global credentials_file
    global credentials_file_sheet_name

    if sheet_name is None:
        sheet_name = credentials_file_sheet_name

    if file_path is None and credentials_file is None:
        file_path = set_credentials_folder(sheet_name=sheet_name)
    elif file_path is None:
        file_path = credentials_file

    current_credentials = pd.read_excel(file_path, sheet_name=sheet_name)
    new_credentials = pd.DataFrame({'App': [app_name], 'Login': [login], 'Senha': [password]})
    new_credentials = pd.concat([current_credentials, new_credentials], ignore_index=True)
    new_credentials.reset_index(inplace=True, drop=True)
    new_credentials.to_excel(credentials_file, sheet_name=sheet_name, index=False)

    print('Credentials created succesfully')

def update_credentials(app_name, login, password, sheet_name=None, file_path=None):
    global credentials_file
    global credentials_file_sheet_name

    if sheet_name is None:
        sheet_name = credentials_file_sheet_name

    if file_path is None and credentials_file is None:
        file_path = set_credentials_folder(sheet_name=sheet_name)
    elif file_path is None:
        file_path = credentials_file

    credentials = pd.read_excel(file_path, sheet_name=sheet_name)
    credentials_others = credentials[(credentials.App != app_name)]
    credentials_to_update = credentials[(credentials.App == app_name)]
    credentials_to_update.reset_index(inplace=True, drop=True)

    credentials_to_update.loc[0, 'Login'] = login
    credentials_to_update.loc[0, 'Senha'] = password

    credentials = pd.concat(
        [credentials_others, credentials_to_update], ignore_index=True)
    credentials.to_excel(credentials_file, sheet_name=sheet_name, index=False)

    print('Credentials updated succesfully')

def get_credentials(app_name, sheet_name=None, file_path=None):
    global login
    global password
    global credentials_file
    global credentials_file_sheet_name

    if sheet_name is None:
        sheet_name = credentials_file_sheet_name

    if file_path is None and credentials_file is None:
        file_path = set_credentials_folder(sheet_name=sheet_name)
    elif file_path is None:
        file_path = credentials_file

    credentials_file = file_path
    credentials = pd.read_excel(credentials_file, sheet_name=sheet_name)
    credentials = credentials[(credentials.App == app_name)]
    credentials.reset_index(inplace=True)
    login = credentials['Login'][0]
    password = credentials['Senha'][0]

    if login == None or password == None:
        print('Login or password not found')
        return False

    return login, password

# COMMAND ----------
logs_folder = None
ctrl_c_pressed = False
credentials_file = None
credentials_file_sheet_name = 'All'
downloads_folder = get_downloads_folder()
exit_alarm_controller = False


