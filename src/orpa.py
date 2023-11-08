# %% [markdown]
# # orpa (RPA Functions)

# %%
import os
import time
import random
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

credentials_file = None
credentials_file_sheet_name = 'All'


def open_mouse_info():
    mouseInfo()


def sleep(seconds):
    time.sleep(seconds)


def press(key):
    pyautogui.press(key)


def hotkey(key1, key2):
    pyautogui.hotkey(key1, key2)


def wait(mintime=3, maxtime=5):
    mintime = mintime * 60
    maxtime = maxtime * 60
    x = random.uniform(mintime, maxtime)
    print('Iniciando próxima ação em ' +
          str(datetime.timedelta(seconds=x)) + ' min \n')
    time.sleep(x)


def copy_clipboard():
    # <- This prevents last copy replacing current copy of null.
    pyperclip.copy("")
    pyautogui.hotkey('ctrl', 'c')
    # ctrl-c is usually very fast but your program may execute faster
    pyautogui.sleep(.01)

    return pyperclip.paste()


def open_microsoft_edge(link=''):
    os.startfile(
        'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe')
    pyautogui.sleep(2)
    pyautogui.getActiveWindow().maximize()

    if link != '':
        pyautogui.typewrite(link)
        pyautogui.press('enter')


def found_all_buttons_and_click(buttons):
    for button in buttons.split(' ', 0) if type(buttons) == str else buttons:
        r = None
        while r == None:
            r = pyautogui.locateOnScreen(button)
        pyautogui.click(r)


def wait_until_download_edge(download_button):
    r = None
    while r == None:
        r = pyautogui.locateOnScreen(download_button, grayscale=False)


def get_screen_size():
    image = ImageGrab.grab()
    return image.height, image.width


def get_screen_height():
    image = ImageGrab.grab()
    return image.height


def get_screen_width():
    image = ImageGrab.grab()
    return image.width


def screenshot_to_clipboard(x_top, y_top, x_bottom, y_bottom):
    image = ImageGrab.grab(bbox=(x_top, y_top, x_bottom, y_bottom))
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
        credentials = pd.DataFrame(columns=['App', 'Login', 'Senha'])
        credentials.to_excel(
            credentials_file, sheet_name=sheet_name, index=False)

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
    new_credentials = pd.DataFrame(
        {'App': [app_name], 'Login': [login], 'Senha': [password]})
    new_credentials = pd.concat(
        [current_credentials, new_credentials], ignore_index=True)
    new_credentials.to_excel(
        credentials_file, sheet_name=sheet_name, index=False)

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


def get_credentials(app_name, sheet_name='All', file_path=r'C:\Users\B117539\Downloads\00. Credentials\Credentials.xlsx'):
    credentials_file = file_path
    credentials = pd.read_excel(credentials_file, sheet_name=sheet_name)
    credentials = credentials[(credentials.App == app_name)]
    credentials.reset_index(inplace=True)

    global login
    global password
    login = credentials['Login'][0]
    password = credentials['Senha'][0]

    if login == None or password == None:
        print('Login or password not found')
        return False

    return login, password


def selenium_action(driver, xpath, action='click', keys=None, tries=10, wait=1):
    tries_count = 0
    while tries_count <= tries:
        try:
            if action == 'click':
                driver.find_element(By.XPATH, xpath).click()
            elif action == 'send_keys':
                driver.find_element(By.XPATH, xpath).send_keys(keys)
            elif action == 'get':
                driver.get(keys)
            tries_count = tries + 1
        except NoSuchElementException:
            tries_count += 1
            sleep(wait)
            continue
    if tries_count >= tries:
        print('Element not found')
        return False


def selenium_clear_dropdown_list(driver, xpath, tries=10, wait=1):
    tries_count = 0
    while tries_count <= tries:
        try:
            dropdown = Select(driver.find_element(By.XPATH, xpath))
            dropdown.deselect_all()
        except NoSuchElementException:
            tries_count += 1
            sleep(wait)
            continue
    if tries_count >= tries:
        print('Element not found')
        return False


def start_saving_logs():
    global logs_df
    logs_df = pd.DataFrame()

# create a function to append the dataframe


def job(job_name='Job'):
    global logs_df
    df = pd.DataFrame({'Date': format(datetime.now())[
                      :19], 'Job': job_name}, index=[0])
    try:
        logs_df = pd.concat([logs_df, df], ignore_index=True)
    except Exception as e:
        start_saving_logs()
        logs_df = pd.concat([logs_df, df], ignore_index=True)


def save_logs(file_path=os.path.join(os.path.expanduser('~'), 'Downloads'), file_prefix='', mode=None):
    global logs_df
    file_prefix = ' ' + file_prefix if file_prefix != '' else file_prefix
    logs_file_name = os.path.join(file_path, format(datetime.now())[:10].replace(
        ' ', '-').replace(':', '-') + file_prefix + ' Logs' + '.txt')
    logs_df.to_csv(logs_file_name, mode='a', index=False,
                   header=not (os.path.isfile(logs_file_name)))
    if mode == 'skipline':
        with open(logs_file_name, 'a') as f:
            f.write('\n')
    logs_df = pd.DataFrame()


def get_active_window():
    return print(GetWindowText(GetForegroundWindow()))


def hide_edge():
    global edge_options
    # Open Microsoft Edge Browser
    edge_options = webdriver.EdgeOptions()
    edge_options.add_experimental_option("prefs", {
                                         "download.prompt_for_download": False, 'profile.default_content_settings.popups': False})
    edge_options.add_argument("--window-size=1920,1080")
    edge_options.add_argument("--start-maximized")
    edge_options.add_argument("--headless=new")


def open_edge(link='https://www.google.com/', hidden=False):
    global driver
    global edge_options

    if hidden:
        hide_edge()
        driver = webdriver.Edge(options=edge_options)
    else:
        driver = webdriver.Edge()
        driver.maximize_window()
    driver.get(link)


def power_automates_notify(notify_url, notify_path):
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


def send_email_notification(subject, body, to=None, importance='normal'):
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
        mail.To = main_account
        mail.Subject = subject
        if importance == 'high':
            mail.Importance = 2
        elif importance == 'normal':
            mail.Importance = 1
        elif importance == 'low':
            mail.Importance = 0
        mail.Body = body
        mail.Send()


def selenium_wait_page_load(driver):
    while True:
        page_state = driver.execute_script('return document.readyState;')
        if page_state == 'complete':
            break
        else:
            sleep(1)


def selenium_perform_action(driver, Keys=None):
    if Keys == None:
        Keys = []
        print('No keys to perform action')

    actions = ActionChains(driver)

    for key in Keys:
        key = key.lower()
        if key == 'up':
            actions.send_keys(Keys.UP)
        elif key == 'down':
            actions.send_keys(Keys.DOWN)
        elif key == 'left':
            actions.send_keys(Keys.LEFT)
        elif key == 'right':
            actions.send_keys(Keys.RIGHT)
        elif key == 'enter':
            actions.send_keys(Keys.ENTER)
        elif key == 'tab':
            actions.send_keys(Keys.TAB)
        elif key == 'esc':
            actions.send_keys(Keys.ESCAPE)
        elif key == 'ctrl':
            actions.key_down(Keys.CONTROL)
        actions.perform()
