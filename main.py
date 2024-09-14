import win32api
import os
from os.path import exists, dirname
import sys, requests
from zipfile import ZipFile
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import SessionNotCreatedException, WebDriverException
import requests


chromedriver_path = r"E:\Soft\chromedriver.exe"
chromedriver_zip_path = r"E:\Soft\chromedriver_win32.zip"


def get_chrome_version():
    chrome_exe = r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    if not os.path.exists(chrome_exe):
        chrome_exe = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

    fixedInfo = win32api.GetFileVersionInfo(chrome_exe, '\\')
    # return "%d.%d.%d.%d" % (fixedInfo['FileVersionMS'] / 65536,
    #     fixedInfo['FileVersionMS'] % 65536, fixedInfo['FileVersionLS'] / 65536,
    #     fixedInfo['FileVersionLS'] % 65536)

    # Major version
    major = "%d" % (fixedInfo['FileVersionMS'] / 65536)
    resp = requests.get(f'https://chromedriver.storage.googleapis.com/LATEST_RELEASE_{major}')
    return resp.content.decode('UTF-8')
    
    
def create_chrome(recurs=False):
    chrome_options = Options()
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                                "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.106 Safari/537.36")
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
    try:
        chrome = webdriver.Chrome(chromedriver_path, options=chrome_options)
    except (SessionNotCreatedException, WebDriverException) as e:
        print(str(e))
        if recurs:
            print('Ошибка при запуске chrome driver:\n', str(e))
            sys.exit(1)
        s_chrome_version = get_chrome_version()
        chromedriver_dl_url = f'https://chromedriver.storage.googleapis.com/{s_chrome_version}/chromedriver_win32.zip'
        print(chromedriver_dl_url)
        resp = requests.get(chromedriver_dl_url)
        with open(chromedriver_zip_path, 'wb') as binary_file:
            binary_file.write(resp.content)

        if not exists(chromedriver_zip_path):
            sys.exit(1)

        with ZipFile(chromedriver_zip_path) as zipfile:
            for filename in zipfile.namelist():
                if filename == 'chromedriver.exe':
                    zipfile.extract(filename, dirname(chromedriver_path))
                    break

        # Запускаем ещё раз эту функцию
        return create_chrome(True)
    
    chrome.set_window_size(1000, 1000)

    return chrome