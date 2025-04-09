import win32api
from os import makedirs, rmdir
from os.path import exists, dirname, basename, join
import shutil
from time import sleep
from zipfile import ZipFile
import sys
import requests
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import SessionNotCreatedException, WebDriverException


CHROMEDRIVER_PATH = r"E:\Soft\chromedriver.exe"
CHROMEDRIVER_ZIP_PATH32 = r"E:\Soft\chromedriver_win32.zip"
CHROMEDRIVER_ZIP_PATH64 = r"E:\Soft\chromedriver_win64.zip"
PATH_INSIDE_ZIP = 'chromedriver-win64'
CHROMEDRIVER_ZIP_PATH = CHROMEDRIVER_ZIP_PATH64


def get_installed_chrome_version():
    chrome_exe = r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    if not exists(chrome_exe):
        chrome_exe = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
    if not exists(chrome_exe):
        print('Chrome is not installed')
        sys.exit(1)

    fixedInfo = win32api.GetFileVersionInfo(chrome_exe, '\\')
    fullVersion = "%d.%d.%d.%d" % (
        fixedInfo['FileVersionMS'] / 65536,
        fixedInfo['FileVersionMS'] % 65536,
        fixedInfo['FileVersionLS'] / 65536,
        fixedInfo['FileVersionLS'] % 65536
    )

    # Major version
    print('Installed Chrome version: ' + fullVersion)

    return fullVersion
    

def download_new_driver(s_chrome_version: str):
    # Как варианты
    # chromedriver_dl_url = 'https://googlechromelabs.github.io/chrome-for-testing/last-known-good-versions-with-downloads.json'
    # https://googlechromelabs.github.io/chrome-for-testing/known-good-versions-with-downloads.json  # ALL VERSIONs
    chromedriver_dl_url = 'https://googlechromelabs.github.io/chrome-for-testing/latest-versions-per-milestone-with-downloads.json'

    j = requests.get(chromedriver_dl_url).json()
    platforms = j['milestones'][s_chrome_version[:s_chrome_version.find('.')]]['downloads']['chromedriver']
    for p in platforms:
        if p['platform'] == 'win64':
            chromedriver_dl_url = p['url']
            break

    print('Chromedriver download URL:', )
    print(f'Downloading chromedriver... ({chromedriver_dl_url})')
    resp = requests.get(chromedriver_dl_url)
    
    # Making folder before writing file and write chromedriver.exe
    out_folder = dirname(CHROMEDRIVER_PATH)
    makedirs(out_folder, exist_ok=True)
    with open(CHROMEDRIVER_ZIP_PATH, 'wb') as binary_file:
        binary_file.write(resp.content)

    if not exists(CHROMEDRIVER_ZIP_PATH):
        print("[Error] Chromedriver doesn't downloaded")
        sys.exit(1)

    # unzipping zip-file
    with ZipFile(CHROMEDRIVER_ZIP_PATH) as zipfile:
        for filename in zipfile.namelist():
            fn = basename(filename)
            if fn == 'chromedriver.exe':
                zipfile.extract(filename, out_folder)
                shutil.move(join(out_folder, filename), CHROMEDRIVER_PATH)
                sleep(1)
                rmdir(join(out_folder, dirname(filename)))
                break

    
def create_chrome(recurs=False):
    s_chrome_version = get_installed_chrome_version()
    chrome_options = Options()
    major_version = s_chrome_version[:s_chrome_version.find('.')]
    chrome_options.add_argument(f'user-agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/{major_version}.0.0.0 Safari/537.36"')
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])

    try:
        chrome = webdriver.Chrome(CHROMEDRIVER_PATH, options=chrome_options)
    except (SessionNotCreatedException, WebDriverException) as e:
        print('AHTUNG! Error: ' + str(e))
        if recurs:
            print('Ошибка при запуске chrome driver:\n', str(e))
            sys.exit(1)
            
        download_new_driver(s_chrome_version)

        # Запускаем ещё раз эту функцию
        return create_chrome(True)
    
    chrome.set_window_size(1000, 1000)
    
    return chrome