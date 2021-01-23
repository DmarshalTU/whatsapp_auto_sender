from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.keys import Keys
import pyperclip
import pandas as pd
import openpyxl
from config import CHROME_PROFILE_PATH
import os
import time


def send_message(group_list, message_list):
    time.sleep(3)
    browser = webdriver.Chrome(
        executable_path=BASE_DIR + '\\web_drivers\\chromedriver', options=options
    )

    browser.get('https://web.whatsapp.com/')
    for group_name in group_list:
        search_xpath = '/html/body/div[1]/div/div/div[3]/div/div[1]/div/label/div/div[2]'

        search_box = WebDriverWait(browser, 50).until(
            ec.presence_of_element_located((By.XPATH, search_xpath))
        )

        search_box.clear()
        print("Contact name: " + group_name)
        pyperclip.copy(group_name)

        search_box.send_keys(Keys.CONTROL + "v")

        group_xpath = '/html/body/div[1]/div/div/div[3]/div/div[2]/div[1]/div/div/div[1]'
        time.sleep(2)

        group_title = WebDriverWait(browser, 50).until(
            ec.presence_of_element_located((By.XPATH, group_xpath))
        )

        group_title.click()

        input_xpath = '/html/body/div[1]/div/div/div[4]/div/footer/div[1]/div[2]/div/div[2]'
        time.sleep(2)
        input_box = WebDriverWait(browser, 50).until(
            ec.presence_of_element_located((By.XPATH, input_xpath))
        )
        time.sleep(2)
        print("Message sent: " + message_list)
        print()
        pyperclip.copy(message_list)
        input_box.send_keys(Keys.CONTROL + "v")
        input_box.send_keys(Keys.ENTER)
        time.sleep(2)


def send_image_and_text(group_list, message_list, image):
    time.sleep(3)
    browser = webdriver.Chrome(
        executable_path=BASE_DIR + '\\web_drivers\\chromedriver', options=options
    )
    browser.get('https://web.whatsapp.com/')

    for group_name in group_list:
        search_xpath = '/html/body/div[1]/div/div/div[3]/div/div[1]/div/label/div/div[2]'

        search_box = WebDriverWait(browser, 50).until(
            ec.presence_of_element_located((By.XPATH, search_xpath))
        )

        search_box.clear()
        print("Contact name: " + group_name)
        pyperclip.copy(group_name)

        search_box.send_keys(Keys.CONTROL + "v")

        group_xpath = '/html/body/div[1]/div/div/div[3]/div/div[2]/div[1]/div/div/div[1]'
        time.sleep(2)

        group_title = WebDriverWait(browser, 50).until(
            ec.presence_of_element_located((By.XPATH, group_xpath))
        )

        group_title.click()

        attachment_xpath = '/html/body/div[1]/div/div/div[4]/div/footer/div[1]/div[1]/div[2]/div/div'
        attachment_box = WebDriverWait(browser, 50).until(
            ec.presence_of_element_located((By.XPATH, attachment_xpath))
        )
        attachment_box.click()
        time.sleep(2)

        image_xpath = '/html/body/div[1]/div/div/div[4]/div/footer/div[1]/div[1]/div[2]/div/span/div/div/ul/li[' \
                      '1]/button/input '
        image_box = WebDriverWait(browser, 50).until(
            ec.presence_of_element_located((By.XPATH, image_xpath))
        )
        image_box.send_keys(image)
        time.sleep(2)

        input_xpath1 = '//*[@id="app"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/div[1]/span/div/div[' \
                       '2]/div/div[3]/div[1]/div[2] '

        input_box1 = WebDriverWait(browser, 500).until(
            ec.presence_of_element_located((By.XPATH, input_xpath1))
        )
        time.sleep(2)

        pyperclip.copy(message_list)
        input_box1.send_keys(Keys.CONTROL + "v")
        input_box1.send_keys(Keys.ENTER)
        time.sleep(2)


def image_only(group_list, image):
    time.sleep(3)
    browser = webdriver.Chrome(
        executable_path=BASE_DIR + '\\web_drivers\\chromedriver', options=options
    )
    browser.get('https://web.whatsapp.com/')

    for group_name in group_list:
        search_xpath = '/html/body/div[1]/div/div/div[3]/div/div[1]/div/label/div/div[2]'

        search_box = WebDriverWait(browser, 50).until(
            ec.presence_of_element_located((By.XPATH, search_xpath))
        )

        search_box.clear()
        print("Contact name: " + group_name)
        pyperclip.copy(group_name)

        search_box.send_keys(Keys.CONTROL + "v")

        group_xpath = '/html/body/div[1]/div/div/div[3]/div/div[2]/div[1]/div/div/div[1]'
        time.sleep(2)

        group_title = WebDriverWait(browser, 50).until(
            ec.presence_of_element_located((By.XPATH, group_xpath))
        )

        group_title.click()

        attachment_xpath = '/html/body/div[1]/div/div/div[4]/div/footer/div[1]/div[1]/div[2]/div/div'
        attachment_box = WebDriverWait(browser, 50).until(
            ec.presence_of_element_located((By.XPATH, attachment_xpath))
        )
        attachment_box.click()
        time.sleep(2)

        image_xpath = '/html/body/div[1]/div/div/div[4]/div/footer/div[1]/div[1]/div[2]/div/span/div/div/ul/li[' \
                      '1]/button/input '
        image_box = WebDriverWait(browser, 50).until(
            ec.presence_of_element_located((By.XPATH, image_xpath))
        )
        image_box.send_keys(image)
        print("Image name: " + image)
        print()
        time.sleep(2)

        input_xpath1 = '//*[@id="app"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/div[1]/span/div/div[' \
                       '2]/div/div[3]/div[1]/div[2] '

        input_box1 = WebDriverWait(browser, 500).until(
            ec.presence_of_element_located((By.XPATH, input_xpath1))
        )
        time.sleep(2)
        input_box1.send_keys(Keys.ENTER)
        time.sleep(2)


def exel(group_list, qwerty, browser):
    time.sleep(3)

    for group_name in group_list:
        search_xpath = '/html/body/div[1]/div/div/div[3]/div/div[1]/div/label/div/div[2]'

        search_box = WebDriverWait(browser, 50).until(
            ec.presence_of_element_located((By.XPATH, search_xpath))
        )

        search_box.clear()
        print("Contact name: " + group_name)
        pyperclip.copy(group_name)

        search_box.send_keys(Keys.CONTROL + "v")

        group_xpath = '/html/body/div[1]/div/div/div[3]/div/div[2]/div[1]/div/div/div[1]'
        time.sleep(2)

        group_title = WebDriverWait(browser, 50).until(
            ec.presence_of_element_located((By.XPATH, group_xpath))
        )

        group_title.click()

        input_xpath = '/html/body/div[1]/div/div/div[4]/div/footer/div[1]/div[2]/div/div[2]'
        time.sleep(2)
        input_box = WebDriverWait(browser, 50).until(
            ec.presence_of_element_located((By.XPATH, input_xpath))
        )
        time.sleep(2)
        print()
        pyperclip.copy(qwerty)
        input_box.send_keys(Keys.CONTROL + "v")
        input_box.send_keys(Keys.ENTER)
        time.sleep(2)


if __name__ == '__main__':

    # Directory path to main file:
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    # Set PATH to save login info, (one time QR code scan)
    options = webdriver.ChromeOptions()
    options.add_argument(CHROME_PROFILE_PATH)

    chose_mode = input("Chose mode: (T)ext, (B)oth, (I)mage , (E)xel >>> ").lower()

    # Only TEXT Mode
    if chose_mode == 't':
        gr = BASE_DIR + '\\assets\\' + 'gr.txt'
        ms = BASE_DIR + '\\assets\\' + 'ms.txt' if gr else print('False')

        try:
            with open(gr, 'r', encoding='utf8') as group:
                groups = [group.strip() for group in group.readlines()]
        except FileNotFoundError:
            print('File for group not found')

        try:
            with open(ms, 'r', encoding='utf8') as message:
                msg = message.read()
        except IndexError:
            print('File not found')

        send_message(groups, msg)

    # Text and Image Mode
    elif chose_mode == 'b':
        gr = BASE_DIR + '\\assets\\' + 'gr.txt'
        ms = BASE_DIR + '\\assets\\' + 'ms.txt' if gr else print('False')
        im = BASE_DIR + '\\assets\\' + 'im.jpg' if ms else print('False')

        try:
            with open(gr, 'r', encoding='utf8') as group:
                groups = [group.strip() for group in group.readlines()]
        except IndexError:
            print('File for group not found')

        try:
            with open(ms, 'r', encoding='utf8') as message:
                msg = message.read()
        except IndexError:
            print('File not found')

        send_image_and_text(groups, msg, im)

    # Only IMAGE Mode
    elif chose_mode == 'i':
        gr = BASE_DIR + '\\assets\\' + 'gr.txt'
        im = BASE_DIR + '\\assets\\' + 'im.jpg' if gr else print('False')

        try:
            with open(gr, 'r', encoding='utf8') as group:
                groups = [group.strip() for group in group.readlines()]
        except IndexError:
            print('File not found')

        image_only(groups, im)

    # Exel format
    elif chose_mode == 'e':
        gr = BASE_DIR + '\\assets\\' + 'gr.txt'
        ms = BASE_DIR + '\\assets\\' + 'ms.txt'
        xl = pd.ExcelFile(input("Enter the path to exel file:\n>>> "))
        df = xl.parse(input("Enter the sheet name:\n>>> "))
        row = len(df)
        col = len(df.columns)
        time.sleep(3)
        browser = webdriver.Chrome(
            executable_path=BASE_DIR + '\\web_drivers\\chromedriver', options=options
        )

        browser.get('https://web.whatsapp.com/')

        try:
            with open(gr, 'r', encoding='utf8') as group:
                groups = [group.strip() for group in group.readlines()]
        except FileNotFoundError:
            print('File for group not found')

        try:
            for i in range(0, row):
                with open(ms, 'w', encoding='utf8') as ms1:
                    costumer = (df['name'][i]) # change the column name to one that contain the costumer name value
                    total = (df['total'][i]) # the total column name

                    message = f"Hello dear {costumer}, the total bill is: ${total}."
                    ms1.write(f"{message}\r\n")
                try:

                    with open(ms, 'r', encoding='utf8') as ex:
                        ex = ex.read()

                except IndexError:
                    print('File not found')

                exel(groups, ex, browser)

        except IndexError:
            print('File not found')

