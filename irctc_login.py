from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from pathlib import Path
from time import sleep
import urllib
import os
import win32com.client
from PIL import Image

src_path = Path.cwd().joinpath('irctc_captcha')

""" Create webdriver instance """


def create_webdriver_instance(browser='firefox'):
    if browser == 'chrome':
        driver = webdriver.Chrome()
        driver.maximize_window()
        return driver

    elif browser == 'firefox':
        return webdriver.Firefox()


"""" handle file selector windows popup using AutoIt """


def handle_file_selector(driver, file_path):
    autoit = win32com.client.Dispatch("AutoItX3.Control")

    if driver.capabilities['browserName'] == 'firefox':
        title = 'File Upload'
    elif driver.capabilities['browserName'] == 'chrome':
        title = 'Open'

    autoit.ControlFocus(title, "", "Edit1")
    autoit.Sleep(250)
    autoit.ControlSetText(title, "", "Edit1", file_path)
    autoit.Sleep(250)
    autoit.ControlClick(title, "", "Button1");
    autoit.Sleep(500)


"""" Complete workflow for ocr """


def do_ocr(image_file_path, browser_name):
    driver = create_webdriver_instance(browser_name)
    driver.get('http://www.free-ocr.com/')
    img_text = ''
    try:

        # click on browse button
        browse_button = WebDriverWait(driver, 120).until(
            ec.presence_of_element_located((By.CSS_SELECTOR, '[for=fileupfield')))
        browse_button.click()

        # handle windows popup
        handle_file_selector(driver, image_file_path)

        # click on submit button
        submit_button = WebDriverWait(driver, 10).until(
            ec.presence_of_element_located((By.CSS_SELECTOR, '#fbut')))
        submit_button.click()

        # wait till 'page loading' image/animation disappears
        WebDriverWait(driver, 120).until(ec.invisibility_of_element_located((By.CSS_SELECTOR, '.spinner')))
        WebDriverWait(driver, 120).until(ec.invisibility_of_element_located((By.CSS_SELECTOR, '#result-waiting')))

        try:

            # throws timeout exception if no text is found in image
            text_area = WebDriverWait(driver, 10).until(
                ec.presence_of_element_located((By.CSS_SELECTOR, '#resultarea')))
            img_text = text_area.text

        except TimeoutException:

            text_area = WebDriverWait(driver, 10).until(
                ec.presence_of_element_located((By.CSS_SELECTOR, '.messagebox.notice')))
            img_text = text_area.text

        # click on home button
        home_button = WebDriverWait(driver, 120).until(
            ec.presence_of_element_located((By.CSS_SELECTOR, 'li.navitem:nth-child(1) > a:nth-child(1)')))

        """" server can block your ip if too many requests are send in very short duration.
        Only purpose of this wait is to increase the time interval between two subsequent requests """

        # driver.implicitly_wait(10)
        home_button.click()

    except Exception as ex:
        driver.save_screenshot(str(src_path.parent.joinpath('exception_screenshot').joinpath('exception1.png')))
        print(ex.message)

    finally:
        driver.quit()
        return img_text


""" Make new folder. If it already exists, delete it and create new one."""


def make_folder(folder_path):
    try:
        os.makedirs(folder_path)
    except OSError:
        os.system("rm -rf " + folder_path)
        os.makedirs(folder_path)


""" download image and save it in a folder """


def download_image(image_url):
    make_folder(str(src_path))
    try:
        urllib.request.urlretrieve(image_url, str(src_path.joinpath('0.png')))
        sleep(0.5)
    except urllib.error.URLError as e:
        print(e)


"""" Complete workflow for irctc  login """


def irctc_login():
    browser_name = input("Enter browser name firefox/chrome : ")
    user_id = input("Enter irctc user_id : ")
    password = input("Enter irctc password : ")

    banner_captcha_flag = True
    driver = create_webdriver_instance(browser_name)
    driver.get('https://irctc.co.in/eticketing/home')

    try:
        # store window handle of current driver window
        main_window = driver.current_window_handle

        while 1:
            # type username
            userid_textfield = WebDriverWait(driver, 60).until(
                ec.presence_of_element_located((By.CSS_SELECTOR, '#usernameId')))
            userid_textfield.send_keys(user_id)

            # type password
            password_textfield = WebDriverWait(driver, 10).until(
                ec.presence_of_element_located((By.CSS_SELECTOR, '.loginPassword')))
            password_textfield.send_keys(password)

            # get captcha image
            try:
                # check for type of captcha
                captcha_type = WebDriverWait(driver, 10).until(
                    ec.presence_of_element_located((By.CSS_SELECTOR, '#nlpAdType')))
                captcha_type = captcha_type.get_attribute("value")

                # Check if captcha type is traditional
                if captcha_type == 'traditional':
                    captcha_image = WebDriverWait(driver, 10).until(
                        ec.presence_of_element_located((By.CSS_SELECTOR, 'img#nlpCaptchaImg')))

                # Check if Banner Captcha is present
                else:
                    captcha_image = WebDriverWait(driver, 10).until(
                        ec.presence_of_element_located((By.CSS_SELECTOR, 'img#captchaImg')))
                image_url = captcha_image.get_attribute("src")

                # Convert theme captcha into banner captcha
                if image_url.split('/')[-1] == 'theme1':
                    image_url = image_url.replace('theme1', 'banner')
                download_image(image_url)

            except TimeoutException:
                captcha_image = WebDriverWait(driver, 5).until(
                    ec.presence_of_element_located((By.CSS_SELECTOR, 'img#cimage')))

                # found basic captcha
                banner_captcha_flag = False

                # Incase of basic captcha, img download is not possible because every time we hit image url it produce
                # a fresh captcha
                # So instead we take screenshot of <img> tag.

                if driver.name.lower() == 'firefox':
                    captcha_image.screenshot(str(src_path.joinpath('0.png')))

                else:
                    # for chrome, take screenshot of full webpage and crop the part in which we are interested

                    location = captcha_image.location
                    size = captcha_image.size

                    driver.save_screenshot(str(src_path.joinpath('0.png')))

                    x = location['x']
                    y = location['y']
                    width = location['x'] + size['width']
                    height = location['y'] + size['height']

                    captcha_image = Image.open(str(src_path.joinpath('0.png')))
                    captcha_image = captcha_image.crop((int(x), int(y), int(width), int(height)))
                    captcha_image.save(str(src_path.joinpath('0.png')), optimize=True, quality=95)

            # get text from captcha image
            img_text = do_ocr(src_path.joinpath('0.png'), browser_name)
            if banner_captcha_flag:
                img_text = img_text.split(':')[-1].replace(' ', '').strip().upper()
            else:
                img_text = img_text.strip().upper()

            # switch back to main window
            driver.switch_to.window(main_window)

            driver.implicitly_wait(5)

            # type text in captcha text field
            if banner_captcha_flag:
                captcha_textfield = WebDriverWait(driver, 10).until(
                    ec.presence_of_element_located((By.CSS_SELECTOR, '#nlpAnswer')))
                captcha_textfield.send_keys(img_text)
            else:
                captcha_textfield = WebDriverWait(driver, 10).until(
                    ec.presence_of_element_located((By.CSS_SELECTOR, 'input.loginCaptcha')))
                captcha_textfield.send_keys(img_text)

            # Click on login button
            login_button = WebDriverWait(driver, 10).until(
                ec.presence_of_element_located((By.CSS_SELECTOR, '#loginbutton')))
            login_button.click()

            try:
                # Handle 'wrong captcha' pop up window
                WebDriverWait(driver, 10).until(
                    ec.presence_of_element_located((By.CSS_SELECTOR, '#loginerrorpanelok')))
                invalid_captcha_button = WebDriverWait(driver, 10).until(

                    # Once I got an error where this particular button was not interactable though it was visible in dom
                    # To handle that below wait was added

                    ec.element_to_be_clickable((By.CSS_SELECTOR, '#loginerrorpanelok')))
                invalid_captcha_button.click()

            except TimeoutException:
                print('Login Successful')
                break

    except Exception as ex:
        driver.save_screenshot(str(src_path.parent.joinpath('exception_screenshot').joinpath('exception1.png')))
        print(ex.message)

    finally:
        driver.quit()
        print('Script Complete')


irctc_login()
