from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
import win32com.client
from pathlib import Path

src_path = Path.cwd().joinpath('captcha_data')

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


def do_ocr():
    driver = webdriver.Firefox()
    driver.get('http://www.free-ocr.com/')
    try:
        # for every image present in the folder, repeat the below flow
        for image_file_path in src_path.glob('*.png'):
            print('Reading data from ' + str(image_file_path), end = '')

            # click on browse button
            browse_button = WebDriverWait(driver, 120).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, '[for=fileupfield')))
            browse_button.click()

            # handle windows popup
            handle_file_selector(driver, image_file_path)

            # click on submit button
            submit_button = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, '#fbut')))
            submit_button.click()

            # wait till 'page loading' image/animation disappears
            WebDriverWait(driver, 120).until(EC.invisibility_of_element_located((By.CSS_SELECTOR, '.spinner')))
            WebDriverWait(driver, 120).until(EC.invisibility_of_element_located((By.CSS_SELECTOR, '#result-waiting')))

            try:
                # throws timeout exception if no text is found in image
                text_area = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, '#resultarea')))
                print('text found : ' + text_area.text)

            except TimeoutException:
                text_area = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, '.messagebox.notice')))
                print('text found : ' + text_area.text)

            # click on home button
            home_button = WebDriverWait(driver, 120).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'li.navitem:nth-child(1) > a:nth-child(1)')))

            """" server can block your ip if too many requests are send in very short duration.
            Only purpose of this wait is to increase the time interval between two subsequent requests """

            driver.implicitly_wait(10)

            home_button.click()

    except Exception as ex:
        # print(type(ex).__name__ + ex.args + ex.message)
        print(ex.message)

    finally:
        # driver.quit()
        print('Script Complete')

do_ocr()


# Observations so far :
# Z --> 2
# G --> 6
# S --> 3
# 6 --> S
# s --> 9
# N (italic)  --> ~

#Using JS to click on button
# def jsClick( wd, element) :
#     wd.execute_script("return arguments[0].click();", element);
# jsClick(driver, browse_button)

# Below code gives : elementnotinteractableexception
# browse_button = driver.find_element_by_css_selector('.clear-filefield-button')

#Print inner and outer html
# inner_source_code = browse_button.get_attribute("innerHTML")
# print(inner_source_code)
# # or
# outer_source_code = browse_button.get_attribute("outerHTML")
# print(outer_source_code)

# Display complete source code
# print(driver.page_source.encode("utf-8"))

# WebDriverWait(driver, 10).until(
#                     EC.alert_is_present())
#                 # (By.CSS_SELECTOR, '#loginerrorpanelok'))
#                 alert = driver.switch_to.alert
#                 alert.accept()

#            userid_textfield.screenshot(str(src_path.joinpath('s.png')))



