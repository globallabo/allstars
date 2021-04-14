import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time
import os
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from zipfile import ZipFile
from dotenv import load_dotenv

load_dotenv()  # take environment variables from .env.

output_path = "/mnt/c/Use/c/Doc/Go/All Stars Second Edition/images-uncropped/Level 4/"

# Set up Selenium Chrome Webdriver Options
webdriver_options = Options()
webdriver_options.headless = True
webdriver_options.add_experimental_option('prefs', {
    "download.default_directory": output_path,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True})

webdriver_path = "/usr/bin/chromedriver"
driver = webdriver.Chrome(webdriver_path, options=webdriver_options)
# Freepik info
login_url = 'https://www.freepik.com/profile/login'
username = os.environ.get('FREEPIK_USERNAME')
password = os.environ.get('FREEPIK_PASSWORD')
# Log in to Freepik
driver.get(login_url)
driver.find_element_by_id('gr_login_username').send_keys(username)
driver.find_element_by_id('gr_login_password').send_keys(password)
driver.find_element_by_id('signin_button').click()
# Login takes a few seconds
time.sleep(120)

# levels = [1, 2, 3]
# units = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16]
levels = [4]
# units = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16]
units = [1]
# lessons = [1, 2, 3, 4]
lessons = [1]

for level in levels:
    print(f'Level {level}')
    # Fetch data from Google Sheet
    scope = ["https://spreadsheets.google.com/feeds",
             "https://www.googleapis.com/auth/spreadsheets",
             "https://www.googleapis.com/auth/drive.file",
             "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)
    client = gspread.authorize(creds)
    sheet = client.open("all_stars_revised_0128").worksheet(f"Level {level}")
    data = sheet.get_all_values()

    # Set the starting point of the gspread output
    row = 3
    column = 3

    # Loop through all Units and Lessons
    #  (range() needs a +1 because it stops at the number before)
    for unit in units:
        print(f'Unit {unit}')
        for lesson in lessons:
            print(f'Lesson {lesson}')
            # Set the row based on the unit and lesson
            # - Image links start on the 3rd row and then every 3rd row from there,
            # - moving ahead 12 rows for each unit
            row = 3 + ((unit - 1) * 12) + ((lesson - 1) * 3)
            request_url = data[row][column]
            # print("URL:")
            # print(request_url)
            if request_url:
                driver.get(request_url)
                driver.find_element_by_class_name('download-button').click()
                # Let the file download
                still_downloading = True
                while still_downloading:
                    time.sleep(10)
                    filename = max(
                        [output_path + f for f in os.listdir(output_path)],
                        key=os.path.getctime)
                    if filename.endswith('.jpg'):
                        still_downloading = False
                new_filename = f"AS{level}U{unit:02}L{lesson:02}-story.jpg"
                os.rename(filename, new_filename)
                # *** Freepik now sends JPGs
                # with ZipFile(zip_filename, 'r') as zip_object:
                #     # Get a list of all archived file names from the zip
                #     list_of_filenames = zip_object.namelist()
                #     # Iterate over the file names
                #     for filename in list_of_filenames:
                #         # find the JPG image
                #         if filename.endswith('.jpg'):
                #             # get the info on the JPG file
                #             file_info = zip_object.getinfo(filename)
                #             # rename the JPG
                #             file_info.filename = f"AS{level}U{unit:02}L{lesson:02}-story.jpg"
                #             # Extract the JPG file from the zip
                #             zip_object.extract(file_info, output_path)
                # Delete the zip file when finished
                os.remove(os.path.join(output_path, zip_filename))
        # stop after one unit
        # break
# Quit the Chrome webdriver instance
driver.quit()
