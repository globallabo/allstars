import gspread
from oauth2client.service_account import ServiceAccountCredentials
# from pprint import pprint
from weasyprint import HTML, CSS
from weasyprint.fonts import FontConfiguration
from string import Template
import logging

# Log WeasyPrint output
logger = logging.getLogger('weasyprint')
logger.addHandler(logging.FileHandler('/tmp/weasyprint.log'))

levels = [1, 2, 3, 4]
# units = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16]
# levels = [4]
units = [1]
# units = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16]

for level in levels:
    # Create HTML template for image flashcards
    font_config = FontConfiguration()
    template_image_filename = "flashcards-images-template.html"
    with open(template_image_filename, "r") as template_image_file:
        template_image_file_contents = template_image_file.read()
    template_image_string = Template(template_image_file_contents)
    # Create HTML template for word flashcards
    template_word_filename = "flashcards-words-template.html"
    template_word_file = open(template_word_filename, "r")
    template_word_file_contents = template_word_file.read()
    template_word_string = Template(template_word_file_contents)
    # Load CSS (not necessary)
    css_filename = "flashcards.css"
    with open(css_filename, "r") as css_file:
        css_string = css_file.read()

    # output_path = f'/Users/cbunn/Documents/Employment/5 Star/Google Drive/All Stars Second Edition/All Stars Second Edition/Flashcards/Level {level}/'
    # output_path = f'/Users/cbunn/Documents/Employment/5 Star/Google Drive/All Stars Second Edition/test-output-fc/Level {level}/'
    output_path = f'/Users/cbunn/Documents/Employment/5 Star/Google Drive/All Stars Second Edition/unit1-output/flashcards/'

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
    row = 1
    column = 4

    # Loop through all Units and Lessons
    #  (range() needs a +1 because it stops at the number before)
    for unit in units:
        # Set the row based on the unit
        row = 1 + (unit-1) * 12
        template_mapping = dict()
        template_mapping["level"] = level
        # This is used for the page header:
        template_mapping["unit"] = unit
        # This is used for image naming:
        template_mapping["unit_zfill"] = str(unit).zfill(2)

        print(f'Unit: {unit}')

        # set vocab vars
        template_mapping["vocab1"] = data[row][column] \
            if '/' not in data[row][column] \
            else f'<ul><li>{data[row][column].split(sep="/")[0]}</li><li>{data[row][column].split(sep="/")[1]}</li></ul>'
        template_mapping["vocab2"] = data[row][column+1] \
            if '/' not in data[row][column+1] \
            else f'<ul><li>{data[row][column+1].split(sep="/")[0]}</li><li>{data[row][column+1].split(sep="/")[1]}</li></ul>'
        template_mapping["vocab3"] = data[row][column+2] \
            if '/' not in data[row][column+2] \
            else f'<ul><li>{data[row][column+2].split(sep="/")[0]}</li><li>{data[row][column+2].split(sep="/")[1]}</li></ul>'
        template_mapping["vocab4"] = data[row][column+3] \
            if '/' not in data[row][column+3] \
            else f'<ul><li>{data[row][column+3].split(sep="/")[0]}</li><li>{data[row][column+3].split(sep="/")[1]}</li></ul>'
        template_mapping["vocab5"] = data[row+3][column] \
            if '/' not in data[row+3][column] \
            else f'<ul><li>{data[row+3][column].split(sep="/")[0]}</li><li>{data[row+3][column].split(sep="/")[1]}</li></ul>'
        template_mapping["vocab6"] = data[row+3][column+1] \
            if '/' not in data[row+3][column+1] \
            else f'<ul><li>{data[row+3][column+1].split(sep="/")[0]}</li><li>{data[row+3][column+1].split(sep="/")[1]}</li></ul>'
        template_mapping["vocab7"] = data[row+3][column+2] \
            if '/' not in data[row+3][column+2] \
            else f'<ul><li>{data[row+3][column+2].split(sep="/")[0]}</li><li>{data[row+3][column+2].split(sep="/")[1]}</li></ul>'
        template_mapping["vocab8"] = data[row+3][column+3] \
            if '/' not in data[row+3][column+3] \
            else f'<ul><li>{data[row+3][column+3].split(sep="/")[0]}</li><li>{data[row+3][column+3].split(sep="/")[1]}</li></ul>'

        # Substitute
        template_word_filled = template_word_string.safe_substitute(template_mapping)
        template_image_filled = template_image_string.safe_substitute(template_mapping)

        # css = CSS(string=css_string, font_config=font_config)
        # html.write_pdf('/tmp/test-py.pdf',
        #                stylesheets=[css],
        #                font_config=font_config)
        f_level = str(level)
        f_unit = str(unit).zfill(2)
        # print(template_string)
        html_word = HTML(string=template_word_filled)
        html_word.write_pdf(f'{output_path}AS{f_level}U{f_unit}-word_flashcards.pdf')
        if level == 1:
            html_image = HTML(string=template_image_filled)
            html_image.write_pdf(
                f'{output_path}AS{f_level}U{f_unit}-image_flashcards.pdf')

        # Advance to the next unit
        # row += 12
        # stop after one Unit
        # break
