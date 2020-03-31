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

# Create HTML template for image flashcards
font_config = FontConfiguration()
template_image_filename = "AS1-flashcards-images-template.html"
template_image_file = open(template_image_filename, "r")
template_image_file_contents = template_image_file.read()
template_image_string = Template(template_image_file_contents)
# Create HTML template for word flashcards
template_word_filename = "AS1-flashcards-words-template.html"
template_word_file = open(template_word_filename, "r")
template_word_file_contents = template_word_file.read()
template_word_string = Template(template_word_file_contents)
# Load CSS (not necessary)
css_filename = "flashcards.css"
css_file = open(css_filename, "r")
css_string = css_file.read()

output_path = "/Users/cbunn/Documents/Employment/5 Star/Google Drive/All Stars Second Edition/test-output/"

# Fetch data from Google Sheet
scope = ["https://spreadsheets.google.com/feeds",
         "https://www.googleapis.com/auth/spreadsheets",
         "https://www.googleapis.com/auth/drive.file",
         "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)
client = gspread.authorize(creds)
sheet = client.open("all_stars_revised_0128").sheet1
data = sheet.get_all_values()

# So far, we're only doing Level 1, but in the future, we'll have to deal
#  with the others
level = 1
numUnits = 16

# Set the starting point of the gspread output
row = 1
column = 4

# Loop through all Units and Lessons
#  (range() needs a +1 because it stops at the number before)
for unit in range(1, numUnits+1):
    template_mapping = dict()
    template_mapping["level"] = level
    # This is used for the page header:
    template_mapping["unit"] = unit
    # This is used for image naming:
    template_mapping["unit_zfill"] = str(unit).zfill(2)

    print("Unit: " + str(unit))

    # set vocab vars
    vocab1 = data[row][column]
    template_mapping["vocab1"] = vocab1
    # print(vocab1)
    vocab2 = data[row][column+1]
    template_mapping["vocab2"] = vocab2
    # print(vocab2)
    vocab3 = data[row][column+2]
    template_mapping["vocab3"] = vocab3
    # print(vocab3)
    vocab4 = data[row][column+3]
    template_mapping["vocab4"] = vocab4
    # print(vocab4)
    vocab5 = data[row+3][column]
    template_mapping["vocab5"] = vocab5
    # print(vocab5)
    vocab6 = data[row+3][column+1]
    template_mapping["vocab6"] = vocab6
    # print(vocab6)
    vocab7 = data[row+3][column+2]
    template_mapping["vocab7"] = vocab7
    # print(vocab7)
    vocab8 = data[row+3][column+3]
    template_mapping["vocab8"] = vocab8
    # print(vocab8)

    # Substitute
    template_word_filled = template_word_string.safe_substitute(template_mapping)
    template_image_filled = template_image_string.safe_substitute(template_mapping)

    # print(template_string)
    html_word = HTML(string=template_word_filled)
    html_image = HTML(string=template_image_filled)
    # css = CSS(string=css_string, font_config=font_config)
    # html.write_pdf('/tmp/test-py.pdf',
    #                stylesheets=[css],
    #                font_config=font_config)
    f_level = str(level)
    f_unit = str(unit).zfill(2)
    html_word.write_pdf('***REMOVED******REMOVED***AS***REMOVED******REMOVED***U***REMOVED******REMOVED***-word_flashcards.pdf'.format(output_path, f_level, f_unit))
    html_image.write_pdf('***REMOVED******REMOVED***AS***REMOVED******REMOVED***U***REMOVED******REMOVED***-image_flashcards.pdf'.format(output_path, f_level, f_unit))

    # Advance to the next unit
    row += 12
    # stop after one Unit
    break
