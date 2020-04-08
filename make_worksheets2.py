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

# Create HTML template
font_config = FontConfiguration()
template_filename = "AS2-worksheets-template.html"
template_file = open(template_filename, "r")
template_file_contents = template_file.read()
template_string = Template(template_file_contents)
css_filename = "worksheets.css"
css_file = open(css_filename, "r")
css_string = css_file.read()

output_path = "/Users/cbunn/Documents/Employment/5 Star/Google Drive/All Stars Second Edition/test-output2/"

# Fetch data from Google Sheet
scope = ["https://spreadsheets.google.com/feeds",
         "https://www.googleapis.com/auth/spreadsheets",
         "https://www.googleapis.com/auth/drive.file",
         "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)
client = gspread.authorize(creds)
sheet = client.open("all_stars_revised_0128").get_worksheet(1)
data = sheet.get_all_values()

# So far, we're only doing Level 1, but in the future, we'll have to deal
#  with the others
level = 2
units = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16]
lessons = [1, 2, 3, 4]

# Set the starting point of the gspread output
row = 1
column = 3

# Loop through all Units and Lessons
#  (range() needs a +1 because it stops at the number before)
for unit in units:
    for lesson in lessons:
        # Create substitution mapping
        template_mapping = dict()
        template_mapping["level"] = level
        # These are used for the page header
        template_mapping["unit"] = unit
        template_mapping["lesson"] = lesson
        # These are used for image naming
        template_mapping["unit_zfill"] = str(unit).zfill(2)
        template_mapping["lesson_zfill"] = str(lesson).zfill(2)

        print("Unit: " + str(unit) + ", Lesson: " + str(lesson))
        story_en = data[row][column]
        template_mapping["story_en"] = story_en
        story_jp = data[row+1][column]
        template_mapping["story_jp"] = story_jp

        # reading and writing sentences
        # - all four sets are used in Level 2, unlike Level 1
        sentence1a_en = data[row][column+5]
        template_mapping["sentence1a_en"] = sentence1a_en
        sentence1b_en = data[row][column+6]
        template_mapping["sentence1b_en"] = sentence1b_en
        sentence2a_en = data[row][column+7]
        template_mapping["sentence2a_en"] = sentence2a_en
        sentence2b_en = data[row][column+8]
        template_mapping["sentence2b_en"] = sentence2b_en
        sentence3a_en = data[row][column+9]
        template_mapping["sentence3a_en"] = sentence3a_en
        sentence3b_en = data[row][column+10]
        template_mapping["sentence3b_en"] = sentence3b_en
        sentence4a_en = data[row][column+11]
        template_mapping["sentence4a_en"] = sentence4a_en
        sentence4b_en = data[row][column+12]
        template_mapping["sentence4b_en"] = sentence4b_en
        sentence1a_jp = data[row+1][column+5]
        template_mapping["sentence1a_jp"] = sentence1a_jp
        sentence1b_jp = data[row+1][column+6]
        template_mapping["sentence1b_jp"] = sentence1b_jp
        sentence2a_jp = data[row+1][column+7]
        template_mapping["sentence2a_jp"] = sentence2a_jp
        sentence2b_jp = data[row+1][column+8]
        template_mapping["sentence2b_jp"] = sentence2b_jp
        sentence3a_jp = data[row+1][column+9]
        template_mapping["sentence3a_jp"] = sentence3a_jp
        sentence3b_jp = data[row+1][column+10]
        template_mapping["sentence3b_jp"] = sentence3b_jp
        sentence4a_jp = data[row+1][column+11]
        template_mapping["sentence4a_jp"] = sentence4a_jp
        sentence4b_jp = data[row+1][column+12]
        template_mapping["sentence4b_jp"] = sentence4b_jp

        # writing sentences - *** Not used in Levels other than 1 ***
        wsentence1_en = data[row][column+13]
        template_mapping["wsentence1_en"] = wsentence1_en
        wsentence2_en = data[row][column+14]
        template_mapping["wsentence2_en"] = wsentence2_en
        wsentence1_jp = data[row+1][column+13]
        template_mapping["wsentence1_jp"] = wsentence1_jp
        wsentence2_jp = data[row+1][column+14]
        template_mapping["wsentence2_jp"] = wsentence2_jp

        # Substitute
        template_filled = template_string.safe_substitute(template_mapping)
        html = HTML(string=template_filled)
        # css = CSS(string=css_string, font_config=font_config)
        # html.write_pdf('/tmp/test-py.pdf',
        #                stylesheets=[css],
        #                font_config=font_config)
        # The numbers used in the filename need to be zero filled
        f_level = str(level)
        f_unit = str(unit).zfill(2)
        f_lesson = str(lesson).zfill(2)
        html.write_pdf('{}AS{}U{}L{}.pdf'.format(output_path, f_level, f_unit, f_lesson))

        # Advance to the next row of stories and vocab
        row += 3
        # stop after one Lesson
        # break
    # stop after one Unit
    # break
