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
template_filename = "AS1-worksheets-template.html"
template_file = open(template_filename, "r")
template_file_contents = template_file.read()
template_string = Template(template_file_contents)
css_filename = "worksheets.css"
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
numLessons = 4

# Set some HTML strings for the yes/no overlays
image_overlay_yes = "<img class=\"yes-no\" src=\"images/yes.png\" alt=\"\">"
image_overlay_no = "<img class=\"yes-no\" src=\"images/no.png\" alt=\"\">"
image_overlay_none = ""

# Set the starting point of the gspread output
row = 1
column = 3

# Loop through all Units and Lessons
#  (range() needs a +1 because it stops at the number before)
for unit in range(1, numUnits+1):
    for lesson in range(1, numLessons+1):
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
        # story_en = sheet.cell(row,column).value
        story_en = data[row][column]
        # print("English story: " + story_en)
        template_mapping["story_en"] = story_en
        # story_jp = sheet.cell(row+1,column).value
        story_jp = data[row+1][column]
        # print("Japanese story: " + story_jp)
        template_mapping["story_jp"] = story_jp

        # set list vars, then loop through vocab
        vocab_en = []
        vocab_jp = []
        if lesson == 1:
            vocab_nums = [1, 2, 3, 4]
            template_mapping["image_overlay"] = image_overlay_none
        elif lesson == 2:
            vocab_nums = [5, 6, 7, 8]
            template_mapping["image_overlay"] = image_overlay_none
        elif lesson == 3:
            vocab_nums = [1, 2, 3, 4]
            template_mapping["image_overlay"] = image_overlay_yes
        elif lesson == 4:
            vocab_nums = [5, 6, 7, 8]
            template_mapping["image_overlay"] = image_overlay_no
        print("vocab_nums: ")
        print(vocab_nums)
        print("Image overlay: ")
        print(template_mapping["image_overlay"])
        for k in range(1, 5):
            vocab_en.append(data[row][column+k])
            # print("English vocab" + str(k) + ": " + vocab_en[k-1])
            # -1 because the range is 1-5, while the list is 0-4
            template_mapping["vocab" + str(k) + "_en"] = vocab_en[k-1]
            vocab_jp.append(data[row+1][column+k])
            # print("Japanese vocab" + str(k) + ": " + vocab_jp[k-1])
            # -1 because the range is 1-5, while the list is 0-4
            template_mapping["vocab" + str(k) + "_jp"] = vocab_jp[k-1]
            # Map the vocab image numbers for pages 1 and 2
            template_mapping["vocab_img" + str(k)] = vocab_nums[k-1]

        # reading sentences
        rsentence1a_en = data[row][column+5]
        # print("English Reading Sentence 1A: " + rsentence1a_en)
        template_mapping["rsentence1a_en"] = rsentence1a_en
        rsentence1b_en = data[row][column+6]
        # print("English Reading Sentence 1B: " + rsentence1b_en)
        template_mapping["rsentence1b_en"] = rsentence1b_en
        rsentence2a_en = data[row][column+7]
        # print("English Reading Sentence 2A: " + rsentence2a_en)
        template_mapping["rsentence2a_en"] = rsentence2a_en
        rsentence2b_en = data[row][column+8]
        # print("English Reading Sentence 2B: " + rsentence2b_en)
        template_mapping["rsentence2b_en"] = rsentence2b_en
        rsentence3a_en = data[row][column+9]
        # print("English Reading Sentence 3A: " + rsentence3a_en)
        template_mapping["rsentence3a_en"] = rsentence3a_en
        rsentence3b_en = data[row][column+10]
        # print("English Reading Sentence 3B: " + rsentence3b_en)
        template_mapping["rsentence3b_en"] = rsentence3b_en
        rsentence4a_en = data[row][column+11]
        # print("English Reading Sentence 4A: " + rsentence4a_en)
        template_mapping["rsentence4a_en"] = rsentence4a_en
        rsentence4b_en = data[row][column+12]
        # print("English Reading Sentence 4B: " + rsentence4b_en)
        template_mapping["rsentence4b_en"] = rsentence4b_en
        rsentence1a_jp = data[row+1][column+5]
        # print("Japanese Reading Sentence 1A: " + rsentence1a_jp)
        template_mapping["rsentence1a_jp"] = rsentence1a_jp
        rsentence1b_jp = data[row+1][column+6]
        # print("Japanese Reading Sentence 1B: " + rsentence1b_jp)
        template_mapping["rsentence1b_jp"] = rsentence1b_jp
        rsentence2a_jp = data[row+1][column+7]
        # print("Japanese Reading Sentence 2A: " + rsentence2a_jp)
        template_mapping["rsentence2a_jp"] = rsentence2a_jp
        rsentence2b_jp = data[row+1][column+8]
        # print("Japanese Reading Sentence 2B: " + rsentence2b_jp)
        template_mapping["rsentence2b_jp"] = rsentence2b_jp
        rsentence3a_jp = data[row+1][column+9]
        # print("Japanese Reading Sentence 3A: " + rsentence3a_jp)
        template_mapping["rsentence3a_jp"] = rsentence3a_jp
        rsentence3b_jp = data[row+1][column+10]
        # print("Japanese Reading Sentence 3B: " + rsentence3b_jp)
        template_mapping["rsentence3b_jp"] = rsentence3b_jp
        rsentence4a_jp = data[row+1][column+11]
        # print("Japanese Reading Sentence 4A: " + rsentence4a_jp)
        template_mapping["rsentence4a_jp"] = rsentence4a_jp
        rsentence4b_jp = data[row+1][column+12]
        # print("Japanese Reading Sentence 4B: " + rsentence4b_jp)
        template_mapping["rsentence4b_jp"] = rsentence4b_jp

        # writing sentences
        wsentence1_en = data[row][column+13]
        # print("English Writing Sentence 1: " + wsentence1_en)
        template_mapping["wsentence1_en"] = wsentence1_en
        wsentence2_en = data[row][column+14]
        # print("English Writing Sentence 2: " + wsentence2_en)
        template_mapping["wsentence2_en"] = wsentence2_en
        wsentence1_jp = data[row+1][column+13]
        # print("Japanese Writing Sentence 1: " + wsentence1_jp)
        template_mapping["wsentence1_jp"] = wsentence1_jp
        wsentence2_jp = data[row+1][column+14]
        # print("Japanese Writing Sentence 2: " + wsentence2_jp)
        template_mapping["wsentence2_jp"] = wsentence2_jp
        print("")

        # Substitute
        template_filled = template_string.safe_substitute(template_mapping)

        # print(template_string)
        html = HTML(string=template_filled)
        # css = CSS(string=css_string, font_config=font_config)
        # html.write_pdf('/tmp/test-py.pdf',
        #                stylesheets=[css],
        #                font_config=font_config)
        f_level = str(level)
        f_unit = str(unit).zfill(2)
        f_lesson = str(lesson).zfill(2)
        html.write_pdf('***REMOVED******REMOVED***AS***REMOVED******REMOVED***U***REMOVED******REMOVED***L***REMOVED******REMOVED***.pdf'.format(output_path, f_level, f_unit, f_lesson))

        # Advance to the next row of stories and vocab
        row += 3
        # stop after one Lesson
        # break
    # stop after one Unit
    break