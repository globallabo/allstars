import gspread
from oauth2client.service_account import ServiceAccountCredentials
# from pprint import pprint
from weasyprint import HTML, CSS
from weasyprint.fonts import FontConfiguration
from string import Template
import logging


def make_blank_string(length, max_length=27):
    # Create string of underscores whose length depends on the given length,
    #  but is not more than the maximum
    ratio = 0.6  # Level 1, max=16
    # ratio = 0.667 # Level 2, max=27
    if round(ratio * length) > max_length:
        print(max_length)
        return '_' * max_length
    else:
        print(round(ratio * length))
        return '_' * round(ratio * length)


def replace_blank(string):
    find_blank = '__'
    new_blank = '<span class="blank">__________</span>'
    if find_blank in string:
        string = string.replace(find_blank, new_blank)
    return string


# Log WeasyPrint output
logger = logging.getLogger('weasyprint')
logger.addHandler(logging.FileHandler('/tmp/weasyprint.log'))

# So far, we're only doing Level 1, but in the future, we'll have to deal
#  with the others
# levels = [1, 2, 3]
levels = [5]
# units = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16]
units = [1]
# units = [4, 6, 10, 14]
lessons = [1, 2, 3, 4]
# lessons = [2]

for level in levels:
    print(f'Level {level}')
    # Create HTML template
    font_config = FontConfiguration()
    template_filename = f'AS{level}-worksheets-template.html'
    with open(template_filename, "r") as template_file:
        template_file_contents = template_file.read()
    template_string = Template(template_file_contents)
    css_filename = "worksheets.css"
    with open(css_filename, "r") as css_file:
        css_string = css_file.read()

    # output_path = f'/Users/cbunn/Documents/Employment/5 Star/Google Drive/All Stars Second Edition/All Stars Second Edition/Worksheets/Level {level}/'
    output_path = f'/Users/cbunn/Documents/Employment/5 Star/Google Drive/All Stars Second Edition/test-output/Level {level}/'

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
    column = 3

    # Loop through all Units and Lessons
    #  (range() needs a +1 because it stops at the number before)
    for unit in units:
        for lesson in lessons:
            # Set the row based on the unit and lesson
            row = 1 + ((unit - 1) * 12) + ((lesson - 1) * 3)

            # Create substitution mapping
            template_mapping = dict()
            template_mapping["level"] = level
            # These are used for the page header
            template_mapping["unit"] = unit
            template_mapping["lesson"] = lesson
            # These are used for image naming
            template_mapping["unit_zfill"] = str(unit).zfill(2)
            template_mapping["lesson_zfill"] = str(lesson).zfill(2)

            print(f'Unit: {unit}, Lesson: {lesson}')

            if level != 5:
                template_mapping["story_en"] = data[row][column]
                template_mapping["story_jp"] = data[row + 1][column]

                # Level 1 is separate because of the vocab section including images
                if level == 1:
                    # set list vars
                    vocab_en = []
                    vocab_jp = []
                    # Set some HTML strings for the yes/no overlays
                    image_overlay_yes = "<img class=\"yes-no\" src=\"images/yes.png\" alt=\"\">"
                    image_overlay_no = "<img class=\"yes-no\" src=\"images/no.png\" alt=\"\">"
                    image_overlay_none = ""
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
                    for k in range(1, 5):
                        vocab_en.append(data[row][column + k])
                        # At least one lesson has fewer than four vocab words
                        if not vocab_en[k - 1]:
                            template_mapping["no_vocab4"] = "no_vocab4"
                        else:
                            # -1 because the range is 1-5, while the list is 0-4
                            template_mapping["vocab" + str(k) + "_en"] = vocab_en[k - 1]
                            vocab_jp.append(data[row + 1][column + k])
                            # -1 because the range is 1-5, while the list is 0-4
                            template_mapping["vocab" + str(k) + "_jp"] = vocab_jp[k - 1]
                            # Map the vocab image numbers for pages 1 and 2
                            template_mapping["vocab_img" + str(k)] = vocab_nums[k - 1]
                    # writing sentences
                    template_mapping["wsentence1_en"] = data[row][column + 13].strip()
                    template_mapping["blank1_en"] = make_blank_string(
                        len(template_mapping["wsentence1_en"]))
                    template_mapping["wsentence2_en"] = data[row][column + 14].strip()
                    template_mapping["blank2_en"] = make_blank_string(
                        len(template_mapping["wsentence2_en"]))
                    template_mapping["wsentence1_jp"] = data[row + 1][column + 13].strip()
                    template_mapping["wsentence2_jp"] = data[row + 1][column + 14].strip()

                # reading sentences
                template_mapping["sentence1a_en"] = data[row][column + 5].strip()
                print(template_mapping["sentence1a_en"])
                print(len(template_mapping["sentence1a_en"]))
                template_mapping["blank1a_en"] = make_blank_string(
                    len(template_mapping["sentence1a_en"]))

                template_mapping["sentence1b_en"] = data[row][column + 6].strip()
                print(template_mapping["sentence1b_en"])
                print(len(template_mapping["sentence1b_en"]))
                template_mapping["blank1b_en"] = make_blank_string(
                    len(template_mapping["sentence1b_en"]))

                template_mapping["sentence2a_en"] = data[row][column + 7].strip()
                print(template_mapping["sentence2a_en"])
                print(len(template_mapping["sentence2a_en"]))
                template_mapping["blank2a_en"] = make_blank_string(
                    len(template_mapping["sentence2a_en"]))

                template_mapping["sentence2b_en"] = data[row][column + 8].strip()
                print(template_mapping["sentence2b_en"])
                print(len(template_mapping["sentence2b_en"]))
                template_mapping["blank2b_en"] = make_blank_string(
                    len(template_mapping["sentence2b_en"]))

                template_mapping["sentence3a_en"] = data[row][column + 9].strip()
                print(template_mapping["sentence3a_en"])
                print(len(template_mapping["sentence3a_en"]))
                template_mapping["blank3a_en"] = make_blank_string(
                    len(template_mapping["sentence3a_en"]))

                template_mapping["sentence3b_en"] = data[row][column + 10].strip()
                print(template_mapping["sentence3b_en"])
                print(len(template_mapping["sentence3b_en"]))
                template_mapping["blank3b_en"] = make_blank_string(
                    len(template_mapping["sentence3b_en"]))

                template_mapping["sentence4a_en"] = data[row][column + 11].strip()
                print(template_mapping["sentence4a_en"])
                print(len(template_mapping["sentence4a_en"]))
                template_mapping["blank4a_en"] = make_blank_string(
                    len(template_mapping["sentence4a_en"]))

                template_mapping["sentence4b_en"] = data[row][column + 12].strip()
                print(template_mapping["sentence4b_en"])
                print(len(template_mapping["sentence4b_en"]))
                template_mapping["blank4b_en"] = make_blank_string(
                    len(template_mapping["sentence4b_en"]))

                template_mapping["sentence1a_jp"] = data[row + 1][column + 5]
                template_mapping["sentence1b_jp"] = data[row + 1][column + 6]
                template_mapping["sentence2a_jp"] = data[row + 1][column + 7]
                template_mapping["sentence2b_jp"] = data[row + 1][column + 8]
                template_mapping["sentence3a_jp"] = data[row + 1][column + 9]
                template_mapping["sentence3b_jp"] = data[row + 1][column + 10]
                template_mapping["sentence4a_jp"] = data[row + 1][column + 11]
                template_mapping["sentence4b_jp"] = data[row + 1][column + 12]

            else:
                # Level 5 has an extra column (situation) before the stories
                template_mapping["story_en"] = data[row][column + 1]
                template_mapping["story_jp"] = data[row + 1][column + 1]
                print(f'Story (EN): {template_mapping["story_en"]}')
                print(f'Story (JP): {template_mapping["story_jp"]}')

                # reading/writing sentences
                template_mapping["sentence1a_en"] = replace_blank(data[row][column + 2].strip())
                print(template_mapping["sentence1a_en"])
                print(len(template_mapping["sentence1a_en"]))
                template_mapping["blank1a_en"] = make_blank_string(
                    len(template_mapping["sentence1a_en"]))

                template_mapping["sentence1b_en"] = replace_blank(data[row][column + 3].strip())
                print(template_mapping["sentence1b_en"])
                print(len(template_mapping["sentence1b_en"]))
                template_mapping["blank1b_en"] = make_blank_string(
                    len(template_mapping["sentence1b_en"]))

                template_mapping["sentence1a_jp"] = data[row + 1][column + 2]
                template_mapping["sentence1b_jp"] = data[row + 1][column + 3]

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
            # with open(f'{output_path}AS{f_level}U{f_unit}L{f_lesson}.html', 'w') as htmlfile:
            #     htmlfile.write(template_filled)
            html.write_pdf(f'{output_path}AS{f_level}U{f_unit}L{f_lesson}.pdf')

            # Advance to the next row of stories and vocab
            # row += 3
            # stop after one Lesson
            # break
        # stop after one Unit
        # break