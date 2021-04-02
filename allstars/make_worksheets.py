import gspread
from oauth2client.service_account import ServiceAccountCredentials
# from pprint import pprint
from weasyprint import HTML
from string import Template
import pathlib
import logging


def make_blank_string(length, max_length=27):
    # Create string of underscores whose length depends on the given length,
    #  but is not more than the maximum
    ratio = 0.6  # Level 1, max=16
    # ratio = 0.667 # Level 2, max=27
    if round(ratio * length) > max_length:
        return '_' * max_length
    else:
        return '_' * round(ratio * length)


def replace_blank(string):
    find_blank = '__'
    new_blank = '<span class="blank">__________</span>'
    if find_blank in string:
        string = string.replace(find_blank, new_blank)
    return string


# Get data from google Sheet (one level at a time)
def get_data_for_level(level: str) -> list[str]:
    # Fetch data from Google Sheet
    scope = ["https://spreadsheets.google.com/feeds",
             "https://www.googleapis.com/auth/spreadsheets",
             "https://www.googleapis.com/auth/drive.file",
             "https://www.googleapis.com/auth/drive"]
    try:
        creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)
        client = gspread.authorize(creds)
        sheet = client.open("all_stars_revised_0128").worksheet(f"Level {level}")
        return sheet.get_all_values()
    except FileNotFoundError as fnf_error:
        print(fnf_error)
        return []



# Create template mapping for one lesson at a time
def create_template_mapping(data: list, level: int, unit: int, lesson: int) -> dict[str, str]:
    # Set the row based on the unit and lesson
    row = 1 + ((unit - 1) * 12) + ((lesson - 1) * 3)
    column = 3

    # Create substitution mapping
    template_mapping = dict()
    template_mapping["template_path"] = pathlib.Path(__file__).parent.absolute()
    template_mapping["level"] = level
    # These are used for the page header
    template_mapping["unit"] = unit
    template_mapping["lesson"] = lesson
    # These are used for image naming
    template_mapping["unit_zfill"] = str(unit).zfill(2)
    template_mapping["lesson_zfill"] = str(lesson).zfill(2)

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
        template_mapping["blank1a_en"] = make_blank_string(
            len(template_mapping["sentence1a_en"]))

        template_mapping["sentence1b_en"] = data[row][column + 6].strip()
        template_mapping["blank1b_en"] = make_blank_string(
            len(template_mapping["sentence1b_en"]))

        template_mapping["sentence2a_en"] = data[row][column + 7].strip()
        template_mapping["blank2a_en"] = make_blank_string(
            len(template_mapping["sentence2a_en"]))

        template_mapping["sentence2b_en"] = data[row][column + 8].strip()
        template_mapping["blank2b_en"] = make_blank_string(
            len(template_mapping["sentence2b_en"]))

        template_mapping["sentence3a_en"] = data[row][column + 9].strip()
        template_mapping["blank3a_en"] = make_blank_string(
            len(template_mapping["sentence3a_en"]))

        template_mapping["sentence3b_en"] = data[row][column + 10].strip()
        template_mapping["blank3b_en"] = make_blank_string(
            len(template_mapping["sentence3b_en"]))

        template_mapping["sentence4a_en"] = data[row][column + 11].strip()
        template_mapping["blank4a_en"] = make_blank_string(
            len(template_mapping["sentence4a_en"]))

        template_mapping["sentence4b_en"] = data[row][column + 12].strip()
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
        # ***NO LONGER TRUE -> Level 5 has an extra column (situation) before the stories
        # Consider refactoring back into the main code
        template_mapping["story_en"] = data[row][column]
        template_mapping["story_jp"] = data[row + 1][column]

        # reading/writing sentences
        template_mapping["sentence1a_en"] = replace_blank(data[row][column + 1].strip())
        template_mapping["blank1a_en"] = make_blank_string(
            len(template_mapping["sentence1a_en"]))

        template_mapping["sentence1b_en"] = replace_blank(data[row][column + 2].strip())
        template_mapping["blank1b_en"] = make_blank_string(
            len(template_mapping["sentence1b_en"]))

        template_mapping["sentence1a_jp"] = data[row + 1][column + 1]
        template_mapping["sentence1b_jp"] = data[row + 1][column + 2]
    return template_mapping


# Open HTML template file and return contents as string
def get_template(filename: str) -> str:
    with open(filename, "r") as template_file:
        template_file_contents = template_file.read()
    return template_file_contents


# Substitute vars in template string
def fill_template(template: str, template_mapping: dict[str, str]) -> str:
    template_string = Template(template)
    return template_string.safe_substitute(template_mapping)


# Output PDF
def output_pdf(contents: str, filename: str):
    # Log WeasyPrint output
    logger = logging.getLogger('weasyprint')
    logger.addHandler(logging.FileHandler('/tmp/weasyprint.log'))
    # Create Weasyprint HTML object
    html = HTML(string=contents)
    # Output PDF via Weasyprint
    html.write_pdf(filename)


def main(levels: list, units: list, lessons: list):
    for level in levels:
        print(f'Level {level}')
        # output_path = f'/Users/cbunn/Documents/Employment/5 Star/Google Drive/All Stars Second Edition/All Stars Second Edition/Worksheets/Level {level}/'
        # output_path = f'/Users/cbunn/Documents/Employment/5 Star/Google Drive/All Stars Second Edition/unit1-output/Level {level}/'
        # output_path = f'/Users/cbunn/Documents/Employment/5 Star/Google Drive/All Stars Second Edition/test-output/Level {level}/'
        output_path = pathlib.Path(__file__).parent.parent.absolute() / f'output/Level {level}/'
        print(f'Output path: {output_path}')
        data = get_data_for_level(level)
        if not data:
            print("Cannot access Google Sheet data.")
            return
        # Loop through all Units and Lessons
        #  (range() needs a +1 because it stops at the number before)
        for unit in units:
            for lesson in lessons:
                print(f'Unit: {unit}, Lesson: {lesson}')
                # Create HTML template
                template_filename = f'AS{level}-worksheets-template.html'
                # Get contents of HTML template file
                template_file_contents = get_template(filename=template_filename)
                # create mapping dict
                template_mapping = create_template_mapping(data=data, level=level, unit=unit, lesson=lesson)
                # Substitute
                template_filled = fill_template(template=template_file_contents, template_mapping=template_mapping)
                # The numbers used in the filename need to be zero filled
                f_level = str(level)
                f_unit = str(unit).zfill(2)
                f_lesson = str(lesson).zfill(2)
                output_filename = f'{output_path}/AS{f_level}U{f_unit}L{f_lesson}.pdf'
                # Output PDF
                output_pdf(contents=template_filled, filename=output_filename)


if __name__ == "__main__":
    # So far, we're only doing Level 1, but in the future, we'll have to deal
    #  with the others
    # levels = [1, 2, 3]
    # levels = [1, 2, 3, 4, 5]
    levels = [5]
    # units = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16]
    units = [1]
    # units = [4, 6, 10, 14]
    lessons = [1, 2, 3, 4]
    # lessons = [2]
    main(levels, units, lessons)
