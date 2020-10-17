import gspread
from oauth2client.service_account import ServiceAccountCredentials
from weasyprint import HTML
from string import Template

# So far, we're only doing Level 1, but in the future, we'll have to deal
#  with the others
levels = [1, 2, 3, 4]
units = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16]
lessons = [1, 2, 3, 4]

for level in levels:
    # with open(f"level{level}.txt", "w") as text_file:
    #     print(f"All Stars - Level {level}", file=text_file)

    print(f'Level {level}')
    # Create HTML template
    template_filename = 'proofing-template-body.html'
    with open(template_filename, "r") as template_file:
        template_file_contents = template_file.read()
    template_string = Template(template_file_contents)

    output_path = '/Users/cbunn/Documents/Employment/5 Star/Google Drive/All Stars Second Edition/All Stars Second Edition/Proofing Documents/'

    # Fetch data from Google Sheet
    scope = ["https://spreadsheets.google.com/feeds",
             "https://www.googleapis.com/auth/spreadsheets",
             "https://www.googleapis.com/auth/drive.file",
             "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)
    client = gspread.authorize(creds)
    # sheet = client.open("all_stars_revised_0128").sheet1
    sheet = client.open("all_stars_revised_0128").worksheet(f"Level {level}")
    data = sheet.get_all_values()

    template_mapping = dict()
    template_mapping["level"] = level

    # initialize page counting
    page_count = 64  # 16 units, 4 lessons per unit, 1 page per lesson
    page_number = 1

    # initialize string for the content
    content = ""

    # Loop through all Units and Lessons
    for unit in units:
        for lesson in lessons:
            # Set the row based on the unit and lesson
            row = 1 + ((unit - 1) * 12) + ((lesson - 1) * 3)
            column = 3

            # Create substitution mapping for each page
            page_template_mapping = dict()
            page_template_mapping["level"] = level
            page_template_mapping["unit"] = unit
            page_template_mapping["lesson"] = lesson
            # for the image file URL
            page_template_mapping["unit_zfill"] = str(unit).zfill(2)
            page_template_mapping["lesson_zfill"] = str(lesson).zfill(2)

            page_template_mapping["story_en"] = data[row][column]
            page_template_mapping["story_jp"] = data[row + 1][column]
            # set list vars
            vocab_en = []
            vocab_jp = []
            for k in range(1, 5):
                vocab_en.append(data[row][column + k])
                vocab_jp.append(data[row + 1][column + k])
                # -1 because the range is 1-5, while the list is 0-4
                page_template_mapping["vocab" + str(k) + "_en"] = vocab_en[k - 1]
                page_template_mapping["vocab" + str(k) + "_jp"] = vocab_jp[k - 1]
            page_template_mapping["sentence1a_en"] = data[row][column + 5]
            page_template_mapping["sentence1b_en"] = data[row][column + 6]
            page_template_mapping["sentence2a_en"] = data[row][column + 7]
            page_template_mapping["sentence2b_en"] = data[row][column + 8]
            page_template_mapping["sentence3a_en"] = data[row][column + 9]
            page_template_mapping["sentence3b_en"] = data[row][column + 10]
            page_template_mapping["sentence4a_en"] = data[row][column + 11]
            page_template_mapping["sentence4b_en"] = data[row][column + 12]
            page_template_mapping["sentence1a_jp"] = data[row + 1][column + 5]
            page_template_mapping["sentence1b_jp"] = data[row + 1][column + 6]
            page_template_mapping["sentence2a_jp"] = data[row + 1][column + 7]
            page_template_mapping["sentence2b_jp"] = data[row + 1][column + 8]
            page_template_mapping["sentence3a_jp"] = data[row + 1][column + 9]
            page_template_mapping["sentence3b_jp"] = data[row + 1][column + 10]
            page_template_mapping["sentence4a_jp"] = data[row + 1][column + 11]
            page_template_mapping["sentence4b_jp"] = data[row + 1][column + 12]

            page_template_mapping["page_number"] = page_number
            page_template_mapping["page_count"] = page_count

            page_template_filename = 'proofing-template-page.html'
            with open(page_template_filename, "r") as page_template_file:
                page_template_file_contents = page_template_file.read()
            page_template_string = Template(page_template_file_contents)
            page_template_filled = page_template_string.safe_substitute(page_template_mapping)
            content += page_template_filled

            # increment page number
            page_number += 1

    # fill in template using content from loop above
    template_mapping["content"] = content
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
    html.write_pdf(f'{output_path}AS{f_level}-proofing.pdf')
