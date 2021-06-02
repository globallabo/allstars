import gspread
from oauth2client.service_account import ServiceAccountCredentials
# from pprint import pprint
from weasyprint import HTML
from weasyprint.fonts import FontConfiguration
from string import Template
import pathlib
import logging


# Get data from google Sheet (one level at a time)
def get_data_for_level(level: str) -> list[str]:
    # Fetch data from Google Sheet
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive.file",
        "https://www.googleapis.com/auth/drive",
    ]
    try:
        creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)
        client = gspread.authorize(creds)
        sheet = client.open("all_stars_revised_0128").worksheet(f"Level {level}")
        return sheet.get_all_values()
    except FileNotFoundError as fnf_error:
        print(fnf_error)
        return []


# Create template mapping for one lesson at a time
def create_template_mapping(data: list, level: int, unit: int) -> dict[str, str]:
    # Set the row based on the unit and lesson
    row = 1 + ((unit - 1) * 12)
    column = 4

    # Create substitution mapping
    template_mapping = dict()
    template_mapping["template_path"] = pathlib.Path(__file__).parent.absolute()
    template_mapping["level"] = level
    # This is used for the page header:
    template_mapping["unit"] = unit
    # This is used for image naming:
    template_mapping["unit_zfill"] = str(unit).zfill(2)

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
    logger = logging.getLogger("weasyprint")
    logger.addHandler(logging.FileHandler("/tmp/weasyprint.log"))
    # Create Weasyprint HTML object
    html = HTML(string=contents)
    # Output PDF via Weasyprint
    html.write_pdf(filename)


def main(levels: list, units: list):
    for level in levels:
        print(f"Level {level}")

        output_path = (
            pathlib.Path(__file__).parent.parent.absolute() / f"output/flashcards/Level {level}/"
        )
        pathlib.Path(output_path).mkdir(parents=True, exist_ok=True)
        print(f"Output path: {output_path}")

        data = get_data_for_level(level)
        if not data:
            print("Cannot access Google Sheet data.")
            return        

        # Loop through all Units and Lessons
        for unit in units:
            print(f'Unit: {unit}')

            # Create HTML templates
            words_template_filename = "flashcards-words-template.html"
            images_template_filename = "flashcards-images-template.html"
            # Get contents of HTML template files
            words_template_file_contents = get_template(filename=words_template_filename)
            images_template_file_contents = get_template(filename=images_template_filename)
            # create mapping dicts
            words_template_mapping = create_template_mapping(
                data=data, level=level, unit=unit
            )
            images_template_mapping = create_template_mapping(
                data=data, level=level, unit=unit
            )
            # Substitute
            words_template_filled = fill_template(
                template=words_template_file_contents, template_mapping=words_template_mapping
            )
            images_template_filled = fill_template(
                template=images_template_file_contents, template_mapping=images_template_mapping
            )

            f_level = str(level)
            f_unit = str(unit).zfill(2)
            words_output_filename = f"{output_path}/AS{f_level}U{f_unit}-word_flashcards.pdf"
            # Output PDF
            output_pdf(contents=words_template_filled, filename=words_output_filename)
            # There are only image-based flashcards for Level 1
            if level == 1:
                images_output_filename = f"{output_path}/AS{f_level}U{f_unit}-image_flashcards.pdf"
                # Output PDF
                output_pdf(contents=images_template_filled, filename=images_output_filename)


if __name__ == "__main__":
    levels = [1, 2, 3, 4]
    units = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16]
    main(levels, units)