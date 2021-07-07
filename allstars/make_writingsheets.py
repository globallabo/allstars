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
        return "_" * max_length
    else:
        return "_" * round(ratio * length)


def replace_blank(string):
    find_blank = "__"
    new_blank = '<span class="blank">__________</span>'
    if find_blank in string:
        string = string.replace(find_blank, new_blank)
    return string


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


def main(levels: list):
    for level in levels:
        print(f"Level {level}")
        # output_path = f'/Users/cbunn/Documents/Employment/5 Star/Google Drive/All Stars Second Edition/All Stars Second Edition/Worksheets/Level {level}/'
        # output_path = f'/Users/cbunn/Documents/Employment/5 Star/Google Drive/All Stars Second Edition/unit1-output/Level {level}/'
        # output_path = f'/Users/cbunn/Documents/Employment/5 Star/Google Drive/All Stars Second Edition/test-output/Level {level}/'
        output_path = (
            pathlib.Path(__file__).parent.parent.absolute() / f"output/Writing Sheets/"
        )
        pathlib.Path(output_path).mkdir(parents=True, exist_ok=True)
        print(f"Output path: {output_path}")

        template_filename = f"AS{level}-worksheets-template-blank.html"
        # Get contents of HTML template file
        template_file_contents = get_template(filename=template_filename)
        # create mapping dict
        # Create substitution mapping
        template_mapping = dict()
        template_mapping["template_path"] = pathlib.Path(__file__).parent.absolute()
        template_mapping["level"] = level
        # For level 1
        template_mapping["blank1_en"] = make_blank_string(23)
        # For all other levels
        template_mapping["blank1a_en"] = make_blank_string(40)
        # Substitute
        template_filled = fill_template(
            template=template_file_contents, template_mapping=template_mapping
        )
        output_filename = f"{output_path}/AS{level}-writing_sheet.pdf"
        # Output PDF
        output_pdf(contents=template_filled, filename=output_filename)


if __name__ == "__main__":
    # So far, we're only doing Level 1, but in the future, we'll have to deal
    #  with the others
    # levels = [1, 2, 3, 4, 5]
    levels = [1, 2, 3, 4, 5]
    main(levels)
