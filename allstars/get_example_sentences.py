import gspread
from oauth2client.service_account import ServiceAccountCredentials

# So far, we're only doing Level 1, but in the future, we'll have to deal
#  with the others
levels = [1, 2, 3]
units = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16]
lessons = [1, 2, 3, 4]

for level in levels:
    # Fetch data from Google Sheet
    scope = ["https://spreadsheets.google.com/feeds",
             "https://www.googleapis.com/auth/spreadsheets",
             "https://www.googleapis.com/auth/drive.file",
             "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)
    client = gspread.authorize(creds)
    # sheet = client.open("all_stars_revised_0128").sheet1
    sheet = client.open("all_stars_revised_0128").get_worksheet(level-1)
    data = sheet.get_all_values()

    # Set the starting point of the gspread output
    row = 1
    column = 8

    print('Level {}'.format(level))

    # Loop through all Units and Lessons
    #  (range() needs a +1 because it stops at the number before)
    for unit in units:
        sentence_a = data[row][column]
        sentence_b = data[row][column+1]

        # Print example A/B
        print('{}. {} / {}'.format(unit, sentence_a, sentence_b))
        # Advance to the next row of stories and vocab
        row += 12
