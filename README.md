# All Stars Curriculum Materials Generation

![Google Sheets + HTML and CSS + Python = PDF](/GSplusHTMLCSS.png)

These scripts make it possible to automatically generate a full set of curriculum materials, including worksheets and flashcards. The templates are made using HTML and CSS for fine control of layout and typography. The lesson content, such as vocabulary words and target sentences, is kept in a Google Sheets spreadsheet for best visibility to the team and so any team member can easily make changes to the content. The Weasyprint package for Python is used to convert the HTML and CSS to PDF files which are then distributed to the team and printed for use in the classroom.

This system makes it easy to create and maintain a large and complex set of materials. For example, in All Stars, there are five distinct levels, each with 16 units of four lessons each. Any member of the team can change the lesson content. And if the appearance or layout needs to be changed, only the template needs to be changed and the scripts will output all the materials in seconds. Doing this by hand would require editing up to 320 separate documents.

## Example Worksheet Output

![All Stars worksheets](/allstars-worksheets-example.png)

## Example Flashcard Output

![All Stars flashcards](/allstars-flashcards-example.png)
