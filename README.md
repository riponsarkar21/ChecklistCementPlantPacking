# ChecklistCementPlantPacking
Check List Form - Tkinter Interface with Excel Integration
This project provides a Tkinter-based user interface that loads questions from an Excel file and allows users to submit answers to these questions. The answers are then stored in a new Excel file, and if the file already exists, new rows are added for each submission. If the date and shift are the same, the existing row is updated.

Features:
Excel Integration:

Questions are loaded dynamically from an Excel file (questions.xlsx), with labels pulled from Column C of the "question" sheet.
User responses are saved in a new Excel file (submissions.xlsx). If the file exists, data is added in new rows.
Date and Shift Management:

A date picker and shift selector (with shifts A, B, C) are provided for each submission.
If the date and shift match an existing entry, the corresponding row is updated with the new responses. If they do not match, a new row is added.
Scrollable Form:

A scrollable checklist allows users to select multiple questions and enter visitor names for each.
Validation:

All fields (questions, date, and shift) must be completed before submission.
How it Works:
Load Questions:

Questions are read from an Excel file (questions.xlsx), and each question is displayed with a checkbox and a visitor name field.
Submit Answers:

After completing the form, users can submit the data.
The data is validated to ensure all fields are filled, then saved in an Excel file (submissions.xlsx).
Update or Create File:

If the submissions.xlsx file exists, new data is added in a new row or the existing row for the same date and shift is updated.
Date & Shift Logic:

If a row with the same date and shift exists, that row is overwritten with the latest submission.
If no matching row is found, a new row is created with the submitted data.
Setup Instructions:
Install the required libraries:

tkinter
openpyxl
tkcalendar
Ensure the Excel file questions.xlsx is in the files/ folder with the required questions and labels in Column C.

Run the script and the Tkinter interface will launch, allowing you to fill out the form and submit answers.

The results will be stored in submissions.xlsx in the files/ folder.
