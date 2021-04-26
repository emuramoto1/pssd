Goal 

The goal of this project is to create a more seamless experience for Babson students to track their courses and credits. The end output is a filled out "Babson Degree Worksheet" based off of the user's past courses.

How to Use the Website:

First, a Babson student needs to obtain their Academic Record from Workday in an xlsx format. They can obtain this information from Academics >> Academic Record >> View My Academic Record >> and then click the little “X” with a spreadsheet icon in the top right corner. Then, the user submits this file in the "Registration" tab and the site returns a filled out PDF "Babson Degree Worksheet" file based off of the .xlsx file submitted.

How to run code on Python:

First, a Babson student needs to obtain their Academic Record from Workday in an xlsx format. They can obtain this information from Academics >> Academic Record >> View My Academic Record >> and then click the little “X” with a spreadsheet icon in the top right corner. They can then drop the file into the main directory and run the function write(<FILE_NAME>) in backend.py to automatically fill out the Babson Degree Worksheet. The final outputted worksheet will be in the “/static/” folder. 

For testing purposes, an Academic Record of a current Babson Student has been provided labeled as “test1.xlsx”. Users can run the code using this excel file.

Process: 
The Write(sheet) function will first produce an output_template.pdf file. This file has all courses under the “discover” section filled out. 


Required dependencies:
python -m pip install pdfrw
python -m pip install xlrd == 1.2.0
python -m pip install flask 

