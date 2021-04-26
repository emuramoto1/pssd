# README.md
### Goal:

The goal of this project is to create a more seamless experience for Babson students to track their courses and credits. The end output is a filled out "Babson Degree Worksheet" based off of the user's past courses.

### How to Use Website:

First, a Babson student needs to his or her their Academic Record from Workday in a .xlsx format. He or she can obtain this information from Academics >> Academic Record >> View My Academic Record >> and then click the little “X” with a spreadsheet icon in the top right corner. Then, the user submits this file in the "Babson Degree Worksheet Auto Fill" tab and the site returns a filled out PDF "Babson Degree Worksheet" file based off of the .xlsx file submitted.

### How to Run Code on Python:

First, a Babson student needs to obtain his or her Academic Record from Workday in an .xlsx format. He or she can obtain this information from Academics >> Academic Record >> View My Academic Record >> and then click the little “X” with a spreadsheet icon in the top right corner. They can then drop the file into the main directory and set the "data" variable in the function "main()" equal to the file name. Finally, they can run backend.py to automatically fill out the Babson Degree Worksheet. The final outputted worksheet will be named "output_template3.pdf" and will be in the “/static/” folder. 

For testing purposes, an Academic Record of a current Babson Student has been provided labeled as “test1.xlsx”. Users can run the code using this excel file if they do not have access to their own.

### How to Run Code on Flask:
First, run the main.py file, which returns a clickable link in the terminal. Then, once again, the Babson student obtains his or her Academic Record from Workday in a .xlsx format. He or she can obtain this information from Academics >> Academic Record >> View My Academic Record >> and then click the little “X” with a spreadsheet icon in the top right corner. Then, the user submits this file in the "Babson Degree Worksheet Auto Fill" tab and the site returns a filled out PDF "Babson Degree Worksheet" file based off of the .xlsx file submitted.

### Required Dependencies:
* python -m pip install pdfrw
* python -m pip install xlrd==1.2.0
* python -m pip install flask 

### Credit:
For the Babson image across all pages except the about.html page, we served this from babson.edu and the link is embedded in our HTML code. For the Babson logo in the footer of all pages, we served this from ytimg.com and the link is embedded in our HTML code. For the images on the howto.html page, these are screenshots taken from Babson Workday.

The navigation bar was created using class materials from Babson Professor Shankar's Webtech course.
