# RSD-Tool
The RSD Tool is a small python application designed to aid in assisting a gas chromatographer to do their job. The tool can do a few things:

  - Calculate relative standard deviation (RSD) values, given peak area inputs (up to 5 values)
  - Enter the RSD value and some additional data into an Excel file
  - Control some settings
  - Provide a simple graphical user interface (GUI) for the above tasks
  

Modules Used

Python Standard Library
-tkinter - Used to provide a GUI
-decimal - Used to provide accurate floating point arithmetic
-time - Used to enter formatted time into the Excel sheet 

Third Party Modules
-openpyxl - Used to connect to Microsoft Excel

Who is it for?

I made this for myself to use at work as a QC chemist at an industrial compressed gas company. It will make my life a little easier. This app could be useful for anyone who needs to calculate RSD values and wants to store that data locally in an Excel file without having to physically open that file and enter the values manually.

What is the goal of this project?

I plan to release updates to this project. I would like to add additional features including but not limited to:
  -Expanding the 'Settings' window to allow for more customization
  -Allowing for MySQL connection as an alternative to Excel
  -General aesthetic improvements 

How to Use

System Requirements
-Python 3
-Microsoft Excel

-Currently to change the Excel file or dictionary file, you must go into the source code and change the two variables at the top of the file.
-The script is saved with a .pyw extension to prevent the console from opening when you click on the script.
