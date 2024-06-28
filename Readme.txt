Project Overview

This project is developed by Gallup Pakistan Digital Analytics to extract trade statistics data published monthly by PBS (Pakistan Bureau of Statistics) in PDF format. The code processes the specified pages of the PDF, extracts tables, performs necessary transformations, and saves the output in an Excel file.

Prerequisites

Before running the script, ensure that you have the following installed:

Python 3.x
Java Development Kit (JDK)
Required Python packages:
pandas
tabula-py
PyPDF2
openpyxl



The script performs the following steps:

Environment Setup: Sets the JAVA_HOME environment variable if needed.

User Input: Prompts the user to enter the page numbers to process.

Table Extraction: Uses tabula.read_pdf to extract tables from the specified pages of the PDF.

Data Transformation:

Drops rows and columns with all NaN values.
Splits cells with multiline content.
Adjusts the number of columns to match the expected format.
Reorders columns and assigns appropriate headers.
Splits specific columns into multiple columns for detailed data.
Output: Saves the processed data to an Excel file with each page's data in a separate sheet.

Output

The final output is saved in an Excel file named sample output.xlsx in the specified output directory.