# Data-formatting-for-Optojump-raw-export
This R script is designed to process and format raw data exported from the Optojump Next device.

It performs several tasks:

Data Import and Cleaning: Reads an Excel file containing raw data, removes empty columns, and handles specific data adjustments such as renaming columns and removing unwanted rows.

Data Transformation: Shifts columns and renames them to match the desired output format. It also converts the data from a long format to a wide format, restructuring it based on time and trigger values.

Data Enrichment: Adds new columns for calculated values like average step length, frequency, contact time, and flight time, and rounds these values to a specified precision.

Data Export: Finally, it exports the cleaned and formatted data to a new Excel file, ready for further analysis or reporting.
