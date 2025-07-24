# Excel Sheets Merger
A Python script that consolidates multiple sheets from an Excel workbook into a single sheet.

# Functionality
- Reads an Excel workbook containing multiple sheets of raw material data
- Excludes unnecessary sheets
- For each sheet:
  - Extracts column names and data from specific rows
  - Transposes the data for proper alignment
  - Cleans empty values and standardizes the data format
- Consolidates all processed sheets into a single DataFrame
- Removes unnecessary columns
- Exports the consolidated data to a new Excel file

Requirements
- Python 3.x
- pandas
Usage
Place the script in the same directory as your input Excel file and run it. The script will generate a new Excel file named 'Consolidated_data.xlsx' with the consolidated data.
