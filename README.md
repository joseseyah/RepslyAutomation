# Repsly Automation Project

This project automates the transformation of a rota into a structured table format in Google Sheets using Google Apps Script. It processes data from an existing sheet, extracts relevant information, and generates a new sheet with a standardized table.

## Features

- **Automated Table Generation**: Converts a rota from a specified sheet into a well-structured table.
- **Employee Mapping**: Converts employee names to representative IDs based on a mapping sheet.
- **Dynamic Data Handling**: Handles variable data ranges and merged cells.
- **Time and Duration Calculation**: Formats visit times and calculates visit durations in minutes.

## Getting Started

### Prerequisites

- Google Account
- Basic knowledge of Google Sheets and Google Apps Script

### Setup

1. **Open Google Sheets**: Create or open an existing Google Sheet where you want to use this script.
2. **Access Apps Script**: Go to `Extensions > Apps Script`.
3. **Create a New Script**: Delete any existing code in the script editor and paste the provided script.

### Script Details

#### Main Function: `generateTable()`

This function performs the following operations:

1. **Retrieve Spreadsheet and Sheets**:
   - Access the active spreadsheet.
   - Get the original sheet named "Sheet1".

2. **Create New Sheet**:
   - Check if a sheet named "Generated Table" exists and delete it if present.
   - Create a new sheet named "Generated Table".
   - Set the header row for the new sheet.

3. **Extract Employee Names**:
   - Call `getEmployeeNames()` to retrieve a list of employee names from the specified sheet.

4. **Process Original Sheet Data**:
   - Loop through the data in the original sheet.
   - Extract relevant information (e.g., representative ID, visit time, duration).
   - Add rows to the new sheet with the extracted data.

5. **Update Schedule**:
   - Call `updateSchedule()` to map employee names to representative IDs in the new table.

#### Helper Functions

1. **`getEmployeeNames()`**:
   - Extracts employee names from a specified sheet.
   - Filters out invalid or header entries.

2. **`updateSchedule()`**:
   - Maps employee names to representative IDs using data from an "EmployeeMapping" sheet.
   - Updates the new table with the mapped IDs.

### How to Run

1. **Run the Script**:
   - In the Apps Script editor, click the play button (`▶️`) to run the `generateTable()` function.

2. **Check the Output**:
   - Open the "Generated Table" sheet to see the transformed rota data.

### Customization

- **Sheet Names**: Modify the sheet names in the script (`Sheet1`, `Generated Table`, `EmployeeMapping`, etc.) to match your actual sheet names.
- **Data Ranges**: Adjust the data ranges and columns in the script to fit your specific data layout.
- **Task Description**: Customize the task description and other static values as needed.

## Notes

- This script assumes a specific layout of the original rota and employee mapping sheets. Ensure your sheets follow the required structure.
- Modify the `holidays()` function to implement custom holiday handling logic if needed.

## Troubleshooting

- **Permissions**: Ensure the script has the necessary permissions to access and modify your Google Sheets.
- **Data Integrity**: Verify that your original data is correctly formatted and contains no unexpected values.

## Contributing

If you encounter any issues or have suggestions for improvements, feel free to open an issue or submit a pull request.

## License

This project is licensed under the MIT License. See the LICENSE file for more details.

---

By using this script, you can automate the tedious process of transforming rota data into a structured table, saving time and reducing errors. Enjoy seamless data management with Google Sheets and Apps Script!
