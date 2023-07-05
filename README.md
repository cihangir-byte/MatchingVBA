# Excel Data Matcher using VBA

This repository contains a project that uses VBA (Visual Basic for Applications) to automatically update specific tables in Excel. The project matches and updates data from several sheets to a 'Matching' sheet.

## Project Structure

This Excel VBA project has four key sheets:
1. `Positions` sheet: This sheet is updated weekly. The data from this sheet is used to update the 'Matching' sheet.
2. `Fleet list` sheet: This sheet is updated occasionally. The data from this sheet is also used to update the 'Matching' sheet.
3. `Fixtures` sheet: This sheet is updated weekly. The data from this sheet is also used to update the 'Matching' sheet.
4. `Matching` sheet: This is the main sheet that gets updated using the data from the other three sheets. 

The data in the 'Matching' sheet gets refreshed each time the macro is run.

## How to Use

To use the code, follow these steps:

1. Open the Excel file.
2. Press ALT + F11 to open the VBA editor.
3. Import the `.bas` file into the VBA editor.
4. Run the `UpdateMatchingSheet` macro.

## Note

Be sure to backup your data before running the macro, as it clears existing data in the 'Matching' sheet before refreshing it with new data.
