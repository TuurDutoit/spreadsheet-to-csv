# Spreadsheet to CSV
A Google Apps Script that exports a whole spreadsheet as a ZIP of CSVs

## Usage
  1) Open the side bar from the `Add-ons` menu > `Spreadsheet to CSV` > `Export whole spreadsheet as CSV`
  2) Fill in the [settings](#settings) as you wish (or leave the defaults)
  3) Click `Download` or `Save`

## Settings
### Separator
The separator to use between fields. By default a comma `,`; Excel uses a semicolon `;`.

### Include `sep=,` line
Whether to include a `sep=,` line. This can be used to change the separator Excel uses (from the standard semicolon to a comma, for example)  
Not all programs support this option.

### Filename
The name of the ZIP file.

### Download method
You can either download the file to your computer (no file will be placed in your Drive), or save the ZIP file to your Drive.
