# Barcode Generator

This application generates barcodes from data stored in an Excel file. It processes all sheets in the Excel file and saves the generated barcodes in a structured directory. The application supports user-specified input for the file path and the column containing the barcode data.

## Features
- Supports multiple sheets in an Excel file.
- Generates Code128 barcodes.
- Saves barcodes as images in organized directories.
- Skips invalid barcode values and reports errors.

## Requirements
The required Python libraries are listed in the `requirements.txt` file. To install them, use:
```bash
pip install -r requirements.txt
```

## Installation
1. Clone this repository or download the script.
2. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage
Run the script from the command line with the following syntax:
```bash
python generate_barcodes.py <file_path> <column_name>
```

### Arguments
- `<file_path>`: The path to the Excel file containing barcode data.
- `<column_name>`: The name of the column in the Excel file that contains the barcode values.

### Example
```bash
python generate_barcodes.py ./warehouse_data.xlsx "BARCODE COLUMN"
```
This will:
1. Read the file `warehouse_data.xlsx`.
2. Look for the column named "BARCODE COLUMN".
3. Generate Code128 barcodes for valid values.
4. Save the barcodes in a `barcodes` directory at the same location as the Excel file.

## Output
- A `barcodes` directory will be created at the same location as the Excel file.
- Subdirectories will be created for each sheet in the Excel file.
- Each barcode will be saved as an image file named after the barcode value.

## Error Handling
- If the specified Excel file does not exist, an error message will be displayed.
- If the specified column does not exist in a sheet, a warning will be displayed for that sheet.
- Invalid barcode values will be skipped, and an error message will be shown in the terminal.
