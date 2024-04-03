# Client Data Transformer

## Overview:
This program intakes a Excel spreadsheet (`.xlsx`) and outputs a new Excel spreadsheet, with the format as requested in `Review Template.xlsx`.
Requires Python 3.10 or higher, as well as the `pandas` library.

## Installation:
Run `pip install -r requirements.txt` to ensure that the necessary dependencies are installed.

## Usage:
Run the script `main.py` with the following command line arguments:
`python3 main.py <file_name> <file_format> [-o <output_file>]`
* `<file_name>`: The name of the input Excel file (`.xlsx`) to be processed.
* `<file_format`>: Specify the format of the input Excel file, either `A` or `B`
* `-o <output_file>` (optional): Specify the name of the output Excel spreadsheet. If not provided, the default ouput file will be named "`output.xlsx`".

## Examples:
`python main.py input_file.xlsx A -o output_file.xlsx`