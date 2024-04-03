#!/usr/bin/env python3
import argparse
import pandas as pd
from data_process import process_excel

def main() -> int:
    """
    Main function for the script.
    Processes an Excel file and writes the processed data to an output file.
    """
    try:
        # Define the command line arguments
        parser: argparse.ArgumentParser  = argparse.ArgumentParser(description = "Read an Excel file")
        parser.add_argument("file_name", type = str, help = "Name of the Excel file")
        parser.add_argument("file_format", choices = ['A', 'B'], help = "Format of Excel file (A or B)")
        parser.add_argument("-o", "--output", type = str, help = "Name of the output file (optional)")
        args: argparse.Namespace = parser.parse_args()

        # Process the Excel spreadsheet
        processed_data: pd.DataFrame = process_excel(args.file_name, args.file_format)

        # Write the processed data to an output file
        output_file: str = args.output if args.output else "output.xlsx"
        processed_data.to_excel(output_file, index = False)
        print(f"DataFrame has been written to {output_file}")
        return 0
    except ValueError as e:
        print(f"Value error: {e}")
        return 1
    except FileNotFoundError as e:
        print(f"File not found: {e}")
        return 2
    except Exception as e:
        print(f"An error occurred: {e}")
        return 3
    except:
        print("An unknown error occurred")
        return -1

if __name__ == '__main__':
    exit_code = main()
    exit(exit_code)
