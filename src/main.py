#!/usr/bin/env python3
import argparse
import pandas as pd
import sys
from data_process import process_excel

def main() -> int:
    """
    Main function for the script.
    Processes an Excel file and writes the processed data to an output file.
    """
    try:
        manual_input: bool = False
        if (len(sys.argv) == 1):
            manual_input = True
            input_file: str = input("Enter the name of the Excel file: ")
            file_format: str = input("Enter the format of the Excel file (A or B): ")
            needs_output: str = input("Do you want to specify an output file? (y/n): ")
            output: bool = True if needs_output.lower() == 'y' else False
            if output:
                output_file: str = input("Enter the name of the output file: ")
        else:            
            if (len(sys.argv) == 2):
                raise ValueError("No file format specified")
            elif (len(sys.argv) == 4):
                raise ValueError("No output file specified")
            elif (len(sys.argv) > 5):
                raise ValueError("Too many arguments")
            else:
                # Define the command line arguments
                parser: argparse.ArgumentParser  = argparse.ArgumentParser(description = "Read an Excel file")
                parser.add_argument("file_name", type = str, help = "Name of the Excel file")
                parser.add_argument("file_format", choices = ['A', 'B'], help = "Format of Excel file (A or B)")
                parser.add_argument("-o", "--output", type = str, help = "Name of the output file (optional)")
                args: argparse.Namespace = parser.parse_args()

        # Process the Excel spreadsheet
        processed_data: pd.DataFrame = process_excel((input_file if manual_input else args.file_name), (file_format if manual_input else args.file_format))

        # Write the processed data to an output file
        if manual_input:
            output_file: str = output_file if output else "output.xlsx"
        else:
            output_file: str = args.output if args.output else "output.xlsx"
        processed_data.to_excel(output_file, index = False)
        print(f"DataFrame has been written to {output_file}")
        return 0
    except ValueError as e:
        print(f"Value error: {e}", file = sys.stderr)
        return 1
    except FileNotFoundError as e:
        print(f"File not found: {e}", file = sys.stderr)
        return 2
    except Exception as e:
        print(f"An error occurred: {e}", file = sys.stderr)
        return 3
    except:
        print("An unknown error occurred", file = sys.stderr)
        return -1

if __name__ == '__main__':
    sys.exit(main())
