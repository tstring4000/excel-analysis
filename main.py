# Filename: main.py
# Author: Tyler Stringer
# Date: 2024-09-03
# Details: Main Python file to analyze an Excel workbook. It finds formula interdependencies between different sheets.

# Local Python files:
import read_sheets
import read_formulas
import graph

# Modules:
import csv
import matplotlib.pyplot as plt
import time
import os


def get_user_input():
    # Ask the user if they want to proceed with the default settings
    user_choice = input("Do you want to proceed with the default settings ('y' = whatever is hard-coded in the "
                        "main.py file, or 'n' if you want to enter your own)? (y/n): ").strip().lower()

    if user_choice == 'n':
        # Prompt user for custom inputs
        input_file = input("(1/4) Enter the path to your input file (e.g., 'data/input/FILENAME.xlsx'): ").strip()
        num_rows = int(input("(2/4) Enter the number of rows (i.e., how many rows down would you expect "
                             "formulas?): ").strip())
        num_cols = int(input("(3/4) Enter the number of columns (702 columns = going out to ZZ): ").strip())
        layout_choice = input("(4/4) Enter your graph layout choice ('circular', 'spring', 'shell', 'planar', "
                              "or 'kamada_kawai' ... 'circular' is recommended): ").strip()

        # Generate the output filename
        input_filename = os.path.basename(input_file)
        file_prefix = input_filename[:4]
        formulas_csv = f"data/output/{file_prefix}_formulas.csv"
        dependencies_csv = f"data/output/{file_prefix}_dependencies.csv"

        # Return the custom inputs
        return input_file, num_rows, num_cols, layout_choice, formulas_csv, dependencies_csv
    else:
        # Use default settings
        input_file = 'data/input/0052_vr_int_dash.xlsx'
        num_rows = 15
        num_cols = 50  # 26 columns = through column Z, 73 = through column CC, 702 columns = through column ZZ
        layout_choice = 'circular'
        formulas_csv = 'data/output/formulas.csv'
        dependencies_csv = 'data/output/dependencies.csv'

        # Return the default settings
        return input_file, num_rows, num_cols, layout_choice, formulas_csv, dependencies_csv


def save_dependencies_to_csv(dependencies, filename):
    # Saves the dependencies dictionary to a CSV file using utf-8 encoding.
    with open(filename, mode='w', newline='', encoding='utf-8', errors='replace') as file:
        writer = csv.writer(file)
        for sheet, deps in dependencies.items():
            writer.writerow([sheet] + deps)


def main():
    # Get user input or default settings
    path, num_rows, num_cols, layout_choice, formulas_csv, dependencies_csv = get_user_input()

    # Pass the variable "path" to the function in read_sheets.py
    sheet_names = read_sheets.read_sheet_names(path)
    print('The Excel file in', path, 'contains sheet names', sheet_names)
    print('Access that variable by using "sheet_names"')

    # Prepare to write to CSV
    with open(formulas_csv, mode='w', newline='', encoding='utf-8', errors='replace') as file:
        writer = csv.writer(file)

        # Write the header row
        header = ['sheet_name'] + [f'formula{str(i).zfill(3)}' for i in range(1, num_cols)]
        writer.writerow(header)

        # Initialize the dependencies dictionary
        dependencies = {sheet: [] for sheet in sheet_names}
        formula_dependencies = {}

        # Loop through all sheet names and get formulas from the first few rows
        for sheet_name in sheet_names:
            print(f"Processing sheet: {sheet_name} ...")  # Status message
            formulas = read_formulas.get_formulas(path, sheet_name, num_rows)
            row = [sheet_name] + list(formulas)
            writer.writerow(row)

            # Analyze formulas to find dependencies
            for formula in formulas:
                formula_node = f'"{sheet_name} - {formula}"'
                formula_dependencies[formula_node] = []
                for other_sheet in sheet_names:
                    if other_sheet in formula and other_sheet != sheet_name:
                        dependencies[sheet_name].append(other_sheet)
                        formula_dependencies[formula_node].append(other_sheet)

    # Save the dependencies dictionary to a CSV file
    save_dependencies_to_csv(dependencies, dependencies_csv)
    print(f"Dependencies have been written to {dependencies_csv}")

    print(f"Formulas have been written to {formulas_csv}")
    print("Main.py successfully ran.")

    # Generate the dependency graph
    graph.generate_graph(dependencies, layout=layout_choice)

    # Generate the formula dependency graph
    # graph.generate_formula_graph(formula_dependencies)  # Uncomment this line to make formula graph.

    # Keep all figures open
    #    plt.show()

    # Time how long it takes main.py to run
    for i in range(1000000):
        pass  # Simulating some work


if __name__ == '__main__':
    start_time = time.time()  # Record the start time
    main()  # Run your main function
    end_time = time.time()  # Record the end time after processing

    elapsed_time_sec = end_time - start_time  # Calculate the elapsed time in seconds
    elapsed_time_min = elapsed_time_sec / 60  # Calculate the elapsed time in minutes
    print(f"Execution time: {elapsed_time_sec:.2f} seconds, or {elapsed_time_min:.2f} minutes")

    plt.show()  # Show the figure after the timing-keeping is done
