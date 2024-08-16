import csv
import os

# Directory containing the CSV files
directory = os.path.dirname(os.path.abspath(__file__))

# List to store data from all CSV files
all_rows = []
header_saved = False

# Check and read each CSV file in the directory
for filename in os.listdir(directory):
    if filename.endswith(".csv"):
        file_path = os.path.join(directory, filename)
        with open(file_path, newline='', encoding='utf-8') as csvfile:
            reader = csv.reader(csvfile)
            header = next(reader)  # Read the header

            if not header_saved:
                all_rows.append(header)  # Save the header from the first file
                header_saved = True
            
            all_rows.extend(reader)  # Append the rest of the data

if all_rows:
    # Create a new CSV file to save the merged content
    output_file = os.path.join(directory, 'merged_output.csv')
    with open(output_file, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerows(all_rows)

    print("All CSV files have been merged successfully, with only the first file's header preserved.")
else:
    print("No CSV files found in the directory.")
