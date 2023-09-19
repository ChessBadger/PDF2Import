import re
from openpyxl import load_workbook

# Define the input and output file paths
input_file_path = 'report.txt'
output_file_path = 'output.txt'

# Define a regular expression pattern to match lines starting with 3 or more numbers
pattern = r'^\d{2,}'

# Open the input file for reading and the output file for writing
with open(input_file_path, 'r') as input_file, open(output_file_path, 'w') as output_file:
    # Loop through each line in the input file
    for line in input_file:
        # Check if the line matches the pattern
        if re.match(pattern, line):
            # If it does, write the line to the output file
            output_file.write(line)

# REMOVE LINES WITH DATES
with open(output_file_path, 'r') as file:
    lines = file.readlines()

# Step 2: Reopen the file in write mode and write back only the lines you want to keep
with open(output_file_path, 'w') as file:
    for line in lines:
        # Check if the line does not contain slashes (2)
        if line.count('/') < 2 and line.count('\\') < 2:
            file.write(line)

# Read the file content
with open(output_file_path, 'r', encoding='utf-8') as file:
    content = file.read()

# Remove all commas
content_without_commas = content.replace(',', '')

# Write the modified content back to the file
with open(output_file_path, 'w', encoding='utf-8') as file:
    file.write(content_without_commas)

# AREA
with open('output.txt', 'r') as infile, open('area.txt', 'w') as outfile:
    buffer_line = None
    non_match_count = 0

    for line in infile:
        if re.search(r'\bSIDE\b', line, re.IGNORECASE):
            match = re.search(r'\b\d{5}\b', line)
            if match:
                if buffer_line:
                    outfile.write(buffer_line * non_match_count)
                # Keep only the 5-digit number
                buffer_line = match.group(0) + '\n'
                non_match_count = 0
        else:
            match = re.match(r'^(\d{5})(?!.*\d)', line.strip())
            if match:
                if buffer_line:
                    outfile.write(buffer_line * non_match_count)
                # Only keep the 5-digit number
                buffer_line = match.group(1) + '\n'
                non_match_count = 0
            elif not re.match(r'^\d{5}', line.strip()) and not re.match(r'^\d*\.\d+', line.strip()):
                non_match_count += 1

    # Handle the last matching line at the end of the file
    if buffer_line:
        outfile.write(buffer_line * non_match_count)

# #LOCATION
with open('output.txt', 'r') as infile, open('location.txt', 'w') as outfile:
    buffer_line = None
    non_match_count = 0

    for line in infile:
        match = re.match(r'^(\d{5}).*\d', line.strip())
        if match:
            if buffer_line:
                outfile.write(buffer_line * non_match_count)
            buffer_line = match.group(1) + '\n'  # Only keep the 5-digit number
            non_match_count = 0
        elif not re.match(r'^\d{5}', line.strip()) and not re.match(r'^\d*\.\d+', line.strip()):
            non_match_count += 1

    # Handle the last matching line at the end of the file
    if buffer_line:
        outfile.write(buffer_line * non_match_count)


# #CATEGORY
# Open the input file for reading, the output files for writing
with open('output.txt', 'r') as input_file, open('category.txt', 'w') as category_file, open('totals.txt', 'w') as totals_file:
    # Iterate through each line in the input file
    for line in input_file:
        # Remove leading and trailing whitespace from the line
        line = line.strip()

        # Split the line into words
        words = line.split()

        # Check if the line has at least two words
        if len(words) >= 2:
            # Check if the first part is a 3-digit number
            if words[0].isdigit() and len(words[0]) == 3 or len(words[0]) == 2:
                # Write the first 3 digits to the category file
                category_file.write(words[0] + '\n')

                # Write the rest of the line to the totals file
                totals_file.write(' '.join(words[1:]) + '\n')

# PRIORS
# Open the "totals" file for reading
with open('totals.txt', 'r') as totals_file:
    # Open the "prior1", "prior2", and "prior3" files for writing
    with open('prior1.txt', 'w') as prior1_file, open('prior2.txt', 'w') as prior2_file, open('prior3.txt', 'w') as prior3_file:
        # Iterate through each line in the "totals" file
        for line in totals_file:
            # Use regular expressions to find all numbers in the line (including commas)
            matches = re.findall(r'\b\d+\.\d+\b', line)

            # Check if at least one number was found in the line
            if matches:
                # Write the first number to "prior1.txt"
                prior1_file.write(matches[0].replace(',', '') + '\n')

                # Check if at least two numbers were found in the line
                if len(matches) >= 2:
                    # Write the second number to "prior2.txt"
                    prior2_file.write(matches[1].replace(',', '') + '\n')

                    # Check if at least three numbers were found in the line
                    if len(matches) >= 3:
                        # Write the third number to "prior3.txt"
                        prior3_file.write(matches[2].replace(',', '') + '\n')


# COPY TO CSV
# Define the mappings between file names and the target starting cells
file_to_cell_map = [
    ('area.txt', 'B16'),
    ('location.txt', 'C16'),
    ('category.txt', 'D16'),
    ('prior1.txt', 'F16'),
]

# Load the existing workbook
workbook = load_workbook(filename="001 CSV IMPORT FORMAT FOR UNITED.xlsx")

# Get the active worksheet
worksheet = workbook.active

# Read content from each file and write it to the respective cell in the Excel file
for file_name, cell in file_to_cell_map:
    with open(file_name, 'r', encoding='utf-8') as f:
        # Get the column letter and starting row number from the cell address
        col_letter, row_num = cell[0], int(cell[1:])
        # Read the file line by line and write each line to a new cell
        for i, line in enumerate(f, start=row_num):
            worksheet[f'{col_letter}{i}'] = line.strip()

# Save the workbook with the modified content
workbook.save(filename="001 CSV IMPORT FORMAT FOR UNITED WORKING.xlsx")

print('Content has been transferred to the Excel file.')
