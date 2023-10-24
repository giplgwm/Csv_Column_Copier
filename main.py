import csv
import os
import string
import configparser


def col2num(col):
    num = 0
    for c in col:
        if c in string.ascii_letters:
            num = num * 26 + (ord(c.upper()) - ord('A')) + 1
    return num


config = configparser.ConfigParser()
config.read('config.ini')
input_path = config['FILE PATHS']['input_path']
output_path = config['FILE PATHS']['output_path']
protected_columns = list(eval(config['WORKING COLUMNS']['protected_columns']))
output_column_names = list(eval(config['WORKING COLUMNS']['output_column_names']))

if type(protected_columns[0]) == str:
    for x in range(len(protected_columns)):
        protected_columns[x] = col2num(protected_columns[x])

input_path = input_path.removesuffix('/').removeprefix('/')
output_path = output_path.removesuffix('/').removeprefix('/')
if len(protected_columns) > len(output_column_names):
    print('You did not enter enough titles for the number of columns you have. For now they will be filled in with "PROTECTED DATA"')
    dif = len(protected_columns) - len(output_column_names)
    for x in range(dif):
        output_column_names.append('Protected Data')
all_files = os.listdir(input_path)
csv_files = list(filter(lambda f: f.endswith('.csv'), all_files))


for csv_file in csv_files:
    buffer = []
    all_rows = []
    try:
        with open(input_path+'/'+csv_file,  newline='') as in_file, open(output_path+'/'+csv_file, 'w', newline='') as out_file:
            input_csv = csv.reader(in_file, dialect='excel')
            output_csv = csv.writer(out_file, dialect='excel')
            for row in input_csv:
                all_rows.append(row)
            for title in output_column_names:
                all_rows[0].append(title)
            for row in all_rows[1:]:
                for column in protected_columns:
                    if row[column-1]:
                        row.append(f'`{row[column-1]}`')
                    else:
                        row.append('\n')
            output_csv.writerows(all_rows)
    except PermissionError:
        print('File Permission Error: Please close all csv files in your input directory and try again, thank you.')
