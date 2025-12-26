# pip install pandas openpyxl
import pandas as pd
import os

def replace_strings(df):
    
    replacement_map = {
        'Vlada Vulovic': 'Влада Вуловић',
        'Petar Spasic': 'Петар Спасић',
        'Branislav Koledin': 'Бранислав Коледин',
        'Ognjen Sobajic': 'Огњен Шобајић',
        'Zeljko Nikolicic': 'Жељко Николичић',
        'Dejan Subotic': 'Дејан Суботић',
        'Jovica Spasic': 'Јовица Спасић',
        'Milan Stefanovic': 'Милан Стефановић',
        'Vojislav Kokeza': 'Војислав Кокеза',
        'Predrag Roso': 'Предраг Росо',
        'Mirko Spasojevic': 'Мирко Спасојевић',
        'Aca Spasojevic': 'Аца Спасојевић',
        'Dalibor Marceta': 'Далибор Марчета',
        'Nikola Rudic': 'Никола Рудић',
        'Aca': 'Аца Спасојевић',
        'Mirko': 'Мирко Спасојевић',
        'Воја Кокеза': 'Војислав Кокеза'
    }

    # Iterate over each item in the mapping
    for old_value, new_value in replacement_map.items():
        # Replace occurrences of old_value with new_value
        df = df.replace(old_value, new_value)
    return df


def izvoz(output_file_id, input_file_name):
    print(f'Processing {input_file_name}')

    # Set the output file path
    output_file = f"..\\Rezultati\\{output_file_id}.csv"

    # Check if the file already exists, and if so, delete it
    if os.path.exists(output_file):
        os.remove(output_file)

    # Create a new DataFrame for the output
    output_df = pd.DataFrame()

    # Create object for Excel file
    xls = pd.ExcelFile(f"..\\{input_file_name}\\{input_file_name}.xlsm")

    # Loop through each sheet
    for sheet_name in xls.sheet_names:
        if sheet_name.isnumeric():
            df = pd.read_excel(xls, sheet_name=sheet_name)

            # Check if the DataFrame is not empty and the first cell is not empty
            if not df.empty and df.iloc[0, 0] != '':
                last_row = df[df.columns[1]].last_valid_index()

                # Select only columns B, C, and D
                filtered_df = df.iloc[0:last_row + 1, 1:4].dropna(subset=[df.columns[1]])

                # Insert the sheet name as the first column
                filtered_df.insert(0, 'SheetName', sheet_name)

                # Append the filtered data to the output DataFrame
                output_df = pd.concat([output_df, filtered_df], ignore_index=True)

    # Apply replacements
    output_df = replace_strings(output_df)

    # Save the DataFrame to a CSV file
    output_df.to_csv(output_file, index=False, header=False)

    print(f"Data exported successfully to {output_file}\n")

def excel_to_csv():
    
    pairs = [
        ("1", r'Turnir 2011'),
        ("2", r'Turnir 2013'),
        ("3", r'Turnir 2015 - 02'),
        ("4", r'Turnir 2015 - 12'),
        ("5", r'Turnir 2016'),
        ("6", r'Turnir 2017'),
        ("7", r'Turnir 2019 - 02'),
        ("8", r'Turnir 2019 - 12'),
        ("9", r'Turnir 2022'),
        ("10", r'Turnir 2023'),
        ("11", r'Turnir 2024')
    ]
    for id, fileName in pairs:
        izvoz(id, fileName)

excel_to_csv()
