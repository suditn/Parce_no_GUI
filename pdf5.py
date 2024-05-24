import os
import pdfplumber
from pathlib import Path
import pandas as pd
import json

ser1 = []
ser2 = []

def there_is_pdf(folder_path, excel_file, excel_sheet):
    wb = pd.read_excel(excel_file, sheet_name=excel_sheet, engine='openpyxl')

    series_paths = {}

    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith(".pdf"):
                file_path = os.path.join(root, file)
                relative_path = os.path.relpath(file_path, folder_path)
                ser1.append(relative_path)
                series_name = os.path.dirname(relative_path).split(os.sep)[-1]
                ser2.append(series_name)
                if series_name not in series_paths:
                    series_paths[series_name] = []
                if relative_path not in series_paths[series_name]:
                    series_paths[series_name].append(relative_path)

def create_part_number(file_name, value, tolerance):
    parts = file_name.split('-')
    if len(parts) < 2:
        return None

    product_family = parts[0]
    size = parts[1]

    if value is not None:
        try:
            value_float = float(str(value).replace(',', '.'))
            if value_float < 10:
                value_str = str(value_float).replace('.', 'R')
            else:
                value_str = str(int(value_float))
                if len(value_str) == 2:
                    value_str += '0'
                elif len(value_str) == 3:
                    value_str = value_str[:-1] + '1'
                elif len(value_str) == 4:
                    value_str = value_str[:-2] + '2'
        except ValueError:
            value_str = "000"
    else:
        value_str = "000"

    if tolerance == 20:
        tol = "M"
    elif tolerance == 30:
        tol = "N"
    else:
        tol = ""

    part_number = f"{product_family}{size}ER{value_str}{tol}"

    if len(parts) > 2:
        part_number += parts[2].split(' ')[0]

    return part_number

def create_csv_with_file_description(save_path, pdf_path, image_path):
    file_name = os.path.basename(pdf_path).replace(".pdf", "")

    table_data = []
    column_names = []
    column_names2 = []
    i = 0

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                if table and len(table) > 1 and table[0][0] and table[0][0] == "STANDARD ELECTRICAL SPECIFICATIONS":
                    if not column_names:
                        column_names = table[1]
                        column_names2 = table[2]
                        while(i < len(column_names)):
                            if column_names[i] == None:
                                if column_names2[i] != None:
                                    column_names[i] = column_names2[i]
                                else:
                                    column_names[i] = ' '
                            i += 1
                    table_data.extend(table[2:])

    if table_data:
        max_columns = max(len(row) for row in table_data)
        if not column_names:
            column_names = [f"Column {i + 1}" for i in range(max_columns)]
        else:
            while len(column_names) < max_columns:
                column_names.append(f"Column {len(column_names) + 1}")

        df = pd.DataFrame(table_data, columns=column_names)
    else:
        df = pd.DataFrame(columns=["File Name", "Description"])

    df["File Name"] = file_name
    df["Image Path"] = os.path.normpath(image_path).replace(os.sep, '/')

    with pdfplumber.open(pdf_path) as pdf:
        full_text = ""
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                full_text += page_text

    lines = [line.strip() for line in full_text.splitlines() if line.strip()]
    df["Description"] = " ".join(lines[2:3]) if len(lines) > 2 else ""

    # Ищем столбцы, содержащие ключевые слова
    value_column = None
    for col in df.columns:
        if col and any(keyword in col for keyword in ["INDUCTANCE", "(nH)", "(μH)", "IND. AT 1 kHz", "L 0 INDUCTANC ± 20 % AT 100 kHz", "INDUCTANC", "IMPEDAN", "L\n500\nMHz\n(nH)", "COMMON MODE IMPEDAN"]):
            value_column = col
            break

    if value_column:
        def validate_value(value):
            try:
                if pd.notna(value):
                    float(value)
                    return True
                return False
            except ValueError:
                return False

        def extract_value(row):
            value = row[value_column] if validate_value(row[value_column]) else None
            return create_part_number(file_name, value, 20)

        part_numbers = df.apply(extract_value, axis=1)
        df["PART NUMBER"] = part_numbers

    if "PART NUMBER" in df.columns:
        columns = ["File Name", "Image Path", "Description", "PART NUMBER"]
        columns += [col for col in df.columns if col not in columns]
        df = df[columns]
    else:
        print(f"PART NUMBER column not found in DataFrame for file: {pdf_path}")

    output_csv_path = os.path.join(os.path.dirname(pdf_path), f"{file_name}.csv")
    df.replace(r'\s+|\\n', ' ', regex=True, inplace=True)
    df.fillna('', inplace=True)

    header = ';'.join(df.columns)
    header = header.replace('\n', ' ').replace('\r', ' ')

    with open(output_csv_path, 'w', encoding='utf-8', newline='\n') as f:
        f.write(header + '\n')
        df.to_csv(f, index=False, header=False, sep=';')

    return output_csv_path

def create_individual_json(pdf_path, csv_path, image_path):
    file_name = os.path.basename(pdf_path).replace(".pdf", "")
    pdf_relative_path = f"/{os.path.normpath(pdf_path).replace(os.sep, '/')}"
    image_relative_path = f"/{os.path.normpath(image_path).replace(os.sep, '/')}"

    json_content = {
        os.path.basename(csv_path): {
            "pdf": pdf_relative_path,
            "img": image_relative_path
        }
    }

    json_output_path = os.path.join(os.path.dirname(pdf_path), f"{file_name}.json")
    with open(json_output_path, 'w', encoding='utf-8') as json_file:
        json.dump(json_content, json_file, indent=4)

    return json_output_path

def create_master_json(json_list, save_path):
    master_json_content = {}
    for json_file in json_list:
        with open(json_file, 'r', encoding='utf-8') as f:
            content = json.load(f)
            master_json_content.update(content)

    master_json_path = os.path.join(save_path, "master.json")
    with open(master_json_path, 'w', encoding='utf-8') as master_json_file:
        json.dump(master_json_content, master_json_file, indent=4)

if __name__ == "__main__":
    save_path = str(Path(__file__).parent.resolve())
    folder_path = save_path
    excel_file = "inductors.xlsx"
    excel_sheet = "Inductors"

    there_is_pdf(folder_path, excel_file, excel_sheet)

    image_paths = pd.read_excel(excel_file, sheet_name=excel_sheet, engine='openpyxl').set_index('Series')['Image path'].to_dict()

    json_files = []

    for s in ser1:
        if s != "Datasheet\\IDCS-5020\\IDCS-5020.pdf":
            pdf_path = os.path.join(s)
            series_name = s.split(os.sep)[-2]
            image_path = image_paths.get(series_name, "Images/default.png")
            output_csv_path = create_csv_with_file_description(save_path, pdf_path, image_path)
            json_file = create_individual_json(pdf_path, output_csv_path, image_path)
            json_files.append(json_file)
            print(f"Processed: {pdf_path}")

    create_master_json(json_files, save_path)
    print("Master JSON created.")
