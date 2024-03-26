import argparse
import os
import shutil
import zipfile
from tempfile import mkdtemp
import xml.etree.ElementTree as ET


def unzip_excel_as_zip(filename, template_dir):
    if not filename.endswith('.xlsx'):
        raise ValueError("File must be an Excel (.xlsx) file")

    temp_dir = mkdtemp()
    extracted_dir = os.path.join(temp_dir, "Extracted")
    extracted_template_dir = os.path.join(temp_dir, "Template")

    try:
        # Copy Template directory
        shutil.copytree(template_dir, extracted_template_dir)

        # Rename the file to .zip and extract to the Extracted subdirectory
        zip_path = filename + '.zip'
        os.rename(filename, zip_path)
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(extracted_dir)
        os.rename(zip_path, filename)
    except Exception as e:
        shutil.rmtree(temp_dir)
        raise e

    return temp_dir, extracted_dir


def parse_sheet_data(filepath):
    tree = ET.parse(filepath)
    root = tree.getroot()
    # Define the namespace map
    namespaces = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

    # Use the namespace in the XPath query
    sheet_data_elements = root.findall('.//ns:sheetData', namespaces=namespaces)

    for sheet_data in sheet_data_elements:
        # Check for a child <row> node within <sheetData>
        if sheet_data.find('.//ns:row', namespaces=namespaces) is not None:
            # If a <sheetData> element with a <row> child is found, serialize and return it
            return ET.tostring(sheet_data, encoding='utf-8')

    raise ValueError(f"No <sheetData> element with a <row> child found in {filepath}")



def process_xml_files(extracted_dir, template_sheet_data, temp_dir):
    external_links_path = os.path.join(extracted_dir, 'xl', 'externalLinks')
    output_dir = os.path.join(temp_dir, "output")
    os.makedirs(output_dir, exist_ok=True)

    if not os.path.exists(external_links_path):
        print("No externalLinks directory found.")
        return

    for i, filename in enumerate(os.listdir(external_links_path), start=1):
        if filename.endswith('.xml'):
            full_path = os.path.join(external_links_path, filename)
            try:
                current_sheet_data = parse_sheet_data(full_path)
                output_filename = os.path.join(output_dir, f"sheet{i}.xml")
                with open(output_filename, 'wb') as f:
                    f.write(template_sheet_data.replace(b"<sheetData/>", current_sheet_data))
                print(f"Processed: {filename} -> {output_filename}")
            except ValueError as e:
                print(e)


def main():
    parser = argparse.ArgumentParser(description="Extract and process cached data from an Excel file.")
    parser.add_argument("filename", type=str, help="Path to the Excel file")
    args = parser.parse_args()

    template_dir = "Template"  # Adjust the path as needed
    try:
        temp_dir, extracted_dir = unzip_excel_as_zip(args.filename, template_dir)
        template_sheet_path = os.path.join(temp_dir, "Template", "xl", "worksheets", "sheet1.xml")
        template_sheet_data = parse_sheet_data(template_sheet_path)
        process_xml_files(extracted_dir, template_sheet_data, temp_dir)
    finally:
        shutil.rmtree(temp_dir)


if __name__ == "__main__":
    main()
