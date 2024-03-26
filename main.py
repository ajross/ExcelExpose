import argparse
import os
import shutil
import zipfile
from tempfile import mkdtemp
import xml.etree.ElementTree as ET

# Register the namespace globally to avoid 'ns0:' prefixes
ET.register_namespace('', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')


def unzip_excel_as_zip(filename, template_dir):
    if not filename.endswith('.xlsx'):
        raise ValueError("File must be an Excel (.xlsx) file")

    temp_dir = mkdtemp()
    extracted_dir = os.path.join(temp_dir, "Extracted")
    extracted_template_dir = os.path.join(temp_dir, "Template")

    try:
        shutil.copytree(template_dir, extracted_template_dir)
        zip_path = filename + '.zip'
        os.rename(filename, zip_path)
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(extracted_dir)
        os.rename(zip_path, filename)
    except Exception as e:
        shutil.rmtree(temp_dir)
        raise e

    return temp_dir, extracted_dir


def parse_template_sheet_data(filepath):
    tree = ET.parse(filepath)
    root = tree.getroot()
    return tree, root  # Return both the tree (for manipulation) and the root


def parse_sheet_data(filepath, namespaces):
    tree = ET.parse(filepath)
    root = tree.getroot()
    sheet_data_elements = root.findall('.//ns:sheetData', namespaces=namespaces)

    for sheet_data in sheet_data_elements:
        if sheet_data.find('.//ns:row', namespaces=namespaces) is not None:
            return sheet_data

    raise ValueError(f"No <sheetData> element with a <row> child found in {filepath}")


def process_xml_files(extracted_dir, template_tree, template_root, namespaces, temp_dir):
    external_links_path = os.path.join(extracted_dir, 'xl', 'externalLinks')
    output_dir = os.path.join(temp_dir, "output")
    os.makedirs(output_dir, exist_ok=True)

    template_sheet_data = template_root.find('.//ns:sheetData', namespaces=namespaces)
    if template_sheet_data is not None:
        for i, filename in enumerate(os.listdir(external_links_path), start=1):
            if filename.endswith('.xml'):
                full_path = os.path.join(external_links_path, filename)
                try:
                    new_sheet_data = parse_sheet_data(full_path, namespaces)
                    template_sheet_data.clear()  # Clear existing children
                    for child in new_sheet_data:
                        for subchild in child:
                            # Check if the node is a 'cell' and rename it to 'c'
                            if subchild.tag == '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}cell':
                                subchild.tag = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'
                        template_sheet_data.append(child)

                    output_filename = os.path.join(output_dir, f"sheet{i:03}.xml")
                    template_tree.write(output_filename, encoding='UTF-8', xml_declaration=True)
                    print(f"Processed: {filename} -> {output_filename}")
                except ValueError as e:
                    print(e)
    else:
        raise ValueError("Template does not contain a <sheetData> element.")

def main():
    parser = argparse.ArgumentParser(description="Extract and process cached data from an Excel file.")
    parser.add_argument("filename", type=str, help="Path to the Excel file")
    args = parser.parse_args()

    template_dir = "Template"
    namespaces = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

    try:
        temp_dir, extracted_dir = unzip_excel_as_zip(args.filename, template_dir)
        template_sheet_path = os.path.join(temp_dir, "Template", "xl", "worksheets", "sheet1.xml")
        template_tree, template_root = parse_template_sheet_data(template_sheet_path)
        process_xml_files(extracted_dir, template_tree, template_root, namespaces, temp_dir)
    finally:
        shutil.rmtree(temp_dir)


if __name__ == "__main__":
    main()
