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
    output_dir = os.path.join(temp_dir, "output")

    try:
        shutil.copytree(template_dir, extracted_template_dir)
        zip_path = filename + '.zip'
        os.rename(filename, zip_path)
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(extracted_dir)
        os.rename(zip_path, filename)

        # Copy Template directory to output directory
        if os.path.exists(output_dir):
            shutil.rmtree(output_dir)
        shutil.copytree(template_dir, output_dir)
    except Exception as e:
        shutil.rmtree(temp_dir)
        raise e

    return temp_dir, extracted_dir, output_dir


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
    output_count = 1
    if template_sheet_data is not None and os.path.exists(external_links_path):
        for i, filename in enumerate(sorted(os.listdir(external_links_path)), start=1):
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

                    output_filename = os.path.join(output_dir, "xl", "worksheets", f"sheet{output_count}.xml")
                    template_tree.write(output_filename, encoding='UTF-8', xml_declaration=True)
                    print(f"Processed: {filename} -> {output_filename}")
                    output_count += 1
                except ValueError as e:
                    print(e)
    else:
        raise ValueError("Template does not contain a <sheetData> element.")

    return output_count - 1


def update_workbook_xml(output_dir, num_sheets):
    workbook_xml_path = os.path.join(output_dir, "xl", "workbook.xml")
    tree = ET.parse(workbook_xml_path)
    root = tree.getroot()
    root.attrib['xmlns:r'] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    namespaces = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

    # Find the <sheets> node
    sheets_node = root.find('ns:sheets', namespaces)
    if sheets_node is not None:
        sheets_node.clear()  # Clear existing <sheet> children

        # Append a new <sheet> node for each generated file
        for i in range(1, num_sheets + 1):
            sheet = ET.SubElement(sheets_node, '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet',
                                  attrib={
                                      'name': f"Sheet{i}",
                                      'sheetId': str(i),
                                      'state': 'visible',
                                      'r:id': f"rId{i+1}"
                                  })

    tree.write(workbook_xml_path, encoding='UTF-8', xml_declaration=True)

def update_workbook_rels(output_dir, num_sheets):
    ET.register_namespace('', 'http://schemas.openxmlformats.org/package/2006/relationships')
    workbook_xml_path = os.path.join(output_dir, "xl", "_rels", "workbook.xml.rels")
    tree = ET.parse(workbook_xml_path)
    root = tree.getroot()

    if root is not None:
        # Append a new <sheet> node for each generated file
        for i in range(1, num_sheets + 1):
            relationship = ET.SubElement(root, '{http://schemas.openxmlformats.org/package/2006/relationships}Relationship',
                                  attrib={
                                      'Id': f"rId{i+1}",
                                      'Type': "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                                      'Target': f"worksheets/sheet{i}.xml"
                                  })

    tree.write(workbook_xml_path, encoding='UTF-8', xml_declaration=True)

def create_new_excel_file(input_filepath, output_dir):
    # Extract the base name and directory of the input file
    base_name = os.path.splitext(os.path.basename(input_filepath))[0]
    input_dir = os.path.dirname(input_filepath)

    # Define the name and path for the zip (temporary) and final Excel file
    temp_zip_path = os.path.join(input_dir, f"{base_name}-extracted")
    final_excel_path = f"{temp_zip_path}.xlsx"

    # Create a zip archive of the "output" directory
    shutil.make_archive(temp_zip_path, 'zip', output_dir)

    # Rename the zip file to have a ".xlsx" extension (overwrites if exists)
    os.rename(f"{temp_zip_path}.zip", final_excel_path)

    print(f"New Excel file created at: {final_excel_path}")
    return final_excel_path

def main():
    parser = argparse.ArgumentParser(description="Extract and process cached data from an Excel file.")
    parser.add_argument("filename", type=str, help="Path to the Excel file")
    args = parser.parse_args()

    template_dir = "Template"
    namespaces = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

    try:
        temp_dir, extracted_dir, output_dir = unzip_excel_as_zip(args.filename, template_dir)
        template_sheet_path = os.path.join(temp_dir, "Template", "xl", "worksheets", "sheet1.xml")
        template_tree, template_root = parse_template_sheet_data(template_sheet_path)
        num_sheets = process_xml_files(extracted_dir, template_tree, template_root, namespaces, temp_dir)

        update_workbook_xml(output_dir, num_sheets)
        update_workbook_rels(output_dir, num_sheets)

        create_new_excel_file(args.filename, output_dir)
    finally:
        shutil.rmtree(temp_dir)


if __name__ == "__main__":
    main()
