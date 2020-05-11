import zipfile
import re
import random
import os
import sys

def unzip_file(fname):
    """
    Unzip zipped file.
    """

    # Is this a ZIP file?
    if (not zipfile.is_zipfile(fname)):
        print("not zip '" + str(fname) + "'")
        return None
        
    # This is a ZIP file. Unzip it.
    unzipped_data = zipfile.ZipFile(fname, 'r')

    # Return the unzipped data.
    return unzipped_data

def _read_excel_2007_sheet(zip_subfile, unzipped_data):

    # Only reading formulas.
    if (not zip_subfile.startswith(b"macrosheets/")):
        print("NOPE")
        print(zip_subfile)
        return None
    
    # Read in the formula file.
    zip_subfile = (b"xl/" + zip_subfile).decode()
    if (zip_subfile not in unzipped_data.namelist()):
        print("NOPE1")
        print(zip_subfile)
        return None
    f1 = unzipped_data.open(zip_subfile)
    contents = f1.read()
    f1.close()

    # <c r="HO1" t="str"><f>CHAR($EC$210-123)</f><v>e</v></c>
    # <c r="EY1"><v>383</v></c>
    # <c r="FE23" t="b"><f>RUN($IK$1673)</f><v>0</v></c>
    #cell_pat = b'<c +r="(\w+)".+<f>(.+)</f>(?:<v>(.+)</v>)?</c>'
    cell_pat = b'<c +r="(\w+)".+?<f>([^<]+)</f>(?:<v>([^<]+)</v>)?</c>'
    print("----")
    print(zip_subfile)
    print(re.findall(cell_pat, contents))

    return None

def read_excel_2007_XLM(fname):

    # Make sure this is an Excel 2007 file.
    
    # Unzip the file.
    unzipped_data = unzip_file(fname)
    if (fname is None):
        return None

    # Get internal ID to sheet name mapping from ./xl/workbook.xml.

    # Read in xl/workbook.xml.
    zip_subfile = 'xl/workbook.xml'
    if (zip_subfile not in unzipped_data.namelist()):
        return None
    f1 = unzipped_data.open(zip_subfile)
    contents = f1.read()
    f1.close()

    # Pull out sheet names and IDs.
    # <sheet name="DynIxoDNvviVwTft" sheetId="2" state="hidden" r:id="rId1"/>
    name_id_pat = b'<sheet +name="(\w+)"[^>]+r:id="(\w+)"'
    name_id_info = re.findall(name_id_pat, contents)

    # Make a map from sheet IDs to names.
    id_to_name_map = {}
    for i in name_id_info:
        id_to_name_map[i[1]] = i[0]
    print(id_to_name_map)

    # Get ID to formula XML file name mapping from xl/_rels/workbook.xml.rels.

    # Read in xl/_rels/workbook.xml.rels.
    zip_subfile = 'xl/_rels/workbook.xml.rels'
    if (zip_subfile not in unzipped_data.namelist()):
        return None
    f1 = unzipped_data.open(zip_subfile)
    contents = f1.read()
    f1.close()

    # Pull out sheet files and IDs.
    # <Relationship Id="rId8" Type="http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet" Target="macrosheets/sheet8.xml"/>
    id_file_pat = b'<Relationship +Id="(\w+)"[^>]+Target="([\w\./]+)"'
    file_id_info = re.findall(id_file_pat, contents)

    # Make a map from sheet IDs to formula file names.
    id_to_file_map = {}
    for i in file_id_info:
        id_to_file_map[i[0]] = i[1]
    print(id_to_file_map)

    # Read in the formulas from each sheet.
    formulas = {}
    for curr_id in id_to_name_map.keys():
        curr_sheet = id_to_name_map[curr_id]
        curr_file = id_to_file_map[curr_id]
        print(curr_id)
        print(curr_sheet)
        print(curr_file)
        curr_formulas = _read_excel_2007_sheet(curr_file, unzipped_data)
        if (curr_formulas is not None):
            formulas[curr_sheet] = curr_formulas
    
    return None

print(read_excel_2007_XLM(sys.argv[1]))

      
