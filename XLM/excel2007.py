"""@package excel2007

Functions for reading XLM macros from Excel 2007+ files.
"""

import zipfile
import re
import random
import os
import sys

import XLM.utils

####################################################################
def unzip_file(fname):
    """
    Unzip a zipped file into memory.

    @param fname (str) The name of the file to unzip.

    @return (ZipFile object) A zipfile.ZipFile object with the unzipped file
    contents.
    """

    # Is this a ZIP file?
    if (not zipfile.is_zipfile(fname)):
        return None
        
    # This is a ZIP file. Unzip it.
    unzipped_data = zipfile.ZipFile(fname, 'r')

    # Return the unzipped data.
    return unzipped_data

####################################################################
def _read_excel_2007_sheet(zip_subfile, unzipped_data):
    """
    Read in the formulas from a given 2007+ sheet file.

    @param zip_subfile (str) The name of the macro sheet file in the Excel 2007+
    ZIP archive.

    @param unzipped_data (ZipFile object) The Excel 2007+ ZIP data.

    @return (dict) A map from a cell index (2 element (row, column) tuple) to a 2 element 
    tuple where the 1st element is the raw cell formula (str) and the 2nd element is the
    computed formula value (str, int, or float if the value is known, None if it is not
    known).
    """
    
    # Only reading formulas.
    if (not zip_subfile.startswith(b"macrosheets/")):
        return None
    
    # Read in the formula file.
    zip_subfile = (b"xl/" + zip_subfile).decode()
    if (zip_subfile not in unzipped_data.namelist()):
        return None
    f1 = unzipped_data.open(zip_subfile)
    contents = f1.read()
    f1.close()

    # Pull out the cell ID, raw formula, and formula computed value (if there is one)
    # for each cell.
    #
    # <c r="HO1" t="str"><f>CHAR($EC$210-123)</f><v>e</v></c>
    # <c r="EY1"><v>383</v></c>
    # <c r="FE23" t="b"><f>RUN($IK$1673)</f><v>0</v></c>
    cell_pat = b'<c +r="(\w+)".+?<f>([^<]+)</f>(?:<v>([^<]+)</v>)?</c>'
    cell_raw_info = re.findall(cell_pat, contents)
    r = {}
    for curr_cell_info in cell_raw_info:

        # Pull out the letter style column ID.
        cell_id_raw = curr_cell_info[0]
        col_raw = ""
        row_pos = 0
        for c in cell_id_raw.decode():
            if (c.isdigit()):
                break
            row_pos += 1
            col_raw += c

        # Convert the letter style column ID to an integer.
        col = XLM.utils.excel_col_letter_to_index(col_raw)

        # Get the row #.
        row = int(cell_id_raw[row_pos:])

        # Tuple cell index.
        cell_index = (row, col)

        # Get the raw formula.
        formula = curr_cell_info[1]

        # Do we know the computed formula value?
        formula_val = None
        if (len(curr_cell_info) > 2):

            # Get the value as a string.
            formula_val = curr_cell_info[2]

            # Try some numeric conversions. If these fail we will just track it as a
            # string.
            try:
                formula_val = int(formula_val)
            except:
                try:
                    formula_val = float(formula_val)
                except:
                    pass
                
        # Save the cell data.
        r[cell_index] = (formula, formula_val)

    return r

####################################################################
def read_excel_2007_XLM(fname):
    """
    Read in the formula cells of each sheet in an Excel 2007+ file.

    @param fname (str) The name of the Excel 2007+ file.

    @return (dict) A map from sheet names (str) to sheet formula information
    (see _read_excel_2007_sheet() for how the cell contents for each sheet are
    represented).
    """
    
    # Make sure this is an Excel 2007 file.
    if (not XLM.utils.is_excel_file_2007(fname)):
        return None
    
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

    # Read in the formulas from each sheet.
    formulas = {}
    for curr_id in id_to_name_map.keys():
        curr_sheet = id_to_name_map[curr_id]
        curr_file = id_to_file_map[curr_id]
        curr_formulas = _read_excel_2007_sheet(curr_file, unzipped_data)
        if ((curr_formulas is not None) and (len(curr_formulas) > 0)):
            formulas[curr_sheet] = curr_formulas
    
    return formulas

###########################################################################
# Main Program (for testing).
###########################################################################
if __name__ == '__main__':
    r = read_excel_2007_XLM(sys.argv[1])
    for sheet in r.keys():
        print("\n------")
        print(sheet)
        print("")
        for c in r[sheet].keys():
            print(str(c) + " ---> " + str(r[sheet][c]))
