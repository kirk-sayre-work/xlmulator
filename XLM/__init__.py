"""@package XLM

Top level Excel XLM macro emulator interface.
"""

from __future__ import print_function
import subprocess
import sys
import re
import os
import string

# https://github.com/kirk-sayre-work/office_dumper.git
import excel

import XLM.color_print
import XLM.stack_transformer
import XLM.XLM_Object
import XLM.xlm_library
import XLM.utils
import XLM.ms_stack_transformer
import XLM.excel2007

## Check installation prerequisites.

# Make sure olevba is installed.
try:
    subprocess.check_output(["olevba", "-h"])
except Exception as e:
    color_print.output('r', "ERROR: It looks like olevba is not installed. " + str(e) + "\n")
    sys.exit(101)

# Debugging flag.
debug = False

####################################################################
def set_debug(flag):
    """
    Turn debugging on or off.

    @param flag (boolean) True means output debugging, False means no.
    """
    global debug
    debug = flag
    XLM.XLM_Object.debug = flag
    XLM.xlm_library.debug = flag
    XLM.ms_stack_transformer.debug = flag
    XLM.stack_transformer.debug = flag
    XLM.excel2007.debug = flag
    
####################################################################
def _extract_xlm(maldoc):
    """
    Run olevba on the given file and extract the XLM macro code lines.

    @param maldoc (str) The fully qualified name of the Excel file to
    analyze.

    @return (str) The XLM macro cells extracted from running olevba on the
    given file. None will be returned on error.
    """

    # Run olevba on the given file.
    olevba_out = None
    FNULL = open(os.devnull, 'w')
    try:
        cmd = "timeout 30 olevba -c \"" + str(maldoc) + "\""
        olevba_out = subprocess.check_output(cmd, shell=True, stderr=FNULL)
    except Exception as e:
        color_print.output('r', "ERROR: Running olevba on " + str(maldoc) + " failed. " + str(e))
        return None

    # Pull out the chunks containing the XLM lines.
    chunk_pat = b"in file: xlm_macro \- OLE stream: 'xlm_macro'\n(?:\- ){39}\n(.+)"
    chunks = re.findall(chunk_pat, olevba_out, re.DOTALL)
    
    # Pull out all the XLM lines from each chunk.
    r = b""
    xlm_pat = br"' \d\d\d\d {1,10}\d{1,6} [^\n]+\n"
    for chunk in chunks:

        # plugin_biff does not escape newlines in strings. Try to find them and fix them.
        mod_chunk = b""
        lines = chunk.split(b"\n")
        pos = -1
        new_line = b""
        while (pos < (len(lines) - 1)):

            # Start putting together an aggregated line?
            pos += 1
            curr_line = lines[pos]
            if (curr_line.startswith(b"' ")):
                mod_chunk += new_line + b"\n"
                new_line = curr_line
                continue

            # This line is part of a string with unescaped newlines.
            new_line += b"\\n" + curr_line
        mod_chunk += b"\n" + new_line
        
        for line in re.findall(xlm_pat, mod_chunk):

            # plugin_biff does not escape double quotes in strings. Try to find them
            # and fix them.
            #
            # ' 0006     72 FORMULA : Cell Formula - R9C1 len=50 ptgRefV R7C49153 ptgStr "Set wsh = CreateObject("WScript.Shell")" ptgFuncV FWRITELN (0x0089) 
            str_pat = b"Str \".*?\" ptg"
            str_pat1 = b"Str \"(.*?)\" ptg"
            for old_str in re.findall(str_pat, line):
                tmp_str = re.findall(str_pat1, old_str)[0]
                if (b'"' in tmp_str):
                    new_str = b"Str '" + old_str[5:-5] + b"' ptg"
                    line = line.replace(old_str, new_str)
            r += line

    # Convert characters so this can be parsed.
    try:
        r = XLM.utils.to_str(r)
    except UnicodeDecodeError:
        r = XLM.utils.strip_unprintable(r)

    # Did we find XLM?
    if (len(r.strip()) == 0):
        color_print.output('y', "WARNING: No XLM found.")
        return None
            
    # Done. Return XLM lines.
    return r

####################################################################
def _guess_xlm_sheet(workbook):
    """
    Guess the sheet containing the XLM macros by finding the sheet with the
    most unresolved "#NAME" cells.

    @param workbook (ExcelSheet object) The Excel spreadsheet to check.

    @return (str) The name of the sheet that might contain the XLM macros.
    """

    # TODO: If plugin_biff.py used by olevba to dump XLM includes sheet names this
    # function will no longer be needed.

    # Look at each sheet.
    xlm_sheet = None
    unresolved_count = -1
    for curr_sheet_name in workbook.sheet_names():
        curr_sheet = workbook.sheet_by_name(curr_sheet_name)
        curr_unresolved_count = 0
        for cell_value in list(curr_sheet.cells.values()):
            cell_value = cell_value.strip()
            if (len(cell_value) == 0):
                continue
            if (cell_value.strip() == "#NAME?"):
                curr_unresolved_count += 1
        if (curr_unresolved_count > unresolved_count):
            unresolved_count = curr_unresolved_count
            xlm_sheet = curr_sheet_name

    # Return the sheet with the most '#NAME' cells.
    return xlm_sheet                
    
####################################################################
def _merge_XLM_cells(maldoc, xlm_cells):
    """
    Merge the given XLM cells into the value cells read from the
    given Excel file.

    @param maldoc (str) The fully qualified name of the Excel file to
    analyze.

    @param xlm_cells (dict) A dict of XLM formula objects (XLM_Object objects) where
    dict[ROW][COL] gives the XLM cell at (ROW, COL).

    @return (tuple) A 3 element tuple where the 1st element is the updated ExcelWorkbook and 
    2nd element is a list of 2 element tuples containing the XLM cell indices on success and 
    the 3rd element is the XLM sheet object, (None, None, None) on error.
    """

    # Read in the Excel workbook data.
    color_print.output('g', "Merging XLM macro cells with data cells ...")
    workbook = excel.read_excel_sheets(maldoc)
    if (workbook is None):
        color_print.output('r', "ERROR: Reading in Excel file " + str(maldoc) + " failed.")
        return (None, None, None)

    # Guess the name of the sheet containing the XLM macros.
    xlm_sheet_name = _guess_xlm_sheet(workbook)
    if debug:
        print("XLM Sheet:")
        print(xlm_sheet_name)
        print("")

    # Insert the XLM macros into the XLM sheet.
    xlm_sheet = workbook.sheet_by_name(xlm_sheet_name)
    xlm_cell_indices = []
    if debug:
        print("=========== START MERGE ==============")
    rows = xlm_cells.keys()
    for row in rows:
        cols = xlm_cells[row].keys()
        for col in cols:
            xlm_cell = xlm_cells[row][col]
            if debug:
                print((row, col))
                print(xlm_cell)
            cell_index = (row, col)
            xlm_sheet.cells[cell_index] = xlm_cell
            xlm_cell_indices.append(cell_index)

    # Debug.
    if debug:
        print("=========== DONE MERGE ==============")
        print(workbook)
            
    # Done. Return the indices of the added XLM cells and the updated
    # workbook.
    color_print.output('g', "Merged XLM macro cells with data cells.")
    return (workbook, xlm_cell_indices, xlm_sheet)

####################################################################
def _read_workbook_2007(maldoc):
    """
    Read in an Excel 2007+ workbook and the XLM macros in the workbook.

    @param maldoc (str) The fully qualified name of the Excel file to
    analyze.

    @return (tuple) A 3 element tuple where the 1st element is the workbook object,
    the 2nd element is a list of XLM cell indices ((row, column) tuples) and the 3rd
    element is a sheet element for the sheet with XLM macros.
    """

    # Read in the 2007+ cells.
    color_print.output('g', "Analyzing Excel 2007+ file ...")
    workbook_info = XLM.excel2007.read_excel_2007_XLM(maldoc)    
    color_print.output('g', "Extracted XLM from ZIP archive.")
    if (workbook_info is None):
        return (None, None, None)
    if (len(workbook_info) == 0):
        color_print.output('y', "WARNING: No XLM macros found.")
        return (None, None, None)

    if debug:
        print("=========== START 2007+ CONTENTS ==============")
        for sheet in workbook_info.keys():
            print("\n------")
            print(sheet)
            print("")
            for c in workbook_info[sheet].keys():
                print(str(c) + " ---> " + str(workbook_info[sheet][c]))
        print("=========== DONE 2007+ CONTENTS ==============")
            
    # Figure out which sheet probably has the XLM macros.
    xlm_sheet_name = None
    max_formulas = -1
    for sheet in workbook_info.keys():
        if (len(workbook_info[sheet]) > max_formulas):
            max_formulas = len(workbook_info[sheet])
            xlm_sheet_name = sheet

    # Parse each formula and add it to a sheet object.
    xlm_cells = {}
    for cell_index in workbook_info[xlm_sheet_name].keys():

        # Value only cell?
        row = cell_index[0]
        col = cell_index[1]
        if (row not in xlm_cells):
            xlm_cells[row] = {}
        raw_formula = workbook_info[xlm_sheet_name][cell_index][0]
        if (raw_formula is None):

            # Do we have a value?
            formula_val = workbook_info[xlm_sheet_name][cell_index][1]
            if (formula_val is not None):

                # Just save the value in the cell.
                xlm_cells[row][col] = formula_val
            continue
            
        # Parse the formula into an XLM object.
        formula_str = b"=" + raw_formula
        formula = XLM.ms_stack_transformer.parse_ms_xlm(formula_str)

        # Set the value of the formula if we know it.
        formula_val = workbook_info[xlm_sheet_name][cell_index][1]
        if (formula_val is not None):
            formula.value = formula_val

        # Save the XLM object.
        formula.update_cell_id(cell_index)
        xlm_cells[row][col] = formula
    color_print.output('g', "Parsed MS XLM macros.")
        
    # Merge the XLM cells with the value cells into a single unified spereadsheet
    # object.
    workbook, xlm_cell_indices, xlm_sheet = _merge_XLM_cells(maldoc, xlm_cells)
    if (workbook is None):
        color_print.output('r', "ERROR: Merging XLM cells failed. Emulation aborted.")
        return (None, None, None)
    
    # Done.
    return (workbook, xlm_cell_indices, xlm_sheet)
    
####################################################################
def _read_workbook_97(maldoc):
    """
    Read in an Excel 97 workbook and the XLM macros in the workbook.

    @param maldoc (str) The fully qualified name of the Excel file to
    analyze.

    @return (tuple) A 3 element tuple where the 1st element is the workbook object,
    the 2nd element is a list of XLM cell indices ((row, column) tuples) and the 3rd
    element is a sheet element for the sheet with XLM macros.
    """

    # Run olevba on the file and extract the XLM macro code lines.
    color_print.output('g', "Analyzing Excel 97 file ...")
    xlm_code = _extract_xlm(maldoc)
    color_print.output('g', "Extracted XLM with olevba.")
    if debug:
        print("=========== START RAW XLM ==============")
        print(xlm_code)
        print("=========== DONE RAW XLM ==============")
    if (xlm_code is None):
        color_print.output('r', "ERROR: Unable to extract XLM. Emulation aborted.")
        return (None, None, None)

    # Parse the XLM text and get XLM objects that can be emulated.
    xlm_cells = XLM.stack_transformer.parse_olevba_xlm(xlm_code)
    color_print.output('g', "Parsed olevba XLM macros.")
    if (xlm_cells is None):
        color_print.output('r', "ERROR: Parsing of XLM failed. Emulation aborted.")
        return (None, None, None)

    # Merge the XLM cells with the value cells into a single unified spereadsheet
    # object.
    workbook, xlm_cell_indices, xlm_sheet = _merge_XLM_cells(maldoc, xlm_cells)
    if (workbook is None):
        color_print.output('r', "ERROR: Merging XLM cells failed. Emulation aborted.")
        return (None, None, None)

    # Done.    
    return (workbook, xlm_cell_indices, xlm_sheet)
    
####################################################################
def emulate(maldoc):
    """
    Emulate the behavior of a given Excel file containing XLM macros.

    @param maldoc (str) The fully qualified name of the Excel file to
    analyze.

    @return (tuple) 1st element is a list of 3 element tuples containing the actions performed
    by the sheet, 2nd element is the human readable XLM code.
    """

    # Excel 97 file?
    if (XLM.utils.is_excel_file_97(maldoc)):
        workbook, xlm_cell_indices, xlm_sheet = _read_workbook_97(maldoc)
        if (workbook is None):
            color_print.output('r', "ERROR: Reading Excel 97 file failed. Emulation aborted.")
            return ([], "")

    # Excel 2007+ file?
    elif (XLM.utils.is_excel_file_2007(maldoc)):
        workbook, xlm_cell_indices, xlm_sheet = _read_workbook_2007(maldoc)
        if (workbook is None):
            color_print.output('r', "ERROR: Reading Excel 2007 file failed. Emulation aborted.")
            return ([], "")

    else:
        color_print.output('y', "WARNING: " + maldoc + " is not an Excel file. Emulation aborted.")
        return ([], "")
        
    # Save the indices of the XLM cells in the workbook. We do this here directly so that
    # the base definition of the ExcelWorkbook class does not need to be changed.
    xlm_sheet.xlm_cell_indices = xlm_cell_indices
    
    # Emulate the XLM.
    color_print.output('g', "Starting XLM emulation ...")
    r = XLM_Object.eval(xlm_sheet)
    color_print.output('g', "Finished XLM emulation.")
    
    # Done.
    return r
