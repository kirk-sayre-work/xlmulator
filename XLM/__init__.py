"""@package XLM

Top level Excel XLM macro emulator interface.
"""

from __future__ import print_function
import subprocess
import sys
import re
import os

# sudo pip install lark-parser
from lark import Lark
from lark import UnexpectedInput
# https://github.com/kirk-sayre-work/office_dumper.git
import excel

import XLM.color_print
from XLM.stack_transformer import StackTransformer 
import XLM.XLM_Object

## Check installation prerequisites.

# Make sure olevba is installed.
try:
    subprocess.check_output(["olevba", "-h"])
except Exception as e:
    color_print.output('r', "ERROR: It looks like olevba is not installed. " + str(e) + "\n")
    sys.exit(101)

## Load the olevba XLM grammar.
xlm_grammar_file = os.path.join(os.path.abspath(os.path.dirname(__file__)), "olevba_xlm.bnf")
xlm_grammar = None
try:
    f = open(xlm_grammar_file, "r")
    xlm_grammar = f.read()
    f.close()
except IOError as e:
    color_print.output('r', "ERROR: Cannot read XLM grammar file " + xlm_grammar_file + ". " + str(e))
    sys.exit(102)

# Debugging flag.
debug = False
XLM.XLM_Object.debug = debug

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
    try:
        cmd = "timeout 30 olevba -c \"" + str(maldoc) + "\""
        olevba_out = subprocess.check_output(cmd, shell=True)
    except Exception as e:
        color_print.output('r', "Error running olevba on " + str(maldoc) + " failed. " + str(e))
        return None

    # Pull out all the XLM lines.
    r = b""
    xlm_pat = br"' \d\d\d\d {1,10}\d{1,6} [^\n]+\n"
    for line in re.findall(xlm_pat, olevba_out):

        # plugin_biff does not escape double quotes in strings. Try to find them
        # and fix them.
        #
        # ' 0006     72 FORMULA : Cell Formula - R9C1 len=50 ptgRefV R7C49153 ptgStr "Set wsh = CreateObject("WScript.Shell")" ptgFuncV FWRITELN (0x0089) 
        str_pat = b"Str \".*?\" ptg"
        str_pat1 = b"Str \"(.*?)\" ptg"
        for old_str in re.findall(str_pat, line):
            tmp_str = re.findall(str_pat1, old_str)[0]
            if (b'"' in tmp_str):
                new_str = "Str '" + old_str[5:-5] + "' ptg"
                line = line.replace(old_str, new_str)
        r += line
    return r

####################################################################
def _extract_xlm_objects(xlm_code):
    """
    Parse the given olevba XLM code into an internal object representation 
    that can be emulated.

    @param xlm_code (str) The olevba XLM code to parse.

    @return (dict) A dict of XLM formula objects (XLM_Object objects) where
    dict[ROW][COL] gives the XLM cell at (ROW, COL).
    """

    # Parse the olevba XLM.
    xlm_ast = None
    try:
        xlm_parser = Lark(xlm_grammar, start="lines", parser='lalr')
        xlm_ast = xlm_parser.parse(xlm_code.decode())
    except UnexpectedInput as e:
        color_print.output('r', "ERROR: Parsing olevba XLM failed.\n" + str(e))
        return None

    # Transform the AST into XLM_Object objects.
    if debug:
        print("=========== START XLM AST ==============")
        print(xlm_ast.pretty())
        print("=========== DONE XLM AST ==============")
    formula_cells = StackTransformer().transform(xlm_ast)
    if debug:
        print("=========== START XLM TRANSFORMED ==============")
        print(formula_cells)
        print("=========== DONE XLM TRANSFORMED ==============")
    return formula_cells

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
    return (workbook, xlm_cell_indices, xlm_sheet)
            
####################################################################
def emulate(maldoc):
    """
    Emulate the behavior of a given Excel file containing XLM macros.

    @param maldoc (str) The fully qualified name of the Excel file to
    analyze.

    @return (list) A list of 3 element tuples containing the actions performed
    by the Excel file.
    """

    # Run olevba on the file and extract the XLM macro code lines.
    xlm_code = _extract_xlm(maldoc)
    if debug:
        print("=========== START RAW XLM ==============")
        print(xlm_code)
        print("=========== DONE RAW XLM ==============")
    if (xlm_code is None):
        return []

    # Parse the XLM text and get XLM objects that can be emulated.
    xlm_cells = _extract_xlm_objects(xlm_code)
    if (xlm_cells is None):
        color_print.output('r', "ERROR: Parsing of XLM failed. Emulation aborted.")
        return []

    # Merge the XLM cells with the value cells into a single unified spereadsheet
    # object.
    workbook, xlm_cell_indices, xlm_sheet = _merge_XLM_cells(maldoc, xlm_cells)
    if (workbook is None):
        color_print.output('r', "ERROR: Merging XLM cells failed. Emulation aborted.")
        return []

    # Save the indices of the XLM cells in the workbook. We do this here directly so that
    # the base definition of the ExcelWorkbook class does not need to be changed.
    xlm_sheet.xlm_cell_indices = xlm_cell_indices
    
    # Emulate the XLM.
    r = XLM_Object.eval(xlm_sheet)
    
    # Done.
    return []
