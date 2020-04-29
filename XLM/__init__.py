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
debug = True
    
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
    xlm_pat = br"' \d\d\d\d     [ \d]{2} [^\n]+\n"
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
    formula_cells = StackTransformer().transform(xlm_ast)
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

    @return (excel object) An excel workbook object on success, None on error.
    """

    # Debug.
    if debug:
        rows = list(xlm_cells.keys())
        rows.sort()
        for row in rows:
            cols = list(xlm_cells[row].keys())
            cols.sort()
            for col in cols:
                print(xlm_cells[row][col])

    # Read in the Excel workbook data.
    workbook = excel.read_excel_sheets(maldoc)
    if (workbook is None):
        color_print.output('r', "ERROR: Reading in Excel file " + str(maldoc) + " failed.")
        return None

    # Guess the name of the sheet containing the XLM macros.
    xlm_sheet_name = _guess_xlm_sheet(workbook)
    if debug:
        print("XLM Sheet:")
        print(xlm_sheet_name)

    return None
            
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
    if (xlm_code is None):
        return []

    # Parse the XLM text and get XLM objects that can be emulated.
    xlm_cells = _extract_xlm_objects(xlm_code)
    if (xlm_cells is None):
        color_print.output('r', "ERROR: Parsing of XLM failed. Emulation aborted.")
        return []

    # Merge the XLM cells with the value cells into a single unified spereadsheet
    # object.
    workbook = _merge_XLM_cells(maldoc, xlm_cells)
    if (workbook is None):
        color_print.output('r', "ERROR: Merging XLM cells failed. Emulation aborted.")
        return []
    
    # Emulate the XLM.
    #r = workbook.trace()
    
    # Done.
    return []
