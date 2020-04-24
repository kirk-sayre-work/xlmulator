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

import XLM.color_print

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
        str_pat = b"ptgStr \"(.*)\" ptg"
        for old_str in re.findall(str_pat, line):
            if (b'"' in old_str):
                new_str = old_str.replace('"', "'")
                line = line.replace(old_str, new_str)        
        r += line
    return r

####################################################################
def _extract_xlm_objects(xlm_code):
    """
    Parse the given olevba XLM code into an internal object representation 
    that can be emulated.

    @param xlm_code (str) The olevba XLM code to parse.

    @return (XLM_Object) An object that can emulate the parsed XLM, or None on
    error
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
    print(xlm_ast.pretty())
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
    xlm_object = _extract_xlm_objects(xlm_code)
    if (xlm_object is None):
        color_print.output('r', "Parsing of XLM failed. Emulation aborted.")
        return []
    
    # Emulate the XLM.
    #r = xlm_object.eval()
    
    # Done.
    return []
