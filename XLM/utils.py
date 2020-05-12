"""@package utils
Utility functions.
"""

import string
import subprocess
from functools import reduce

import XLM.color_print

####################################################################
def convert_num(num_str):
    """
    Convert a float or int string to a float or int.

    @param num_str (str) The numeric string.

    @return (int or float) The converted number.
    """

    # Try 1st as an int.
    num_str = str(num_str)
    try:
        return int(num_str)
    except ValueError:
        pass

    # Now try as a float.
    try:
        return float(num_str)
    except ValueError:
        XLM.color_print.output('r', "ERROR: Cannot convert '" + num_str + "' to a number. Returning 0.")
        return 0

####################################################################
def strip_unprintable(the_str):

    # Grr. Python2 unprintable stripping.
    r = the_str
    if ((isinstance(r, str)) or (not isinstance(r, bytes))):
        r = ''.join(filter(lambda x:x in string.printable, r))
        
    # Grr. Python3 unprintable stripping.
    else:
        tmp_r = ""
        for char_code in filter(lambda x:chr(x) in string.printable, r):
            tmp_r += chr(char_code)
        r = tmp_r

    # Done.
    return r

###################################################################################################
def is_excel_file(maldoc):
    """
    Check to see if the given file is an Excel (97 or 2007+) file.

    @param name (str) The name of the file to check.

    @return (bool) True if the file is an Excel file, False if not.
    """
    return (is_excel_file_97(maldoc) or is_excel_file_2007(maldoc))

###################################################################################################
def is_excel_file_2007(maldoc):
    """
    Check to see if the given file is a 2007+ Excel file.

    @param name (str) The name of the file to check.

    @return (bool) True if the file is an Excel file, False if not.
    """
    typ = subprocess.check_output(["file", maldoc])
    if (b"Excel 2007+" in typ):
        return True
    return False

###################################################################################################
def is_excel_file_97(maldoc):
    """
    Check to see if the given file is an Excel 97 file.

    @param name (str) The name of the file to check.

    @return (bool) True if the file is an Excel file, False if not.
    """
    typ = str(subprocess.check_output(["file", maldoc]))
    if ("Composite Document File" not in typ):
        return False
    if ("Excel" in typ):
        return True
    typ = str(subprocess.check_output(["exiftool", maldoc]))
    return (("ms-excel" in typ) or ("Worksheets" in typ))

####################################################################
def excel_col_letter_to_index(x): 
    """
    Convert a 'AH','C', etc. style Excel column reference to its integer
    equivalent.

    @param x (str) The letter style Excel column reference.

    @return (int) The integer version of the column reference.
    """
    return reduce(lambda s,a:s*26+ord(a)-ord('A')+1, x, 0)

####################################################################
def parse_cell_index(cell_id_raw):
    """
    Parse out a MS cell reference like 'AD234' into an integer (row, column)
    tuple.

    @param cell_id_raw (str) The MS Excel letters for column, integer for row
    style cell reference.

    @return (tuple) An integer (row, column) tuple.
    """

    # Pull out the letter style column ID.
    col_raw = ""
    row_pos = 0
    for c in to_str(cell_id_raw):
        if (c.isdigit()):
            break
        row_pos += 1
        col_raw += c

    # Convert the letter style column ID to an integer.
    #print(curr_cell_info)
    col = excel_col_letter_to_index(col_raw)

    # Get the row #.
    row = int(cell_id_raw[row_pos:])

    # Done.
    return (row, col)

####################################################################
def to_str(s):
    """
    Convert a bytes like object to a str.

    param s (bytes) The string to convert to str. If this is already str
    the original string will be returned.

    @return (str) s as a str.
    """

    # Needs conversion?
    if (isinstance(s, bytes)):
        return s.decode()
    return s

