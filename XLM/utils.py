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
    typ = subprocess.check_output(["file", maldoc])
    if ("Composite Document File" not in typ):
        return False
    if (b"Excel" in typ):
        return True
    typ = subprocess.check_output(["exiftool", maldoc])
    return ((b"ms-excel" in typ) or (b"Worksheets" in typ))

####################################################################
def excel_col_letter_to_index(x): 
    """
    Convert a 'AH','C', etc. style Excel column reference to its integer
    equivalent.

    @param x (str) The letter style Excel column reference.

    @return (int) The integer version of the column reference.
    """
    #return (reduce(lambda s,a:s*26+ord(a)-ord('A')+1, x, 0) - 1)
    return reduce(lambda s,a:s*26+ord(a)-ord('A')+1, x, 0)
