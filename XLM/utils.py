"""@package utils
Utility functions.
"""

import string

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
