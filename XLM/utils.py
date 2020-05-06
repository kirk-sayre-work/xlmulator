"""@package utils
Utility functions.
"""

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
