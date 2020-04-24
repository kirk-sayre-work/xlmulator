#!/usr/bin/env python

import argparse

import XLM

"""@package xlmulator
Top level command line tool for emulating Excel XLM macros.
"""

###########################################################################
def emulate_XLM(maldoc):
    """
    Emulate the behavior of the XLM macros in the given Excel file.

    @param maldoc (str) The fully qualified name of the Excel file to
    analyze.

    @return (boolean) Return True if analysis succeeded, False if it failed.
    """

    # Emulate the XLM macros.
    r = XLM.emulate(maldoc)

    return True


###########################################################################
# Main Program
###########################################################################
if __name__ == '__main__':

    # Read command line arguments.
    help_msg = "Emulate the behavior of the XLM macros in a Excel file."
    parser = argparse.ArgumentParser(description=help_msg)
    parser.add_argument("maldocs", help="The Excel file to analyze.")
    args = parser.parse_args()

    # Emulate the XLM macros.
    emulate_XLM(args.maldocs)
    
