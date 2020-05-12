#!/usr/bin/env python

import argparse

import prettytable
import excel

import XLM
import XLM.color_print
import XLM.utils

"""@package xlmulator
Top level command line tool for emulating Excel XLM macros.
"""

###########################################################################
def emulate_XLM(maldoc, debug=False):
    """
    Emulate the behavior of the XLM macros in the given Excel file.

    @param maldoc (str) The fully qualified name of the Excel file to
    analyze.

    @param debug (boolean) Whether to print LOTS of debug to STDOUT.

    @return (tuple) 1st element is a list of 3 element tuples containing the actions performed
    by the sheet, 2nd element is the human readable XLM code.
    """

    # Does the file exist?
    try:
        f = open(maldoc, 'r')
        f.close()
    except IOError:
        XLM.color_print.output('r', "ERROR: File '" + str(maldoc) + "' cannot be opened. Not emulating.")
        return ([], "")
        
    # Only emulate Excel files.
    if (not XLM.utils.is_excel_file(maldoc)):
        XLM.color_print.output('y', "WARNING: '" + str(maldoc) + "' is not an Excel file. Not emulating.")
        return ([], "")
    
    # Emulate the XLM macros.
    XLM.set_debug(debug)
    r = XLM.emulate(maldoc)
    return r

###########################################################################
def dump_actions(actions):
    """
    return a table of all actions recorded by trace(), as a prettytable object
    that can be printed or reused.
    """
    t = prettytable.PrettyTable(('Action', 'Parameters', 'Description'))
    t.align = 'l'
    t.max_width['Action'] = 20
    t.max_width['Parameters'] = 65
    t.max_width['Description'] = 25
    for action in actions:
        # Cut insanely large results down to size.
        str_action = str(action)
        if (len(str_action) > 50000):
            new_params = str(action[1])
            if (len(new_params) > 50000):
                new_params = new_params[:25000] + "... <SNIP> ..." + new_params[-25000:]
            action = (action[0], new_params, action[2])
        t.add_row(action)
    return t

###########################################################################
# Main Program
###########################################################################
if __name__ == '__main__':

    # Read command line arguments.
    help_msg = "Emulate the behavior of the XLM macros in a Excel file."
    parser = argparse.ArgumentParser(description=help_msg)
    parser.add_argument("maldocs", help="The Excel file to analyze.")
    parser.add_argument('-d', '--debug',
                        action='store_true',
                        help="Print lots of debug information.")
    args = parser.parse_args()

    # Emulate the XLM macros.
    print("Emulating XLM macros in " + str(args.maldocs) + " ...")
    actions, xlm_code = emulate_XLM(args.maldocs, args.debug)
    print("Done emulating XLM macros in " + str(args.maldocs) + " .")

    # Display results.

    # XLM Macros.
    if (len(xlm_code) > 0):
        print('-'*79)
        print('XLM MACRO %s ' % args.maldocs)
        print('- '*39)
        print(xlm_code)

    if (len(actions) > 0):
        print('\nRecorded Actions:')
        print(dump_actions(actions))
    
