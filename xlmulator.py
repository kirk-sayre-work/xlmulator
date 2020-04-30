#!/usr/bin/env python

import argparse

import prettytable

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
    args = parser.parse_args()

    # Emulate the XLM macros.
    actions, xlm_code = emulate_XLM(args.maldocs)

    # Display results.

    # XLM Macros.
    print('-'*79)
    print('XLM MACRO %s ' % args.maldocs)
    print('- '*39)
    print(xlm_code)
    
    print('\nRecorded Actions:')
    print(dump_actions(actions))
    
