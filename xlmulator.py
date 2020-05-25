#!/usr/bin/env python

import argparse
import json
import re
from collections import Counter

import prettytable

import XLM
import XLM.color_print
import XLM.utils

"""@package xlmulator
Top level command line tool for emulating Excel XLM macros.
"""

###########################################################################
def emulate_XLM(maldoc, debug=False, out_file_name=None):
    """
    Emulate the behavior of the XLM macros in the given Excel file.

    @param maldoc (str) The fully qualified name of the Excel file to
    analyze.

    @param debug (boolean) Whether to print LOTS of debug to STDOUT.

    @param out_file_name (str) The name of the file in which to save JSON analysis results.

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
    
    print(':'.join(get_xlmfuncset(r[1])))

    # Save analysis to a JSON file?
    if (out_file_name is not None):
        json_report = r + ({"DllFuncSetAlpha": get_dllfuncset(r[0])}, {"XLMFuncSetAlpha": get_xlmfuncset(r[1])}, {"XLMFuncSetFreq": get_xlmfuncset_frequency(r[1])},)
        with open(out_file_name, 'w') as outfile:
            json.dump(json_report, outfile)
        
    # Done.
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
def get_dllfuncset(actions):
    """
    return an alpha-sorted set of DLLs and function calls
    """

    function_set = set()
    for action in actions:
        if action[0] == 'CALL':
            function_set.add(str(action[2].split("'")[1] + "." + action[1].split("(")[0]))

    return ":".join(sorted(function_set))

###########################################################################
def get_xlmfuncset(xlm_execution_report):
    """
    return an alpha-sorted set of XLM function calls
    """
    
    xlm_funcset = set()

    xlm_func_patterns = (
            r'---> ([A-Z\.]{1,25})\(',
            r'\(([A-Z\.]{1,25})\('
            )

    for xlm_func_pattern in xlm_func_patterns:
        results = re.findall(xlm_func_pattern, xlm_execution_report)
        for result in results:
            xlm_funcset.add(result)
    return ":".join(sorted(xlm_funcset))

###########################################################################
def get_xlmfuncset_frequency(xlm_execution_report):
    """
    return an frequency-sorted set of XLM function calls
    """

    xlm_funclist = list()

    xlm_func_patterns = (
            r'---> ([A-Z\.]{1,25})\(',
            r'\(([A-Z\.]{1,25})\('
            )

    for xlm_func_pattern in xlm_func_patterns:
        results = re.findall(xlm_func_pattern, xlm_execution_report)
        xlm_funclist = xlm_funclist + results

    result = [item for items, c in Counter(xlm_funclist).most_common()
                                      for item in [items] * c]

    xlmfuncset_freq = list()
    for func in result:
        if not func in xlmfuncset_freq:
            xlmfuncset_freq.append(func)
    return ":".join(xlmfuncset_freq)


###########################################################################
# Main Program
###########################################################################
if __name__ == '__main__':

    # Read command line arguments.
    help_msg = "Emulate the behavior of the XLM macros in a Excel file."
    parser = argparse.ArgumentParser(description=help_msg)
    parser.add_argument("maldocs", help="The Excel file to analyze.")
    parser.add_argument("-o", "--out-file", action="store", default=None,
                        help="JSON output file containing resulting actions and XLM macro code.")
    parser.add_argument('-d', '--debug',
                        action='store_true',
                        help="Print lots of debug information.")
    parser.add_argument('-q', '--quiet',
                        action='store_true',
                        help="Do not print progress or error messages.")
    args = parser.parse_args()

    # Disabling progress output?
    if (args.quiet):
        XLM.color_print.quiet = True
    
    # Emulate the XLM macros.
    if (not args.quiet):
        print("Emulating XLM macros in " + str(args.maldocs) + " ...")
    actions, xlm_code = emulate_XLM(args.maldocs, args.debug, args.out_file)
    if (not args.quiet):
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
    
    if (len(actions) > 0):
        print('\nFunction Sets:')
        print(str("\nDLL Function Calls: " + get_dllfuncset(actions)))
        print(str("\nAlpha sorted XLM Functions: " + get_xlmfuncset(xlm_code)))
        print(str("\nFrequency sorted XLM functions: " + get_xlmfuncset_frequency(xlm_code)))

    if (args.out_file is not None):
        XLM.color_print.output('g', "Saved analysis results to " + str(args.out_file) + " .")
