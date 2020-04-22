"""@package XLM

Top level Excel XLM macro emulator interface.
"""

import subprocess
import sys
import re

import color_print

## Check installation prerequisites.

# Make sure olevba is installed.
try:
    subprocess.check_output(["olevba", "-h"])
except Exception as e:
    print "ERROR: It looks like olevba is not installed. " + str(e) + "\n"
    sys.exit(101)

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
    r = ""
    xlm_pat = r"' \d\d\d\d     [ \d]{2} [^\n]+\n"
    for line in re.findall(xlm_pat, olevba_out):
        r += line
    return r
    
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

    print xlm_code
    
    # Done.
    return []
