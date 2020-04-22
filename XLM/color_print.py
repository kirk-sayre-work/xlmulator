"""@package color_print
Print colored text to stdout.
"""

from __future__ import print_function

###########################################################################
def safe_print(text):
    """
    Sometimes printing large strings when running in a Docker container triggers exceptions.
    This function just wraps a print in a try/except block to not crash when this happens.

    @param text (str) The text to print.
    """
    try:
        print(text)
    except Exception as e:
        msg = "ERROR: Printing text failed (len text = " + str(len(text)) + ". " + str(e)
        if (len(msg) > 100):
            msg = msg[:100]
        try:
            print(msg)
        except:
            # At this point output is so messed up we can't do anything.
            pass

# Used to colorize printed text.
colors = {
    'g' : '\033[92m',
    'y' : '\033[93m',
    'r' : '\033[91m'
}
ENDC = '\033[0m'

###########################################################################
def output(color, text):
    """
    Print colored text to stdout.

    color - (str) The color to use. 'g' = green, 'r' = red, 'y' = yellow.

    text - (str) The text to print.
    """

    # Is this a color we handle?
    color = str(color).lower()
    if (color not in colors):
        raise ValueError("Color '" + color + "' not known.")

    # Print the text with the color.
    safe_print(colors[color] + str(text) + ENDC)
