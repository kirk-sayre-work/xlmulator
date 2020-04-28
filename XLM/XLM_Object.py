"""@package XLM_Object

Class for representing a single XLM formula (1 cell).
"""

from stack_item import *

####################################################################
def _get_str(stack):
    """
    Get the string representation for a single function on the top of the stack.

    @param stack (list) The current stack
        
    @return (tuple) A 2 element tuple with the 1st element being the 
    string representation of the topmost stack item and the 2nd item being 
    the remaining stack.
    """

    # Sanity check. Explicit checks used to differentiate error cases.
    if (stack is None):
        raise ValueError("The stack is None.")
    if (len(stack) == 0):
        raise ValueError("The stack is empty.")

    # Get the current stack item. Make sure the original stack is not modified.
    tmp_stack = list(stack)
    curr_item = tmp_stack.pop()
        
    # If this is not a function there is nothing to do.
    if (not curr_item.is_function()):

        # Just convert the top stack item to a string and we are done.
        r = str(curr_item)
        return (r, tmp_stack)

    # We have a function.

    # Infix function? These always have 2 arguments.
    if (curr_item.is_infix_function()):

        # Sanity check.
        if (len(tmp_stack) < 2):
            raise ValueError("Infix operator '" + str(curr_item) + "' requires 2 arguments.")

        # Resolve the strings for the 2 function arguments.
        arg2, tmp_stack = _get_str(tmp_stack)
        arg1, tmp_stack = _get_str(tmp_stack)

        # Return the string for the function now that we have the arguments.
        r = str(arg1) + str(curr_item) + str(arg2)
        return (r, tmp_stack)

    # Non-infix function. These have a variable # of arguments.
    num_args = curr_item.get_num_args()

    # Sanity check.
    if (len(tmp_stack) < num_args):
        print(tmp_stack)
        raise ValueError("Operator '" + str(curr_item) + "' requires " + str(num_args) + " arguments.")

    # Resolve the strings for all the arguments.
    first = True
    args = ""
    for i in range(0,num_args):
        if (not first):
            args = "," + args
        first = False
        arg, tmp_stack = _get_str(tmp_stack)
        args = str(arg) + args

    # Return the string for the function.
    r = str(curr_item) + "(" + args + ")"
    return (r, tmp_stack)

####################################################################
class XLM_Object(object):
    """
    Class for representing a single XLM formula (1 cell).
    """

    ####################################################################
    def __init__(self, row, col, stack):
        """
        Create a XLM formula object.

        @param row (int) The row of the cell containing the formula.

        @param col (int) The column of the cell containing the formula.

        @param stack (list) List of stack_item objects representing the XLM formula elements on
        the evaluation stack.
        """
        self.row = row
        self.col = col
        self.stack = []
        for item in stack:
            self.stack.append(item)
        self.gloss = None
        
    ####################################################################
    def full_str(self):
        """
        A human readable version of this XLM formula.
        """

        # Have we already computed the string version?
        if (self.gloss is not None):
            return self.gloss

        # Work through the stack to compute the human readable string.
        self.gloss, _ = _get_str(self.stack)
        self.gloss = "$R" + str(self.row) + "$C" + str(self.col) + ":\t\t=" + self.gloss
        return self.gloss

    ####################################################################
    def raw_str(self):
        """
        A debug version of this XLM formula.
        """
        return "$R" + str(self.row) + "$C" + str(self.col) + ":\t\t" + str(self.stack)
        
    ####################################################################
    def __repr__(self):
        return self.full_str()
