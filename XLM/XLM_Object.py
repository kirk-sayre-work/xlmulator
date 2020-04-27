"""@package XLM_Object

Class for representing a single XLM formula (1 cell).
"""

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
        self.stack = list(stack)

    ####################################################################
    def __repr__(self):
        return "$R" + str(self.row) + "$C" + str(self.col) + ":\t" + str(self.stack)
