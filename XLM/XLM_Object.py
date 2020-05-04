"""@package XLM_Object

Class for representing a single XLM formula (1 cell).
"""

import json

# https://github.com/kirk-sayre-work/office_dumper.git
import excel

from XLM.stack_item import *
import XLM.xlm_library
import XLM.color_print

####################################################################
def _eval_stack(stack, sheet, cell_stack):
    """
    Evaluate XLM stack items.

    @param stack (list) The stack to emulate.
    @param sheet (ExcelSheet object) The sheet containing the XLM cell with the given stack.
    @param cell_stack (list) Stack of cells being analyzed. Used to break infinite recursion.

    @return (tuple) A 2 element tuple where the 1st element is the fully emulated result value
    of the top stack item and the 2nd element is the updated stack.
    """

    # Sanity check.
    if (cell_stack is None):
        raise ValueError("Stack of cells being emulated is None.")
    if (stack is None):
        raise ValueError("The XLM cell stack is None.")
    if (len(stack) == 0):
        raise ValueError("The XLM cell stack is empty.")

    # Get the current stack item. Make sure the original stack is not modified.
    tmp_stack = list(stack)
    curr_item = tmp_stack.pop()

    if debug:
        print("===== START MID LEVEL EVAL " + str(curr_item) + " =======")
        print(curr_item)
        print(tmp_stack)

    # If this has already been ersolved to a constant we are done.
    if (not hasattr(curr_item, "is_function")):
        if debug:
            print("===== DONE MID LEVEL EVAL " + str(curr_item) + " =======")
        return (curr_item, tmp_stack)
        
    # If this is not a function there is nothing much to do.
    if (not curr_item.is_function()):

        # Just return the stack item if it is fully resolved.
        if (hasattr(curr_item, "eval")):
            curr_item = curr_item.eval(sheet)
        if debug:
            print("????????")
            print(curr_item)
            print(type(curr_item))

        # Is this a reference to a new XLM cell?
        if (isinstance(curr_item, XLM_Object)):

            # Yes, eval that cell.
            curr_item = _eval_cell(curr_item, sheet, cell_stack)

        # Done emulating this item?
        if ((not hasattr(curr_item, "is_function")) or (not curr_item.is_function())):
            if debug:
                print("===== DONE MID LEVEL EVAL " + str(curr_item) + " =======")
            return (curr_item, tmp_stack)

    # We have a function.

    # Sanity check.
    num_args = curr_item.get_num_args()
    if (len(tmp_stack) < num_args):
        print(tmp_stack)
        raise ValueError("Operator '" + str(curr_item) + "' requires " + str(num_args) + " arguments.")

    # If we are currently looking at a 2 element FORMLUA function we will need to tweak the top
    # argument on the stack. This is actually the destination to where to write the formula value,
    # so we don't want to read the current value of that cell and pass that as an argument to
    # FORMULA.
    if ((curr_item.name == "FORMULA") and (num_args == 2)):
        ref_item = tmp_stack.pop()
        tmp_stack.append(str(ref_item))
    
    # Resolve all the arguments.
    args = []
    for i in range(0, num_args):
        arg, tmp_stack = _eval_stack(tmp_stack, sheet, cell_stack)
        args.insert(0, arg)

    # Evaluate the function.
    r = XLM.xlm_library.eval(curr_item.name, args, sheet)
    if debug:
        print(r)
        print(tmp_stack)
        print("===== DONE MID LEVEL EVAL " + str(curr_item) + " =======")
    return (r, tmp_stack)
    
####################################################################
def _eval_cell(xlm_cell, sheet, cell_stack):
    """
    Emulate the behavior of a single XLM cell.

    @param xlm_cell (XLM_Object object) The XLM cell to emulate.

    @param sheet (ExcelSheet object) The sheet containing the cell. Intermediate
    XLM cell values will be updated in this object.

    @param cell_stack (list) Stack of cells being analyzed. Used to break infinite recursion.

    @return (str) The final value of the XLM function.
    """

    # Are we getting into infinite recursion?
    if (xlm_cell in cell_stack):
        msg = "WARNING: Infinite recursion detected when resolving '" + xlm_cell.cell_id + ":" + str(xlm_cell) + "'."
        XLM.color_print.output('y', msg)
        return 0
    
    # Evaluate the XLM stack for the cell.
    if (not hasattr(xlm_cell, "stack")):
        return str(xlm_cell)
    stack = xlm_cell.stack
    if debug:
        print("==== START TOP LEVEL EVAL " + str(xlm_cell) + " ======")
        print(xlm_cell)
        print(stack)
    cell_stack.append(xlm_cell)
    final_val, _ = _eval_stack(stack, sheet, cell_stack)
    if debug:
        print("==== DONE TOP LEVEL EVAL " + str(xlm_cell) + " ======")

    # Done.
    cell_stack.pop()
    return str(final_val)

####################################################################
def _pull_actions(sheet):
    """
    Pull the actions from the given sheet of resolved XLM.

    @param sheet (ExcelSheet object) The resolved Excel sheet.

    @return (list) A list of 3 element tuples where the 1st element is the general 
    action type, the 2nd element is details of the action, and the 3rd element is a 
    general note.
    """

    # Cycle through the XLM cells in numeric order.
    indices = sheet.xlm_cell_indices
    indices.sort()
    r = []
    for index in indices:

        # Get the current cell value.
        curr_val = str(sheet.cell(index[0], index[1]))

        # Is this an action?
        # 'ACTION: CALL(['URLDownloadToFileA', 0, 'foo', 0, 'http:/bar.com', 'C:\\ProgramData\\junk', 0, 0])'
        if (curr_val.startswith("ACTION: ")):

            # Break out the action.
            curr_val = curr_val.replace("ACTION: ", "")

            # Function call?
            if (curr_val.startswith("CALL(")):

                # Pull out the call name and args.
                tmp = curr_val.replace("CALL(", "")[:-1].replace("'", '"')
                fields = json.loads(tmp)
                dll_name = fields[0]
                func_name = fields[1]
                if (isinstance(func_name, int)):
                    func_name = fields[0]
                    dll_name = ""
                call = func_name + "("
                first = True
                for arg in fields[3:]:
                    if (not first):
                        call += ", "
                    first = False
                    call += str(arg)
                call += ")"

                # Save the action.
                dll_info = "From DLL '" + dll_name + "'"
                if (len(dll_name) == 0):
                    dll_info = ""
                r.append(("CALL", call, dll_info))

            # Halt?
            if ((curr_val.startswith("HALT")) or (curr_val.startswith("CLOSE"))):
                r.append(("HALT", curr_val, "Done."))

            # File action?
            if (curr_val.startswith("FILE:")):
                curr_val = curr_val.replace("FILE:", "")
                call = curr_val
                func_name = curr_val[:curr_val.index("(")]
                r.append(("FILE", call, func_name))

            # User output?
            if (curr_val.startswith("OUTPUT:")):
                curr_val = curr_val.replace("OUTPUT:", "")
                call = curr_val
                func_name = curr_val[:curr_val.index("(")]
                r.append(("OUTPUT", call, func_name))

    # Done.
    return r
    
####################################################################
def eval(sheet):
    """
    Emulate the XLM behavior of an Excel sheet.

    @param sheet (ExcelSheet object) The workbook to emulate.

    @return (tuple) 1st element is a list of 3 element tuples containing the actions performed
    by the sheet, 2nd element is the human readable XLM code.
    """

    # Sanity check.
    if (not isinstance(sheet, excel.ExcelSheet)):
        raise ValueError("sheet arg is a '" + str(type(sheet)) + "', not a 'ExcelSheet'.")
    if (not hasattr(sheet, "xlm_cell_indices")):
        raise ValueError("sheet arg does not have 'xlm_cell_indices' field.")

    # Emulate the XLM cells.
    #
    # XLM macros are fancy Excel formulas, so we should be able to ignore the execution
    # flow of the XLM macros and just evaluate every XLM cell in an arbitrary order. This
    # assumes that the XLM macros do not modify the program state, where program state is tracked
    # by updating constant values stored in cells.

    # We will be updating cells values as we emulate the XLM. Make a copy of the sheet so
    # we don't stomp on the original sheet.
    result_sheet = excel.ExcelSheet(sheet)
    result_sheet.xlm_cell_indices = sheet.xlm_cell_indices

    # Evaluate all the FORMULA() cells first since they can modify cell values.
    done_cells = set()
    xlm_code = ""
    result_sheet.xlm_cell_indices.sort()
    for cell_index in result_sheet.xlm_cell_indices:

        # Get the XLM cell (XLM_Object) to emulate.
        xlm_cell = result_sheet.cell(cell_index[0], cell_index[1])
        xlm_code += xlm_cell.cell_id + " ---> " + str(xlm_cell) + "\n"
        
        # Is this a FORMULA() cell?
        if ("FORMULA(" not in str(xlm_cell)):
            continue

        # This is a FORMULA() cell. Evaluate it.
        resolved_cell = _eval_cell(xlm_cell, result_sheet, [])
        result_sheet.cells[cell_index] = resolved_cell
        done_cells.add(str(xlm_cell))
    
    # Cycle through each remaining XLM cell.
    result_sheet.xlm_cell_indices.sort()
    for cell_index in result_sheet.xlm_cell_indices:

        # Get the XLM cell (XLM_Object) to emulate.
        xlm_cell = result_sheet.cell(cell_index[0], cell_index[1])
        if (str(xlm_cell) in done_cells):
            continue
        resolved_cell = _eval_cell(xlm_cell, result_sheet, [])
        if debug:
            print("-------")
            print(xlm_cell)
            print(resolved_cell)
        result_sheet.cells[cell_index] = resolved_cell
        
    if debug:
        print("------- FINAL SHEET --------")
        print(result_sheet)

    # Pull the actions from the resolved XLM cells.
    r = _pull_actions(result_sheet)

    # Done.
    return (r, xlm_code)
        
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
        self.cell_id = "$R" + str(self.row) + "$C" + str(self.col) + ":"
        for item in stack:
            self.stack.append(item)
        self.gloss = None

    ####################################################################
    def update_cell_id(self, new_id):
        """
        Change the row and column of the cell.

        @param new_id (tuple) A 2 element tuple of the from (row, column).
        """
        self.row = new_id[0]
        self.col = new_id[1]
        self.cell_id = "$R" + str(self.row) + "$C" + str(self.col) + ":"
        
    ####################################################################
    def is_function(self):
        """
        Determine if this XLM formula is a function call.

        @return (boolean) True if it is a call, False if not.
        """
        if (len(self.stack) == 0):
            return False
        return self.stack[-1].is_function()
        
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
