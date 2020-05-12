"""@package ms_stack_transformer

Lark AST transformer to generate XLM_Objects from an AST generated by ms_xlm.bnf.
"""

from __future__ import print_function
import string
import os
import sys

from lark import Lark
from lark import UnexpectedInput
from lark import Transformer

from XLM.stack_item import *
from XLM.XLM_Object import *

import XLM.color_print

# Load the MS XLM grammar.
ms_xlm_grammar_file = os.path.join(os.path.abspath(os.path.dirname(__file__)), "ms_xlm.bnf")
ms_xlm_grammar = None
try:
    f = open(ms_xlm_grammar_file, "r")
    ms_xlm_grammar = f.read()
    f.close()
except IOError as e:
    XLM.color_print.output('r', "ERROR: Cannot read MS XLM grammar file " + ms_xlm_grammar_file + ". " + str(e))
    sys.exit(102)

# Debugging flag.
debug = True
#debug = False
    
####################################################################
def parse_ms_xlm(expression):
    """
    Parse a XLM expression in the real MS XLM format (not plugin_biff) to an
    XLM_Object.

    @param expression (str) The MS XLM to parse.

    @return (XLM_Object) On success return the parsed XLM object, On failure
    return None.
    """

    # Convert to MS XLM is needed.
    orig_expression = expression
    if (not expression.startswith("=")):
        expression = '="' + str(expression) + '"'
    
    # Parse the given MS XLM.
    xlm_parser = Lark(ms_xlm_grammar, start="start", parser='lalr')
    xlm_ast = None
    try:
        xlm_ast = xlm_parser.parse(expression)
    except UnexpectedInput as e:

        # Maybe there is a problem with nested double quotes.
        expression = "=" + str(orig_expression)
        try:
            xlm_ast = xlm_parser.parse(expression)
        except UnexpectedInput as e:

            # Parsing failed. Just return this as a string.
            XLM.color_print.output('r', "ERROR: Cannot parse MS XLM expression '" + orig_expression + "'. " + str(e))
            tmp_str = stack_str(orig_expression)
            r = XLM_Object(-1, -1, [tmp_str])
            return r

    # Convert the AST to a XLM_Object.
    r = MsStackTransformer().transform(xlm_ast)

    # If we did not get a XLM_Object (just a stack_item), make an XLM_Object with the
    # single stack item on the stack.
    if (isinstance(r, stack_item)):
        stack = [r]
        r = XLM_Object(-1, -1, stack)
    if (isinstance(r, list)):
        stack = _load_stack_args([r], [])
        r = XLM_Object(-1, -1, stack)
        
    # Done.
    return r
    
####################################################################
def _load_stack_args(args, stack):
    """
    Load the arguments of a function call onto the given stack.

    @param args (list) stack_items for the function arguments.
    @param stack (list) The stack on which to push the call arguments.

    @return (list) Return the updated stack.
    """

    # Process each argument.
    while (len(args) > 0):
    
        # Pull an arg.
        item = args[0]
        args = args[1:]
        
        # If this is an XLM object just push the stack of the XLM object.
        if (isinstance(item, XLM_Object)):
            for s in item.stack:
                stack.append(s)
            continue
        
        # Simple case is a literal.
        if (not isinstance(item, list)):
            stack.append(item)
            continue

        # Single item as 2nd argument to an infix operator.
        if (len(item) == 1):
            stack.append(item[0])
            continue
            
        # 2nd item (infix) in the list is the operator.
        op = item[1]
        first_arg = item[0]
        second_arg = item[2:]
        if (len(second_arg) == 1):
            second_arg = second_arg[0]
            
        # Nested expression involving another operator with args.
        stack = _load_stack([op, [first_arg, second_arg]], stack)

    # Return the updated stack.
    return stack

####################################################################
def _load_stack(items, stack):
    """
    Work through a function call and its arguments and load them as stack_item
    objects onto a stack.

    @param items (list) The items making up the CALL.
    @param stack (list) The stack onto which to push the items.

    @return (list) A stack of stack_item objects representing the CALL.
    """

    # ['ALERT', ["The workbook cannot be opened or repaired by Microsoft Excel because it's corrupt.", 2, [1, +, [3, *, 2], +, [[2, -, 5], *, 4]], ["ff", &, "dd", &, "ss"]]]
    # ['IF', [[GET.WORKSPACE(13), <, 770], CLOSE(FALSE), ]]

    # The function args are a list in the 2nd element.
    func_args = items[1]

    # The name of the function being called is the 1st argument.
    func_name = items[0]
    num_args = len(func_args)
    if (not isinstance(func_name, stack_item)):
        top_func = stack_func_var(func_name, num_args)
    else:
        top_func = func_name

    # Push the arguments on the stack.
    stack = _load_stack_args(func_args, stack)

    # Call goes on the top of the stack.
    stack.append(top_func)

    # Done.
    return stack
    
####################################################################
class MsStackTransformer(Transformer):
    """
    Lark AST transformer to generate XLM_Objects from an AST generated by ms_xlm.bnf.
    """

    ##########################################################
    ## Non-terminal Transformers
    ##########################################################

    def start(self, items):
        return items[0]

    def function_call(self, items):

        # Make the function call stack.
        stack = []
        stack = _load_stack(items, stack)
        
        # Nested calls are represented as XLM_Objects by the transformer. Pull out
        # the stacks of these XLM objects and add them to the single stack for the
        # expression.
        new_stack = []
        for s in stack:
            if (isinstance(s, XLM_Object)):
                for s1 in s.stack:
                    new_stack.append(s1)
            else:
                new_stack.append(s)
        stack = new_stack
        
        # Create the XLM object for the call. We don't know the column and row at this point.
        xlm_object = XLM_Object(-1, -1, stack)
        
        return xlm_object
    
    def method_call(self, items):
        new_items = [items[0] + "." + items[1], items[2]]
        
        # Make the function call stack.
        stack = []
        stack = _load_stack(new_items, stack)
        
        # Create the XLM object for the call. We don't know the column and row at this point.
        xlm_object = XLM_Object(-1, -1, stack)
        return xlm_object
    
    def arglist(self, items):
        return items
    
    def argument(self, items):
        if (len(items) > 0):
            return items[0]
        return stack_str("")
    
    def cell(self, items):
        return items
    
    def a1_notation_cell(self, items):
        if (isinstance(items, list)):
            return items[0]
        return items
    
    def r1c1_notation_cell(self, items):
        row = 0
        col = 0
        if (len(items) >= 2):
            row = int(items[1])
        if (len(items) >= 4):
            col = int(items[3])
        r = stack_cell_ref(row, col)
        return r

    def expression(self, items):
        return items
    
    def concat_expression(self, items):
        return items
    
    def additive_expression(self, items):
        return items
    
    def multiplicative_expression(self, items):
        return items
    
    def final(self, items):
        return items
    
    def atom(self, items):
        return items

    ##########################################################
    ## Terminal Transformers
    ##########################################################
    
    def ADDITIVEOP(self, items):
        op = items[0]
        if (op == "+"):
            return stack_add()
        if (op == "-"):
            return stack_sub()
        raise ValueError("Unkown operator '" + str(op) + "'.")

    def MULTIOP(self, items):
        op = items[0]
        if (op == "*"):
            return stack_mul()
        if (op == "/"):
            return stack_div()
        raise ValueError("Unkown operator '" + str(op) + "'.")
    
    def CMPOP(self, items):
        op = items[0]
        if (op == "<"):
            return stack_less_than()
        if (op == ">"):
            return stack_greater_than()
        if (op == "="):
            return stack_equal()
        if (op == "<>"):
            return stack_not_equal()
        raise ValueError("Unkown operator '" + str(op) + "'.")
    
    def CONCATOP(self, items):
        return stack_concat()
    
    def STRING(self, items):
        return stack_str(items[1:-1])
    
    def BOOLEAN(self, items):
        return stack_bool(str(items))
    
    def ROW(self, items):
        return items[0]
    
    def COL(self, items):
        return items[0]
        
    def REF(self, items):
        r = int(str(items)[1:-1])
        return r
    
    def NAME(self, items):
        return str(items)
    
    def SIGNED_INT(self, items):
        return int(str(items))
    
    def INT(self, items):
        return int(str(items))
    
    def DECIMAL(self, items):
        return stack_float(str(items))
    
    def SIGNED_DECIMAL(self, items):
        return float(str(items))
    
    def NUMBER(self, items):
        return stack_int(str(items))
    
    def DOLLAR_CELL_REF(self, items):
        fields = str(items)[1:].split("$")
        col = XLM.utils.excel_col_letter_to_index(fields[0])
        row = int(fields[1])
        r = stack_cell_ref(row, col)
        return r
