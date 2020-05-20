import re
import sys

from constraint import *

import XLM.color_print

var_map = {}

def _parse_cell_index(s):
    index_pat = "\$R(\d+)\$C(\d+)"
    index_info = re.findall(index_pat, s)[0]
    return (int(index_info[0]), int(index_info[1]))

def _get_compute_items(char_comp):

    # $R21846$C80-$R65031$C63
    comp_pat = "(\$R\d+\$C\d+)([\-\+\*/])(\$R\d+\$C\d+)"
    comp_info = re.findall(comp_pat, char_comp)
    if (len(comp_info) == 0):
        return None
    return comp_info[0]

def _get_vars_and_constants(compute_exprs, sheet):

    # Count up the # times each cell reference is used in an expression.
    cell_ref_count = {}
    for expr in compute_exprs:
        print(expr)
        lhs = expr[0]
        if (lhs not in cell_ref_count):
            cell_ref_count[lhs] = 0
        cell_ref_count[lhs] += 1
        rhs = expr[2]
        if (rhs not in cell_ref_count):
            cell_ref_count[rhs] = 0
        cell_ref_count[rhs] += 1

    # Decode keys tend to be referenced multiple times.
    decode_key_cells = set()
    data_cells = set()
    for cell_ref in cell_ref_count.keys():
        if (cell_ref_count[cell_ref] > 1):
            decode_key_cells.add(cell_ref)
        else:
            data_cells.add(cell_ref)

    print("DECODE KEYS:")
    print(decode_key_cells)
    print("DATA KEYS:")
    print(data_cells)

    # Read in the values of the data cells.
    print("DATA VALS:")
    data_map = {}
    for cell_ref in data_cells:
        cell_index = _parse_cell_index(cell_ref)
        data_map[cell_ref] = sheet.cell(cell_index[0], cell_index[1])
        print(cell_ref)
        print(data_map[cell_ref])

    # Done.
    return (decode_key_cells, data_map)

def _extract_char_computations(xlm_cell):

    # Pull out the arguments for each CHAR call.
    # $R57844$C104: ---> FORMULA.FILL(CHAR($R22336$C160-$R48103$C255)&CHAR($R19921$C185/$R49030$C240)&CHAR($R22336$C160*$R6843$C142) ... )
    xlm_str = str(xlm_cell)
    print("DECODE:")
    print(xlm_str)
    first_char_pat = "FORMULA(?:\.FILL)\(\s*CHAR\(([^&]+)\)&"
    other_char_pat = "CHAR\(([^&]+)\)&"
    tmp_strs = re.findall(first_char_pat, xlm_str)
    if (len(tmp_strs) == 0):
        return None
    first_char = tmp_strs[0]
    tmp_strs = re.findall(other_char_pat, xlm_str)
    other_chars = tmp_strs[1:]
    print(first_char)
    print(other_chars)

    # TODO: For now just handling the simple case of 'CELL_REF op CELL_REF'.
    other_exprs = []
    first_expr = _get_compute_items(first_char)
    for other_char in other_chars:
        tmp_info = _get_compute_items(other_char)
        if (tmp_info is not None):
            other_exprs.append(tmp_info)

    # Return the char expressions.
    return (first_expr, other_exprs)

def _group_by_decode_key(decode_key_cells, compute_exprs):

    # Group the expressions that use the same decode key.
    grouped_exprs = {}
    for decode_key in decode_key_cells:
        if (decode_key not in grouped_exprs):
            grouped_exprs[decode_key] = set()
        for expr in compute_exprs:
            if ((decode_key == expr[0]) or (decode_key == expr[2])):
                grouped_exprs[decode_key].add(expr)

    # Done.
    return grouped_exprs

def _gen_single_constraint(decode_key, exp, constraint_str, data_map):

    # TODO: Generalize this beyond simple expressions like 'a+1'.
    
    # Rename the decode key cell reference.
    lhs = exp[0]
    if (lhs == decode_key):
        lhs = "a"
    rhs = exp[2]
    if (rhs == decode_key):
        rhs = "a"

    # Resolve values of cell references.
    num_pat = "^\-?\d+(?:\.\d+)?$"
    range_hint = None
    if (lhs in data_map):

        # For later safety ensure that this is a number. This is done since the constraint functions
        # are going to be exec'ed later and we don't want malicious content getting exec'ed.
        tmp_lhs = data_map[lhs.strip()]
        if (re.match(num_pat, tmp_lhs) is None):
            XLM.color_print.output('y', "WARNING: CHAR() expression item " + str(tmp_lhs) + " is not a number. Not using.")
        else:
            lhs = tmp_lhs
            range_hint = XLM.utils.convert_num(lhs)
    if (rhs in data_map):

        # For later safety ensure that this is a number. This is done since the constraint functions
        # are going to be exec'ed later and we don't want malicious content getting exec'ed.
        tmp_rhs = data_map[rhs.strip()]
        if (re.match(num_pat, tmp_rhs) is None):
            XLM.color_print.output('y', "WARNING: CHAR() expression item " + str(tmp_rhs) + " is not a number. Not using.")
        else:
            rhs = tmp_rhs
            range_hint = XLM.utils.convert_num(rhs)
            
    # Make the expression string to which to apply a constraint.
    expr_str = lhs + exp[1] + rhs

    # Figure out the range of values the decode key can take, if possible.
    op = exp[1]
    key_range = None
    print("HINT:")
    print(range_hint)
    if (range_hint is not None):
        if (op == "+"):
            key_range = (int(32 - range_hint), int(126 - range_hint))
        if (op == "-"):
            key_range = (int(32 + range_hint), int(126 + range_hint))
    
    # Return the constraint Python expression and decode key range.
    return (constraint_str % expr_str, key_range)
    
def _gen_constraint_funcs(grouped_exprs, first_exprs, data_map):

    # Look through each decode key.
    r = {}
    for decode_key in grouped_exprs.keys():

        # Is this key used to decode an initial '=' in a FORMULA()?
        constraint_exp = ""
        first = True
        for curr_exp in first_exprs:
            if ((decode_key == curr_exp[0]) or (decode_key == curr_exp[2])):

                # Make a constraint that this value must decode to the ASCII for '='.
                first = False
                tmp = _gen_single_constraint(decode_key, curr_exp, "((%s) == 61)", data_map)
                constraint_exp += tmp[0]
                break

        # Add a constraint expression for each of the other expressions that use
        # this decode key. 
        key_range = None
        for curr_exp in grouped_exprs[decode_key]:

            # Add in the constraint check.
            # Also figure out the range of values to try for the decode key if possible.
            if (not first):
                constraint_exp += " and "
            first = False
            tmp = _gen_single_constraint(decode_key, curr_exp, "(32 <= (%s) <= 126)", data_map)
            constraint_exp += tmp[0]
            if key_range is None:
                key_range = tmp[1]

        # Generate the constraint function for this decode key.
        constraint_func = "def ascii_constraint(a):\n    return (%s)\n" % constraint_exp
        r[decode_key] = (constraint_func, key_range)

    # Done.
    return r

def _compute_decode_keys(constraint_funcs):

    # Loop through each decode key.
    r = {}
    for decode_key in constraint_funcs.keys():

        # Get the range of possible decode key values and constraint function for the key.
        constraint_func = constraint_funcs[decode_key][0]
        key_range = constraint_funcs[decode_key][1]

        # Exec the constraint function definition to define the constraint function.
        # We checked earlier that all input values used to create this function are
        # numbers so we won't have malicious side effects by exec'ing this code.
        exec(constraint_func)

        # Use the constraint solver to get possible CHAR() values.
        problem = Problem()
        problem.addVariable(decode_key, range(key_range[0], key_range[1]))
        problem.addConstraint(ascii_constraint)
        solution = problem.getSolutions()

        # Save the possible decode key values.
        r[decode_key] = solution

    # Done.
    return r

def resolve_char_keys(sheet):

    # Find all the dynamically created FORMULA() and FORMULA.FILL() cells in the
    # sheet.
    first_exprs = set()
    other_exprs = set()
    for cell_index in sheet.xlm_cell_indices:

        # Is the current cell a FORMULA cell?
        xlm_cell = None
        try:
            xlm_cell = sheet.cell(cell_index[0], cell_index[1])
        except KeyError:
            XLM.color_print.output('y', "WARNING: Cell " + str(cell_index) + " not found. Skipping.")
            continue
        cell_str = str(xlm_cell)
        if (not ((cell_str.startswith("FORMULA")) and ("CHAR" in cell_str))):
            continue

        # We have a formula cell. Pull out the raw expressions for computing each char.
        first_expr, curr_other_exprs = _extract_char_computations(xlm_cell)
        first_exprs.add(first_expr)
        other_exprs = other_exprs.union(set(curr_other_exprs))

    # Get the cell references that are variables to compute and cell references that
    # are data (constant values)
    compute_exprs = first_exprs.union(other_exprs)
    decode_key_cells, data_map = _get_vars_and_constants(compute_exprs, sheet)

    # Group the expressions that use the same decode key.
    grouped_exprs = _group_by_decode_key(decode_key_cells, compute_exprs)

    # Generate constraint functions for each decode key.
    constraint_funcs = _gen_constraint_funcs(grouped_exprs, first_exprs, data_map)
    print("CONSTRAINT FUNCS:")
    for k in constraint_funcs.keys():
        print(k)
        print(constraint_funcs[k][0])
        print(constraint_funcs[k][1])
        print("\n")

    # Compute the possible values for each decode key.
    poss_vals = _compute_decode_keys(constraint_funcs)
    print("POSSIBLE VALS:")
    for k in poss_vals.keys():
        print(k)
        print(poss_vals[k])
        print("\n")
    
    sys.exit(0)
    
