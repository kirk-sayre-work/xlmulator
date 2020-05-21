import re
import sys
import itertools

from constraint import *
import sympy

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

    #print("DECODE KEYS:")
    #print(decode_key_cells)
    #print("DATA KEYS:")
    #print(data_cells)

    # Read in the values of the data cells.
    #print("DATA VALS:")
    data_map = {}
    for cell_ref in data_cells:
        cell_index = _parse_cell_index(cell_ref)
        data_map[cell_ref] = sheet.cell(cell_index[0], cell_index[1])
        #print(cell_ref)
        #print(data_map[cell_ref])

    # Done.
    return (decode_key_cells, data_map)

def _extract_char_computations(xlm_cell):

    # Pull out the arguments for each CHAR call.
    # $R57844$C104: ---> FORMULA.FILL(CHAR($R22336$C160-$R48103$C255)&CHAR($R19921$C185/$R49030$C240)&CHAR($R22336$C160*$R6843$C142) ... )
    xlm_str = str(xlm_cell)
    first_char_pat = "FORMULA(?:\.FILL)\(\s*CHAR\(([^&]+)\)&"
    other_char_pat = "CHAR\(([^&]+)\)&"
    tmp_strs = re.findall(first_char_pat, xlm_str)
    if (len(tmp_strs) == 0):
        return None
    first_char = tmp_strs[0]
    tmp_strs = re.findall(other_char_pat, xlm_str)
    other_chars = tmp_strs[1:]

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

def _gen_single_constraint(decode_key, exp, constraint_str, data_map, int_vals):

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
    if int_vals:
        expr_str = "int(" + lhs + exp[1] + rhs + ")"
    else:
        expr_str = "round(" + lhs + exp[1] + rhs + ")"
    print(expr_str)
        
    # Figure out the range of values the decode key can take, if possible.
    op = exp[1]
    key_range = None
    #print("HINT:")
    #print(range_hint)
    if (range_hint is not None):
        if (op == "+"):
            key_range = (int(32 - range_hint), int(126 - range_hint))
        if (op == "-"):
            key_range = (int(32 + range_hint), int(126 + range_hint))

    # Figure out if this decode key should be a float.
    float_decode_keys = set()
    if ((((lhs != "a") and isinstance(XLM.utils.convert_num(lhs), float)) or
         ((rhs != "a") and isinstance(XLM.utils.convert_num(rhs), float))) and
        ((exp[1] == "*") or (exp[1] == "/"))):
        float_decode_keys.add((decode_key, (XLM.utils.convert_num(lhs, no_error=True),
                                            exp[1],
                                            XLM.utils.convert_num(rhs, no_error=True))))
            
    # Return the constraint Python expression and decode key range.
    return (float_decode_keys, (constraint_str % expr_str, key_range))
    
def _gen_constraint_funcs(grouped_exprs, first_exprs, data_map):

    # Look through each decode key.
    r = {}
    for decode_key in grouped_exprs.keys():

        # Is this key used to decode an initial '=' in a FORMULA()?
        constraint_exp = ""
        first = True
        float_decode_keys = set()
        for curr_exp in first_exprs:
            if ((decode_key == curr_exp[0]) or (decode_key == curr_exp[2])):

                # Make a constraint that this value must decode to the ASCII for '='.
                first = False
                curr_float_decode_keys, tmp = _gen_single_constraint(decode_key, curr_exp, "((%s) == 61)", data_map, True)
                float_decode_keys = float_decode_keys.union(curr_float_decode_keys)
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
            curr_float_decode_keys, tmp = _gen_single_constraint(decode_key, curr_exp, "(32 <= (%s) <= 126)", data_map, False)
            float_decode_keys = float_decode_keys.union(curr_float_decode_keys)
            constraint_exp += tmp[0]
            if key_range is None:
                key_range = tmp[1]

        # Generate the constraint function for this decode key.
        constraint_func = "def ascii_constraint(a):\n    return (%s)\n" % constraint_exp
        r[decode_key] = (constraint_func, key_range)

    # Done.
    return (r, float_decode_keys)

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

def _get_equation(expr, decode_key, data_map):

    # Build up a new expression with constants resolved and the
    # decode key renamed to 'x'.
    new_expr = [None, expr[1], None]    
    if (expr[0] == decode_key):
        new_expr[0] = sympy.Symbol('x')
        new_expr[2] = sympy.sympify(data_map[expr[2]])
    if (expr[2] == decode_key):
        new_expr[2] = sympy.Symbol('x')
        new_expr[0] = sympy.sympify(data_map[expr[0]])
    #print("NEW EXPR:")
    #print(new_expr)

    # Build the sympy expression.
    sym_expr = None
    if (new_expr[1] == "/"):
        new_expr[2] = sympy.Pow(new_expr[2], -1)
        new_expr[1] = "*"
    if (new_expr[1] == "-"):
        new_expr[2] = sympy.Mul(new_expr[2], -1)
        new_expr[1] = "+"
    if (new_expr[1] == "*"):
        sym_expr = sympy.Mul(new_expr[0], new_expr[2])
    if (new_expr[1] == "+"):
        sym_expr = sympy.Add(new_expr[0], new_expr[2])

    # Build the equation to solve. We want this to be equal to the ASCII
    # code for '='.
    r = sympy.Eq(sym_expr, 61)
    #print("SYMPY EQ:")
    #print(r)
    #print("")
    return r

def _add_slop(poss_vals, float_decode_keys):

    # Pick out the expressions we will use for each decode key to introduce slop.
    slop_exprs = {}
    for float_info in float_decode_keys:
        decode_key = float_info[0]
        float_expr = float_info[1]
        if (decode_key not in slop_exprs):
            slop_exprs[decode_key] = float_expr
    
    # Add possible values for float decode keys so they will resolve 1 higher and
    # 1 lower.
    r = {}
    for decode_key in poss_vals.keys():

        # Do we need to add slop for this key?
        curr_vals = poss_vals[decode_key]
        if (decode_key not in slop_exprs):
            r[decode_key] = curr_vals
            continue

        # Get the float value from the expression.
        float_expr = slop_exprs[decode_key]
        float_val = None
        if (isinstance(float_expr[0], float)):
            float_val = float_expr[0]
        else:
            float_val = float_expr[2]

        # Compute the change offset.
        slop_val = None
        if (float_expr[1] == "*"):
            slop_val = 1/float_val
        if (float_expr[1] == "/"):
            slop_val = float_val
        
        # Get new wider range of values.
        new_vals = set()
        for v in curr_vals:
            curr_val = v[decode_key]
            new_vals.add(curr_val)
            new_vals.add(curr_val + slop_val)
            new_vals.add(curr_val - slop_val)

        # Package up the expanded range of decode key values.
        final_new_vals = []
        for v in new_vals:
            final_new_vals.append({decode_key : v})
        r[decode_key] = final_new_vals

    # Done.
    return r

def _handle_missing_keys(poss_vals, first_exprs, data_map):

    # Look for decode keys with no possible values.
    for decode_key in poss_vals.keys():

        # Got values?
        if (len(poss_vals[decode_key]) > 0):
            continue

        # This one has 0 values. Is it a decode key that is supposed to resolve
        # a CHAR() to '='?
        expr = None
        for curr_expr in first_exprs:
            if ((decode_key == curr_expr[0]) or (decode_key == curr_expr[2])):
                expr = curr_expr
                break

        # Do we have an expression we can use to manually compute the decode key?
        if (expr is None):
            XLM.color_print.output('r', "ERROR: Cannot compute CHAR() decode key " + decode_key + ".")
            return None

        # Convert the decode key expression to a sympy equation.
        eq = _get_equation(expr, decode_key, data_map)

        # Solve for the decode key.
        key_val = sympy.solve(eq)[0]

        # Save the decode key.
        poss_vals[decode_key] = [{decode_key : key_val}]

    # Return the updated decode keys.
    return poss_vals

def _get_all_decode_key_combos(cell_index, formula_cells, poss_vals):

    # Pull out the CHAR() expressions for the cell.
    char_exprs = formula_cells[cell_index]

    # Get the possible values for each decode key used in the current
    # dynamic formula.
    key_val_list = []
    added_keys = set()
    for curr_expr in char_exprs:

        # Figure out the decode key.
        decode_key = None
        if (curr_expr[0] in poss_vals):
            decode_key = curr_expr[0]
        if (curr_expr[2] in poss_vals):
            decode_key = curr_expr[2]

        # Save all the possible values for this key.
        if (decode_key not in added_keys):
            curr_vals = []
            for val in poss_vals[decode_key]:
                curr_vals.append((decode_key, val[decode_key]))
            key_val_list.append(curr_vals)
            added_keys.add(decode_key)

    # Get all possible combinations of decode keys.
    r = []
    print("ALL COMBOS RAW:")
    for key_combo in itertools.product(*key_val_list):
        print(key_combo)
        curr_key_map = {}
        for key_pair in key_combo:
            curr_key_map[key_pair[0]] = key_pair[1]
        r.append(curr_key_map)
            
    # Done
    print("ALL COMBOS MAP:")
    print(r)
    return r

def _resolve_char_expr(char_expr, sheet, decode_keys):

    # Resolve decode keys and data cell references.
    r = [None, char_expr[1], None]
    if (char_expr[0] in decode_keys):
        r[0] = decode_keys[char_expr[0]]
        cell_index = _parse_cell_index(char_expr[2])
        r[2] = XLM.utils.convert_num(sheet.cell(cell_index[0], cell_index[1]))
    if (char_expr[2] in decode_keys):
        r[2] = decode_keys[char_expr[2]]
        cell_index = _parse_cell_index(char_expr[0])
        r[0] = XLM.utils.convert_num(sheet.cell(cell_index[0], cell_index[1]))

    return r

def _resolve_formulas(sheet, formula_cells, poss_vals):

    # Resolve each formula cell.
    for cell_index in formula_cells.keys():

        # Get all possible combinations of decode keys.
        key_combos = _get_all_decode_key_combos(cell_index, formula_cells, poss_vals)

        # Get the formula as a string.
        formula_cell = sheet.cell(cell_index[0], cell_index[1])
        formula_str = str(formula_cell)
        print("\nFORMULA CELL:")
        print(formula_str)

        # Try all the decode key combos.
        print("NUM COMBOS: " + str(len(key_combos)))
        for key_combo in key_combos:
            print("KEY COMBO:")
            print(key_combo)

            # Get raw CHAR() argument expressions.
            first_expr, char_exprs = _extract_char_computations(formula_cell)
            char_exprs.append(first_expr)

            # Resolve the CHAR() expressions in the formula string.
            curr_formula_str = formula_str
            for char_expr in char_exprs:

                # Resolve cell references in the expression
                resolved_char_expr = _resolve_char_expr(char_expr, sheet, key_combo)

                # Compute the CHAR value.
                char_ord = None
                if (resolved_char_expr[1] == "+"):
                    char_ord = resolved_char_expr[0] + resolved_char_expr[2]
                if (resolved_char_expr[1] == "-"):
                    char_ord = resolved_char_expr[0] - resolved_char_expr[2]
                if (resolved_char_expr[1] == "*"):
                    char_ord = resolved_char_expr[0] * resolved_char_expr[2]
                if (resolved_char_expr[1] == "/"):
                    char_ord = resolved_char_expr[0] / resolved_char_expr[2]
                char_ord = int(round(char_ord))
                char = chr(char_ord)

                # Replace CHAR() in formula.
                old_char = "CHAR(" + char_expr[0] + char_expr[1] + char_expr[2] + ")"
                curr_formula_str = curr_formula_str.replace(old_char, char)
            curr_formula_str = curr_formula_str.replace("&", "")
            print(curr_formula_str)

def resolve_char_keys(sheet):

    # Find all the dynamically created FORMULA() and FORMULA.FILL() cells in the
    # sheet.
    first_exprs = set()
    other_exprs = set()
    formula_cells = {}
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
        formula_cells[cell_index] = [first_expr] + curr_other_exprs

    # Get the cell references that are variables to compute and cell references that
    # are data (constant values)
    compute_exprs = first_exprs.union(other_exprs)
    decode_key_cells, data_map = _get_vars_and_constants(compute_exprs, sheet)

    # Group the expressions that use the same decode key.
    grouped_exprs = _group_by_decode_key(decode_key_cells, compute_exprs)

    # Generate constraint functions for each decode key.
    constraint_funcs, float_decode_keys = _gen_constraint_funcs(grouped_exprs, first_exprs, data_map)
    print("FLOAT DECODE KEYS:")
    print(float_decode_keys)
    #print("CONSTRAINT FUNCS:")
    #for k in constraint_funcs.keys():
    #    print(k)
    #    print(constraint_funcs[k][0])
    #    print(constraint_funcs[k][1])
    #    print("\n")

    # Compute the possible values for each decode key.
    poss_vals = _compute_decode_keys(constraint_funcs)

    # Add some slop around the decode keys that should be floating point values.
    poss_vals = _add_slop(poss_vals, float_decode_keys)
    
    # Try to directly compute decode keys that we did not get values for from
    # the constraint solver.
    poss_vals = _handle_missing_keys(poss_vals, first_exprs, data_map)
    if (poss_vals is None):
        return None

    print("POSSIBLE VALS:")
    for k in poss_vals.keys():
        print(k)
        print(poss_vals[k])
        print("\n")

    # Resolve the formulas in each formula cell using the computed decode keys.
    _resolve_formulas(sheet, formula_cells, poss_vals)
        
    sys.exit(0)
    
