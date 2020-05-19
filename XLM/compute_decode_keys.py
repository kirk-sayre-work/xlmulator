import re
import sys

import XLM.color_print

var_map = {}

def _parse_cell_index(s):
    index_pat = "\$R(\d+)\$C(\d+)"
    index_info = re.findall(index_pat, s)[0]
    return (int(index_info[0]), int(index_info[1]))

def _get_compute_items(char_comp, sheet):

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
        if (not lhs in cell_ref_count):
            cell_ref_count[lhs] = 0
        cell_ref_count[lhs] += 1
        rhs = expr[2]
        if (not rhs in cell_ref_count):
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

def _extract_char_computations(xlm_cell, sheet):

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
    first_expr = _get_compute_items(first_char, sheet)
    for other_char in other_chars:
        tmp_info = _get_compute_items(other_char, sheet)
        if (tmp_info is not None):
            other_exprs.append(tmp_info)

    # Return the char expressions.
    return (first_expr, other_exprs)
    
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
        first_expr, curr_other_exprs = _extract_char_computations(xlm_cell, sheet)
        first_exprs.add(first_expr)
        other_exprs = other_exprs.union(set(curr_other_exprs))

    # Get the cell references that are variables to compute and cell references that
    # are data (constant values)
    compute_exprs = first_exprs.union(other_exprs)
    decode_key_cells, data_map = _get_vars_and_constants(compute_exprs, sheet)

    # Group the expressions that use the same decode key.
    grouped_exprs = {}
    for decode_key in decode_key_cells:
        if (decode_key not in grouped_exprs):
            grouped_exprs[decode_key] = set()
        for expr in compute_exprs:
            if ((decode_key == expr[0]) or (decode_key == expr[2])):
                grouped_exprs[decode_key].add(expr)

    print("GROUPED:")
    for e in grouped_exprs.keys():
        print(e)
        print(grouped_exprs[e])
            
    sys.exit(0)
        
