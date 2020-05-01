"""@package xlm_library
Emulation support for XLM functions.
"""

import XLM.XLM_Object
import XLM.stack_item

####################################################################
## Function emulators
####################################################################

func_lookup = {}

def _concat(params):
    r = ""
    for p in params:
        r += str(p)
    return r
func_lookup["_concat"] = _concat

def _plus(params):
    r = 0
    try:
        for p in params:
            r += int(p)
    except ValueError:
        # Just do string concat.
        r = ""
        for p in params:
            r += str(p)
    return r
func_lookup["_plus"] = _plus

def _minus(params):
    r = 0
    r = int(params[0]) - int(params[1])
    return r
func_lookup["_minus"] = _minus

def _less_than(params):
    r = False
    r = int(params[0]) < int(params[1])
    return r
func_lookup["_less_than"] = _less_than

def _not_equal(params):
    r = False
    r = int(params[0]) != int(params[1])
    return r
func_lookup["_not_equal"] = _not_equal

def _times(params):
    r = 0
    r = int(params[0]) * int(params[1])
    return r
func_lookup["_times"] = _times

def _equals(params):
    r = False
    r = int(params[0]) == int(params[1])
    return r
func_lookup["_equals"] = _equals

def _greater_than(params):
    r = False
    r = int(params[0]) > int(params[1])
    return r
func_lookup["_greater_than"] = _greater_than

def _divide(params):
    r = 0
    r = int(params[0]) / int(params[1])
    return r
func_lookup["_divide"] = _divide

def _unsigned_minus(params):
    r = 0
    # TODO: This is probably wrong.
    r = int(params[0]) - int(params[1])
    return r
func_lookup["_unsigned_minus"] = _unsigned_minus

def _greater_or_equal(params):
    r = False
    r = int(params[0]) >= int(params[1])
    return r
func_lookup["_greater_or_equal"] = _greater_or_equal

def CHAR(params):
    r = chr(int(params[0]))
    return r
func_lookup["CHAR"] = CHAR

def RUN(params):
    return "RUN"
func_lookup["RUN"] = RUN

def CONCATENATE(params):
    return _concat(params)
func_lookup["CONCATENATE"] = CONCATENATE

def CALL(params):
    r = "ACTION: CALL(" + str(params) + ")"
    return r
func_lookup["CALL"] = CALL

def HALT(params):
    r = "ACTION: HALT"
    return r
func_lookup["HALT"] = HALT

def FORMULA(params):
    r = str(params[0])
    return r
func_lookup["FORMULA"] = FORMULA

def WORKBOOK_HIDE(params):
    return "WORKBOOK.HIDE"
func_lookup["WORKBOOK.HIDE"] = WORKBOOK_HIDE

def GOTO(params):
    return "GOTO"
func_lookup["GOTO"] = GOTO

def GET_WORKSPACE(params):
    return "GET.WORKSPACE"
func_lookup["GET.WORKSPACE"] = GET_WORKSPACE

def NOW(params):
    return 12345
func_lookup["NOW"] = NOW

def WAIT(params):
    return "WAIT"
func_lookup["WAIT"] = WAIT

def FOPEN(params):
    return "ACTION: FILE:FOPEN(" + str(params) + ")"
func_lookup["FOPEN"] = FOPEN

def FPOS(params):
    return "ACTION: FILE:FPOS(" + str(params) + ")"
func_lookup["FPOS"] = FPOS

def FREAD(params):
    return "ACTION: FILE:FREAD(" + str(params) + ")"
func_lookup["FREAD"] = FREAD

def FCLOSE(params):
    return "ACTION: FILE:FCLOSE(" + str(params) + ")"
func_lookup["FCLOSE"] = FCLOSE

def FILE_DELETE(params):
    return "ACTION: FILE:FILE.DELETE(" + str(params) + ")"
func_lookup["FILE.DELETE"] = FILE_DELETE

def IF(params):
    return "IF"
func_lookup["IF"] = IF

def CLOSE(params):
    r = "ACTION: CLOSE"
    return r
func_lookup["CLOSE"] = CLOSE

def SEARCH(params):
    r = "SEARCH"
    return r
func_lookup["SEARCH"] = SEARCH

def ISNUMBER(params):
    r = "ISNUMBER"
    return r
func_lookup["ISNUMBER"] = ISNUMBER

def ALERT(params):
    return "ACTION: OUTPUT:ALERT(" + str(params) + ")"
func_lookup["ALERT"] = ALERT

####################################################################
def eval(func_name, params, sheet):
    """
    Emulate the behavior of a given XLM function.

    @param func_name (str) The name of the function.
    @param params (list) The mostly resolved parameters to pass to the function.
    @param sheet (ExcelSheet) The sheet containing the paramaters to fully resolve.
    
    @return The result of running the function.
    """

    # Do we know how to emulate this?
    if (func_name not in func_lookup.keys()):
        raise ValueError("Do not know how to emulate '" + str(func_name) + "'.")

    # Get the emulator for the function and run it.
    func = func_lookup[func_name]

    # FORMULA can compute a value and update a cell with that computed value. Track
    # the cell to update in this case.
    update_index = None
    if ((func_name == "FORMULA") and (len(params) > 1)):
        # $R54$C54
        update_fields = str(params[1]).replace("$C", ":").replace("$R", "").split(":")
        update_index = (int(update_fields[0]), int(update_fields[1]))
    
    # Evaluate the parameters.
    eval_params = []
    for p in params:
        if (hasattr(p, "eval")):
            eval_params.append(p.eval(sheet))
        else:
            eval_params.append(p)
    r = func(eval_params)

    # Does this value get written to a cell?
    if (update_index is not None):

        # r is something in the real XLM format. Need to parse real XLM to an XLM object.
        import XLM.ms_stack_transformer
        new_cell = XLM.ms_stack_transformer.parse_ms_xlm(str(r))
        new_cell.row = update_index[0]
        new_cell.col = update_index[1]
        sheet.cells[update_index] = new_cell
        if (update_index not in sheet.xlm_cell_indices):
            sheet.xlm_cell_indices.append(update_index)

    # Return the result.
    return r

    
