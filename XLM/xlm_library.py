"""@package xlm_library
Emulation support for XLM functions.
"""

import random

import XLM.XLM_Object
import XLM.stack_item
import XLM.utils
import XLM.color_print

debug = False

####################################################################
## Function emulators
####################################################################

func_lookup = {}

def _concat(params, sheet):
    r = ""
    for p in params:
        r += str(p)
    return r
func_lookup["_concat"] = _concat

def _plus(params, sheet):
    r = 0
    try:
        for p in params:
            r += XLM.utils.convert_num(p)
    except ValueError:
        # Just do string concat.
        r = ""
        for p in params:
            r += str(p)
    return r
func_lookup["_plus"] = _plus

def _minus(params, sheet):
    r = XLM.utils.convert_num(params[0]) - XLM.utils.convert_num(params[1])
    return r
func_lookup["_minus"] = _minus

def _less_than(params, sheet):
    r = False
    r = XLM.utils.convert_num(params[0]) < XLM.utils.convert_num(params[1])
    return r
func_lookup["_less_than"] = _less_than

def _not_equal(params, sheet):
    r = False
    r = XLM.utils.convert_num(params[0]) != XLM.utils.convert_num(params[1])
    return r
func_lookup["_not_equal"] = _not_equal

def _times(params, sheet):
    r = 0
    r = XLM.utils.convert_num(params[0]) * XLM.utils.convert_num(params[1])
    return r
func_lookup["_times"] = _times

def _equals(params, sheet):
    r = False
    r = XLM.utils.convert_num(params[0]) == XLM.utils.convert_num(params[1])
    return r
func_lookup["_equals"] = _equals

def _greater_than(params, sheet):
    r = False
    r = XLM.utils.convert_num(params[0]) > XLM.utils.convert_num(params[1])
    return r
func_lookup["_greater_than"] = _greater_than

def _divide(params, sheet):
    r = 0
    r = XLM.utils.convert_num(params[0]) / XLM.utils.convert_num(params[1])
    return r
func_lookup["_divide"] = _divide

def _unsigned_minus(params, sheet):
    r = 0
    # TODO: This is probably wrong.
    r = XLM.utils.convert_num(params[0]) - XLM.utils.convert_num(params[1])
    return r
func_lookup["_unsigned_minus"] = _unsigned_minus

def _greater_or_equal(params, sheet):
    r = False
    r = XLM.utils.convert_num(params[0]) >= XLM.utils.convert_num(params[1])
    return r
func_lookup["_greater_or_equal"] = _greater_or_equal

def CHAR(params, sheet):
    try:
        r = chr(int(XLM.utils.convert_num(params[0])))
        return r
    except ValueError:
        XLM.color_print.output('y', "WARNING: Invalid CHAR() code " + str(XLM.utils.convert_num(params[0])) + ".")
        return "?"
func_lookup["CHAR"] = CHAR

def RUN(params, sheet):
    return "RUN"
func_lookup["RUN"] = RUN

def CONCATENATE(params, sheet):
    return _concat(params, sheet)
func_lookup["CONCATENATE"] = CONCATENATE

def CALL(params, sheet):
    r = "ACTION: CALL(" + str(params) + ")"
    return r
func_lookup["CALL"] = CALL

def EXEC(params, sheet):
    r = "ACTION: EXEC: " + str(params)
    return r
func_lookup["EXEC"] = EXEC

def HALT(params, sheet):
    r = "ACTION: HALT"
    return r
func_lookup["HALT"] = HALT

def FORMULA(params, sheet):
    r = str(params[0])
    return r
func_lookup["FORMULA"] = FORMULA

def WORKBOOK_HIDE(params, sheet):
    return "WORKBOOK.HIDE"
func_lookup["WORKBOOK.HIDE"] = WORKBOOK_HIDE

def GOTO(params, sheet):
    return "GOTO"
func_lookup["GOTO"] = GOTO

def GET_WORKSPACE(params, sheet):
    return "GET.WORKSPACE"
func_lookup["GET.WORKSPACE"] = GET_WORKSPACE

def NOW(params, sheet):
    return 12345
func_lookup["NOW"] = NOW

def WAIT(params, sheet):
    return "WAIT"
func_lookup["WAIT"] = WAIT

def FOPEN(params, sheet):
    return "ACTION: FILE:FOPEN(" + str(params) + ")"
func_lookup["FOPEN"] = FOPEN

def OPEN(params, sheet):
    return FOPEN(params, sheet)
func_lookup["OPEN"] = OPEN

def FPOS(params, sheet):
    return "ACTION: FILE:FPOS(" + str(params) + ")"
func_lookup["FPOS"] = FPOS

def FREAD(params, sheet):
    return "ACTION: FILE:FREAD(" + str(params) + ")"
func_lookup["FREAD"] = FREAD

def FCLOSE(params, sheet):
    return "ACTION: FILE:FCLOSE(" + str(params) + ")"
func_lookup["FCLOSE"] = FCLOSE

def FILE_CLOSE(params, sheet):
    return FCLOSE(params, sheet)
func_lookup["FILE.CLOSE"] = FILE_CLOSE

def FILE_DELETE(params, sheet):
    return "ACTION: FILE:FILE.DELETE(" + str(params) + ")"
func_lookup["FILE.DELETE"] = FILE_DELETE

def IF(params, sheet):
    return "IF"
func_lookup["IF"] = IF

def END_IF(params, sheet):
    return "END.IF"
func_lookup["END.IF"] = END_IF

def CLOSE(params, sheet):
    r = "ACTION: CLOSE"
    return r
func_lookup["CLOSE"] = CLOSE

def SEARCH(params, sheet):
    r = "SEARCH"
    return r
func_lookup["SEARCH"] = SEARCH

def ISNUMBER(params, sheet):
    r = "ISNUMBER"
    return r
func_lookup["ISNUMBER"] = ISNUMBER

def ALERT(params, sheet):
    return "ACTION: OUTPUT:ALERT(" + str(params) + ")"
func_lookup["ALERT"] = ALERT

def ALIGNMENT(params, sheet):
    r = "ALIGNMENT"
    return r
func_lookup["ALIGNMENT"] = ALIGNMENT

def ERROR(params, sheet):
    r = "ERROR"
    return r
func_lookup["ERROR"] = ERROR

def BORDER(params, sheet):
    r = "BORDER"
    return r
func_lookup["BORDER"] = BORDER

def WORKBOOK_SELECT(params, sheet):
    r = "WORKBOOK.SELECT"
    return r
func_lookup["WORKBOOK.SELECT"] = WORKBOOK_SELECT

def PATTERNS(params, sheet):
    r = "PATTERNS"
    return r
func_lookup["PATTERNS"] = PATTERNS

def WINDOW_RESTORE(params, sheet):
    r = "WINDOW.RESTORE"
    return r
func_lookup["WINDOW.RESTORE"] = WINDOW_RESTORE

def FORMAT_FONT(params, sheet):
    r = "FORMAT.FONT"
    return r
func_lookup["FORMAT.FONT"] = FORMAT_FONT

def WINDOW_SIZE(params, sheet):
    r = "WINDOW.SIZE"
    return r
func_lookup["WINDOW.SIZE"] = WINDOW_SIZE

def RETURN(params, sheet):
    r = "RETURN"
    return r
func_lookup["RETURN"] = RETURN

def EDIT_COLOR(params, sheet):
    r = "EDIT.COLOR"
    return r
func_lookup["EDIT.COLOR"] = EDIT_COLOR

def DELETE_NAME(params, sheet):
    r = "DELETE.NAME"
    return r
func_lookup["DELETE.NAME"] = DELETE_NAME

def SELECT(params, sheet):
    r = "SELECT"
    return r
func_lookup["SELECT"] = SELECT

def COLUMN_WIDTH(params, sheet):
    r = "COLUMN.WIDTH"
    return r
func_lookup["COLUMN.WIDTH"] = COLUMN_WIDTH

def ROW_HEIGHT(params, sheet):
    r = "ROW.HEIGHT"
    return r
func_lookup["ROW.HEIGHT"] = ROW_HEIGHT

def WINDOW_MAXIMIZE(params, sheet):
    r = "WINDOW.MAXIMIZE"
    return r
func_lookup["WINDOW.MAXIMIZE"] = WINDOW_MAXIMIZE

def FORMAT_NUMBER(params, sheet):
    r = "FORMAT.NUMBER"
    return r
func_lookup["FORMAT.NUMBER"] = FORMAT_NUMBER

def OFFSET(params, sheet):
    r = "OFFSET"
    return r
func_lookup["OFFSET"] = OFFSET

def WORKBOOK_UNHIDE(params, sheet):
    r = "WORKBOOK.UNHIDE"
    return r
func_lookup["WORKBOOK.UNHIDE"] = WORKBOOK_UNHIDE

def FILL_AUTO(params, sheet):
    r = "FILL.AUTO"
    return r
func_lookup["FILL.AUTO"] = FILL_AUTO

def SET_NAME(params, sheet):
    r = "SET.NAME"
    return r
func_lookup["SET.NAME"] = SET_NAME

def ERROR_TYPE(params, sheet):
    return 1.0
func_lookup["ERROR.TYPE"] = ERROR_TYPE

def DOCUMENTS(params, sheet):
    return ["workbook.xls"]
func_lookup["DOCUMENTS"] = DOCUMENTS

def MATCH(params, sheet):
    # TODO: Find some coherent documentation on what MATCH is supposed to do.
    #print("MATCH")
    #print(params)
    return("MATCH")
func_lookup["MATCH"] = MATCH

def GET_CELL(params, sheet):
    
    # This gets information about a cell. For now this will just return hardcoded
    # values.

    # 1st param is the lookup type, 2nd param is the cell reference.
    info_type = XLM.utils.convert_num(params[0])
    cell_ref = params[1]

    # Row height of cell, in points ?
    if (info_type == 17):
        return 16.5

    # Size of font, in points ?
    if (info_type == 19):
        return 9

    # Unhandled info type.
    XLM.color_print.output('y', "WARNING: GET.CELL() information type " + str(info_type) + " is not handled. Defaulting to 1.")
    return 1
    
func_lookup["GET.CELL"] = GET_CELL

def DAY(params, sheet):
    r = 2
    return r
func_lookup["DAY"] = DAY

def APP_MAXIMIZE(params, sheet):
    r = "APP.MAXIMIZE"
    return r
func_lookup["APP.MAXIMIZE"] = APP_MAXIMIZE

def RANDBETWEEN(params, sheet):
    start = XLM.utils.convert_num(params[0])
    end = XLM.utils.convert_num(params[1])
    return random.randrange(start, end)
func_lookup["RANDBETWEEN"] = RANDBETWEEN

def ISERROR(params, sheet):
    # All good.
    return False
func_lookup["ISERROR"] = ISERROR

def CLEAR(params, sheet):
    r = "CLEAR"
    return r
func_lookup["CLEAR"] = CLEAR

def ON_TIME(params, sheet):
    r = "ON.TIME"
    return r
func_lookup["ON.TIME"] = ON_TIME

def User_Defined_Function(params, sheet):
    r = "User Defined Function"
    return r
func_lookup["User Defined Function"] = User_Defined_Function

def LOWER(params, sheet):
    return str(params[0]).lower()
func_lookup["LOWER"] = LOWER

def SUMIF(params, sheet):
    return "SUMIF"
func_lookup["SUMIF"] = SUMIF

def SEND_KEYS(params, sheet):
    return "SEND.KEYS"
func_lookup["SEND.KEYS"] = SEND_KEYS

def APP_ACTIVATE(params, sheet):
    return "APP.ACTIVATE"
func_lookup["APP.ACTIVATE"] = APP_ACTIVATE

def FWRITELN(params, sheet):
    return "ACTION: FILE:FWRITELN(" + str(params) + ")"
func_lookup["FWRITELN"] = FWRITELN

def FILES(params, sheet):
    return "FILES"
func_lookup["FILES"] = FILES

def WHILE(params, sheet):
    return "WHILE"
func_lookup["WHILE"] = WHILE

def NEXT(params, sheet):
    return "NEXT"
func_lookup["NEXT"] = NEXT

def COPY(params, sheet):
    return "COPY"
func_lookup["COPY"] = COPY

def PASTE_SPECIAL(params, sheet):
    return "PASTE.SPECIAL"
func_lookup["PASTE.SPECIAL"] = PASTE_SPECIAL

def CANCEL_COPY(params, sheet):
    return "CANCEL.COPY"
func_lookup["CANCEL.COPY"] = CANCEL_COPY

"""
def (params, sheet):
    return ""
func_lookup[""] = 
"""

####################################################################
def _is_interesting_cell(cell):
    """
    See if a cell value looks interesting.

    @param cell (XLM_Object) The cell to check.

    @return (boolean) True if the cell value is interesting, False if not.
    """
    cell_str = str(cell).replace('"', '').strip()
    return (len(cell_str) > 0)

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
    r = func(eval_params, sheet)

    # Does this value get written to a cell?
    if (update_index is not None):

        # r is something in the real XLM format. Need to parse real XLM to an XLM object.
        import XLM.ms_stack_transformer
        new_cell = XLM.ms_stack_transformer.parse_ms_xlm(str(r))
        new_cell.update_cell_id(update_index)

        # Only do the update if it looks interesting.
        if (_is_interesting_cell(new_cell)):
            sheet.cells[update_index] = new_cell
            if debug:
                print("FORMULA:")
                print("'" + str(new_cell) + "'")
                print(new_cell.cell_id)

        # Track this if it is a new cell.
        if (update_index not in sheet.xlm_cell_indices):
            sheet.xlm_cell_indices.append(update_index)

    # Return the result.
    return r

    
