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

# Functions to always emulate.
funcs_of_interest = []

# Table mapping function names to emulator functions.
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
    num = XLM.utils.convert_num(params[0])
    den = XLM.utils.convert_num(params[1])
    if (den == 0):
        XLM.color_print.output('r', "ERROR: Division by zero.")
        # Not right, but better than crashing.
        return 0
    r = num / den
    return r
func_lookup["_divide"] = _divide

def _unsigned_minus(params, sheet):
    r = 0
    # TODO: This is probably wrong.
    r = XLM.utils.convert_num(params[0]) - XLM.utils.convert_num(params[1])
    return r
func_lookup["_unsigned_minus"] = _unsigned_minus

def _unsigned_plus(params, sheet):
    r = 0
    # TODO: This is probably wrong.
    r = XLM.utils.convert_num(params[0]) + XLM.utils.convert_num(params[1])
    return r
func_lookup["_unsigned_plus"] = _unsigned_plus

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
    # STUBBED
    return "RUN"
func_lookup["RUN"] = RUN

def CONCATENATE(params, sheet):
    return _concat(params, sheet)
func_lookup["CONCATENATE"] = CONCATENATE

def CALL(params, sheet):
    r = "ACTION: CALL(" + str(params)[1:-1] + ")"
    return r
func_lookup["CALL"] = CALL
funcs_of_interest.append("CALL")

def EXEC(params, sheet):
    r = "ACTION: EXEC: " + str(params)[1:-1]
    return r
func_lookup["EXEC"] = EXEC
funcs_of_interest.append("EXEC")

def HALT(params, sheet):
    r = "ACTION: HALT"
    return r
func_lookup["HALT"] = HALT
funcs_of_interest.append("HALT")

def FORMULA(params, sheet):
    r = str(params[0])
    return r
func_lookup["FORMULA"] = FORMULA
funcs_of_interest.append("FORMULA")

def WORKBOOK_HIDE(params, sheet):
    return "WORKBOOK.HIDE"
func_lookup["WORKBOOK.HIDE"] = WORKBOOK_HIDE

def GOTO(params, sheet):
    # STUBBED
    return "GOTO"
func_lookup["GOTO"] = GOTO

def GET_WORKSPACE(params, sheet):
    return "GET.WORKSPACE"
func_lookup["GET.WORKSPACE"] = GET_WORKSPACE

def NOW(params, sheet):
    return 12345
func_lookup["NOW"] = NOW

def WAIT(params, sheet):
    # STUBBED
    return "WAIT"
func_lookup["WAIT"] = WAIT

def FOPEN(params, sheet):
    return "ACTION: FILE:FOPEN(" + str(params)[1:-1] + ")"
func_lookup["FOPEN"] = FOPEN
funcs_of_interest.append("FOPEN")

def OPEN(params, sheet):
    return FOPEN(params, sheet)
func_lookup["OPEN"] = OPEN

def FPOS(params, sheet):
    return "ACTION: FILE:FPOS(" + str(params)[1:-1] + ")"
func_lookup["FPOS"] = FPOS
funcs_of_interest.append("FPOS")

def FREAD(params, sheet):
    return "ACTION: FILE:FREAD(" + str(params)[1:-1] + ")"
func_lookup["FREAD"] = FREAD
funcs_of_interest.append("FREAD")

def FCLOSE(params, sheet):
    return "ACTION: FILE:FCLOSE(" + str(params)[1:-1] + ")"
func_lookup["FCLOSE"] = FCLOSE
funcs_of_interest.append("FCLOSE")

def FILE_CLOSE(params, sheet):
    return FCLOSE(params, sheet)
func_lookup["FILE.CLOSE"] = FILE_CLOSE

def FILE_DELETE(params, sheet):
    return "ACTION: FILE:FILE.DELETE(" + str(params)[1:-1] + ")"
func_lookup["FILE.DELETE"] = FILE_DELETE
funcs_of_interest.append("FILE.DELETE")

def IF(params, sheet):
    # STUBBED
    return "IF"
func_lookup["IF"] = IF

def END_IF(params, sheet):
    # STUBBED
    return "END.IF"
func_lookup["END.IF"] = END_IF

def CLOSE(params, sheet):
    r = "ACTION: CLOSE"
    return r
func_lookup["CLOSE"] = CLOSE
funcs_of_interest.append("CLOSE")

def SEARCH(params, sheet):
    # STUBBED
    r = "SEARCH"
    return r
func_lookup["SEARCH"] = SEARCH

def ISNUMBER(params, sheet):
    # STUBBED
    r = "ISNUMBER"
    return r
func_lookup["ISNUMBER"] = ISNUMBER

def ALERT(params, sheet):
    return "ACTION: OUTPUT:ALERT(" + str(params)[1:-1] + ")"
func_lookup["ALERT"] = ALERT
funcs_of_interest.append("ALERT")

def ALIGNMENT(params, sheet):
    # STUBBED
    r = "ALIGNMENT"
    return r
func_lookup["ALIGNMENT"] = ALIGNMENT

def ERROR(params, sheet):
    # STUBBED
    r = "ERROR"
    return r
func_lookup["ERROR"] = ERROR

def BORDER(params, sheet):
    # STUBBED
    r = "BORDER"
    return r
func_lookup["BORDER"] = BORDER

def WORKBOOK_SELECT(params, sheet):
    # STUBBED
    r = "WORKBOOK.SELECT"
    return r
func_lookup["WORKBOOK.SELECT"] = WORKBOOK_SELECT

def PATTERNS(params, sheet):
    # STUBBED
    r = "PATTERNS"
    return r
func_lookup["PATTERNS"] = PATTERNS

def WINDOW_RESTORE(params, sheet):
    # STUBBED
    r = "WINDOW.RESTORE"
    return r
func_lookup["WINDOW.RESTORE"] = WINDOW_RESTORE

def FORMAT_FONT(params, sheet):
    # STUBBED
    r = "FORMAT.FONT"
    return r
func_lookup["FORMAT.FONT"] = FORMAT_FONT

def WINDOW_SIZE(params, sheet):
    # STUBBED
    r = "WINDOW.SIZE"
    return r
func_lookup["WINDOW.SIZE"] = WINDOW_SIZE

def RETURN(params, sheet):
    # STUBBED
    r = "RETURN"
    return r
func_lookup["RETURN"] = RETURN

def EDIT_COLOR(params, sheet):
    # STUBBED
    r = "EDIT.COLOR"
    return r
func_lookup["EDIT.COLOR"] = EDIT_COLOR

def DELETE_NAME(params, sheet):
    # STUBBED
    r = "DELETE.NAME"
    return r
func_lookup["DELETE.NAME"] = DELETE_NAME

def SELECT(params, sheet):
    # STUBBED
    r = "SELECT"
    return r
func_lookup["SELECT"] = SELECT

def COLUMN_WIDTH(params, sheet):
    # STUBBED
    r = "COLUMN.WIDTH"
    return r
func_lookup["COLUMN.WIDTH"] = COLUMN_WIDTH

def ROW_HEIGHT(params, sheet):
    # STUBBED
    r = "ROW.HEIGHT"
    return r
func_lookup["ROW.HEIGHT"] = ROW_HEIGHT

def WINDOW_MAXIMIZE(params, sheet):
    # STUBBED
    r = "WINDOW.MAXIMIZE"
    return r
func_lookup["WINDOW.MAXIMIZE"] = WINDOW_MAXIMIZE

def FORMAT_NUMBER(params, sheet):
    # STUBBED
    r = "FORMAT.NUMBER"
    return r
func_lookup["FORMAT.NUMBER"] = FORMAT_NUMBER

def OFFSET(params, sheet):
    # STUBBED
    r = "OFFSET"
    return r
func_lookup["OFFSET"] = OFFSET

def WORKBOOK_UNHIDE(params, sheet):
    # STUBBED
    r = "WORKBOOK.UNHIDE"
    return r
func_lookup["WORKBOOK.UNHIDE"] = WORKBOOK_UNHIDE

def FILL_AUTO(params, sheet):
    # STUBBED
    r = "FILL.AUTO"
    return r
func_lookup["FILL.AUTO"] = FILL_AUTO

def SET_NAME(params, sheet):
    # STUBBED
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
    # STUBBED
    return("MATCH")
func_lookup["MATCH"] = MATCH

def GET_CELL(params, sheet):
    
    # This gets information about a cell. For now this will just return hardcoded
    # values.

    # 1st param is the lookup type, 2nd param is the cell reference.
    info_type = XLM.utils.convert_num(params[0])
    cell_ref = params[1]

    # Cell horizontal alignment?
    if (info_type == 8):
        return 5
    
    # Row height of cell, in points ?
    if (info_type == 17):
        return 16.5

    # Size of font, in points ?
    if (info_type == 19):
        return 9

    # Font color of 1st char in cell?
    if (info_type == 24):
        return 13

    # Shade foreground color of cell?
    if (info_type == 38):
        return 15

    # Cell vertical alignment?
    if (info_type == 50):
        return 5
    
    # Unhandled info type.
    XLM.color_print.output('y', "WARNING: GET.CELL() information type " + str(info_type) + " is not handled. Defaulting to 1.")
    return 1    
func_lookup["GET.CELL"] = GET_CELL

def DAY(params, sheet):
    r = 4
    return r
func_lookup["DAY"] = DAY

def APP_MAXIMIZE(params, sheet):
    # STUBBED
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
    # STUBBED
    r = "CLEAR"
    return r
func_lookup["CLEAR"] = CLEAR

def ON_TIME(params, sheet):
    # STUBBED
    r = "ON.TIME"
    return r
func_lookup["ON.TIME"] = ON_TIME

def User_Defined_Function(params, sheet):
    # STUBBED
    r = "User Defined Function"
    return r
func_lookup["User Defined Function"] = User_Defined_Function

def LOWER(params, sheet):
    return str(params[0]).lower()
func_lookup["LOWER"] = LOWER

def SUMIF(params, sheet):
    # STUBBED
    return "SUMIF"
func_lookup["SUMIF"] = SUMIF

def SUM(params, sheet):
    # STUBBED
    return "SUM"
func_lookup["SUM"] = SUM

def SEND_KEYS(params, sheet):
    return "ACTION: INPUT:SEND.KEYS(" + str(params)[1:-1] + ")"
func_lookup["SEND.KEYS"] = SEND_KEYS
funcs_of_interest.append("SEND.KEYS")

def APP_ACTIVATE(params, sheet):
    # STUBBED
    return "APP.ACTIVATE"
func_lookup["APP.ACTIVATE"] = APP_ACTIVATE

def FWRITELN(params, sheet):
    return "ACTION: FILE:FWRITELN(" + str(params)[1:-1] + ")"
func_lookup["FWRITELN"] = FWRITELN
funcs_of_interest.append("FWRITELN")

def FILES(params, sheet):
    # STUBBED
    return "FILES"
func_lookup["FILES"] = FILES

def WHILE(params, sheet):
    # STUBBED
    return "WHILE"
func_lookup["WHILE"] = WHILE

def NEXT(params, sheet):
    # STUBBED
    return "NEXT"
func_lookup["NEXT"] = NEXT

def COPY(params, sheet):
    # STUBBED
    return "COPY"
func_lookup["COPY"] = COPY

def PASTE_SPECIAL(params, sheet):
    # STUBBED
    return "PASTE.SPECIAL"
func_lookup["PASTE.SPECIAL"] = PASTE_SPECIAL

def CANCEL_COPY(params, sheet):
    # STUBBED
    return "CANCEL.COPY"
func_lookup["CANCEL.COPY"] = CANCEL_COPY

def INDEX(params, sheet):
    # STUBBED
    return "INDEX"
func_lookup["INDEX"] = INDEX

def SET_VALUE(params, sheet):
    update_fields = str(params[0]).replace("$C", ":").replace("$R", "").split(":")
    update_index = (int(update_fields[0]), int(update_fields[1]))
    sheet.cells[update_index] = params[1]
    return "SET.VALUE"
func_lookup["SET.VALUE"] = SET_VALUE

def ON_SHEET(params, sheet):
    # STUBBED
    return "ON.SHEET"
func_lookup["ON.SHEET"] = ON_SHEET

def GET_DOCUMENT(params, sheet):
    # STUBBED
    return "GET.DOCUMENT"
func_lookup["GET.DOCUMENT"] = GET_DOCUMENT

def NEW(params, sheet):
    # STUBBED
    return "NEW"
func_lookup["NEW"] = NEW

def WORKBOOK_INSERT(params, sheet):
    # STUBBED
    return "WORKBOOK.INSERT"
func_lookup["WORKBOOK.INSERT"] = WORKBOOK_INSERT

def ACTIVATE_PREV(params, sheet):
    # STUBBED
    return "ACTIVATE.PREV"
func_lookup["ACTIVATE.PREV"] = ACTIVATE_PREV

def WORKBOOK_COPY(params, sheet):
    # STUBBED
    return "WORKBOOK.COPY"
func_lookup["WORKBOOK.COPY"] = WORKBOOK_COPY

def WORKBOOK_NAME(params, sheet):
    # STUBBED
    return "WORKBOOK.NAME"
func_lookup["WORKBOOK.NAME"] = WORKBOOK_NAME

def PROTECT_DOCUMENT(params, sheet):
    # STUBBED
    return "PROTECT.DOCUMENT"
func_lookup["PROTECT.DOCUMENT"] = PROTECT_DOCUMENT

def WORKBOOK_PREV(params, sheet):
    # STUBBED
    return "WORKBOOK.PREV"
func_lookup["WORKBOOK.PREV"] = WORKBOOK_PREV

def SAVE_AS(params, sheet):
    return "ACTION: FILE:SAVE.AS(" + str(params)[1:-1] + ")"
func_lookup["SAVE.AS"] = SAVE_AS
funcs_of_interest.append("SAVE.AS")

def APP_TITLE(params, sheet):
    return "ACTION: OUTPUT:APP.TITLE(" + str(params)[1:-1] + ")"
func_lookup["APP.TITLE"] = APP_TITLE
funcs_of_interest.append("APP.TITLE")

def MESSAGE(params, sheet):
    return "ACTION: OUTPUT:MESSAGE(" + str(params)[1:-1] + ")"
func_lookup["MESSAGE"] = MESSAGE
funcs_of_interest.append("MESSAGE")

def FORMULA_FILL(params, sheet):
    # STUBBED
    if debug:
        print("FORMULA.FILL!!")
        print(params)
    return "FORMULA.FILL"
func_lookup["FORMULA.FILL"] = FORMULA_FILL

def FOR_CELL(params, sheet):
    # STUBBED
    return "FOR.CELL"
func_lookup["FOR.CELL"] = FOR_CELL

def VBA_INSERT_FILE(params, sheet):
    return "ACTION: FILE:VBA.INSERT.FILE(" + str(params)[1:-1] + ")"
func_lookup["VBA.INSERT.FILE"] = VBA_INSERT_FILE
funcs_of_interest.append("VBA.INSERT.FILE")

def OR(params, sheet):
    r = False
    for p in params:
        if (isinstance(p, bool)):
            r = r or p
    return r
func_lookup["OR"] = OR

def NOT(params, sheet):
    if ((len(params) > 0) and (isinstance(params[0], bool))):
        return not params[0]
    XLM.color_print.output('y', "WARNING: Cannot perform NOT on '" + str(params) + "'. Empty or not a boolean.")
    return False
func_lookup["NOT"] = NOT

def GET_WORKBOOK(params, sheet):
    # STUBBED
    return "GET.WORKBOOK"
func_lookup["GET.WORKBOOK"] = GET_WORKBOOK

def REGISTER_ID(params, sheet):
    # STUBBED
    return "REGISTER.ID"
func_lookup["REGISTER.ID"] = REGISTER_ID

def ACTIVATE(params, sheet):
    # STUBBED
    return "ACTIVATE"
func_lookup["ACTIVATE"] = ACTIVATE

def ARGUMENT(params, sheet):
    # STUBBED
    return "ARGUMENT"
func_lookup["ARGUMENT"] = ARGUMENT

def ACTIVE_CELL(params, sheet):
    # STUBBED
    return "ACTIVE.CELL"
func_lookup["ACTIVE.CELL"] = ACTIVE_CELL

def LEN(params, sheet):
    return len(str(params)) - 2
func_lookup["LEN"] = LEN

def ELSE(params, sheet):
    # STUBBED
    return "ELSE"
func_lookup["ELSE"] = ELSE

def COUNTA(params, sheet):
    # STUBBED
    return "COUNTA"
func_lookup["COUNTA"] = COUNTA

def MID(params, sheet):
    if (len(params) < 3):
        XLM.color_print.output('r', "ERROR: MID() expects 3 parameters. " + str(params) + " given.")
        return ""
    start = XLM.utils.convert_num(params[1]) - 1
    end = start + XLM.utils.convert_num(params[2])
    the_str = str(params[0])
    if (start > len(the_str)):
        XLM.color_print.output('r', "ERROR: Start index of MID() (" + str(start) + ") not in string '" + the_str + "'.")
        return ""
    return the_str[start:end]
func_lookup["MID"] = MID

def CODE(params, sheet):
    # STUBBED
    return "CODE"
func_lookup["CODE"] = CODE

def GET_WINDOW(params, sheet):
    # STUBBED
    return "GET.WINDOW"
func_lookup["GET.WINDOW"] = GET_WINDOW

"""
def (params, sheet):
    # STUBBED
    return ""
func_lookup[""] = 
"""

####################################################################
def should_emulate_cell(cell):
    """
    Check to see if a given XLM cell runs something that we want to track 
    in the actions.

    @param cell (XLM_Object) The cell to check.

    @return (boolean) True if the cell should be emulated, False if not.
    """
    cell_str = str(cell)    
    for func in funcs_of_interest:
        if (cell_str.startswith(func)):
            return True
    return False

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
        if debug:
            print("FORMULA:")
            print("orig string: '" + str(r) + "'")
        if (_is_interesting_cell(new_cell)):
            sheet.cells[update_index] = new_cell
            if debug:
                print("'" + str(new_cell) + "'")
                print(new_cell.cell_id)

        # Track this if it is a new cell.
        if (update_index not in sheet.xlm_cell_indices):
            sheet.xlm_cell_indices.append(update_index)

    # Return the result.
    return r

    
