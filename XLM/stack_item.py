"""@package stack_item

Class for representing a single item on the XLM stack machine.
"""

import XLM.utils
import XLM.color_print

####################################################################
class stack_item(object):
    """
    Parent class for representing a single item on the XLM stack machine.
    """

    ####################################################################
    def eval(self, sheet):
        """
        Evaluate the current value of the object.
        
        @param sheet (ExcelSheet object) The Excel sheet in which to evaluate this
        XLM stack item.

        @return The result of emualting this stack item.
        """
        raise NotImplementedError("eval() not implemented in " + str(type(self)))

    ####################################################################
    def full_str(self):
        """
        A human readable version of this stack item.
        """
        raise NotImplementedError

    ####################################################################
    def __repr__(self):
        """
        String representation.
        """
        return self.full_str()
    
    ####################################################################
    def is_function(self):
        """
        Test the current stack item to see if it is a function that operates on other
        items pushed on the stack.

        @return (boolean) True if this is a function, False if not.
        """
        return hasattr(self, "num_args")

    ####################################################################
    def get_num_args(self):
        """
        If this stack item is a function, get the number of arguments the function
        takes.

        @return (int) 0 if this is not a function, the number of arguments if it is a 
        function.
        """
        if (self.is_function()):
            return self.num_args
        return 0

    ####################################################################
    def is_infix_function(self):
        """
        Tests this stack item to see if it should be represented as an infix function in
        a human readable string.

        @return (boolean) True if this is an infix function, False if not.
        """
        return ((hasattr(self, "is_infix_func")) and self.is_infix_func)
    
####################################################################    
## Implementation of specific stack items.
####################################################################

####################################################################
class unparsed(stack_item):
    """
    Unparsed bytes.
    """
    
    ####################################################################
    def __init__(self):
        """
        Constructor.
        """
        pass
    
    ####################################################################
    def full_str(self):
        """
        A human readable version of this stack item.
        """
        return "...UNPARSED..."

    ####################################################################
    def eval(self, sheet):
        XLM.color_print.output('y', "WARNING: Unparsed XLM being skipped during emulation.")
        return 0
    
####################################################################
class stack_int(stack_item):
    """
    Integer constant on the stack.
    """

    ####################################################################
    def __init__(self, value):
        """
        Constructor.

        @param value (str or int) The value of the stack item.
        """

        # Convert strings.
        if (isinstance(value, str)):
            value = int(value)
        self.value = value
    
    ####################################################################
    def full_str(self):
        """
        A human readable version of this stack item.
        """
        return str(self.value)

    ####################################################################
    def eval(self, sheet):
        return self.value
    
####################################################################
## Number of arguments required by functions referenced with ptgFuncV.
num_funcv_args = {"CHAR" : (1),
                  "END.IF" : (0),
                  "ERROR.TYPE" : (1),
                  "FCLOSE" : (1),
                  "FREAD" : (2,3),
                  "FWRITELN" : (2),
                  "GET.WORKSPACE" : (1),
                  "GOTO" : (1),
                  "ISERROR" : (1),
                  "ISNUMBER" : (1),
                  "LOWER" : (1),
                  "NEXT" : (0),
                  "NOW" : (0),
                  "WHILE" : (1),
                  "SET.VALUE" : (2),
                  "NOT" : (1),
                  "DAY" : (1),
                  "ACTIVE.CELL" : (0),
                  "ELSE" : (0),
                  "LEN" : (1),
                  "MID" : (3) }
    
####################################################################
class stack_funcv(stack_item):
    """
    Name of called funtion on stack.
    """
    
    ####################################################################
    def __init__(self, name, hexcode):
        """
        Constructor.

        @param name (str) The name of the function.
        @param hexcode (str) The hex associated with this item by olevba.
        """
        self.name = str(name)
        self.hexcode = str(hexcode)

        # Figure out how many args the function takes.
        if (self.name in num_funcv_args.keys()):
            self.num_args = num_funcv_args[self.name]
            # TODO: Handle functions that can take a variable # of arguments.
            if (isinstance(self.num_args, tuple)):
                self.num_args = self.num_args[0]
        else:
            XLM.color_print.output('y', "WARNING: # of arguments for '" + self.name + "' unknown. Defaulting to 1.")
            self.num_args = 1
    
    ####################################################################
    def full_str(self):
        """
        A human readable version of this stack item.
        """
        return self.name

####################################################################
class stack_concat(stack_item):
    """
    String concatenate operator on the stack.
    """
    
    ####################################################################
    def __init__(self):
        """
        Constructor.
        """
        self.num_args = 2
        self.is_infix_func = True
        self.name = "_concat"
    
    ####################################################################
    def full_str(self):
        """
        A human readable version of this stack item.
        """
        return "&"

####################################################################
class stack_cell_ref(stack_item):
    """
    Cell reference on the stack.
    """
    
    ####################################################################
    def __init__(self, row, column):
        """
        Constructor.

        @param row (str or int) The row number of the cell.
        @param column (str or int) The column number of the cell.
        """
        if (isinstance(row, str)):
            row = int(row)
        self.row = row
        if (isinstance(column, str)):
            column = int(column)
        self.column = column
        if (self.column > 49152):
            self.column -= 49152
        self.is_relative = False
    
    ####################################################################
    def full_str(self):
        """
        A human readable version of this stack item.
        """
        return "$R" + str(self.row) + "$C" + str(self.column)

    ####################################################################
    def eval(self, sheet):
        try:
            return sheet.cell(self.row, self.column)
        except KeyError:
            XLM.color_print.output('y', "WARNING: Cell '" + str(self) + "' not found. Defaulting to ''.")
            return ''
    
####################################################################
class stack_str(stack_item):
    """
    String constant on stack.
    """
    
    ####################################################################
    def __init__(self, value):
        """
        Constructor.

        @param value (str) Value of the string.
        """
        try:
            self.value = str(value).replace("&apos;", "'")
        except UnicodeEncodeError:
            self.value = XLM.utils.strip_unprintable(value)
    
    ####################################################################
    def full_str(self):
        """
        A human readable version of this stack item.
        """
        return '"' + self.value + '"'

    ####################################################################
    def eval(self, sheet):
        return self.value

####################################################################
class stack_bool(stack_item):
    """
    Boolean constant on the stack.
    """

    ####################################################################
    def __init__(self, value):
        """
        Constructor.

        @param value (str or int) The value of the stack item.
        """

        # Convert strings.
        if (isinstance(value, str)):
            if (value.lower() == "true"):
                value = True
            else:
                value = False
        self.value = value
    
    ####################################################################
    def full_str(self):
        """
        A human readable version of this stack item.
        """
        return str(self.value).upper()

    ####################################################################
    def eval(self, sheet):
        return self.value
    
####################################################################
class stack_attr(stack_item):

    ####################################################################
    def full_str(self):
        """
        A human readable version of this stack item.
        """
        return ""

    ####################################################################
    def eval(self, sheet):
        # TODO: What should this actually evaluate to?
        return ''
    
####################################################################
class stack_add(stack_item):
    """
    Addition operator on the stack.
    """
    
    ####################################################################
    def __init__(self):
        """
        Constructor.
        """
        self.num_args = 2
        self.is_infix_func = True
        self.name = "_plus"
    
    ####################################################################
    def full_str(self):
        """
        A human readable version of this stack item.
        """
        return "+"    

####################################################################
class stack_sub(stack_item):
    """
    Subtraction operator on the stack.
    """
    
    ####################################################################
    def __init__(self):
        """
        Constructor.
        """
        self.num_args = 2
        self.is_infix_func = True
        self.name = "_minus"
    
    ####################################################################
    def full_str(self):
        """
        A human readable version of this stack item.
        """
        return "-"    

####################################################################
class stack_exp(stack_item):
    """
    Exp on stack.
    """

    ####################################################################
    def __init__(self, row, column):
        """
        Constructor.

        @param row (str or int) The row number of the cell.
        @param column (str or int) The column number of the cell.
        """
        if (isinstance(row, str)):
            row = int(row)
        self.row = row
        if (isinstance(column, str)):
            column = int(column)
        self.column = column
        if (self.column > 49152):
            self.column -= 49152
    
    ####################################################################
    def full_str(self):
        """
        A human readable version of this stack item.
        """
        # TODO: Figure out proper representation.
        return "EXP?? $R" + str(self.row) + "$C" + str(self.column)

    ####################################################################
    def eval(self, sheet):
        # TODO: There may be a problem with plugin_biff.py where it incorrectly
        # parses out ptgExp when there should be a function call in the cell.
        return ''
    
####################################################################
class stack_name(stack_item):
    """
    Name?? of called funtion on stack.
    """
    
    ####################################################################
    def __init__(self, hexcode):
        """
        Constructor.

        @param hexcode (str) The hex associated with this item by olevba.
        """
        self.hexcode = str(hexcode)
    
    ####################################################################
    def full_str(self):
        """
        A human readable version of this stack item.
        """
        # TODO: Figure out the actual representation.
        return "NAME?? " + self.hexcode

    ####################################################################
    def eval(self, sheet):
        # TODO: What should this actually evaluate to?
        return 0
    
####################################################################
class stack_num(stack_item):
    """
    ptgNum on stack.
    """
    
    ####################################################################
    def __init__(self, name):
        """
        Constructor.

        @param name (str) The name of the function.
        """
        self.name = str(name)
    
    ####################################################################
    def full_str(self):
        """
        A human readable version of this stack item.
        """
        return self.name

    ####################################################################
    def eval(self, sheet):
        # TODO: What should this actually evaluate to?
        return self.name
    
####################################################################
class stack_missing_arg(stack_item):
    """
    Missing argument on stack.
    """
    
    ####################################################################
    def full_str(self):
        """
        A human readable version of this stack item.
        """
        return "...MISSING_ARG..."

    ####################################################################
    def eval(self, sheet):
        # TODO: What should this actually evaluate to?
        return 0
    
####################################################################
class stack_func(stack_item):
    """
    ptgFunc on stack.
    """
    
    ####################################################################
    def __init__(self, name):
        """
        Constructor.

        @param name (str) The name of the function.
        """
        self.name = str(name)
    
    ####################################################################
    def full_str(self):
        """
        A human readable version of this stack item.
        """
        return self.name

    ####################################################################
    def eval(self, sheet):
        # TODO: What should this actually evaluate to?
        return self.name
    
####################################################################
class stack_func_var(stack_item):
    """
    Generic function call on the stack.
    """

    ####################################################################
    def __init__(self, name, num_args, hexcode="???"):
        """
        Constructor.

        @param name (str) The name of the function.
        @param num_args (str or int) The number of arguments the function takes.
        @param hexcode (str) The hex associated with this item by olevba.
        """
        self.name = str(name)
        self.num_args = int(num_args)
        self.hexcode = str(hexcode)
    
    ####################################################################
    def full_str(self):
        """
        A human readable version of this stack item.
        """
        return self.name

####################################################################
class stack_namev(stack_item):

    ####################################################################
    def full_str(self):
        """
        A human readable version of this stack item.
        """
        return ""

####################################################################
class stack_area(stack_item):
    """
    Area on stack.
    """

    ####################################################################
    def __init__(self, row, column):
        """
        Constructor.

        @param row (str or int) The row number of the cell.
        @param column (str or int) The column number of the cell.
        """
        # TODO: Handle multiple cells.
        if (isinstance(row, str)):
            row = int(row)
        self.row = row
        if (isinstance(column, str)):
            column = int(column)
        self.column = column
        if (self.column > 49152):
            self.column -= 49152
    
    ####################################################################
    def full_str(self):
        """
        A human readable version of this stack item.
        """
        # TODO: Figure out proper representation.
        return "AREA?? $R" + str(self.row) + "$C" + str(self.column)    

    ####################################################################
    def eval(self, sheet):
        try:
            return sheet.cell(self.row, self.column)
        except KeyError:
            XLM.color_print.output('y', "WARNING: Cell '" + str(self) + "' not found. Defaulting to ''.")
            return ''
    
####################################################################
class stack_less_than(stack_item):
    """
    Less than operator on the stack.
    """
    
    ####################################################################
    def __init__(self):
        """
        Constructor.
        """
        self.num_args = 2
        self.is_infix_func = True
        self.name = "_less_than"
    
    ####################################################################
    def full_str(self):
        """
        A human readable version of this stack item.
        """
        return "<"

####################################################################
class stack_namex(stack_item):
    """
    ptgNamex on stack.
    """
    
    ####################################################################
    def __init__(self, name, number):
        """
        Constructor.
        """
        self.name = str(name)
        if (isinstance(number, str)):
            number = int(number)
        self.number = number
    
    ####################################################################
    def full_str(self):
        """
        A human readable version of this stack item.
        """
        # TODO: Figure out the proper representation for this.
        return "ptgNameX " + self.name + " " + str(self.number)

    ####################################################################
    def eval(self, sheet):
        # TODO: Figure out what this should do.
        return self.name
    
####################################################################
class stack_not_equal(stack_item):
    """
    Not equal operator on the stack.
    """
    
    ####################################################################
    def __init__(self):
        """
        Constructor.
        """
        self.num_args = 2
        self.is_infix_func = True
        self.name = "_not_equal"
    
    ####################################################################
    def full_str(self):
        """
        A human readable version of this stack item.
        """
        return "!="    

####################################################################
class stack_mul(stack_item):
    """
    Multiplication than operator on the stack.
    """
    
    ####################################################################
    def __init__(self):
        """
        Constructor.
        """
        self.num_args = 2
        self.is_infix_func = True
        self.name = "_times"
    
    ####################################################################
    def full_str(self):
        """
        A human readable version of this stack item.
        """
        return "*"

####################################################################
class stack_paren(stack_item):

    ####################################################################
    def full_str(self):
        """
        A human readable version of this stack item.
        """
        return "("

####################################################################
class unknown_token(unparsed):
    pass

####################################################################
class stack_array(stack_item):

    ####################################################################
    def full_str(self):
        """
        A human readable version of this stack item.
        """
        return "ARRAY"

####################################################################
class stack_equal(stack_item):
    """
    Equal operator on the stack.
    """
    
    ####################################################################
    def __init__(self):
        """
        Constructor.
        """
        self.num_args = 2
        self.is_infix_func = True
        self.name = "_equals"
    
    ####################################################################
    def full_str(self):
        """
        A human readable version of this stack item.
        """
        return "=="

####################################################################
class stack_greater_than(stack_item):
    """
    Greater than operator on the stack.
    """
    
    ####################################################################
    def __init__(self):
        """
        Constructor.
        """
        self.num_args = 2
        self.is_infix_func = True
        self.name = "_greater_than"
    
    ####################################################################
    def full_str(self):
        """
        A human readable version of this stack item.
        """
        return ">"

####################################################################
class stack_mem_func(stack_item):

    ####################################################################
    def full_str(self):
        """
        A human readable version of this stack item.
        """
        return "MEMFUNC"

####################################################################
class stack_power(stack_item):

    ####################################################################
    def full_str(self):
        """
        A human readable version of this stack item.
        """
        return "POWER"

####################################################################
class stack_ref_error(stack_item):

    ####################################################################
    def full_str(self):
        """
        A human readable version of this stack item.
        """
        return "REFERROR"

####################################################################
class stack_mem_no_mem(stack_item):

    ####################################################################
    def full_str(self):
        """
        A human readable version of this stack item.
        """
        return "NOMEM"

####################################################################
class stack_area_error(stack_item):

    ####################################################################
    def full_str(self):
        """
        A human readable version of this stack item.
        """
        return "AREAERROR"

####################################################################
class stack_div(stack_item):
    """
    Division operator on the stack.
    """
    
    ####################################################################
    def __init__(self):
        """
        Constructor.
        """
        self.num_args = 2
        self.is_infix_func = True
        self.name = "_divide"
    
    ####################################################################
    def full_str(self):
        """
        A human readable version of this stack item.
        """
        return "/"

####################################################################
class stack_uminus(stack_item):
    """
    Unsigned minus operator on the stack.
    """
    
    ####################################################################
    def __init__(self):
        """
        Constructor.
        """
        self.num_args = 2
        self.is_infix_func = True
        self.name = "_unsigned_minus"
    
    ####################################################################
    def full_str(self):
        """
        A human readable version of this stack item.
        """
        return "-u"

####################################################################
class stack_uplus(stack_item):
    """
    Unsigned plus operator on the stack.
    """
    
    ####################################################################
    def __init__(self):
        """
        Constructor.
        """
        self.num_args = 2
        self.is_infix_func = True
        self.name = "_unsigned_plus"
    
    ####################################################################
    def full_str(self):
        """
        A human readable version of this stack item.
        """
        return "+u"    

####################################################################
class stack_greater_equal(stack_item):
    """
    Greater than or equal operator on the stack.
    """
    
    ####################################################################
    def __init__(self):
        """
        Constructor.
        """
        self.num_args = 2
        self.is_infix_func = True
        self.name = "_greater_or_equal"
    
    ####################################################################
    def full_str(self):
        """
        A human readable version of this stack item.
        """
        return ">="

####################################################################
class stack_area_3d(stack_item):

    ####################################################################
    def full_str(self):
        """
        A human readable version of this stack item.
        """
        return "AREA3D"

####################################################################
class stack_end_sheet(stack_item):

    ####################################################################
    def full_str(self):
        """
        A human readable version of this stack item.
        """
        return "ENDSHEET"

####################################################################
class stack_mem_error(stack_item):

    ####################################################################
    def full_str(self):
        """
        A human readable version of this stack item.
        """
        return "MEMERROR"

####################################################################
class stack_percent(stack_item):

    ####################################################################
    def full_str(self):
        """
        A human readable version of this stack item.
        """
        return "PERCENT"

####################################################################
class stack_mem_area(stack_item):

    ####################################################################
    def full_str(self):
        """
        A human readable version of this stack item.
        """
        return "MEMAREA"

####################################################################
class stack_range(stack_item):

    ####################################################################
    def full_str(self):
        """
        A human readable version of this stack item.
        """
        return "RANGE"
    
