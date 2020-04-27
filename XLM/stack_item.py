"""@package stack_item

Class for representing a single item on the XLM stack machine.
"""

####################################################################
class stack_item(object):
    """
    Parent class for representing a single item on the XLM stack machine.
    """

    ####################################################################
    def eval(self, workbook):
        """
        Evaluate the current value of the object.
        
        @param workbook (Excel workbook object) The Excel workbook in which to evaluate this
        XLM stack item.

        @return The result of emualting this stack item.
        """
        raise NotImplementedError

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
    
    ####################################################################
    def full_str(self):
        """
        A human readable version of this stack item.
        """
        return "$R" + str(self.row) + "$C" + str(self.column)

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
        self.value = str(value)
    
    ####################################################################
    def full_str(self):
        """
        A human readable version of this stack item.
        """
        return '"' + self.value + '"'

####################################################################
class stack_bool(stack_item):
    pass

####################################################################
class stack_attr(stack_item):

    ####################################################################
    def full_str(self):
        """
        A human readable version of this stack item.
        """
        return ""

####################################################################
class stack_add(stack_item):
    pass

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
    
    ####################################################################
    def full_str(self):
        """
        A human readable version of this stack item.
        """
        return "-"    

####################################################################
class stack_exp(stack_item):
    pass

####################################################################
class stack_name(stack_item):
    pass

####################################################################
class stack_num(stack_item):
    pass

####################################################################
class stack_missing_arg(stack_item):
    pass

####################################################################
class stack_func(stack_item):
    pass

####################################################################
class stack_func_var(stack_item):
    """
    Generic function call on the stack.
    """

    ####################################################################
    def __init__(self, name, num_args, hexcode):
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
    pass

####################################################################
class stack_area(stack_item):
    pass

####################################################################
class stack_less_than(stack_item):
    pass

####################################################################
class stack_namex(stack_item):
    pass

####################################################################
class stack_not_equal(stack_item):
    pass

####################################################################
class stack_mul(stack_item):
    pass

####################################################################
class stack_paren(stack_item):
    pass

####################################################################
class unknown_token(stack_item):
    pass

####################################################################
class stack_array(stack_item):
    pass

####################################################################
class stack_equal(stack_item):
    pass

####################################################################
class stack_greater_than(stack_item):
    pass

####################################################################
class stack_mem_func(stack_item):
    pass

####################################################################
class stack_power(stack_item):
    pass

####################################################################
class stack_ref_error(stack_item):
    pass

####################################################################
class stack_mem_no_mem(stack_item):
    pass

####################################################################
class stack_area_error(stack_item):
    pass

####################################################################
class stack_div(stack_item):
    pass

####################################################################
class stack_uminus(stack_item):
    pass

####################################################################
class stack_greater_equal(stack_item):
    pass

####################################################################
class stack_area_3d(stack_item):
    pass

####################################################################
class stack_end_sheet(stack_item):
    pass

####################################################################
class stack_mem_error(stack_item):
    pass
