"""@package stack_transformer

Lark AST transformer to generate XLM_Objects from an AST.
"""

from lark import Transformer

from stack_item import *

####################################################################
class StackTransformer(Transformer):
    """
    Lark AST transformer to generate XLM_Objects from an AST.
    """
    
    def NUMBER(self, items):
        return int(str(items))

    def NAME(self, items):
        return str(items)

    def HEX_NUMBER(self, items):
        return str(items)

    def DOUBLE_QUOTE_STRING(self, items):
        tmp = str(items)
        return tmp[1:-1]

    def SINGLE_QUOTE_STRING(self, items):
        tmp = str(items)
        return tmp[1:-1]

    def stack_item(self, items):
        return items[0]

    def cell_formula(self, items):
        print items
        # Skip length.        
        r = [items[0]] + items[2:]
        return r
    
    def cell(self, items):
        return (items[0], items[1])
    
    def stack_int(self, items):
        return stack_int(items[0])
    
    def stack_funcv(self, items):
        # ['CHAR', '0x006f']
        return stack_funcv(items[0], items[1])
    
    def stack_concat(self, items):
        return stack_concat()
    
    def stack_cell_ref(self, items):
        return stack_cell_ref(items[0][0], items[0][1])
    
    def stack_str(self, items):
        return stack_str(items[0])
    
    def stack_bool(self, items):
        return stack_bool()
    
    def stack_attr(self, items):
        return stack_attr()
    
    def stack_add(self, items):
        return stack_add()
    
    def stack_sub(self, items):
        return stack_sub()
    
    def stack_exp(self, items):
        return stack_exp()
    
    def stack_name(self, items):
        return stack_name()
    
    def stack_num(self, items):
        return stack_num()
    
    def stack_missing_arg(self, items):
        return stack_missing_arg()
    
    def stack_func(self, items):
        return stack_func()
    
    def stack_func_var(self, items):
        # [1, 'RUN', '0x8011']
        return stack_func_var(items[1], items[0], items[2])
    
    def stack_namev(self, items):
        return stack_namev()
    
    def stack_area(self, items):
        return stack_area()
    
    def stack_less_than(self, items):
        return stack_less_than()
    
    def stack_namex(self, items):
        return stack_namex()
    
    def stack_not_equal(self, items):
        return stack_not_equal()
    
    def stack_mul(self, items):
        return stack_mul()
    
    def stack_paren(self, items):
        return stack_paren()
    
    def stack_array(self, items):
        return stack_array()
    
    def stack_equal(self, items):
        return stack_equal()
    
    def stack_greater_than(self, items):
        return stack_greater_than()
    
    def stack_mem_func(self, items):
        return stack_mem_func()
    
    def stack_power(self, items):
        return stack_power()
    
    def stack_ref_error(self, items):
        return stack_ref_error()
    
    def stack_mem_no_mem(self, items):
        return stack_mem_no_mem()
    
    def stack_area_error(self, items):
        return stack_area_error()
    
    def stack_div(self, items):
        return stack_div()
    
    def stack_uminus(self, items):
        return stack_uminus()
    
    def stack_greater_equal(self, items):
        return stack_greater_equal()
    
    def stack_area_3d(self, items):
        return stack_area_3d()
    
    def stack_end_sheet(self, items):
        return stack_end_sheet()
    
    def stack_mem_error(self, items):
        return stack_mem_error()
    
