import decimal
import datetime
from functools import singledispatch

from openpyxl.compat import basestring
from openpyxl.styles import numbers
from openpyxl.styles.proxy import StyledObject
from openpyxl.utils.datetime import timedelta_to_days, to_excel, days_to_time


class Cell(StyledObject):

    __slots__ = ('column',
                 'row',
                 'coordinate',
                 '_value',
                 'data_type',
                 'parent',
                 'xf_index',
                 '_hyperlink_rel',
                 '_comment',
                 '_style_id',
                 '_font_id',
                 '_fill_id',
                 '_alignment_id',
                 '_border_id',
                 '_number_format_id',
                 '_protection_id',
                 )

    TYPE_STRING = 's'
    TYPE_FORMULA = 'f'
    TYPE_NUMERIC = 'n'
    TYPE_BOOL = 'b'
    TYPE_NULL = 'n'
    TYPE_INLINE = 'inlineStr'
    TYPE_ERROR = 'e'
    TYPE_FORMULA_CACHE_STRING = 'str'


    def __init__(self, worksheet, column, row, value=None):
        super(Cell, self).__init__()
        self.column = column.upper()
        self.row = row
        self.coordinate = '%s%d' % (self.column, self.row)
        # _value is the stored value, while value is the displayed value
        self._value = None
        self._hyperlink_rel = None
        self.data_type = self.TYPE_NULL
        self.parent = worksheet
        if value is not None:
            self.value = value
        self.xf_index = 0
        self._comment = None


    @property
    def value(self):
        return self._value

    @value.setter
    def value(self, value):
        _bind(value, self)

    @property
    def base_date(self):
        return None


@singledispatch
def _bind(value, cell):
    raise ValueError()


@_bind.register(bool)
def _bind_bool(value, cell):
    cell.data_type = Cell.TYPE_BOOL
    cell._value = value


@_bind.register(decimal.Decimal)
@_bind.register(int)
@_bind.register(float)
def _bind_number(value, cell):
    cell.data_type = Cell.TYPE_BOOL
    cell._value = value


@_bind.register(datetime.datetime)
def _bind_datetime(value, cell):
    cell.data_type = Cell.TYPE_NUMERIC
    cell.value = to_excel(value, cell.base_date)
    cell.number_format = numbers.FORMAT_DATE_DATETIME


@_bind.register(datetime.date)
def _bind_date(value, cell):
    cell.data_type = Cell.TYPE_NUMERIC
    cell._value = to_excel(value, cell.base_date)
    cell.number_format = numbers.FORMAT_DATE_YYYYMMDD2


@_bind.register(datetime.time)
def _bind_time(value, cell):
    cell.data_type = Cell.TYPE_NUMERIC
    cell._value = time_to_days(value)
    cell.number_format = numbers.FORMAT_DATE_TIME6


@_bind.register(datetime.timedelta)
def _bind_timedelta(value, cell):
    cell.data_type = Cell.TYPE_NUMERIC
    cell._value = timedelta_to_days(value)
    cell.number_format = numbers.FORMAT_DATE_TIMEDELTA


@_bind.register(basestring)
def _bind_string(value, cell):
    data_type = Cell.TYPE_STRING
    if len(value) > 1 and value.startswith("="):
        data_type = Cell.TYPE_FORMULA
    elif value in self.ERROR_CODES:
        data_type = Cell.TYPE_ERROR
