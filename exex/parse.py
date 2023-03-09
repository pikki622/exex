from openpyxl.cell.cell import Cell

import types


def values(val):
    if isinstance(val, types.GeneratorType):
        val = tuple(val)

    if isinstance(val, Cell):
        return from_cell(val)
    elif isinstance(val, tuple):
        return from_array(val) if isinstance(val[0], Cell) else from_range(val)
    else:
        return val


def from_cell(cell):
    return cell.value if type(cell) is Cell else cell


def from_array(row):
    return [from_cell(cell) for cell in list(row)]


def from_range(range_):
    return [from_array(row) for row in list(range_)]
