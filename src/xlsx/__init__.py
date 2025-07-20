from ._core import hello_from_bin, load_workbook, Book, Sheet, Cell

__all__ = ["hello", "load_workbook", "Book", "Sheet", "Cell"]


def hello() -> str:
    return hello_from_bin()
