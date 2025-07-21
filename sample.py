import os
from xlsx import load_workbook, Book, Font, PatternFill, Cell, Sheet, Xml


def test_font_style():
    # Create a new workbook in memory
    book: Book = Book()
    sheet: Sheet = book.create_sheet("Styled Sheet", 0)

    # Set value to a cell
    cell: Cell = sheet["A1"]
    cell.value = "Hello, Styled World!"

    # Create a font and apply it to the cell
    font: Font = Font(name="Arial", size=14, bold=True, color="FF0000")
    cell.font = font

    # Create a fill and apply it to another cell
    fill_cell: Cell = sheet["B1"]
    fill_cell.value = "Hello, Filled World!"
    fill: PatternFill = PatternFill(pattern_type="solid", fg_color="FFFF00")
    fill_cell.fill = fill

    # Save the workbook to a file
    test_file: str = "test_style.xlsx"
    if os.path.exists(test_file):
        os.remove(test_file)
    book.copy(test_file)

    print(f"Workbook '{test_file}' created. Please inspect it manually for styles.")


if __name__ == "__main__":
    test_font_style()
