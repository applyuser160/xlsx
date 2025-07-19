mod xlsx {
    pub mod book;
    pub mod cell;
    pub mod sheet;
    pub mod test_book;
    pub mod test_xml;
    pub mod xml;
}

use pyo3::prelude::*;
use umya_spreadsheet::reader;

use xlsx::book::Book;
use xlsx::cell::Cell;
use xlsx::sheet::Sheet;
use xlsx::xml::{Xml, XmlElement};

#[pyfunction]
pub fn hello_from_bin() -> String {
    "Hello from sample-ext-lib!".to_string()
}

#[pyfunction]
pub fn read_file(path: String, sheet: String, address: String) -> String {
    let path = std::path::Path::new(&path);
    let book = reader::xlsx::read(path).unwrap();
    book.get_sheet_by_name(sheet.as_str())
        .unwrap()
        .get_value(address)
}

#[pyfunction]
pub fn load_workbook(path: String) -> Book {
    Book::new(path)
}

#[pymodule]
pub fn _core(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(hello_from_bin, m)?)?;
    m.add_function(wrap_pyfunction!(read_file, m)?)?;
    m.add_function(wrap_pyfunction!(load_workbook, m)?)?;
    m.add_class::<Book>()?;
    m.add_class::<Sheet>()?;
    m.add_class::<Cell>()?;
    m.add_class::<Xml>()?;
    m.add_class::<XmlElement>()?;
    Ok(())
}
