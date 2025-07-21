#[path = "xlsx/book.rs"]
pub mod book;
#[path = "xlsx/cell.rs"]
pub mod cell;
#[path = "xlsx/sheet.rs"]
pub mod sheet;
#[path = "xlsx/style.rs"]
pub mod style;
#[path = "xlsx/xml.rs"]
pub mod xml;

#[cfg(test)]
#[path = "xlsx/test_book.rs"]
mod test_book;
#[cfg(test)]
#[path = "xlsx/test_cell.rs"]
mod test_cell;
#[cfg(test)]
#[path = "xlsx/test_sheet.rs"]
mod test_sheet;
#[cfg(test)]
#[path = "xlsx/test_xml.rs"]
mod test_xml;

use pyo3::prelude::*;

use book::Book;
use cell::Cell;
use sheet::Sheet;
use style::{Font, PatternFill};
use xml::{Xml, XmlElement};

#[pyfunction]
pub fn hello_from_bin() -> String {
    "Hello from sample-ext-lib!".to_string()
}

#[pyfunction]
pub fn load_workbook(path: String) -> Book {
    Book::new(&path)
}

#[pymodule]
fn xlsx(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(hello_from_bin, m)?)?;
    m.add_function(wrap_pyfunction!(load_workbook, m)?)?;
    m.add_class::<Book>()?;
    m.add_class::<Sheet>()?;
    m.add_class::<Cell>()?;
    m.add_class::<Font>()?;
    m.add_class::<PatternFill>()?;
    m.add_class::<Xml>()?;
    m.add_class::<XmlElement>()?;
    Ok(())
}
