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

#[cfg(test)]
fn run_all_tests() {
    // test_book
    test_book::tests::test_new_book();
    test_book::tests::test_copy_book();
    test_book::tests::test_sheetnames();
    test_book::tests::test_contains__();
    test_book::tests::test_create_sheet();
    test_book::tests::test_merge_xmls();
    test_book::tests::test_write_file_indirectly();
    test_book::tests::test_sheet_tags();
    test_book::tests::test_relationships();
    test_book::tests::test_sheet_paths();
    test_book::tests::test_delete_sheet();
    test_book::tests::test_sheet_index();
    test_book::tests::test_create_sheet_with_index();
    test_book::tests::test_add_table();
}

#[cfg(test)]
#[pyfunction]
fn run_tests_py() -> PyResult<()> {
    run_all_tests();
    Ok(())
}

#[cfg(test)]
#[pymodule]
fn run_tests_mod(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(run_tests_py, m)?)?;
    Ok(())
}

#[cfg(test)]
pub fn main() {
    pyo3::prepare_freethreaded_python();
    Python::with_gil(|py| {
        let module = pyo3::wrap_pymodule!(run_tests_mod)(py);
        py.import_bound("sys")
            .and_then(|sys| sys.getattr("modules"))
            .and_then(|modules| modules.set_item("run_tests_mod", module))
            .unwrap();

        let code = "import run_tests_mod; run_tests_mod.run_tests_py()";
        py.run_bound(code, None, None).unwrap();
    });
}
