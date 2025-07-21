use std::sync::{Arc, Mutex};

use pyo3::prelude::*;

use crate::cell::Cell;
use crate::xml::Xml;

/// Represents a worksheet in an Excel workbook.
#[pyclass]
pub struct Sheet {
    /// The name of the worksheet.
    #[pyo3(get)]
    pub name: String,
    /// The XML of the worksheet.
    xml: Arc<Mutex<Xml>>,
    /// The shared strings XML.
    shared_strings: Arc<Mutex<Xml>>,
    /// The styles XML.
    styles: Arc<Mutex<Xml>>,
}

#[pymethods]
impl Sheet {
    /// Gets a cell by its address (e.g., "A1").
    pub fn __getitem__(&self, key: &str) -> Cell {
        Cell::new(
            self.xml.clone(),
            self.shared_strings.clone(),
            self.styles.clone(),
            key.to_string(),
        )
    }

    /// Gets a cell by its row and column number.
    #[pyo3(signature = (row, column))]
    pub fn cell(&self, row: usize, column: usize) -> Cell {
        let address = Self::coordinate_to_string(row, column);
        Cell::new(
            self.xml.clone(),
            self.shared_strings.clone(),
            self.styles.clone(),
            address,
        )
    }
}

impl Sheet {
    /// Creates a new `Sheet` instance.
    pub fn new(
        name: String,
        xml: Arc<Mutex<Xml>>,
        shared_strings: Arc<Mutex<Xml>>,
        styles: Arc<Mutex<Xml>>,
    ) -> Self {
        Sheet {
            name,
            xml,
            shared_strings,
            styles,
        }
    }

    #[cfg(test)]
    pub(crate) fn get_xml(&self) -> Arc<Mutex<Xml>> {
        self.xml.clone()
    }

    /// Converts row and column numbers to a cell address string.
    fn coordinate_to_string(row: usize, col: usize) -> String {
        let mut col_str = String::new();
        let mut col_num = col;
        while col_num > 0 {
            let remainder = (col_num - 1) % 26;
            col_str.insert(0, (b'A' + remainder as u8) as char);
            col_num = (col_num - 1) / 26;
        }
        format!("{col_str}{row}")
    }
}
