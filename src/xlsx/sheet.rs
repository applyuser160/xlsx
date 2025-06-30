use pyo3::prelude::*;
use umya_spreadsheet::Worksheet;

use crate::xlsx::cell::Cell;

#[pyclass]
pub struct Sheet {
    #[pyo3(get, set)]
    pub name: String,
    value: Worksheet
}

#[pymethods]
impl Sheet {

    pub fn __repr__(&self) -> String {
        format!("<Sheet name='{}'>", self.name)
    }

    pub fn __getitem__(&mut self, key: String, py: Python) -> PyObject {
        if key.contains(':') {
            let vec: Vec<String> = self.value.get_cell_value_by_range(&key)
                .into_iter().map(|v| v.get_value().to_string()).collect();
            return vec.into_py(py);
        }
        let cell_ref = self.value.get_cell_mut(key);
        let cell = Cell::new(cell_ref);
        cell.into_py(py)
    }

    pub fn cell(&mut self, row: usize, col: usize) -> Cell {
        let cell_ref = self.value.get_cell_mut((row as u32, col as u32));
        Cell::new(cell_ref)
    }

    pub fn set_value(&mut self, row: usize, col: usize, value: String) {
        let cell = self.value.get_cell_mut((row as u32, col as u32));
        cell.set_value(value);
    }
}

impl Sheet {
    pub fn new(name: String, value: Worksheet) -> Self {
        Sheet { name, value }
    }

    pub fn get_name(&self) -> &str {
        &self.name
    }

}
