use pyo3::prelude::*;
use umya_spreadsheet::{Cell as UCell};

#[pyclass]
pub struct Cell {
    value: UCell
}

#[pymethods]
impl Cell {
    pub fn __repr__(&self) -> String {
        format!("<Cell value='{}'>", self.value.get_value())
    }

    pub fn get_value(&self) -> String {
        self.value.get_value().to_string()
    }

    pub fn set_value(&mut self, value: String) {
        self.value.set_value(value);
    }
}

impl Cell {
    pub fn new(value: &mut UCell) -> Self {
        Cell { value: value.clone() }
    }
}
