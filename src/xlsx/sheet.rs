use std::sync::{Arc, Mutex};

use pyo3::prelude::*;

use crate::xlsx::xml::Xml;

#[pyclass]
pub struct Sheet {
    pub name: String,
    xml: Arc<Mutex<Xml>>,
}

#[pymethods]
impl Sheet {
    // pub fn __repr__(&self) -> String {
    //     format!("<Sheet name='{}'>", self.name)
    // }

    // pub fn __getitem__(&mut self, key: String) -> Cell {
    //     let cell_ref = self.value.get_cell_mut(key);
    //     Cell::new(cell_ref)
    // }

    // pub fn cell(&mut self, row: usize, col: usize) -> Cell {
    //     let cell_ref = self.value.get_cell_mut((row as u32, col as u32));
    //     Cell::new(cell_ref)
    // }

    // pub fn set_value(&mut self, row: usize, col: usize, value: String) {
    //     let cell = self.value.get_cell_mut((row as u32, col as u32));
    //     cell.set_value(value);
    // }
}

impl Sheet {
    pub fn new(name: String, xml: Arc<Mutex<Xml>>) -> Self {
        Sheet { name, xml }
    }

    // pub fn get_name(&self) -> &str {
    //     &self.name
    // }
}
