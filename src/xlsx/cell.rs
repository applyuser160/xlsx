use std::sync::{Arc, Mutex};
use pyo3::prelude::*;
use crate::xlsx::xml::Xml;

#[pyclass]
pub struct Cell {
    sheet_xml: Arc<Mutex<Xml>>,
    shared_strings: Arc<Mutex<Xml>>,
    address: String,
}

#[pymethods]
impl Cell {
    #[getter]
    pub fn value(&self) -> Option<String> {
        let xml = self.sheet_xml.lock().unwrap();
        if let Some(worksheet) = xml.elements.first() {
            if let Some(sheet_data) = worksheet.children.iter().find(|e| e.name == "sheetData") {
                for row in &sheet_data.children {
                    if row.name == "row" {
                        for cell_element in &row.children {
                            if cell_element.name == "c" {
                                if let Some(r_attr) = cell_element.attributes.get("r") {
                                    if r_attr == &self.address {
                                        if let Some(t_attr) = cell_element.attributes.get("t") {
                                            if t_attr == "s" {
                                                if let Some(v_element) = cell_element.children.iter().find(|e| e.name == "v") {
                                                    if let Some(text) = &v_element.text {
                                                        if let Ok(idx) = text.parse::<usize>() {
                                                            let shared_strings_xml = self.shared_strings.lock().unwrap();
                                                            if let Some(sst) = shared_strings_xml.elements.first() {
                                                                if let Some(si) = sst.children.get(idx) {
                                                                    if let Some(t) = si.children.first() {
                                                                        return t.text.clone();
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            } else if t_attr == "inlineStr" {
                                                if let Some(is_element) = cell_element.children.iter().find(|e| e.name == "is") {
                                                    if let Some(t_element) = is_element.children.iter().find(|e| e.name == "t") {
                                                        return t_element.text.clone();
                                                    }
                                                }
                                            }
                                        }
                                        // t属性がないか、"s"以外の場合は、vタグの値を直接返す
                                        if let Some(v_element) = cell_element.children.iter().find(|e| e.name == "v") {
                                            return v_element.text.clone();
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        None
    }

    #[setter]
    pub fn set_value(&mut self, value: String) {
        let mut xml = self.sheet_xml.lock().unwrap();
        if let Some(worksheet) = xml.elements.first_mut() {
            if let Some(sheet_data) = worksheet.children.iter_mut().find(|e| e.name == "sheetData") {
                for row in &mut sheet_data.children {
                    if row.name == "row" {
                        for cell_element in &mut row.children {
                            if cell_element.name == "c" {
                                if let Some(r_attr) = cell_element.attributes.get("r") {
                                    if r_attr == &self.address {
                                        // t="s" (shared string)の場合の処理は未実装
                                        if let Some(v_element) = cell_element.children.iter_mut().find(|e| e.name == "v") {
                                            v_element.text = Some(value);
                                            // 既存のセルのt属性を削除して、数値を想定
                                            cell_element.attributes.remove("t");
                                            return;
                                        }
                                        // TODO: vタグがない場合の処理
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        // TODO: セルが存在しない場合の処理
    }
}

impl Cell {
    pub fn new(sheet_xml: Arc<Mutex<Xml>>, shared_strings: Arc<Mutex<Xml>>, address: String) -> Self {
        Cell {
            sheet_xml,
            shared_strings,
            address,
        }
    }
}
