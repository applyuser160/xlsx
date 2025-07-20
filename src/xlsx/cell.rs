use crate::xlsx::xml::{Xml, XmlElement};
use pyo3::prelude::*;
use std::sync::{Arc, Mutex};

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
                                                if let Some(v_element) = cell_element
                                                    .children
                                                    .iter()
                                                    .find(|e| e.name == "v")
                                                {
                                                    if let Some(text) = &v_element.text {
                                                        if let Ok(idx) = text.parse::<usize>() {
                                                            let shared_strings_xml =
                                                                self.shared_strings.lock().unwrap();
                                                            if let Some(sst) =
                                                                shared_strings_xml.elements.first()
                                                            {
                                                                if let Some(si) =
                                                                    sst.children.get(idx)
                                                                {
                                                                    if let Some(t) =
                                                                        si.children.first()
                                                                    {
                                                                        return t.text.clone();
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            } else if t_attr == "inlineStr" {
                                                if let Some(is_element) = cell_element
                                                    .children
                                                    .iter()
                                                    .find(|e| e.name == "is")
                                                {
                                                    if let Some(t_element) = is_element
                                                        .children
                                                        .iter()
                                                        .find(|e| e.name == "t")
                                                    {
                                                        return t_element.text.clone();
                                                    }
                                                }
                                            }
                                        }
                                        // t属性がないか、"s"以外の場合は、vタグの値を直接返す
                                        if let Some(v_element) =
                                            cell_element.children.iter().find(|e| e.name == "v")
                                        {
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
        // 数値への変換を試みる
        if let Ok(number) = value.parse::<f64>() {
            self.set_number_value(number);
        } else {
            self.set_string_value(&value);
        }
    }
}

impl Cell {
    pub fn new(
        sheet_xml: Arc<Mutex<Xml>>,
        shared_strings: Arc<Mutex<Xml>>,
        address: String,
    ) -> Self {
        Cell {
            sheet_xml,
            shared_strings,
            address,
        }
    }

    pub fn set_number_value(&mut self, value: f64) {
        let mut xml = self.sheet_xml.lock().unwrap();
        let cell_element = self.get_or_create_cell_element(&mut xml);
        cell_element.attributes.remove("t"); // 文字列型ではないのでt属性を削除
        if let Some(v) = cell_element.children.iter_mut().find(|c| c.name == "v") {
            v.text = Some(value.to_string());
        } else {
            let mut v_element = XmlElement::new("v");
            v_element.text = Some(value.to_string());
            cell_element.children.push(v_element);
        }
    }

    pub fn set_string_value(&mut self, value: &str) {
        let sst_index = self.get_or_create_shared_string(value);
        let mut xml = self.sheet_xml.lock().unwrap();
        let cell_element = self.get_or_create_cell_element(&mut xml);
        cell_element
            .attributes
            .insert("t".to_string(), "s".to_string());
        if let Some(v) = cell_element.children.iter_mut().find(|c| c.name == "v") {
            v.text = Some(sst_index.to_string());
        } else {
            let mut v_element = XmlElement::new("v");
            v_element.text = Some(sst_index.to_string());
            cell_element.children.push(v_element);
        }
    }

    fn get_or_create_cell_element<'a>(&self, xml: &'a mut Xml) -> &'a mut XmlElement {
        let (row_num, _) = self.decode_address();
        let sheet_data = xml
            .elements
            .first_mut()
            .unwrap()
            .children
            .iter_mut()
            .find(|e| e.name == "sheetData")
            .unwrap();

        // Rowを探す
        let row_position = sheet_data
            .children
            .iter()
            .position(|r| r.name == "row" && r.attributes.get("r") == Some(&row_num.to_string()));

        // Rowがなければ作成
        let row_index = match row_position {
            Some(pos) => pos,
            None => {
                let mut new_row = XmlElement::new("row");
                new_row
                    .attributes
                    .insert("r".to_string(), row_num.to_string());
                sheet_data.children.push(new_row);
                sheet_data.children.len() - 1
            }
        };
        let row_element = &mut sheet_data.children[row_index];

        // Cellを探す
        let cell_position = row_element
            .children
            .iter()
            .position(|c| c.name == "c" && c.attributes.get("r") == Some(&self.address));

        // Cellがなければ作成
        let cell_index = match cell_position {
            Some(pos) => pos,
            None => {
                let mut new_cell = XmlElement::new("c");
                new_cell
                    .attributes
                    .insert("r".to_string(), self.address.clone());
                row_element.children.push(new_cell);
                row_element.children.len() - 1
            }
        };
        &mut row_element.children[cell_index]
    }

    fn get_or_create_shared_string(&mut self, text: &str) -> usize {
        let mut shared_strings_xml = self.shared_strings.lock().unwrap();

        // sst要素がなければ作成
        if shared_strings_xml.elements.is_empty() {
            let sst_element = XmlElement::new("sst");
            shared_strings_xml.elements.push(sst_element);
        }
        let sst_element = shared_strings_xml.elements.first_mut().unwrap();

        // 既存の文字列を探す
        for (i, si) in sst_element.children.iter().enumerate() {
            if let Some(t) = si.children.first() {
                if t.text.as_deref() == Some(text) {
                    return i;
                }
            }
        }

        // 新しい文字列を追加
        let mut t_element = XmlElement::new("t");
        t_element.text = Some(text.to_string());
        let mut si_element = XmlElement::new("si");
        si_element.children.push(t_element);
        sst_element.children.push(si_element);
        sst_element.children.len() - 1
    }

    fn decode_address(&self) -> (u32, u32) {
        let mut col_str = String::new();
        let mut row_str = String::new();
        for ch in self.address.chars() {
            if ch.is_alphabetic() {
                col_str.push(ch);
            } else {
                row_str.push(ch);
            }
        }
        let row = row_str.parse::<u32>().unwrap();
        let mut col = 0;
        for (i, ch) in col_str.to_uppercase().chars().rev().enumerate() {
            col += (ch as u32 - 'A' as u32 + 1) * 26u32.pow(i as u32);
        }
        (row, col)
    }
}
