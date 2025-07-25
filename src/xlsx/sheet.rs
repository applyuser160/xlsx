use std::sync::{Arc, Mutex};

use pyo3::prelude::*;

use crate::cell::Cell;
use crate::xml::{Xml, XmlElement};

use std::collections::HashMap;

/// Excelワークブック内のワークシート
#[pyclass]
pub struct Sheet {
    /// ワークシートの名前
    #[pyo3(get)]
    pub name: String,
    /// ワークシートのXML
    xml: Arc<Mutex<Xml>>,
    /// 共有文字列のXML
    shared_strings: Arc<Mutex<Xml>>,
    /// 共有文字列のマップ
    shared_strings_map: Arc<Mutex<HashMap<String, usize>>>,
    /// スタイルのXML
    styles: Arc<Mutex<Xml>>,
}

#[pymethods]
impl Sheet {
    /// アドレスによるセルの取得 (例: "A1")
    pub fn __getitem__(&self, key: &str) -> Cell {
        Cell::new(
            self.xml.clone(),
            self.shared_strings.clone(),
            self.styles.clone(),
            key.to_string(),
        )
    }

    /// 行と列の番号によるセルの取得
    #[pyo3(signature = (row, column))]
    pub fn cell(&self, row: usize, column: usize) -> Cell {
        let address: String = Self::coordinate_to_string(row, column);
        Cell::new(
            self.xml.clone(),
            self.shared_strings.clone(),
            self.styles.clone(),
            address,
        )
    }

    /// シートへの行の追加
    pub fn append(&self, row_data: Vec<String>) {
        if let Ok(mut xml) = self.xml.lock() {
            if let Some(worksheet) = xml.elements.first_mut() {
                let sheet_data = worksheet.get_element_mut("sheetData");
                let new_row_num: usize = sheet_data
                    .get_elements("row")
                    .last()
                    .and_then(|last_row| last_row.get_attribute("r"))
                    .and_then(|r| r.parse::<usize>().ok())
                    .map_or(1, |num| num + 1);

                let mut row_element = XmlElement::new("row");
                row_element
                    .attributes
                    .insert("r".to_string(), new_row_num.to_string());

                for (i, cell_data) in row_data.iter().enumerate() {
                    let col_str: String = Self::col_to_string(i + 1);
                    let mut cell_element = XmlElement::new("c");
                    cell_element
                        .attributes
                        .insert("r".to_string(), format!("{col_str}{new_row_num}"));

                    // 共有文字列テーブルへの追加
                    let shared_string_id: usize = self.add_shared_string(cell_data);

                    cell_element
                        .attributes
                        .insert("t".to_string(), "s".to_string());
                    let mut v_element = XmlElement::new("v");
                    v_element.text = Some(shared_string_id.to_string());
                    cell_element.children.push(v_element);
                    row_element.children.push(cell_element);
                }
                sheet_data.children.push(row_element);
            }
        }
    }

    /// シート内の行のイテレータの取得
    pub fn iter_rows(&self) -> Vec<Vec<String>> {
        self.xml
            .lock()
            .ok()
            .and_then(|xml| {
                xml.elements.first().map(|worksheet| {
                    worksheet
                        .get_element("sheetData")
                        .get_elements("row")
                        .iter()
                        .map(|row| {
                            row.get_elements("c")
                                .iter()
                                .map(|cell| self.get_cell_value(cell).unwrap_or_default())
                                .collect()
                        })
                        .collect::<Vec<Vec<String>>>()
                })
            })
            .unwrap_or_default()
    }
}

impl Sheet {
    /// 新しい `Sheet` インスタンスの作成
    pub fn new(
        name: String,
        xml: Arc<Mutex<Xml>>,
        shared_strings: Arc<Mutex<Xml>>,
        shared_strings_map: Arc<Mutex<HashMap<String, usize>>>,
        styles: Arc<Mutex<Xml>>,
    ) -> Self {
        Sheet {
            name,
            xml,
            shared_strings,
            shared_strings_map,
            styles,
        }
    }

    /// 共有文字列テーブルへの文字列の追加
    fn add_shared_string(&self, s: &str) -> usize {
        if let Ok(mut map) = self.shared_strings_map.lock() {
            if let Some(&id) = map.get(s) {
                return id;
            }
            if let Ok(mut strings) = self.shared_strings.lock() {
                if strings.elements.is_empty() {
                    let mut sst = XmlElement::new("sst");
                    sst.attributes.insert(
                        "xmlns".to_string(),
                        "http://schemas.openxmlformats.org/spreadsheetml/2006/main".to_string(),
                    );
                    strings.elements.push(sst);
                }
                let sst = &mut strings.elements[0];
                let mut si = XmlElement::new("si");
                let mut t = XmlElement::new("t");
                t.text = Some(s.to_string());
                si.children.push(t);
                sst.children.push(si);

                let count = sst.children.len();
                sst.attributes
                    .insert("count".to_string(), count.to_string());
                sst.attributes
                    .insert("uniqueCount".to_string(), count.to_string());

                let new_id = count - 1;
                map.insert(s.to_string(), new_id);
                return new_id;
            }
        }
        0
    }

    #[cfg(test)]
    pub(crate) fn get_xml(&self) -> Arc<Mutex<Xml>> {
        self.xml.clone()
    }

    /// 行と列の番号のセルアドレス文字列への変換
    fn coordinate_to_string(row: usize, col: usize) -> String {
        // A1形式で返却
        format!("{}{}", Self::col_to_string(col), row)
    }

    /// 列番号のアルファベットへの変換
    fn col_to_string(col: usize) -> String {
        let mut result = String::new();
        let mut n = col;
        while n > 0 {
            let rem = (n - 1) % 26;
            result.insert(0, (b'A' + rem as u8) as char);
            n = (n - 1) / 26;
        }
        result
    }

    /// セルの値を取得
    fn get_cell_value(&self, cell: &XmlElement) -> Option<String> {
        let value_element = cell.get_element("v");
        let value = value_element.text.as_ref()?.clone();
        match cell.get_attribute("t").map(String::as_str) {
            Some("s") => {
                let index: usize = value.parse().ok()?;
                self.get_shared_string_by_index(index)
            }
            _ => Some(value),
        }
    }

    /// インデックスによる共有文字列の取得
    fn get_shared_string_by_index(&self, index: usize) -> Option<String> {
        self.shared_strings.lock().ok().and_then(|shared_strings| {
            shared_strings.elements.first().and_then(|sst| {
                sst.children
                    .get(index)
                    .and_then(|si| si.get_element("t").text.clone())
            })
        })
    }
}
