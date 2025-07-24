use std::sync::{Arc, Mutex};

use pyo3::prelude::*;

use crate::cell::Cell;
use crate::xml::Xml;

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
        use crate::xml::XmlElement;
        let mut xml: std::sync::MutexGuard<Xml> = self.xml.lock().unwrap();
        let worksheet: &mut crate::xml::XmlElement = &mut xml.elements[0];
        let sheet_data: &mut crate::xml::XmlElement = worksheet.get_element_mut("sheetData");
        let new_row_num: usize = if let Some(last_row) = sheet_data.get_elements("row").last() {
            last_row
                .get_attribute("r")
                .unwrap()
                .parse::<usize>()
                .unwrap()
                + 1
        } else {
            1
        };

        let mut row_element: XmlElement = XmlElement::new("row");
        row_element
            .attributes
            .insert("r".to_string(), new_row_num.to_string());

        for (i, cell_data) in row_data.iter().enumerate() {
            let col_str: String = Self::col_to_string(i + 1);
            let mut cell_element: XmlElement = XmlElement::new("c");
            cell_element
                .attributes
                .insert("r".to_string(), format!("{col_str}{new_row_num}"));

            // 共有文字列テーブルへの追加
            let shared_string_id: usize = self.add_shared_string(&cell_data.to_string());

            cell_element
                .attributes
                .insert("t".to_string(), "s".to_string());
            let mut v_element: XmlElement = XmlElement::new("v");
            v_element.text = Some(shared_string_id.to_string());
            cell_element.children.push(v_element);
            row_element.children.push(cell_element);
        }
        sheet_data.children.push(row_element);
    }

    /// シート内の行のイテレータの取得
    pub fn iter_rows(&self) -> Vec<Vec<Cell>> {
        let xml: std::sync::MutexGuard<Xml> = self.xml.lock().unwrap();
        let worksheet: &crate::xml::XmlElement = &xml.elements[0];
        let sheet_data: &crate::xml::XmlElement = worksheet.get_element("sheetData");
        let rows: Vec<crate::xml::XmlElement> = sheet_data
            .get_elements("row")
            .iter()
            .map(|&x| x.clone())
            .collect();
        rows.iter()
            .map(|row| {
                row.get_elements("c")
                    .iter()
                    .map(|cell| {
                        let address: String = cell.get_attribute("r").unwrap().to_string();
                        Cell::new(
                            self.xml.clone(),
                            self.shared_strings.clone(),
                            self.styles.clone(),
                            address,
                        )
                    })
                    .collect()
            })
            .collect()
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
        let mut shared_strings_map = self.shared_strings_map.lock().unwrap();
        if let Some(&id) = shared_strings_map.get(s) {
            return id;
        }

        let mut shared_strings = self.shared_strings.lock().unwrap();
        if shared_strings.elements.is_empty() {
            use crate::xml::XmlElement;
            let mut sst = XmlElement::new("sst");
            sst.attributes.insert(
                "xmlns".to_string(),
                "http://schemas.openxmlformats.org/spreadsheetml/2006/main".to_string(),
            );
            shared_strings.elements.push(sst);
        }
        let sst = &mut shared_strings.elements[0];

        // 新しい文字列の追加
        use crate::xml::XmlElement;
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
        shared_strings_map.insert(s.to_string(), new_id);
        new_id
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
}
