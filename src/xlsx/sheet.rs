use std::sync::{Arc, Mutex};

use pyo3::prelude::*;

use crate::cell::Cell;
use crate::xml::Xml;

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
        let mut xml = self.xml.lock().unwrap();
        let worksheet = &mut xml.elements[0];
        let sheet_data = worksheet.get_element_mut("sheetData");
        let new_row_num = if let Some(last_row) = sheet_data.get_elements("row").last() {
            last_row
                .get_attribute("r")
                .unwrap()
                .parse::<usize>()
                .unwrap()
                + 1
        } else {
            1
        };

        let mut row_element = XmlElement::new("row");
        row_element
            .attributes
            .insert("r".to_string(), new_row_num.to_string());

        for (i, cell_data) in row_data.iter().enumerate() {
            let col_str = Self::col_to_string(i + 1);
            let mut cell_element = XmlElement::new("c");
            cell_element
                .attributes
                .insert("r".to_string(), format!("{col_str}{new_row_num}"));
            cell_element
                .attributes
                .insert("t".to_string(), "inlineStr".to_string());

            let mut is_element = XmlElement::new("is");
            let mut t_element = XmlElement::new("t");
            t_element.text = Some(cell_data.clone());
            is_element.children.push(t_element);
            cell_element.children.push(is_element);
            row_element.children.push(cell_element);
        }
        sheet_data.children.push(row_element);
    }

    /// シート内の行のイテレータの取得
    #[pyo3(signature = (values_only = false))]
    pub fn iter_rows(&self, values_only: bool) -> PyResult<Vec<Vec<String>>> {
        let xml = self.xml.lock().unwrap();
        let worksheet = &xml.elements[0];
        let sheet_data = worksheet.get_element("sheetData");
        let rows = sheet_data.get_elements("row");
        let mut result = Vec::new();

        for row in rows {
            let mut row_data = Vec::new();
            let cells = row.get_elements("c");
            for cell in cells {
                let value = if values_only {
                    let val = cell.get_element("is>t").get_text();
                    val.to_string().to_owned()
                } else {
                    // NOTE:現時点ではCellオブジェクトは返さず、値のみを返す
                    let val = cell.get_element("is>t").get_text();
                    val.to_string().to_owned()
                };
                row_data.push(value);
            }
            result.push(row_data);
        }
        Ok(result)
    }
}

impl Sheet {
    /// 新しい `Sheet` インスタンスの作成
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
