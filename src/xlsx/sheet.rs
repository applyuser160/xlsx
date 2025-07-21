use std::sync::{Arc, Mutex};

use pyo3::prelude::*;

use crate::cell::Cell;
use crate::xml::Xml;

/// Excelワークブック内のワークシートを表す
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
    /// アドレス (例: "A1") でセルを取得
    pub fn __getitem__(&self, key: &str) -> Cell {
        Cell::new(
            self.xml.clone(),
            self.shared_strings.clone(),
            self.styles.clone(),
            key.to_string(),
        )
    }

    /// 行と列の番号でセルを取得
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
    /// 新しい `Sheet` インスタンスを作成
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

    /// 行と列の番号をセルアドレス文字列に変換
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
