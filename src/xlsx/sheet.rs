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
        // 列番号をアルファベットに変換
        let col_str: String =
            std::iter::successors(
                Some(col),
                |&c| {
                    if c > 0 {
                        Some((c - 1) / 26)
                    } else {
                        None
                    }
                },
            )
            .take_while(|&c| c > 0)
            .map(|c| ((c - 1) % 26) as u8 + b'A')
            .map(|c| c as char)
            .collect::<String>()
            .chars()
            .rev()
            .collect();
        // A1形式で返却
        format!("{col_str}{row}")
    }
}
