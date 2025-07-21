use crate::cell::Cell;
use crate::xml::Xml;
use pyo3::prelude::*;
use std::sync::{Arc, Mutex};

/// Excelのワークシートを表します。
///
/// この構造体は、シート名へのアクセスや、セルを取得するためのメソッドを提供します。
#[pyclass]
pub struct Sheet {
    /// シート名。
    #[pyo3(get)]
    pub name: String,
    /// シートのXMLデータへの参照。
    xml: Arc<Mutex<Xml>>,
    /// 共有文字列テーブルへの参照。
    shared_strings: Arc<Mutex<Xml>>,
    /// スタイル情報への参照。
    styles: Arc<Mutex<Xml>>,
}

#[pymethods]
impl Sheet {
    /// アドレス文字列（例: "A1"）を指定して、特定のセルを取得します。
    /// Pythonの `sheet["A1"]` のようにアクセスできます。
    pub fn __getitem__(&self, key: &str) -> Cell {
        self.cell_from_address(key)
    }

    /// 行番号と列番号（1から始まる）を指定して、特定のセルを取得します。
    #[pyo3(signature = (row, column))]
    pub fn cell(&self, row: u32, column: u32) -> Cell {
        let address = Self::coordinates_to_address(row, column);
        self.cell_from_address(&address)
    }
}

impl Sheet {
    /// 新しい `Sheet` インスタンスを作成します（内部使用）。
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

    /// アドレス文字列から`Cell`インスタンスを生成するヘルパー関数。
    fn cell_from_address(&self, address: &str) -> Cell {
        Cell::new(
            self.xml.clone(),
            self.shared_strings.clone(),
            self.styles.clone(),
            address.to_string(),
        )
    }

    /// 行番号と列番号（1から始まる）をExcelのアドレス文字列（例: "A1"）に変換します。
    fn coordinates_to_address(row: u32, col: u32) -> String {
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
