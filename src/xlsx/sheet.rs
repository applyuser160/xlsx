use std::sync::{Arc, Mutex};

use pyo3::prelude::*;

use crate::cell::Cell;
use crate::xml::Xml;

/// Excelワークブック内の単一ワークシートの表現
///
/// アドレス（例：「A1」）または行と列のインデックスによる
/// セルへのアクセスと変更インターフェースの提供
#[pyclass]
pub struct Sheet {
    /// ワークシート名
    #[pyo3(get)]
    pub name: String,
    /// ワークシートXMLデータへの共有参照
    xml: Arc<Mutex<Xml>>,
    /// ワークブック共有文字列テーブルへの共有参照
    shared_strings: Arc<Mutex<Xml>>,
    /// ワークブックスタイルへの共有参照
    styles: Arc<Mutex<Xml>>,
}

#[pymethods]
impl Sheet {
    /// 角括弧表記（例：`sheet["A1"]`）によるセルへのアクセス
    ///
    /// # 引数
    ///
    /// * `key` - セルアドレス（例：「A1」）
    ///
    /// # 戻り値
    ///
    /// 指定されたアドレスの`Cell`インスタンス
    pub fn __getitem__(&self, key: &str) -> Cell {
        Cell::new(
            self.xml.clone(),
            self.shared_strings.clone(),
            self.styles.clone(),
            key.to_string(),
        )
    }

    /// 行番号と列番号によるセルへのアクセス
    ///
    /// # 引数
    ///
    /// * `row` - 行番号（1から始まる）
    /// * `column` - 列番号（1から始まる）
    ///
    /// # 戻り値
    ///
    /// 指定された行と列の`Cell`インスタンス
    #[pyo3(signature = (row, column))]
    pub fn cell(&self, row: usize, column: usize) -> Cell {
        let address = Self::coordinate_to_string(row, column);
        self.__getitem__(&address)
    }
}

impl Sheet {
    /// 新しい`Sheet`インスタンスの作成
    ///
    /// `Book`構造体による内部的な使用
    pub fn new(
        name: String,
        xml: Arc<Mutex<Xml>>,
        shared_strings: Arc<Mutex<Xml>>,
        styles: Arc<Mutex<Xml>>,
    ) -> Self {
        Self {
            name,
            xml,
            shared_strings,
            styles,
        }
    }

    /// 行と列番号のExcel形式アドレス文字列（例：A1）への変換
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

    #[cfg(test)]
    pub(crate) fn get_xml(&self) -> Arc<Mutex<Xml>> {
        self.xml.clone()
    }
}
