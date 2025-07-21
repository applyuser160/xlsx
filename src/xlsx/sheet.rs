// Copyright (c) 2024-present, zcayh.
// All rights reserved.
//
// This source code is licensed under the MIT license found in the
// LICENSE file in the root directory of this source tree.

use std::sync::{Arc, Mutex};

use pyo3::prelude::*;

use crate::cell::Cell;
use crate::xml::Xml;

/// Excelワークブック内の単一のワークシートを表します。
///
/// この構造体は、アドレス（例：「A1」）または行と列のインデックスによって
/// セルにアクセスし、変更するためのインターフェースを提供します。
#[pyclass]
pub struct Sheet {
    /// ワークシートの名前。
    #[pyo3(get)]
    pub name: String,
    /// ワークシートのXMLデータへの共有参照。
    xml: Arc<Mutex<Xml>>,
    /// ワークブックの共有文字列テーブルへの共有参照。
    shared_strings: Arc<Mutex<Xml>>,
    /// ワークブックのスタイルへの共有参照。
    styles: Arc<Mutex<Xml>>,
}

#[pymethods]
impl Sheet {
    /// 角括弧表記（例：`sheet["A1"]`）を使用してセルにアクセスできます。
    ///
    /// # 引数
    ///
    /// * `key` - セルのアドレス（例：「A1」）。
    ///
    /// # 戻り値
    ///
    /// 指定されたアドレスの`Cell`インスタンス。
    pub fn __getitem__(&self, key: &str) -> Cell {
        Cell::new(
            self.xml.clone(),
            self.shared_strings.clone(),
            self.styles.clone(),
            key.to_string(),
        )
    }

    /// 行番号と列番号でセルにアクセスします。
    ///
    /// # 引数
    ///
    /// * `row` - 行番号（1から始まる）。
    /// * `column` - 列番号（1から始まる）。
    ///
    /// # 戻り値
    ///
    /// 指定された行と列の`Cell`インスタンス。
    #[pyo3(signature = (row, column))]
    pub fn cell(&self, row: usize, column: usize) -> Cell {
        let address = Self::coordinate_to_string(row, column);
        self.__getitem__(&address)
    }
}

impl Sheet {
    /// 新しい`Sheet`インスタンスを作成します。
    ///
    /// これは`Book`構造体によって内部的に使用されます。
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

    /// 行と列の番号をExcel形式のアドレス文字列（例：A1）に変換します。
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
