// Copyright (c) 2024-present, zcayh.
// All rights reserved.
//
// This source code is licensed under the MIT license found in the
// LICENSE file in the root directory of this source tree.

use pyo3::prelude::*;

/// Excelシートのセルのフォントプロパティを表します。
///
/// この構造体は、名前、サイズ、太字、イタリック、色など、
/// 詳細なフォントのカスタマイズを可能にします。
#[pyclass]
#[derive(Clone, Debug, PartialEq, Default)]
pub struct Font {
    /// フォント名（例：「Calibri」）。
    #[pyo3(get, set)]
    pub name: Option<String>,
    /// フォントのサイズ。
    #[pyo3(get, set)]
    pub size: Option<f64>,
    /// フォントが太字かどうかを示すブール値。
    #[pyo3(get, set)]
    pub bold: Option<bool>,
    /// フォントがイタリックかどうかを示すブール値。
    #[pyo3(get, set)]
    pub italic: Option<bool>,
    /// フォントの色をARGB形式で指定（例：「FF000000」）。
    #[pyo3(get, set)]
    pub color: Option<String>,
}

#[pymethods]
impl Font {
    /// オプションのプロパティを持つ新しい`Font`インスタンスを作成します。
    #[new]
    #[pyo3(signature = (name=None, size=None, bold=None, italic=None, color=None))]
    fn new(
        name: Option<String>,
        size: Option<f64>,
        bold: Option<bool>,
        italic: Option<bool>,
        color: Option<String>,
    ) -> Self {
        Self {
            name,
            size,
            bold,
            italic,
            color,
        }
    }
}

/// セルの罫線プロパティを表します。
///
/// この構造体は、セルの罫線の四方のスタイルと色を定義します。
#[pyclass]
#[derive(Clone, Debug, PartialEq, Default)]
pub struct Border {
    /// 左罫線のスタイル。
    #[pyo3(get, set)]
    pub left: Option<Side>,
    /// 右罫線のスタイル。
    #[pyo3(get, set)]
    pub right: Option<Side>,
    /// 上罫線のスタイル。
    #[pyo3(get, set)]
    pub top: Option<Side>,
    /// 下罫線のスタイル。
    #[pyo3(get, set)]
    pub bottom: Option<Side>,
}

#[pymethods]
impl Border {
    /// オプションの辺スタイルを持つ新しい`Border`インスタンスを作成します。
    #[new]
    #[pyo3(signature = (left=None, right=None, top=None, bottom=None))]
    fn new(
        left: Option<Side>,
        right: Option<Side>,
        top: Option<Side>,
        bottom: Option<Side>,
    ) -> Self {
        Self {
            left,
            right,
            top,
            bottom,
        }
    }
}

/// 罫線の片側のスタイルを表します。
#[pyclass]
#[derive(Clone, Debug, PartialEq, Default)]
pub struct Side {
    /// 罫線のスタイル（例：「thin」、「medium」、「thick」）。
    #[pyo3(get, set)]
    pub style: Option<String>,
    /// 罫線の色をARGB形式で指定。
    #[pyo3(get, set)]
    pub color: Option<String>,
}

#[pymethods]
impl Side {
    /// オプションのスタイルと色を持つ新しい`Side`インスタンスを作成します。
    #[new]
    #[pyo3(signature = (style=None, color=None))]
    fn new(style: Option<String>, color: Option<String>) -> Self {
        Self { style, color }
    }
}

/// セルのパターン塗りを表します。
///
/// この構造体は、セルの塗りつぶしパターン、前景色、および背景色を定義します。
#[pyclass]
#[derive(Clone, Debug, PartialEq, Default)]
pub struct PatternFill {
    /// パターンの種類（例：「solid」、「gray125」）。
    #[pyo3(get, set)]
    pub pattern_type: Option<String>,
    /// 塗りつぶしの前景色をARGB形式で指定。
    #[pyo3(get, set)]
    pub fg_color: Option<String>,
    /// 塗りつぶしの背景色をARGB形式で指定。
    #[pyo3(get, set)]
    pub bg_color: Option<String>,
}

#[pymethods]
impl PatternFill {
    /// オプションのプロパティを持つ新しい`PatternFill`インスタンスを作成します。
    #[new]
    #[pyo3(signature = (pattern_type=None, fg_color=None, bg_color=None))]
    fn new(
        pattern_type: Option<String>,
        fg_color: Option<String>,
        bg_color: Option<String>,
    ) -> Self {
        Self {
            pattern_type,
            fg_color,
            bg_color,
        }
    }
}
