use pyo3::prelude::*;

/// Excelセルのフォントプロパティ表現
///
/// 名前、サイズ、太字、イタリック、色など、詳細なフォントカスタマイズの提供
#[pyclass]
#[derive(Clone, Debug, PartialEq, Default)]
pub struct Font {
    /// フォント名（例：「Calibri」）
    #[pyo3(get, set)]
    pub name: Option<String>,
    /// フォントサイズ
    #[pyo3(get, set)]
    pub size: Option<f64>,
    /// フォントが太字かどうかのブール値
    #[pyo3(get, set)]
    pub bold: Option<bool>,
    /// フォントがイタリックかどうかのブール値
    #[pyo3(get, set)]
    pub italic: Option<bool>,
    /// ARGB形式のフォント色（例：「FF000000」）
    #[pyo3(get, set)]
    pub color: Option<String>,
}

#[pymethods]
impl Font {
    /// オプションのプロパティを持つ新しい`Font`インスタンスの作成
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

/// セルの罫線プロパティ表現
///
/// セル罫線の四方のスタイルと色の定義
#[pyclass]
#[derive(Clone, Debug, PartialEq, Default)]
pub struct Border {
    /// 左罫線のスタイル
    #[pyo3(get, set)]
    pub left: Option<Side>,
    /// 右罫線のスタイル
    #[pyo3(get, set)]
    pub right: Option<Side>,
    /// 上罫線のスタイル
    #[pyo3(get, set)]
    pub top: Option<Side>,
    /// 下罫線のスタイル
    #[pyo3(get, set)]
    pub bottom: Option<Side>,
}

#[pymethods]
impl Border {
    /// オプションの辺スタイルを持つ新しい`Border`インスタンスの作成
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

/// 罫線の片側のスタイル表現
#[pyclass]
#[derive(Clone, Debug, PartialEq, Default)]
pub struct Side {
    /// 罫線のスタイル（例：「thin」、「medium」、「thick」）
    #[pyo3(get, set)]
    pub style: Option<String>,
    /// ARGB形式の罫線色
    #[pyo3(get, set)]
    pub color: Option<String>,
}

#[pymethods]
impl Side {
    /// オプションのスタイルと色を持つ新しい`Side`インスタンスの作成
    #[new]
    #[pyo3(signature = (style=None, color=None))]
    fn new(style: Option<String>, color: Option<String>) -> Self {
        Self { style, color }
    }
}

/// セルのパターン塗り表現
///
/// セルの塗りつぶしパターン、前景色、背景色の定義
#[pyclass]
#[derive(Clone, Debug, PartialEq, Default)]
pub struct PatternFill {
    /// パターンの種類（例：「solid」、「gray125」）
    #[pyo3(get, set)]
    pub pattern_type: Option<String>,
    /// ARGB形式の塗りつぶし前景色
    #[pyo3(get, set)]
    pub fg_color: Option<String>,
    /// ARGB形式の塗りつぶし背景色
    #[pyo3(get, set)]
    pub bg_color: Option<String>,
}

#[pymethods]
impl PatternFill {
    /// オプションのプロパティを持つ新しい`PatternFill`インスタンスの作成
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
