use pyo3::prelude::*;

/// セルのフォントプロパティ
#[pyclass]
#[derive(Clone, Debug, PartialEq, Default)]
pub struct Font {
    /// フォント名
    #[pyo3(get, set)]
    pub name: Option<String>,
    /// フォントサイズ
    #[pyo3(get, set)]
    pub size: Option<f64>,
    /// 太字
    #[pyo3(get, set)]
    pub bold: Option<bool>,
    /// 斜体
    #[pyo3(get, set)]
    pub italic: Option<bool>,
    /// ARGB形式のフォントの色 (例: "FF000000")
    #[pyo3(get, set)]
    pub color: Option<String>,
}

#[pymethods]
impl Font {
    /// 新しい `Font` インスタンスの作成
    #[new]
    #[pyo3(signature = (name=None, size=None, bold=None, italic=None, color=None))]
    fn new(
        name: Option<String>,
        size: Option<f64>,
        bold: Option<bool>,
        italic: Option<bool>,
        color: Option<String>,
    ) -> Self {
        Font {
            name,
            size,
            bold,
            italic,
            color,
        }
    }
}

/// セルの罫線プロパティ
#[pyclass]
#[derive(Clone, Debug, PartialEq, Default)]
pub struct Border {
    /// 左罫線
    #[pyo3(get, set)]
    pub left: Option<Side>,
    /// 右罫線
    #[pyo3(get, set)]
    pub right: Option<Side>,
    /// 上罫線
    #[pyo3(get, set)]
    pub top: Option<Side>,
    /// 下罫線
    #[pyo3(get, set)]
    pub bottom: Option<Side>,
}

#[pymethods]
impl Border {
    /// 新しい `Border` インスタンスの作成
    #[new]
    #[pyo3(signature = (left=None, right=None, top=None, bottom=None))]
    fn new(
        left: Option<Side>,
        right: Option<Side>,
        top: Option<Side>,
        bottom: Option<Side>,
    ) -> Self {
        Border {
            left,
            right,
            top,
            bottom,
        }
    }
}

/// 1つの罫線のプロパティ
#[pyclass]
#[derive(Clone, Debug, PartialEq, Default)]
pub struct Side {
    /// 罫線のスタイル (例: "thin", "medium", "thick")
    #[pyo3(get, set)]
    pub style: Option<String>,
    /// ARGB形式の罫線の色
    #[pyo3(get, set)]
    pub color: Option<String>,
}

#[pymethods]
impl Side {
    /// 新しい `Side` インスタンスの作成
    #[new]
    #[pyo3(signature = (style=None, color=None))]
    fn new(style: Option<String>, color: Option<String>) -> Self {
        Side { style, color }
    }
}

/// セルの塗りつぶしパターンプロパティ
#[pyclass]
#[derive(Clone, Debug, PartialEq, Default)]
pub struct PatternFill {
    /// パターンの種類 (例: "solid", "gray125")
    #[pyo3(get, set)]
    pub pattern_type: Option<String>,
    /// ARGB形式の前面色
    #[pyo3(get, set)]
    pub fg_color: Option<String>,
    /// ARGB形式の背景色
    #[pyo3(get, set)]
    pub bg_color: Option<String>,
}

#[pymethods]
impl PatternFill {
    /// 新しい `PatternFill` インスタンスの作成
    #[new]
    #[pyo3(signature = (pattern_type=None, fg_color=None, bg_color=None))]
    fn new(
        pattern_type: Option<String>,
        fg_color: Option<String>,
        bg_color: Option<String>,
    ) -> Self {
        PatternFill {
            pattern_type,
            fg_color,
            bg_color,
        }
    }
}
