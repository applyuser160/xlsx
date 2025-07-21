use pyo3::prelude::*;

/// Represents the font properties for a cell.
#[pyclass]
#[derive(Clone, Debug, PartialEq, Default)]
pub struct Font {
    /// The name of the font.
    #[pyo3(get, set)]
    pub name: Option<String>,
    /// The size of the font.
    #[pyo3(get, set)]
    pub size: Option<f64>,
    /// Whether the font is bold.
    #[pyo3(get, set)]
    pub bold: Option<bool>,
    /// Whether the font is italic.
    #[pyo3(get, set)]
    pub italic: Option<bool>,
    /// The color of the font in ARGB format (e.g., "FF000000").
    #[pyo3(get, set)]
    pub color: Option<String>,
}

#[pymethods]
impl Font {
    /// Creates a new `Font` instance.
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

/// Represents the border properties for a cell.
#[pyclass]
#[derive(Clone, Debug, PartialEq, Default)]
pub struct Border {
    /// The left border.
    #[pyo3(get, set)]
    pub left: Option<Side>,
    /// The right border.
    #[pyo3(get, set)]
    pub right: Option<Side>,
    /// The top border.
    #[pyo3(get, set)]
    pub top: Option<Side>,
    /// The bottom border.
    #[pyo3(get, set)]
    pub bottom: Option<Side>,
}

#[pymethods]
impl Border {
    /// Creates a new `Border` instance.
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

/// Represents the properties of a single border side.
#[pyclass]
#[derive(Clone, Debug, PartialEq, Default)]
pub struct Side {
    /// The style of the border (e.g., "thin", "medium", "thick").
    #[pyo3(get, set)]
    pub style: Option<String>,
    /// The color of the border in ARGB format.
    #[pyo3(get, set)]
    pub color: Option<String>,
}

#[pymethods]
impl Side {
    /// Creates a new `Side` instance.
    #[new]
    #[pyo3(signature = (style=None, color=None))]
    fn new(style: Option<String>, color: Option<String>) -> Self {
        Side { style, color }
    }
}

/// Represents the pattern fill properties for a cell.
#[pyclass]
#[derive(Clone, Debug, PartialEq, Default)]
pub struct PatternFill {
    /// The type of the pattern (e.g., "solid", "gray125").
    #[pyo3(get, set)]
    pub pattern_type: Option<String>,
    /// The foreground color in ARGB format.
    #[pyo3(get, set)]
    pub fg_color: Option<String>,
    /// The background color in ARGB format.
    #[pyo3(get, set)]
    pub bg_color: Option<String>,
}

#[pymethods]
impl PatternFill {
    /// Creates a new `PatternFill` instance.
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
