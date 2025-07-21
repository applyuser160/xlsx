use pyo3::prelude::*;

#[pyclass]
#[derive(Clone, Debug, PartialEq, Default)]
pub struct Font {
    #[pyo3(get, set)]
    pub name: Option<String>,
    #[pyo3(get, set)]
    pub size: Option<f64>,
    #[pyo3(get, set)]
    pub bold: Option<bool>,
    #[pyo3(get, set)]
    pub italic: Option<bool>,
    #[pyo3(get, set)]
    pub color: Option<String>,
}

#[pymethods]
impl Font {
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

#[pyclass]
#[derive(Clone, Debug, PartialEq, Default)]
pub struct Border {
    #[pyo3(get, set)]
    pub left: Option<Side>,
    #[pyo3(get, set)]
    pub right: Option<Side>,
    #[pyo3(get, set)]
    pub top: Option<Side>,
    #[pyo3(get, set)]
    pub bottom: Option<Side>,
}

#[pymethods]
impl Border {
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

#[pyclass]
#[derive(Clone, Debug, PartialEq, Default)]
pub struct Side {
    #[pyo3(get, set)]
    pub style: Option<String>,
    #[pyo3(get, set)]
    pub color: Option<String>,
}

#[pymethods]
impl Side {
    #[new]
    #[pyo3(signature = (style=None, color=None))]
    fn new(style: Option<String>, color: Option<String>) -> Self {
        Side { style, color }
    }
}

#[pyclass]
#[derive(Clone, Debug, PartialEq, Default)]
pub struct PatternFill {
    #[pyo3(get, set)]
    pub pattern_type: Option<String>,
    #[pyo3(get, set)]
    pub fg_color: Option<String>,
    #[pyo3(get, set)]
    pub bg_color: Option<String>,
}

#[pymethods]
impl PatternFill {
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
