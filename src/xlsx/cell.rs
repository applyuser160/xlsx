use crate::style::{Font, PatternFill};
use crate::xml::{Xml, XmlElement};
use chrono::{NaiveDateTime};
use pyo3::prelude::*;
use std::sync::{Arc, Mutex};

#[pyclass]
pub struct Cell {
    sheet_xml: Arc<Mutex<Xml>>,
    shared_strings: Arc<Mutex<Xml>>,
    styles: Arc<Mutex<Xml>>,
    address: String,
    font: Option<Font>,
    fill: Option<PatternFill>,
}

#[pymethods]
impl Cell {
    #[getter]
    pub fn value(&self) -> Option<String> {
        let xml = self.sheet_xml.lock().unwrap();
        if let Some(worksheet) = xml.elements.first() {
            if let Some(sheet_data) = worksheet.children.iter().find(|e| e.name == "sheetData") {
                for row in &sheet_data.children {
                    if row.name == "row" {
                        for cell_element in &row.children {
                            if cell_element.name == "c" {
                                if let Some(r_attr) = cell_element.attributes.get("r") {
                                    if r_attr == &self.address {
                                        if let Some(t_attr) = cell_element.attributes.get("t") {
                                            if t_attr == "s" {
                                                if let Some(v_element) = cell_element
                                                    .children
                                                    .iter()
                                                    .find(|e| e.name == "v")
                                                {
                                                    if let Some(text) = &v_element.text {
                                                        if let Ok(idx) = text.parse::<usize>() {
                                                            let shared_strings_xml =
                                                                self.shared_strings.lock().unwrap();
                                                            if let Some(sst) =
                                                                shared_strings_xml.elements.first()
                                                            {
                                                                if let Some(si) =
                                                                    sst.children.get(idx)
                                                                {
                                                                    if let Some(t) =
                                                                        si.children.first()
                                                                    {
                                                                        return t.text.clone();
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            } else if t_attr == "inlineStr" {
                                                if let Some(is_element) = cell_element
                                                    .children
                                                    .iter()
                                                    .find(|e| e.name == "is")
                                                {
                                                    if let Some(t_element) = is_element
                                                        .children
                                                        .iter()
                                                        .find(|e| e.name == "t")
                                                    {
                                                        return t_element.text.clone();
                                                    }
                                                }
                                            }
                                        }
                                        // t属性がないか、"s"以外の場合は、vタグの値を直接返す
                                        if let Some(v_element) =
                                            cell_element.children.iter().find(|e| e.name == "v")
                                        {
                                            return v_element.text.clone();
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        None
    }

    #[setter]
    pub fn set_value(&mut self, value: String) {
        // 数式かどうかを判定
        if value.starts_with('=') {
            self.set_formula_value(&value[1..]);
        // 日付時刻への変換を試みる
        } else if let Ok(datetime) = NaiveDateTime::parse_from_str(&value, "%Y-%m-%d %H:%M:%S") {
            self.set_datetime_value(datetime);
        // 数値への変換を試みる
        } else if let Ok(number) = value.parse::<f64>() {
            self.set_number_value(number);
        // ブール値への変換を試みる
        } else if let Ok(boolean) = value.parse::<bool>() {
            self.set_bool_value(boolean);
        } else {
            self.set_string_value(&value);
        }
    }

    #[getter]
    fn get_font(&self) -> PyResult<Option<Font>> {
        Ok(self.font.clone())
    }

    #[setter]
    fn set_font(&mut self, font: Font) {
        self.font = Some(font.clone());
        let font_id = self.add_font_to_styles(&font);
        let fill_id = self.add_fill_to_styles(&self.fill.clone().unwrap_or_default());
        let xf_id = self.add_xf_to_styles(font_id, fill_id, 0, 0);
        let mut xml = self.sheet_xml.lock().unwrap();
        let cell_element = self.get_or_create_cell_element(&mut xml);
        cell_element.attributes.insert("s".to_string(), xf_id.to_string());
    }

    #[getter]
    fn get_fill(&self) -> PyResult<Option<PatternFill>> {
        Ok(self.fill.clone())
    }

    #[setter]
    fn set_fill(&mut self, fill: PatternFill) {
        self.fill = Some(fill.clone());
        let font_id = self.add_font_to_styles(&self.font.clone().unwrap_or_default());
        let fill_id = self.add_fill_to_styles(&fill);
        let xf_id = self.add_xf_to_styles(font_id, fill_id, 0, 0);
        let mut xml = self.sheet_xml.lock().unwrap();
        let cell_element = self.get_or_create_cell_element(&mut xml);
        cell_element.attributes.insert("s".to_string(), xf_id.to_string());
    }
}

impl Cell {
    fn add_font_to_styles(&self, font: &Font) -> usize {
        let mut styles_xml = self.styles.lock().unwrap();
        let fonts_tag = styles_xml.get_mut_or_create_child_by_tag("fonts");

        // Check if the font already exists
        for (i, f) in fonts_tag.children.iter().enumerate() {
            let mut existing_font = Font { name: None, size: None, bold: None, italic: None, color: None };
            for child in &f.children {
                match child.name.as_str() {
                    "name" => existing_font.name = child.attributes.get("val").cloned(),
                    "sz" => existing_font.size = child.attributes.get("val").and_then(|s| s.parse().ok()),
                    "b" => existing_font.bold = Some(true),
                    "i" => existing_font.italic = Some(true),
                    "color" => existing_font.color = child.attributes.get("rgb").cloned(),
                    _ => {}
                }
            }
            if font == &existing_font {
                return i;
            }
        }

        let mut font_element = XmlElement::new("font");
        if let Some(name) = &font.name {
            let mut name_element = XmlElement::new("name");
            name_element.attributes.insert("val".to_string(), name.clone());
            font_element.children.push(name_element);
        }
        if let Some(size) = font.size {
            let mut size_element = XmlElement::new("sz");
            size_element.attributes.insert("val".to_string(), size.to_string());
            font_element.children.push(size_element);
        }
        if let Some(true) = font.bold {
            font_element.children.push(XmlElement::new("b"));
        }
        if let Some(true) = font.italic {
            font_element.children.push(XmlElement::new("i"));
        }
        if let Some(color) = &font.color {
            let mut color_element = XmlElement::new("color");
            color_element.attributes.insert("rgb".to_string(), color.clone());
            font_element.children.push(color_element);
        }

        fonts_tag.children.push(font_element);
        let count = fonts_tag.children.len();
        fonts_tag.attributes.insert("count".to_string(), count.to_string());
        count - 1
    }

    fn add_fill_to_styles(&self, fill: &PatternFill) -> usize {
        let mut styles_xml = self.styles.lock().unwrap();
        let fills_tag = styles_xml.get_mut_or_create_child_by_tag("fills");

        let mut fill_element = XmlElement::new("fill");
        let mut pattern_fill_element = XmlElement::new("patternFill");

        if let Some(pattern_type) = &fill.pattern_type {
            pattern_fill_element.attributes.insert("patternType".to_string(), pattern_type.clone());
        }
        if let Some(fg_color) = &fill.fg_color {
            let mut fg_color_element = XmlElement::new("fgColor");
            fg_color_element.attributes.insert("rgb".to_string(), fg_color.clone());
            pattern_fill_element.children.push(fg_color_element);
        }
        if let Some(bg_color) = &fill.bg_color {
            let mut bg_color_element = XmlElement::new("bgColor");
            bg_color_element.attributes.insert("rgb".to_string(), bg_color.clone());
            pattern_fill_element.children.push(bg_color_element);
        }

        fill_element.children.push(pattern_fill_element);
        fills_tag.children.push(fill_element);
        let count = fills_tag.children.len();
        fills_tag.attributes.insert("count".to_string(), count.to_string());
        count - 1
    }

    fn add_xf_to_styles(&self, font_id: usize, fill_id: usize, border_id: usize, alignment_id: usize) -> usize {
        let mut styles_xml = self.styles.lock().unwrap();
        let cell_xfs_tag = styles_xml.get_mut_or_create_child_by_tag("cellXfs");

        let mut xf_element = XmlElement::new("xf");
        xf_element.attributes.insert("numFmtId".to_string(), "0".to_string());
        xf_element.attributes.insert("fontId".to_string(), font_id.to_string());
        xf_element.attributes.insert("fillId".to_string(), fill_id.to_string());
        xf_element.attributes.insert("borderId".to_string(), border_id.to_string());
        if font_id > 0 {
            xf_element.attributes.insert("applyFont".to_string(), "1".to_string());
        }
        if fill_id > 0 {
            xf_element.attributes.insert("applyFill".to_string(), "1".to_string());
        }
        if border_id > 0 {
            xf_element.attributes.insert("applyBorder".to_string(), "1".to_string());
        }
        if alignment_id > 0 {
            xf_element.attributes.insert("applyAlignment".to_string(), "1".to_string());
        }

        // Check if the xf already exists
        for (i, xf) in cell_xfs_tag.children.iter().enumerate() {
            if xf.attributes.get("fontId") == Some(&font_id.to_string())
                && xf.attributes.get("fillId") == Some(&fill_id.to_string())
                && xf.attributes.get("borderId") == Some(&border_id.to_string()) {
                let has_alignment = xf.children.iter().any(|c| c.name == "alignment");
                if alignment_id > 0 && has_alignment {
                     return i;
                }
                if alignment_id == 0 && !has_alignment {
                    return i;
                }
            }
        }

        cell_xfs_tag.children.push(xf_element);
        let count = cell_xfs_tag.children.len();
        cell_xfs_tag.attributes.insert("count".to_string(), count.to_string());
        count - 1
    }

    pub fn new(
        sheet_xml: Arc<Mutex<Xml>>,
        shared_strings: Arc<Mutex<Xml>>,
        styles: Arc<Mutex<Xml>>,
        address: String,
    ) -> Self {
        Cell {
            sheet_xml,
            shared_strings,
            styles,
            address,
            font: None,
            fill: None,
        }
    }

    pub fn set_number_value(&mut self, value: f64) {
        let mut xml = self.sheet_xml.lock().unwrap();
        let cell_element = self.get_or_create_cell_element(&mut xml);
        cell_element.attributes.remove("t");
        cell_element.children.retain(|c| c.name != "f");
        if let Some(v) = cell_element.children.iter_mut().find(|c| c.name == "v") {
            v.text = Some(value.to_string());
        } else {
            let mut v_element = XmlElement::new("v");
            v_element.text = Some(value.to_string());
            cell_element.children.push(v_element);
        }
    }

    pub fn set_string_value(&mut self, value: &str) {
        let sst_index = self.get_or_create_shared_string(value);
        let mut xml = self.sheet_xml.lock().unwrap();
        let cell_element = self.get_or_create_cell_element(&mut xml);
        cell_element.attributes.insert("t".to_string(), "s".to_string());
        cell_element.children.retain(|c| c.name != "f");
        if let Some(v) = cell_element.children.iter_mut().find(|c| c.name == "v") {
            v.text = Some(sst_index.to_string());
        } else {
            let mut v_element = XmlElement::new("v");
            v_element.text = Some(sst_index.to_string());
            cell_element.children.push(v_element);
        }
    }

    pub fn set_datetime_value(&mut self, value: NaiveDateTime) {
        // Based on https://stackoverflow.com/questions/61546133/int-to-datetime-excel
        let excel_epoch = chrono::NaiveDate::from_ymd_opt(1899, 12, 30).unwrap().and_hms_opt(0,0,0).unwrap();
        let duration = value.signed_duration_since(excel_epoch);
        let serial = duration.num_seconds() as f64 / 86400.0;
        self.set_number_value(serial);
        // TODO: スタイルで日付フォーマットを設定する
    }

    pub fn set_bool_value(&mut self, value: bool) {
        let mut xml = self.sheet_xml.lock().unwrap();
        let cell_element = self.get_or_create_cell_element(&mut xml);
        cell_element.attributes.insert("t".to_string(), "b".to_string());
        cell_element.children.retain(|c| c.name != "f");
        if let Some(v) = cell_element.children.iter_mut().find(|c| c.name == "v") {
            v.text = Some((if value { "1" } else { "0" }).to_string());
        } else {
            let mut v_element = XmlElement::new("v");
            v_element.text = Some((if value { "1" } else { "0" }).to_string());
            cell_element.children.push(v_element);
        }
    }

    pub fn set_formula_value(&mut self, formula: &str) {
        let mut xml = self.sheet_xml.lock().unwrap();
        let cell_element = self.get_or_create_cell_element(&mut xml);
        cell_element.attributes.remove("t");
        cell_element.children.retain(|c| c.name != "v");
        if let Some(f) = cell_element.children.iter_mut().find(|c| c.name == "f") {
            f.text = Some(formula.to_string());
        } else {
            let mut f_element = XmlElement::new("f");
            f_element.text = Some(formula.to_string());
            cell_element.children.push(f_element);
        }
    }

    fn get_or_create_cell_element<'a>(&self, xml: &'a mut Xml) -> &'a mut XmlElement {
        let (row_num, _) = self.decode_address();
        let sheet_data = xml
            .elements
            .first_mut()
            .unwrap()
            .children
            .iter_mut()
            .find(|e| e.name == "sheetData")
            .unwrap();

        // Rowを探す
        let row_position = sheet_data
            .children
            .iter()
            .position(|r| r.name == "row" && r.attributes.get("r") == Some(&row_num.to_string()));

        // Rowがなければ作成
        let row_index = match row_position {
            Some(pos) => pos,
            None => {
                let mut new_row = XmlElement::new("row");
                new_row
                    .attributes
                    .insert("r".to_string(), row_num.to_string());
                sheet_data.children.push(new_row);
                sheet_data.children.len() - 1
            }
        };
        let row_element = &mut sheet_data.children[row_index];

        // Cellを探す
        let cell_position = row_element
            .children
            .iter()
            .position(|c| c.name == "c" && c.attributes.get("r") == Some(&self.address));

        // Cellがなければ作成
        let cell_index = match cell_position {
            Some(pos) => pos,
            None => {
                let mut new_cell = XmlElement::new("c");
                new_cell
                    .attributes
                    .insert("r".to_string(), self.address.clone());
                row_element.children.push(new_cell);
                row_element.children.len() - 1
            }
        };
        &mut row_element.children[cell_index]
    }

    fn get_or_create_shared_string(&mut self, text: &str) -> usize {
        let mut shared_strings_xml = self.shared_strings.lock().unwrap();

        // sst要素がなければ作成
        if shared_strings_xml.elements.is_empty() {
            let sst_element = XmlElement::new("sst");
            shared_strings_xml.elements.push(sst_element);
        }
        let sst_element = shared_strings_xml.elements.first_mut().unwrap();

        // 既存の文字列を探す
        for (i, si) in sst_element.children.iter().enumerate() {
            if let Some(t) = si.children.first() {
                if t.text.as_deref() == Some(text) {
                    return i;
                }
            }
        }

        // 新しい文字列を追加
        let mut t_element = XmlElement::new("t");
        t_element.text = Some(text.to_string());
        let mut si_element = XmlElement::new("si");
        si_element.children.push(t_element);
        sst_element.children.push(si_element);
        sst_element.children.len() - 1
    }

    fn decode_address(&self) -> (u32, u32) {
        let mut col_str = String::new();
        let mut row_str = String::new();
        for ch in self.address.chars() {
            if ch.is_alphabetic() {
                col_str.push(ch);
            } else {
                row_str.push(ch);
            }
        }
        let row = row_str.parse::<u32>().unwrap();
        let mut col = 0;
        for (i, ch) in col_str.to_uppercase().chars().rev().enumerate() {
            col += (ch as u32 - 'A' as u32 + 1) * 26u32.pow(i as u32);
        }
        (row, col)
    }
}
