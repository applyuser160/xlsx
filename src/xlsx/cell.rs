use crate::style::{Font, PatternFill};
use crate::xml::{Xml, XmlElement};
use chrono::NaiveDateTime;
use pyo3::prelude::*;
use std::sync::{Arc, Mutex};

/// ワークシートの単一セル
#[pyclass]
pub struct Cell {
    /// このセルが属するワークシートのXML
    sheet_xml: Arc<Mutex<Xml>>,
    /// 共有文字列のXML
    shared_strings: Arc<Mutex<Xml>>,
    /// スタイルのXML
    styles: Arc<Mutex<Xml>>,
    /// セルのアドレス (例: "A1")
    address: String,
    /// セルのフォント
    font: Option<Font>,
    /// セルの塗りつぶし
    fill: Option<PatternFill>,
}

#[pymethods]
impl Cell {
    /// セルの値の取得
    #[getter]
    pub fn value(&self) -> Option<String> {
        let xml = self.sheet_xml.lock().unwrap();
        let worksheet = xml.elements.first()?;
        let sheet_data = worksheet.children.iter().find(|e| e.name == "sheetData")?;

        sheet_data
            .children
            .iter()
            .filter(|r| r.name == "row")
            .flat_map(|row| row.children.iter().filter(|c| c.name == "c"))
            .find(|cell_element| cell_element.attributes.get("r") == Some(&self.address))
            .and_then(|cell_element| self.get_value_from_cell_element(cell_element))
    }

    /// セルの値の設定
    ///
    /// 値の型は自動的に検出
    #[setter]
    pub fn set_value(&mut self, value: String) {
        if let Some(formula) = value.strip_prefix('=') {
            self.set_formula_value(formula);
        } else if let Ok(number) = value.parse::<f64>() {
            self.set_number_value(number);
        } else if let Ok(boolean) = value.parse::<bool>() {
            self.set_bool_value(boolean);
        } else if let Ok(datetime) = NaiveDateTime::parse_from_str(&value, "%Y-%m-%d %H:%M:%S") {
            self.set_datetime_value(datetime);
        } else {
            self.set_string_value(&value);
        }
    }

    /// セルのフォントの取得
    #[getter]
    fn get_font(&self) -> PyResult<Option<Font>> {
        Ok(self.font.clone())
    }

    /// セルのフォントの設定
    #[setter]
    fn set_font(&mut self, font: Font) {
        self.font = Some(font.clone());
        let font_id = self.add_font_to_styles(&font);
        let fill_id = self.add_fill_to_styles(&self.fill.clone().unwrap_or_default());
        let xf_id = self.add_xf_to_styles(font_id, fill_id, 0, 0);
        let mut xml = self.sheet_xml.lock().unwrap();
        let cell_element = self.get_or_create_cell_element(&mut xml);
        cell_element
            .attributes
            .insert("s".to_string(), xf_id.to_string());
    }

    /// セルの塗りつぶしの取得
    #[getter]
    fn get_fill(&self) -> PyResult<Option<PatternFill>> {
        Ok(self.fill.clone())
    }

    /// セルの塗りつぶしの設定
    #[setter]
    fn set_fill(&mut self, fill: PatternFill) {
        self.fill = Some(fill.clone());
        let font_id = self.add_font_to_styles(&self.font.clone().unwrap_or_default());
        let fill_id = self.add_fill_to_styles(&fill);
        let xf_id = self.add_xf_to_styles(font_id, fill_id, 0, 0);
        let mut xml = self.sheet_xml.lock().unwrap();
        let cell_element = self.get_or_create_cell_element(&mut xml);
        cell_element
            .attributes
            .insert("s".to_string(), xf_id.to_string());
    }
}

impl Cell {
    /// 新しい `Cell` インスタンスの作成
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

    /// セル要素からの値の取得
    fn get_value_from_cell_element(&self, cell_element: &XmlElement) -> Option<String> {
        match cell_element.attributes.get("t").map(|s| s.as_str()) {
            Some("s") => self.get_shared_string_value(cell_element),
            Some("inlineStr") => self.get_inline_string_value(cell_element),
            _ => cell_element
                .children
                .iter()
                .find(|e| e.name == "v")
                .and_then(|v| v.text.clone()),
        }
    }

    /// 共有文字列の値の取得
    fn get_shared_string_value(&self, cell_element: &XmlElement) -> Option<String> {
        let v_element = cell_element.children.iter().find(|e| e.name == "v")?;
        let idx = v_element.text.as_ref()?.parse::<usize>().ok()?;
        let shared_strings_xml = self.shared_strings.lock().unwrap();
        let sst = shared_strings_xml.elements.first()?;
        let si = sst.children.get(idx)?;
        si.children.first().and_then(|t| t.text.clone())
    }

    /// インライン文字列の値の取得
    fn get_inline_string_value(&self, cell_element: &XmlElement) -> Option<String> {
        let is_element = cell_element.children.iter().find(|e| e.name == "is")?;
        let t_element = is_element.children.iter().find(|e| e.name == "t")?;
        t_element.text.clone()
    }

    /// スタイルXMLへのフォントの追加とフォントIDの返却
    fn add_font_to_styles(&self, font: &Font) -> usize {
        let mut styles_xml = self.styles.lock().unwrap();
        let fonts_tag = styles_xml.get_mut_or_create_child_by_tag("fonts");

        if let Some(index) = fonts_tag.children.iter().position(|f| {
            let mut existing_font = Font::default();
            for child in &f.children {
                match child.name.as_str() {
                    "name" => existing_font.name = child.attributes.get("val").cloned(),
                    "sz" => {
                        existing_font.size =
                            child.attributes.get("val").and_then(|s| s.parse().ok())
                    }
                    "b" => existing_font.bold = Some(true),
                    "i" => existing_font.italic = Some(true),
                    "color" => existing_font.color = child.attributes.get("rgb").cloned(),
                    _ => {}
                }
            }
            font == &existing_font
        }) {
            return index;
        }

        let mut font_element = XmlElement::new("font");
        if let Some(name) = &font.name {
            let mut name_element = XmlElement::new("name");
            name_element
                .attributes
                .insert("val".to_string(), name.clone());
            font_element.children.push(name_element);
        }
        if let Some(size) = font.size {
            let mut size_element = XmlElement::new("sz");
            size_element
                .attributes
                .insert("val".to_string(), size.to_string());
            font_element.children.push(size_element);
        }
        if font.bold.unwrap_or(false) {
            font_element.children.push(XmlElement::new("b"));
        }
        if font.italic.unwrap_or(false) {
            font_element.children.push(XmlElement::new("i"));
        }
        if let Some(color) = &font.color {
            let mut color_element = XmlElement::new("color");
            color_element
                .attributes
                .insert("rgb".to_string(), color.clone());
            font_element.children.push(color_element);
        }

        fonts_tag.children.push(font_element);
        let count = fonts_tag.children.len();
        fonts_tag
            .attributes
            .insert("count".to_string(), count.to_string());
        count - 1
    }

    /// スタイルXMLへの塗りつぶしの追加と塗りつぶしIDの返却
    fn add_fill_to_styles(&self, fill: &PatternFill) -> usize {
        let mut styles_xml = self.styles.lock().unwrap();
        let fills_tag = styles_xml.get_mut_or_create_child_by_tag("fills");

        let mut fill_element = XmlElement::new("fill");
        let mut pattern_fill_element = XmlElement::new("patternFill");

        if let Some(pattern_type) = &fill.pattern_type {
            pattern_fill_element
                .attributes
                .insert("patternType".to_string(), pattern_type.clone());
        }
        if let Some(fg_color) = &fill.fg_color {
            let mut fg_color_element = XmlElement::new("fgColor");
            fg_color_element
                .attributes
                .insert("rgb".to_string(), fg_color.clone());
            pattern_fill_element.children.push(fg_color_element);
        }
        if let Some(bg_color) = &fill.bg_color {
            let mut bg_color_element = XmlElement::new("bgColor");
            bg_color_element
                .attributes
                .insert("rgb".to_string(), bg_color.clone());
            pattern_fill_element.children.push(bg_color_element);
        }

        fill_element.children.push(pattern_fill_element);

        if let Some(index) = fills_tag.children.iter().position(|f| f == &fill_element) {
            return index;
        }

        fills_tag.children.push(fill_element);
        let count = fills_tag.children.len();
        fills_tag
            .attributes
            .insert("count".to_string(), count.to_string());
        count - 1
    }

    /// スタイルXMLへのcellXfsの追加とxf IDの返却
    fn add_xf_to_styles(
        &self,
        font_id: usize,
        fill_id: usize,
        border_id: usize,
        alignment_id: usize,
    ) -> usize {
        let mut styles_xml = self.styles.lock().unwrap();
        let cell_xfs_tag = styles_xml.get_mut_or_create_child_by_tag("cellXfs");

        if let Some(index) = cell_xfs_tag.children.iter().position(|xf| {
            let has_alignment = xf.children.iter().any(|c| c.name == "alignment");
            let alignment_check =
                (alignment_id > 0 && has_alignment) || (alignment_id == 0 && !has_alignment);

            xf.attributes.get("fontId") == Some(&font_id.to_string())
                && xf.attributes.get("fillId") == Some(&fill_id.to_string())
                && xf.attributes.get("borderId") == Some(&border_id.to_string())
                && alignment_check
        }) {
            return index;
        }

        let mut xf_element = XmlElement::new("xf");
        xf_element
            .attributes
            .insert("numFmtId".to_string(), "0".to_string());
        xf_element
            .attributes
            .insert("fontId".to_string(), font_id.to_string());
        xf_element
            .attributes
            .insert("fillId".to_string(), fill_id.to_string());
        xf_element
            .attributes
            .insert("borderId".to_string(), border_id.to_string());
        if font_id > 0 {
            xf_element
                .attributes
                .insert("applyFont".to_string(), "1".to_string());
        }
        if fill_id > 0 {
            xf_element
                .attributes
                .insert("applyFill".to_string(), "1".to_string());
        }
        if border_id > 0 {
            xf_element
                .attributes
                .insert("applyBorder".to_string(), "1".to_string());
        }
        if alignment_id > 0 {
            xf_element
                .attributes
                .insert("applyAlignment".to_string(), "1".to_string());
        }

        cell_xfs_tag.children.push(xf_element);
        let count = cell_xfs_tag.children.len();
        cell_xfs_tag
            .attributes
            .insert("count".to_string(), count.to_string());
        count - 1
    }

    /// セルの値の数値としての設定
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

    /// セルの値の文字列としての設定
    pub fn set_string_value(&mut self, value: &str) {
        let sst_index = self.get_or_create_shared_string(value);
        let mut xml = self.sheet_xml.lock().unwrap();
        let cell_element = self.get_or_create_cell_element(&mut xml);
        cell_element
            .attributes
            .insert("t".to_string(), "s".to_string());
        cell_element.children.retain(|c| c.name != "f");
        if let Some(v) = cell_element.children.iter_mut().find(|c| c.name == "v") {
            v.text = Some(sst_index.to_string());
        } else {
            let mut v_element = XmlElement::new("v");
            v_element.text = Some(sst_index.to_string());
            cell_element.children.push(v_element);
        }
    }

    /// セルの値の日時としての設定
    pub fn set_datetime_value(&mut self, value: NaiveDateTime) {
        // 日時をExcelのシリアル値に変換
        // https://stackoverflow.com/questions/61546133/int-to-datetime-excel に基づく
        let excel_epoch = chrono::NaiveDate::from_ymd_opt(1899, 12, 30)
            .unwrap()
            .and_hms_opt(0, 0, 0)
            .unwrap();
        let duration = value.signed_duration_since(excel_epoch);
        let serial = duration.num_seconds() as f64 / 86400.0;
        self.set_number_value(serial);
        // TODO: 日付フォーマットのスタイルを設定
    }

    /// セルの値のブール値としての設定
    pub fn set_bool_value(&mut self, value: bool) {
        let mut xml = self.sheet_xml.lock().unwrap();
        let cell_element = self.get_or_create_cell_element(&mut xml);
        cell_element
            .attributes
            .insert("t".to_string(), "b".to_string());
        cell_element.children.retain(|c| c.name != "f");
        if let Some(v) = cell_element.children.iter_mut().find(|c| c.name == "v") {
            v.text = Some((if value { "1" } else { "0" }).to_string());
        } else {
            let mut v_element = XmlElement::new("v");
            v_element.text = Some((if value { "1" } else { "0" }).to_string());
            cell_element.children.push(v_element);
        }
    }

    /// セルの値の数式としての設定
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

    /// ワークシートXML内のセル要素の取得または作成
    fn get_or_create_cell_element<'a>(&self, xml: &'a mut Xml) -> &'a mut XmlElement {
        let (row_num, _) = self.decode_address();
        let sheet_data = xml
            .elements
            .first_mut()
            .and_then(|ws| ws.children.iter_mut().find(|e| e.name == "sheetData"))
            .unwrap();

        let row_index = sheet_data
            .children
            .iter()
            .position(|r| r.name == "row" && r.attributes.get("r") == Some(&row_num.to_string()))
            .unwrap_or_else(|| {
                let mut new_row = XmlElement::new("row");
                new_row
                    .attributes
                    .insert("r".to_string(), row_num.to_string());
                sheet_data.children.push(new_row);
                sheet_data.children.len() - 1
            });

        let cell_index = sheet_data.children[row_index]
            .children
            .iter()
            .position(|c| c.name == "c" && c.attributes.get("r") == Some(&self.address))
            .unwrap_or_else(|| {
                let mut new_cell = XmlElement::new("c");
                new_cell
                    .attributes
                    .insert("r".to_string(), self.address.clone());
                sheet_data.children[row_index].children.push(new_cell);
                sheet_data.children[row_index].children.len() - 1
            });

        &mut sheet_data.children[row_index].children[cell_index]
    }

    /// 共有文字列XML内の共有文字列の取得または作成
    fn get_or_create_shared_string(&mut self, text: &str) -> usize {
        let mut shared_strings_xml = self.shared_strings.lock().unwrap();

        if shared_strings_xml.elements.is_empty() {
            shared_strings_xml.elements.push(XmlElement::new("sst"));
        }
        let sst_element = shared_strings_xml.elements.first_mut().unwrap();

        if let Some(index) = sst_element
            .children
            .iter()
            .position(|si| si.children.first().and_then(|t| t.text.as_deref()) == Some(text))
        {
            return index;
        }

        let mut t_element = XmlElement::new("t");
        t_element.text = Some(text.to_string());
        let mut si_element = XmlElement::new("si");
        si_element.children.push(t_element);
        sst_element.children.push(si_element);
        sst_element.children.len() - 1
    }

    /// セルアドレス (例: "A1") の行と列の番号へのデコード
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
