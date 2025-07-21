use crate::style::{Font, PatternFill};
use crate::xml::{Xml, XmlElement};
use chrono::{NaiveDate, NaiveDateTime};
use pyo3::prelude::*;
use std::sync::{Arc, Mutex, MutexGuard};

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
        let xml: MutexGuard<Xml> = self.sheet_xml.lock().expect("Failed to lock sheet xml");
        let worksheet: &XmlElement = xml.elements.first()?;
        let sheet_data: &XmlElement = worksheet.children.iter().find(|e| e.name == "sheetData")?;

        sheet_data
            .children
            .iter()
            .filter(|r| r.name == "row")
            .flat_map(|row| &row.children)
            .filter(|c| c.name == "c")
            .find(|cell_element| cell_element.attributes.get("r") == Some(&self.address))
            .map(|cell_element| {
                self.get_value_from_cell_element(cell_element)
                    .unwrap_or_default()
            })
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
        let font_id: usize = self.add_font_to_styles(&font);
        let fill_id: usize =
            self.add_fill_to_styles(self.fill.as_ref().unwrap_or(&PatternFill::default()));
        let xf_id: usize = self.add_xf_to_styles(font_id, fill_id, 0, 0);
        let mut xml: MutexGuard<Xml> = self.sheet_xml.lock().expect("Failed to lock sheet xml");
        let cell_element: &mut XmlElement = self.get_or_create_cell_element(&mut xml);
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
        let font_id: usize =
            self.add_font_to_styles(self.font.as_ref().unwrap_or(&Font::default()));
        let fill_id: usize = self.add_fill_to_styles(&fill);
        let xf_id: usize = self.add_xf_to_styles(font_id, fill_id, 0, 0);
        let mut xml: MutexGuard<Xml> = self.sheet_xml.lock().expect("Failed to lock sheet xml");
        let cell_element: &mut XmlElement = self.get_or_create_cell_element(&mut xml);
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
        let v_element: &XmlElement = cell_element.children.iter().find(|e| e.name == "v")?;
        let idx: usize = v_element.text.as_ref()?.parse::<usize>().ok()?;
        let shared_strings_xml: MutexGuard<Xml> = self
            .shared_strings
            .lock()
            .expect("Failed to lock shared strings");
        let sst: &XmlElement = shared_strings_xml.elements.first()?;
        let si: &XmlElement = sst.children.get(idx)?;
        si.children.first().and_then(|t| t.text.clone())
    }

    /// インライン文字列の値の取得
    fn get_inline_string_value(&self, cell_element: &XmlElement) -> Option<String> {
        let is_element: &XmlElement = cell_element.children.iter().find(|e| e.name == "is")?;
        let t_element: &XmlElement = is_element.children.iter().find(|e| e.name == "t")?;
        t_element.text.clone()
    }

    /// スタイルXMLへのフォントの追加とフォントIDの返却
    fn add_font_to_styles(&self, font: &Font) -> usize {
        let mut styles_xml: MutexGuard<Xml> = self.styles.lock().expect("Failed to lock styles");
        let fonts_tag: &mut XmlElement = styles_xml.get_mut_or_create_child_by_tag("fonts");

        if let Some(index) = fonts_tag.children.iter().position(|f| {
            let mut existing_font: Font = Font::default();
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

        let mut font_element: XmlElement = XmlElement::new("font");
        if let Some(name) = &font.name {
            let mut name_element: XmlElement = XmlElement::new("name");
            name_element
                .attributes
                .insert("val".to_string(), name.clone());
            font_element.children.push(name_element);
        }
        if let Some(size) = font.size {
            let mut size_element: XmlElement = XmlElement::new("sz");
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
            let mut color_element: XmlElement = XmlElement::new("color");
            color_element
                .attributes
                .insert("rgb".to_string(), color.clone());
            font_element.children.push(color_element);
        }

        fonts_tag.children.push(font_element);
        let count: usize = fonts_tag.children.len();
        fonts_tag
            .attributes
            .insert("count".to_string(), count.to_string());
        count - 1
    }

    /// スタイルXMLへの塗りつぶしの追加と塗りつぶしIDの返却
    fn add_fill_to_styles(&self, fill: &PatternFill) -> usize {
        let mut styles_xml: MutexGuard<Xml> = self.styles.lock().expect("Failed to lock styles");
        let fills_tag: &mut XmlElement = styles_xml.get_mut_or_create_child_by_tag("fills");

        let mut fill_element: XmlElement = XmlElement::new("fill");
        let mut pattern_fill_element: XmlElement = XmlElement::new("patternFill");

        if let Some(pattern_type) = &fill.pattern_type {
            pattern_fill_element
                .attributes
                .insert("patternType".to_string(), pattern_type.clone());
        }
        if let Some(fg_color) = &fill.fg_color {
            let mut fg_color_element: XmlElement = XmlElement::new("fgColor");
            fg_color_element
                .attributes
                .insert("rgb".to_string(), fg_color.clone());
            pattern_fill_element.children.push(fg_color_element);
        }
        if let Some(bg_color) = &fill.bg_color {
            let mut bg_color_element: XmlElement = XmlElement::new("bgColor");
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
        let count: usize = fills_tag.children.len();
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
        let mut styles_xml: MutexGuard<Xml> = self.styles.lock().expect("Failed to lock styles");
        let cell_xfs_tag: &mut XmlElement = styles_xml.get_mut_or_create_child_by_tag("cellXfs");

        if let Some(index) = cell_xfs_tag.children.iter().position(|xf| {
            let has_alignment: bool = xf.children.iter().any(|c| c.name == "alignment");
            let alignment_check: bool =
                (alignment_id > 0 && has_alignment) || (alignment_id == 0 && !has_alignment);

            xf.attributes.get("fontId") == Some(&font_id.to_string())
                && xf.attributes.get("fillId") == Some(&fill_id.to_string())
                && xf.attributes.get("borderId") == Some(&border_id.to_string())
                && alignment_check
        }) {
            return index;
        }

        let mut xf_element: XmlElement = XmlElement::new("xf");
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
        let count: usize = cell_xfs_tag.children.len();
        cell_xfs_tag
            .attributes
            .insert("count".to_string(), count.to_string());
        count - 1
    }

    /// セルの値の数値としての設定
    pub fn set_number_value(&mut self, value: f64) {
        let mut xml: MutexGuard<Xml> = self.sheet_xml.lock().expect("Failed to lock sheet xml");
        let cell_element: &mut XmlElement = self.get_or_create_cell_element(&mut xml);
        cell_element.attributes.remove("t");
        cell_element.children.retain(|c| c.name != "f");
        if let Some(v) = cell_element.children.iter_mut().find(|c| c.name == "v") {
            v.text = Some(value.to_string());
        } else {
            let mut v_element: XmlElement = XmlElement::new("v");
            v_element.text = Some(value.to_string());
            cell_element.children.push(v_element);
        }
    }

    /// セルの値の文字列としての設定
    pub fn set_string_value(&mut self, value: &str) {
        let sst_index: usize = self.get_or_create_shared_string(value);
        let mut xml: MutexGuard<Xml> = self.sheet_xml.lock().expect("Failed to lock sheet xml");
        let cell_element: &mut XmlElement = self.get_or_create_cell_element(&mut xml);
        cell_element
            .attributes
            .insert("t".to_string(), "s".to_string());
        cell_element.children.retain(|c| c.name != "f");
        if let Some(v) = cell_element.children.iter_mut().find(|c| c.name == "v") {
            v.text = Some(sst_index.to_string());
        } else {
            let mut v_element: XmlElement = XmlElement::new("v");
            v_element.text = Some(sst_index.to_string());
            cell_element.children.push(v_element);
        }
    }

    /// セルの値の日時としての設定
    pub fn set_datetime_value(&mut self, value: NaiveDateTime) {
        // 日時をExcelのシリアル値に変換
        // https://stackoverflow.com/questions/61546133/int-to-datetime-excel に基づく
        let excel_epoch: NaiveDateTime = NaiveDate::from_ymd_opt(1899, 12, 30)
            .expect("Invalid date")
            .and_hms_opt(0, 0, 0)
            .expect("Invalid time");
        let duration = value.signed_duration_since(excel_epoch);
        let serial: f64 = duration.num_seconds() as f64 / 86400.0;
        self.set_number_value(serial);
        // TODO: 日付フォーマットのスタイルを設定
    }

    /// セルの値のブール値としての設定
    pub fn set_bool_value(&mut self, value: bool) {
        let mut xml: MutexGuard<Xml> = self.sheet_xml.lock().expect("Failed to lock sheet xml");
        let cell_element: &mut XmlElement = self.get_or_create_cell_element(&mut xml);
        cell_element
            .attributes
            .insert("t".to_string(), "b".to_string());
        cell_element.children.retain(|c| c.name != "f");
        if let Some(v) = cell_element.children.iter_mut().find(|c| c.name == "v") {
            v.text = Some((if value { "1" } else { "0" }).to_string());
        } else {
            let mut v_element: XmlElement = XmlElement::new("v");
            v_element.text = Some((if value { "1" } else { "0" }).to_string());
            cell_element.children.push(v_element);
        }
    }

    /// セルの値の数式としての設定
    pub fn set_formula_value(&mut self, formula: &str) {
        let mut xml: MutexGuard<Xml> = self.sheet_xml.lock().expect("Failed to lock sheet xml");
        let cell_element: &mut XmlElement = self.get_or_create_cell_element(&mut xml);
        cell_element.attributes.remove("t");
        cell_element.children.retain(|c| c.name != "v");
        if let Some(f) = cell_element.children.iter_mut().find(|c| c.name == "f") {
            f.text = Some(formula.to_string());
        } else {
            let mut f_element: XmlElement = XmlElement::new("f");
            f_element.text = Some(formula.to_string());
            cell_element.children.push(f_element);
        }
    }

    /// ワークシートXML内のセル要素の取得または作成
    fn get_or_create_cell_element<'a>(&self, xml: &'a mut Xml) -> &'a mut XmlElement {
        let (row_num, _): (u32, u32) = self.decode_address();
        let sheet_data: &mut XmlElement = xml
            .elements
            .first_mut()
            .and_then(|ws| ws.children.iter_mut().find(|e| e.name == "sheetData"))
            .expect("sheetData not found in sheet xml");

        let row_index: usize =
            match sheet_data.children.iter().position(|r| {
                r.name == "row" && r.attributes.get("r") == Some(&row_num.to_string())
            }) {
                Some(idx) => idx,
                None => {
                    let mut new_row: XmlElement = XmlElement::new("row");
                    new_row
                        .attributes
                        .insert("r".to_string(), row_num.to_string());
                    sheet_data.children.push(new_row);
                    sheet_data.children.len() - 1
                }
            };

        let cell_index: usize = match sheet_data.children[row_index]
            .children
            .iter()
            .position(|c| c.name == "c" && c.attributes.get("r") == Some(&self.address))
        {
            Some(idx) => idx,
            None => {
                let mut new_cell: XmlElement = XmlElement::new("c");
                new_cell
                    .attributes
                    .insert("r".to_string(), self.address.clone());
                sheet_data.children[row_index].children.push(new_cell);
                sheet_data.children[row_index].children.len() - 1
            }
        };

        &mut sheet_data.children[row_index].children[cell_index]
    }

    /// 共有文字列XML内の共有文字列の取得または作成
    fn get_or_create_shared_string(&mut self, text: &str) -> usize {
        let mut shared_strings_xml: MutexGuard<Xml> = self
            .shared_strings
            .lock()
            .expect("Failed to lock shared strings");

        if shared_strings_xml.elements.is_empty() {
            shared_strings_xml.elements.push(XmlElement::new("sst"));
        }
        let sst_element: &mut XmlElement = shared_strings_xml
            .elements
            .first_mut()
            .expect("sst element not found in shared strings");

        if let Some(index) = sst_element
            .children
            .iter()
            .position(|si| si.children.first().and_then(|t| t.text.as_deref()) == Some(text))
        {
            return index;
        }

        let mut t_element: XmlElement = XmlElement::new("t");
        t_element.text = Some(text.to_string());
        let mut si_element: XmlElement = XmlElement::new("si");
        si_element.children.push(t_element);
        sst_element.children.push(si_element);
        sst_element.children.len() - 1
    }

    /// セルアドレス (例: "A1") の行と列の番号へのデコード
    fn decode_address(&self) -> (u32, u32) {
        let col_str: String = self.address.chars().filter(|c| c.is_alphabetic()).collect();
        let row_str: String = self
            .address
            .chars()
            .filter(|c| c.is_ascii_digit())
            .collect();
        let row: u32 = row_str
            .parse::<u32>()
            .expect("Invalid row number in address");
        let col: u32 = col_str
            .to_uppercase()
            .chars()
            .rev()
            .enumerate()
            .fold(0, |acc, (i, ch)| {
                acc + (ch as u32 - 'A' as u32 + 1) * 26u32.pow(i as u32)
            });
        (row, col)
    }
}
