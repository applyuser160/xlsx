// Copyright (c) 2024-present, zcayh.
// All rights reserved.
//
// This source code is licensed under the MIT license found in the
// LICENSE file in the root directory of this source tree.

use crate::style::{Font, PatternFill};
use crate::xml::{Xml, XmlElement};
use chrono::NaiveDateTime;
use pyo3::prelude::*;
use std::sync::{Arc, Mutex};

/// ワークシート内の単一のセルを表します。
///
/// この構造体は、セルの値とスタイルを読み書きするためのインターフェースを提供します。
/// セルデータを管理するために、シートのXML、共有文字列、およびスタイルへの参照を保持します。
#[pyclass]
pub struct Cell {
    /// ワークシートのXML構造への共有参照。
    sheet_xml: Arc<Mutex<Xml>>,
    /// ワークブックの共有文字列テーブルへの共有参照。
    shared_strings: Arc<Mutex<Xml>>,
    /// ワークブックのスタイルへの共有参照。
    styles: Arc<Mutex<Xml>>,
    /// セルのアドレス（例：「A1」）。
    address: String,
    /// セルのフォントスタイル。
    font: Option<Font>,
    /// セルの塗りつぶしスタイル。
    fill: Option<PatternFill>,
}

#[pymethods]
impl Cell {
    /// セルの値を文字列として取得します。
    ///
    /// このメソッドは、ワークシートXMLからセルの値を取得します。共有文字列、
    /// インライン文字列、数値など、さまざまなセルタイプを処理します。
    #[getter]
    pub fn value(&self) -> Option<String> {
        let xml = self.sheet_xml.lock().unwrap();
        let sheet_data = xml
            .elements
            .first()?
            .children
            .iter()
            .find(|e| e.name == "sheetData")?;

        for row in &sheet_data.children {
            if row.name != "row" {
                continue;
            }
            for cell_element in &row.children {
                if cell_element.name == "c"
                    && cell_element.attributes.get("r") == Some(&self.address)
                {
                    return self.extract_cell_value(cell_element);
                }
            }
        }
        None
    }

    /// セルの値を設定します。
    ///
    /// このメソッドは、値の型（数式、日時、数値、ブール値、文字列など）を
    /// インテリジェントに判断し、適切なセッターを呼び出します。
    #[setter]
    pub fn set_value(&mut self, value: String) {
        if let Some(formula) = value.strip_prefix('=') {
            self.set_formula_value(formula);
        } else if let Ok(datetime) = NaiveDateTime::parse_from_str(&value, "%Y-%m-%d %H:%M:%S") {
            self.set_datetime_value(datetime);
        } else if let Ok(number) = value.parse::<f64>() {
            self.set_number_value(number);
        } else if let Ok(boolean) = value.parse::<bool>() {
            self.set_bool_value(boolean);
        } else {
            self.set_string_value(&value);
        }
    }

    /// セルのフォントを取得します。
    #[getter]
    fn get_font(&self) -> PyResult<Option<Font>> {
        Ok(self.font.clone())
    }

    /// セルのフォントを設定します。
    ///
    /// このメソッドは、セルに`Font`スタイルを適用し、ワークブックに必要な
    /// スタイルレコードを作成または更新します。
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

    /// セルの塗りつぶしを取得します。
    #[getter]
    fn get_fill(&self) -> PyResult<Option<PatternFill>> {
        Ok(self.fill.clone())
    }

    /// セルの塗りつぶしを設定します。
    ///
    /// このメソッドは、セルに`PatternFill`を適用し、ワークブックに必要な
    /// スタイルレコードを作成または更新します。
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
    /// 新しいフォントをスタイルXMLに追加するか、既存のフォントのインデックスを返します。
    fn add_font_to_styles(&self, font: &Font) -> usize {
        let mut styles_xml = self.styles.lock().unwrap();
        let fonts_tag = styles_xml.get_mut_or_create_child_by_tag("fonts");

        // フォントが既に存在するか確認
        for (i, f) in fonts_tag.children.iter().enumerate() {
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
            if font == &existing_font {
                return i;
            }
        }

        // 新しいフォント要素を作成
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

    /// 新しい塗りつぶしをスタイルXMLに追加するか、既存の塗りつぶしのインデックスを返します。
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
        fills_tag.children.push(fill_element);
        let count = fills_tag.children.len();
        fills_tag
            .attributes
            .insert("count".to_string(), count.to_string());
        count - 1
    }

    /// 新しいセル書式（xf）をスタイルXMLに追加するか、既存の書式のインデックスを返します。
    fn add_xf_to_styles(
        &self,
        font_id: usize,
        fill_id: usize,
        border_id: usize,
        alignment_id: usize,
    ) -> usize {
        let mut styles_xml = self.styles.lock().unwrap();
        let cell_xfs_tag = styles_xml.get_mut_or_create_child_by_tag("cellXfs");

        // 適切なxfが既に存在するか確認
        for (i, xf) in cell_xfs_tag.children.iter().enumerate() {
            let has_alignment = xf.children.iter().any(|c| c.name == "alignment");
            let font_match = xf.attributes.get("fontId") == Some(&font_id.to_string());
            let fill_match = xf.attributes.get("fillId") == Some(&fill_id.to_string());
            let border_match = xf.attributes.get("borderId") == Some(&border_id.to_string());

            if font_match && fill_match && border_match && ((alignment_id > 0) == has_alignment) {
                return i;
            }
        }

        // 新しいxf要素を作成
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

    /// 新しい`Cell`インスタンスを作成します。
    pub fn new(
        sheet_xml: Arc<Mutex<Xml>>,
        shared_strings: Arc<Mutex<Xml>>,
        styles: Arc<Mutex<Xml>>,
        address: String,
    ) -> Self {
        Self {
            sheet_xml,
            shared_strings,
            styles,
            address,
            font: None,
            fill: None,
        }
    }

    /// セルの値を数値に設定します。
    pub fn set_number_value(&mut self, value: f64) {
        let mut xml = self.sheet_xml.lock().unwrap();
        let cell_element = self.get_or_create_cell_element(&mut xml);
        cell_element.attributes.remove("t");
        cell_element.children.retain(|c| c.name != "f");
        self.update_or_create_child_element(cell_element, "v", value.to_string());
    }

    /// セルの値を共有文字列テーブルを使用して文字列に設定します。
    pub fn set_string_value(&mut self, value: &str) {
        let sst_index = self.get_or_create_shared_string(value);
        let mut xml = self.sheet_xml.lock().unwrap();
        let cell_element = self.get_or_create_cell_element(&mut xml);
        cell_element
            .attributes
            .insert("t".to_string(), "s".to_string());
        cell_element.children.retain(|c| c.name != "f");
        self.update_or_create_child_element(cell_element, "v", sst_index.to_string());
    }

    /// セルの値を日時に設定し、Excelのシリアル値に変換します。
    pub fn set_datetime_value(&mut self, value: NaiveDateTime) {
        // Excelのエポックは1899-12-30から始まります。
        let excel_epoch = chrono::NaiveDate::from_ymd_opt(1899, 12, 30)
            .unwrap()
            .and_hms_opt(0, 0, 0)
            .unwrap();
        let duration = value.signed_duration_since(excel_epoch);
        let serial = duration.num_seconds() as f64 / 86400.0;
        self.set_number_value(serial);
        // TODO: 日付書式スタイルを適用する。
    }

    /// セルの値をブール値に設定します。
    pub fn set_bool_value(&mut self, value: bool) {
        let mut xml = self.sheet_xml.lock().unwrap();
        let cell_element = self.get_or_create_cell_element(&mut xml);
        cell_element
            .attributes
            .insert("t".to_string(), "b".to_string());
        cell_element.children.retain(|c| c.name != "f");
        self.update_or_create_child_element(
            cell_element,
            "v",
            (if value { "1" } else { "0" }).to_string(),
        );
    }

    /// セルの値を数式に設定します。
    pub fn set_formula_value(&mut self, formula: &str) {
        let mut xml = self.sheet_xml.lock().unwrap();
        let cell_element = self.get_or_create_cell_element(&mut xml);
        cell_element.attributes.remove("t");
        cell_element.children.retain(|c| c.name != "v");
        self.update_or_create_child_element(cell_element, "f", formula.to_string());
    }

    /// このセルの`XmlElement`を取得または作成します。
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

        // 行要素を検索または作成
        let row_position = sheet_data
            .children
            .iter()
            .position(|r| r.name == "row" && r.attributes.get("r") == Some(&row_num.to_string()));
        let row_index = row_position.unwrap_or_else(|| {
            let mut new_row = XmlElement::new("row");
            new_row
                .attributes
                .insert("r".to_string(), row_num.to_string());
            sheet_data.children.push(new_row);
            sheet_data.children.len() - 1
        });
        let row_element = &mut sheet_data.children[row_index];

        // セル要素を検索または作成
        let cell_position = row_element
            .children
            .iter()
            .position(|c| c.name == "c" && c.attributes.get("r") == Some(&self.address));
        let cell_index = cell_position.unwrap_or_else(|| {
            let mut new_cell = XmlElement::new("c");
            new_cell
                .attributes
                .insert("r".to_string(), self.address.clone());
            row_element.children.push(new_cell);
            row_element.children.len() - 1
        });

        &mut row_element.children[cell_index]
    }

    /// 親要素内の子要素を更新または作成します。
    fn update_or_create_child_element(&self, parent: &mut XmlElement, name: &str, text: String) {
        if let Some(child) = parent.children.iter_mut().find(|c| c.name == name) {
            child.text = Some(text);
        } else {
            let mut new_child = XmlElement::new(name);
            new_child.text = Some(text);
            parent.children.push(new_child);
        }
    }

    /// 共有文字列テーブル内の文字列のインデックスを取得し、存在しない場合は作成します。
    fn get_or_create_shared_string(&mut self, text: &str) -> usize {
        let mut shared_strings_xml = self.shared_strings.lock().unwrap();

        if shared_strings_xml.elements.is_empty() {
            shared_strings_xml.elements.push(XmlElement::new("sst"));
        }
        let sst_element = shared_strings_xml.elements.first_mut().unwrap();

        // 既存の文字列を検索
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

    /// セルアドレス（例：「A1」）を行番号と列番号にデコードします。
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
        let col = col_str
            .to_uppercase()
            .chars()
            .rev()
            .enumerate()
            .fold(0, |acc, (i, ch)| {
                acc + (ch as u32 - 'A' as u32 + 1) * 26u32.pow(i as u32)
            });
        (row, col)
    }

    /// セルのXML要素から値を抽出します。
    fn extract_cell_value(&self, cell_element: &XmlElement) -> Option<String> {
        match cell_element.attributes.get("t").map(|s| s.as_str()) {
            Some("s") => self.get_shared_string(cell_element),
            Some("inlineStr") => self.get_inline_string(cell_element),
            _ => cell_element
                .children
                .iter()
                .find(|e| e.name == "v")
                .and_then(|v| v.text.clone()),
        }
    }

    /// 共有文字列の値を取得します。
    fn get_shared_string(&self, cell_element: &XmlElement) -> Option<String> {
        let v_element = cell_element.children.iter().find(|e| e.name == "v")?;
        let idx = v_element.text.as_ref()?.parse::<usize>().ok()?;
        let shared_strings_xml = self.shared_strings.lock().unwrap();
        let sst = shared_strings_xml.elements.first()?;
        let si = sst.children.get(idx)?;
        si.children.first().and_then(|t| t.text.clone())
    }

    /// インライン文字列の値を取得します。
    fn get_inline_string(&self, cell_element: &XmlElement) -> Option<String> {
        let is_element = cell_element.children.iter().find(|e| e.name == "is")?;
        let t_element = is_element.children.iter().find(|e| e.name == "t")?;
        t_element.text.clone()
    }
}
