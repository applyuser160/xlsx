use crate::style::{Font, PatternFill};
use crate::xml::{Xml, XmlElement};
use chrono::NaiveDateTime;
use pyo3::prelude::*;
use std::sync::{Arc, Mutex};

/// Excelの単一セル
///
/// 値、書式設定、その他のプロパティを操作するメソッドを提供
#[pyclass]
pub struct Cell {
    /// シートのXMLデータへの参照
    sheet_xml: Arc<Mutex<Xml>>,
    /// 共有文字列テーブルへの参照
    shared_strings: Arc<Mutex<Xml>>,
    /// スタイル情報への参照
    styles: Arc<Mutex<Xml>>,
    /// セルのアドレス（例: "A1"）
    address: String,
    /// セルのフォント情報
    font: Option<Font>,
    /// セルの塗りつぶし情報
    fill: Option<PatternFill>,
}

#[pymethods]
impl Cell {
    /// セルの値取得
    ///
    /// 値は常に文字列として返却
    #[getter]
    pub fn value(&self) -> Option<String> {
        let xml = self.sheet_xml.lock().unwrap();
        let worksheet = xml.elements.first()?;
        let sheet_data = worksheet.children.iter().find(|e| e.name == "sheetData")?;

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

    /// セルへの値設定
    ///
    /// value の内容に基づき、適切な型（数値、文字列、日付、真偽値、数式）を自動判断
    /// - `=` で始まる場合: 数式
    /// - `YYYY-MM-DD HH:MM:SS` 形式の場合: 日付時刻
    /// - 数値に変換可能な場合: 数値
    /// - "true" または "false" の場合: 真偽値
    /// - 上記以外の場合: 文字列
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

    /// セルのフォント取得
    #[getter]
    fn get_font(&self) -> PyResult<Option<Font>> {
        Ok(self.font.clone())
    }

    /// セルへのフォント設定
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

    /// セルの塗りつぶし取得
    #[getter]
    fn get_fill(&self) -> PyResult<Option<PatternFill>> {
        Ok(self.fill.clone())
    }

    /// セルへの塗りつぶし設定
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
    /// 新しい `Cell` インスタンス作成（内部使用）
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

    /// セルへの数値設定
    pub fn set_number_value(&mut self, value: f64) {
        let mut xml = self.sheet_xml.lock().unwrap();
        let cell_element = self.get_or_create_cell_element(&mut xml);
        cell_element.attributes.remove("t");
        cell_element.children.retain(|c| c.name != "f"); // 数式タグ削除
        self.update_or_create_child_element(cell_element, "v", &value.to_string());
    }

    /// セルへの文字列設定
    /// 文字列は共有文字列テーブルに追加され、セルにはそのインデックスを格納
    pub fn set_string_value(&mut self, value: &str) {
        let sst_index = self.get_or_create_shared_string(value);
        let mut xml = self.sheet_xml.lock().unwrap();
        let cell_element = self.get_or_create_cell_element(&mut xml);
        cell_element
            .attributes
            .insert("t".to_string(), "s".to_string());
        cell_element.children.retain(|c| c.name != "f"); // 数式タグ削除
        self.update_or_create_child_element(cell_element, "v", &sst_index.to_string());
    }

    /// セルへの日付時刻設定
    /// 日付時刻はExcelのシリアル値に変換して格納
    pub fn set_datetime_value(&mut self, value: NaiveDateTime) {
        // Excelの日付は1900-01-01を1とするが、1900年が閏年と誤認識されるバグがあるため、
        // 1899-12-30をエポックとするのが一般的
        let excel_epoch = chrono::NaiveDate::from_ymd_opt(1899, 12, 30)
            .unwrap()
            .and_hms_opt(0, 0, 0)
            .unwrap();
        let duration = value.signed_duration_since(excel_epoch);
        let serial = duration.num_seconds() as f64 / 86400.0;
        self.set_number_value(serial);
        // TODO: 日付フォーマットのスタイルを自動的に適用
    }

    /// セルへの真偽値設定
    pub fn set_bool_value(&mut self, value: bool) {
        let mut xml = self.sheet_xml.lock().unwrap();
        let cell_element = self.get_or_create_cell_element(&mut xml);
        cell_element
            .attributes
            .insert("t".to_string(), "b".to_string());
        cell_element.children.retain(|c| c.name != "f"); // 数式タグ削除
        let bool_str = if value { "1" } else { "0" };
        self.update_or_create_child_element(cell_element, "v", bool_str);
    }

    /// セルへの数式設定
    pub fn set_formula_value(&mut self, formula: &str) {
        let mut xml = self.sheet_xml.lock().unwrap();
        let cell_element = self.get_or_create_cell_element(&mut xml);
        cell_element.attributes.remove("t");
        cell_element.children.retain(|c| c.name != "v"); // 値タグ削除
        self.update_or_create_child_element(cell_element, "f", formula);
    }

    // --- Private Helper Functions ---

    /// `<c>` (cell) 要素から実際の値抽出
    fn extract_cell_value(&self, cell_element: &XmlElement) -> Option<String> {
        match cell_element.attributes.get("t").map(|s| s.as_str()) {
            Some("s") => self.get_shared_string(cell_element),
            Some("inlineStr") => self.get_inline_string(cell_element),
            _ => self.get_direct_value(cell_element),
        }
    }

    /// 共有文字列テーブルから文字列取得
    fn get_shared_string(&self, cell_element: &XmlElement) -> Option<String> {
        let v_element = cell_element.children.iter().find(|e| e.name == "v")?;
        let idx = v_element.text.as_ref()?.parse::<usize>().ok()?;
        let shared_strings_xml = self.shared_strings.lock().unwrap();
        let sst = shared_strings_xml.elements.first()?;
        let si = sst.children.get(idx)?;
        si.children.first()?.text.clone()
    }

    /// インライン文字列取得
    fn get_inline_string(&self, cell_element: &XmlElement) -> Option<String> {
        let is_element = cell_element.children.iter().find(|e| e.name == "is")?;
        let t_element = is_element.children.iter().find(|e| e.name == "t")?;
        t_element.text.clone()
    }

    /// `<v>` タグから直接値取得
    fn get_direct_value(&self, cell_element: &XmlElement) -> Option<String> {
        cell_element
            .children
            .iter()
            .find(|e| e.name == "v")
            .and_then(|v| v.text.clone())
    }

    /// スタイルXMLにフォントを追加し、そのIDを返却
    fn add_font_to_styles(&self, font: &Font) -> usize {
        let mut styles_xml = self.styles.lock().unwrap();
        let fonts_tag = styles_xml.get_mut_or_create_child_by_tag("fonts");

        // 既存フォント検索
        for (i, f) in fonts_tag.children.iter().enumerate() {
            if let Some(existing_font) = Font::from_xml_element(f) {
                if font == &existing_font {
                    return i;
                }
            }
        }

        // 新規フォント追加
        fonts_tag.children.push(font.to_xml_element());
        let count = fonts_tag.children.len();
        fonts_tag
            .attributes
            .insert("count".to_string(), count.to_string());
        count - 1
    }

    /// スタイルXMLに塗りつぶしを追加し、そのIDを返却
    fn add_fill_to_styles(&self, fill: &PatternFill) -> usize {
        let mut styles_xml = self.styles.lock().unwrap();
        let fills_tag = styles_xml.get_mut_or_create_child_by_tag("fills");
        // TODO: 既存fill検索ロジック追加
        fills_tag.children.push(fill.to_xml_element());
        let count = fills_tag.children.len();
        fills_tag
            .attributes
            .insert("count".to_string(), count.to_string());
        count - 1
    }

    /// スタイルXMLにセル書式(xf)を追加し、そのIDを返却
    fn add_xf_to_styles(
        &self,
        font_id: usize,
        fill_id: usize,
        border_id: usize,
        _alignment_id: usize,
    ) -> usize {
        let mut styles_xml = self.styles.lock().unwrap();
        let cell_xfs_tag = styles_xml.get_mut_or_create_child_by_tag("cellXfs");

        // 既存xf検索
        for (i, xf) in cell_xfs_tag.children.iter().enumerate() {
            if xf.attributes.get("fontId") == Some(&font_id.to_string())
                && xf.attributes.get("fillId") == Some(&fill_id.to_string())
                && xf.attributes.get("borderId") == Some(&border_id.to_string())
            // TODO: alignmentも比較
            {
                return i;
            }
        }

        // 新規xf追加
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
        // border, alignmentも同様に追加
        cell_xfs_tag.children.push(xf_element);
        let count = cell_xfs_tag.children.len();
        cell_xfs_tag
            .attributes
            .insert("count".to_string(), count.to_string());
        count - 1
    }

    /// 指定されたアドレスのセル要素を取得または作成
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

        // 行(row)を取得または作成
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
        let row_element = &mut sheet_data.children[row_index];

        // セル(c)を取得または作成
        let cell_index = row_element
            .children
            .iter()
            .position(|c| c.name == "c" && c.attributes.get("r") == Some(&self.address))
            .unwrap_or_else(|| {
                let mut new_cell = XmlElement::new("c");
                new_cell
                    .attributes
                    .insert("r".to_string(), self.address.clone());
                row_element.children.push(new_cell);
                row_element.children.len() - 1
            });
        &mut row_element.children[cell_index]
    }

    /// 共有文字列テーブルに文字列を追加または取得し、そのインデックスを返却
    fn get_or_create_shared_string(&mut self, text: &str) -> usize {
        let mut shared_strings_xml = self.shared_strings.lock().unwrap();

        // sst要素がなければ作成
        if shared_strings_xml.elements.is_empty() {
            shared_strings_xml.elements.push(XmlElement::new("sst"));
        }
        let sst_element = shared_strings_xml.elements.first_mut().unwrap();

        // 既存文字列検索
        if let Some(index) = sst_element
            .children
            .iter()
            .position(|si| si.children.first().and_then(|t| t.text.as_deref()) == Some(text))
        {
            return index;
        }

        // 新規文字列追加
        let mut t_element = XmlElement::new("t");
        t_element.text = Some(text.to_string());
        let mut si_element = XmlElement::new("si");
        si_element.children.push(t_element);
        sst_element.children.push(si_element);
        sst_element.children.len() - 1
    }

    /// 親要素の子要素を更新または作成
    fn update_or_create_child_element(
        &self,
        parent: &mut XmlElement,
        tag_name: &str,
        text: &str,
    ) {
        if let Some(child) = parent.children.iter_mut().find(|c| c.name == tag_name) {
            child.text = Some(text.to_string());
        } else {
            let mut new_element = XmlElement::new(tag_name);
            new_element.text = Some(text.to_string());
            parent.children.push(new_element);
        }
    }

    /// セルアドレス（"A1"など）を行番号と列番号にデコード
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

// --- XML変換のためのヘルパー実装 ---

impl Font {
    /// XmlElementからFont生成
    fn from_xml_element(element: &XmlElement) -> Option<Font> {
        let mut font = Font::default();
        for child in &element.children {
            match child.name.as_str() {
                "name" => font.name = child.attributes.get("val").cloned(),
                "sz" => font.size = child.attributes.get("val").and_then(|s| s.parse().ok()),
                "b" => font.bold = Some(true),
                "i" => font.italic = Some(true),
                "color" => font.color = child.attributes.get("rgb").cloned(),
                _ => {}
            }
        }
        Some(font)
    }

    /// FontをXmlElementに変換
    fn to_xml_element(&self) -> XmlElement {
        let mut font_element = XmlElement::new("font");
        if let Some(name) = &self.name {
            let mut name_element = XmlElement::new("name");
            name_element
                .attributes
                .insert("val".to_string(), name.clone());
            font_element.children.push(name_element);
        }
        if let Some(size) = self.size {
            let mut size_element = XmlElement::new("sz");
            size_element
                .attributes
                .insert("val".to_string(), size.to_string());
            font_element.children.push(size_element);
        }
        if self.bold == Some(true) {
            font_element.children.push(XmlElement::new("b"));
        }
        if self.italic == Some(true) {
            font_element.children.push(XmlElement::new("i"));
        }
        if let Some(color) = &self.color {
            let mut color_element = XmlElement::new("color");
            color_element
                .attributes
                .insert("rgb".to_string(), color.clone());
            font_element.children.push(color_element);
        }
        font_element
    }
}

impl PatternFill {
    /// PatternFillをXmlElementに変換
    fn to_xml_element(&self) -> XmlElement {
        let mut fill_element = XmlElement::new("fill");
        let mut pattern_fill_element = XmlElement::new("patternFill");

        if let Some(pattern_type) = &self.pattern_type {
            pattern_fill_element
                .attributes
                .insert("patternType".to_string(), pattern_type.clone());
        }
        if let Some(fg_color) = &self.fg_color {
            let mut fg_color_element = XmlElement::new("fgColor");
            fg_color_element
                .attributes
                .insert("rgb".to_string(), fg_color.clone());
            pattern_fill_element.children.push(fg_color_element);
        }
        if let Some(bg_color) = &self.bg_color {
            let mut bg_color_element = XmlElement::new("bgColor");
            bg_color_element
                .attributes
                .insert("rgb".to_string(), bg_color.clone());
            pattern_fill_element.children.push(bg_color_element);
        }

        fill_element.children.push(pattern_fill_element);
        fill_element
    }
}
