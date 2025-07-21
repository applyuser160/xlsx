// Copyright (c) 2024-present, zcayh.
// All rights reserved.
//
// This source code is licensed under the MIT license found in the
// LICENSE file in the root directory of this source tree.

use std::collections::HashMap;
use std::fs::{File, OpenOptions};
use std::io::{Cursor, Read, Write};
use std::sync::{Arc, Mutex};
use zip::write::FileOptions;
use zip::{ZipArchive, ZipWriter};

use pyo3::prelude::*;

use crate::sheet::Sheet;
use crate::xml::{Xml, XmlElement};

/// XMLファイルのファイル拡張子。
const XML_SUFFIX: &str = ".xml";
/// リレーションシップファイルのファイル拡張子。
const XML_RELS_SUFFIX: &str = ".xml.rels";
/// VBAプロジェクトバイナリへのパス。
const VBA_PROJECT_FILENAME: &str = "xl/vbaProject.bin";

/// ワークブックXMLファイルへのパス。
const WORKBOOK_FILENAME: &str = "xl/workbook.xml";
/// スタイルXMLファイルへのパス。
const STYLES_FILENAME: &str = "xl/styles.xml";
/// 共有文字列XMLファイルへのパス。
const SHARED_STRINGS_FILENAME: &str = "xl/sharedStrings.xml";

/// ワークブックリレーションシップファイルのプレフィックス。
const WORKBOOK_RELS_PREFIX: &str = "xl/_rels/";
/// ワークシートリレーションシップファイルのプレフィックス。
const WORKSHEETS_RELS_PREFIX: &str = "xl/worksheets/_rels/";
/// 描画ファイルのプレフィックス。
const DRAWINGS_PREFIX: &str = "xl/drawings/";
/// テーマファイルのプレフィックス。
const THEME_PREFIX: &str = "xl/theme/";
/// ワークシートファイルのプレフィックス。
const WORKSHEETS_PREFIX: &str = "xl/worksheets/";
/// テーブルファイルのプレフィックス。
const TABLES_PREFIX: &str = "xl/tables/";
/// ピボットテーブルファイルのプレフィックス。
const PIVOT_TABLES_PREFIX: &str = "xl/pivotTables/";
/// ピボットキャッシュファイルのプレフィックス。
const PIVOT_CACHES_PREFIX: &str = "xl/pivotCache/";

/// Excelワークブックを表します。
///
/// これは、.xlsxファイルの作成、読み取り、変更のメインエントリポイントです。
/// シート、スタイル、共有文字列など、ワークブックのすべてのコンポーネントを保持します。
#[pyclass]
pub struct Book {
    /// ワークブックのファイルパス。新しいワークブックの場合は空です。
    #[pyo3(get, set)]
    pub path: String,
    /// `xl/_rels/`にあるリレーションシップファイル。
    pub rels: HashMap<String, Xml>,
    /// `xl/drawings/`にある描画ファイル。
    pub drawings: HashMap<String, Xml>,
    /// `xl/tables/`にあるテーブルファイル。
    pub tables: HashMap<String, Xml>,
    /// `xl/pivotTables/`にあるピボットテーブルファイル。
    pub pivot_tables: HashMap<String, Xml>,
    /// `xl/pivotCache/`にあるピボットキャッシュファイル。
    pub pivot_caches: HashMap<String, Xml>,
    /// `xl/theme/`にあるテーマファイル。
    pub themes: HashMap<String, Xml>,
    /// `xl/worksheets/`にあるワークシートファイル。
    pub worksheets: HashMap<String, Arc<Mutex<Xml>>>,
    /// `xl/worksheets/_rels/`にあるワークシートリレーションシップファイル。
    pub sheet_rels: HashMap<String, Xml>,
    /// 共有文字列テーブル（`xl/sharedStrings.xml`）。
    pub shared_strings: Arc<Mutex<Xml>>,
    /// スタイルテーブル（`xl/styles.xml`）。
    pub styles: Arc<Mutex<Xml>>,
    /// メインのワークブックXML（`xl/workbook.xml`）。
    pub workbook: Xml,
    /// VBAプロジェクトバイナリ（存在する場合）。
    pub vba_project: Option<Vec<u8>>,
}

#[pymethods]
impl Book {
    /// 新しい`Book`インスタンスを作成します。
    ///
    /// `path`が空の場合、デフォルト設定で新しいワークブックが作成されます。
    /// それ以外の場合は、指定されたパスからワークブックを読み込みます。
    #[new]
    #[pyo3(signature = (path = ""))]
    pub fn new(path: &str) -> Self {
        if path.is_empty() {
            Self::new_workbook()
        } else {
            Self::from_path(path)
        }
    }

    /// ワークブック内のシート名のリストを返します。
    #[getter]
    pub fn sheetnames(&self) -> Vec<String> {
        self.sheet_tags()
            .iter()
            .map(|x| x.attributes.get("name").unwrap().clone())
            .collect()
    }

    /// シート名を反復処理できます。
    pub fn __iter__(&self) -> Vec<String> {
        self.sheetnames()
    }

    /// 指定された名前のシートが存在するかどうかを確認します。
    pub fn __contains__(&self, key: String) -> bool {
        self.sheetnames().contains(&key)
    }

    /// 名前でシートを取得します。
    pub fn __getitem__(&self, key: String) -> PyResult<Sheet> {
        if let Some(sheet) = self.get_sheet_by_name(key.as_str()) {
            Ok(sheet)
        } else {
            Err(pyo3::exceptions::PyKeyError::new_err(format!(
                "No sheet named '{key}'"
            )))
        }
    }

    /// ワークシートにテーブルを追加します。
    pub fn add_table(&mut self, sheet_name: String, name: String, table_ref: String) {
        let table_id = self.tables.len() + 1;
        let table_filename = format!("xl/tables/table{table_id}.xml");

        // 新しいテーブルXMLを作成
        let new_table_xml = Xml::new(&format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="{table_id}" name="{name}" displayName="{name}" ref="{table_ref}" totalsRowShown="0">
    <autoFilter ref="{table_ref}"/>
    <tableColumns count="3">
        <tableColumn id="1" name="Column1"/>
        <tableColumn id="2" name="Column2"/>
        <tableColumn id="3" name="Column3"/>
    </tableColumns>
    <tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
</table>"#
        ));
        self.tables.insert(table_filename.clone(), new_table_xml);

        // ワークシートXMLにテーブルパーツを追加
        let sheet_path = self.get_sheet_paths().get(&sheet_name).unwrap().clone();
        if let Some(sheet_xml) = self.worksheets.get_mut(&sheet_path) {
            let mut sheet_xml_locked = sheet_xml.lock().unwrap();
            let worksheet = &mut sheet_xml_locked.elements[0];
            let mut table_parts = XmlElement::new("tableParts");
            table_parts
                .attributes
                .insert("count".to_string(), "1".to_string());
            let mut table_part = XmlElement::new("tablePart");
            table_part
                .attributes
                .insert("r:id".to_string(), format!("rId{table_id}"));
            table_parts.children.push(table_part);
            worksheet.children.push(table_parts);
        }

        // テーブルのリレーションシップをワークシートのrelsファイルに追加
        self.add_relationship_to_sheet(
            &sheet_path,
            &format!("../tables/table{table_id}.xml"),
            "table",
            table_id,
        );
    }

    /// 名前でシートを削除します。
    pub fn __delitem__(&mut self, key: String) -> PyResult<()> {
        if let Some(sheet) = self.get_sheet_by_name(key.as_str()) {
            self.remove(&sheet);
            Ok(())
        } else {
            Err(pyo3::exceptions::PyKeyError::new_err(format!(
                "No sheet named '{key}'"
            )))
        }
    }

    /// シートのインデックスを返します。
    pub fn index(&self, sheet: &Sheet) -> PyResult<usize> {
        let sheet_name = &sheet.name;
        self.sheetnames()
            .iter()
            .position(|x| x == sheet_name)
            .ok_or_else(|| {
                pyo3::exceptions::PyValueError::new_err(format!("No sheet named '{sheet_name}'"))
            })
    }

    /// ワークブックからシートを削除します。
    pub fn remove(&mut self, sheet: &Sheet) {
        let sheet_paths = self.get_sheet_paths();
        let sheet_path = match sheet_paths.get(&sheet.name) {
            Some(path) => path,
            None => panic!("No sheet named '{}'", sheet.name),
        };

        if self.worksheets.remove(sheet_path).is_some() {
            let mut rid_to_remove = String::new();
            if let Some(sheets_tag) = self
                .workbook
                .elements
                .first_mut()
                .and_then(|wb| wb.children.iter_mut().find(|c| c.name == "sheets"))
            {
                if let Some(pos) = sheets_tag
                    .children
                    .iter()
                    .position(|s| s.attributes.get("name") == Some(&sheet.name))
                {
                    let sheet_element = &sheets_tag.children[pos];
                    if let Some(rid) = sheet_element.attributes.get("r:id") {
                        rid_to_remove = rid.clone();
                    }
                    sheets_tag.children.remove(pos);
                }
            }

            if !rid_to_remove.is_empty() {
                if let Some(rels) = self
                    .rels
                    .get_mut("xl/_rels/workbook.xml.rels")
                    .and_then(|r| r.elements.first_mut())
                {
                    rels.children
                        .retain(|r| r.attributes.get("Id") != Some(&rid_to_remove));
                }
            }
        }
    }

    /// 新しいシートを作成し、ワークブックに追加します。
    pub fn create_sheet(&mut self, title: String, index: usize) -> Sheet {
        let sheet_tags = self.sheet_tags();
        let next_sheet_id = sheet_tags.len() + 1;
        let next_rid = format!("rId{}", self.get_relationships().len() + 1);
        let sheet_path = format!("xl/worksheets/sheet{next_sheet_id}.xml");

        // 新しいワークシートXMLを作成
        let worksheet_xml = Xml::new(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheetData/>
</worksheet>"#);
        let arc_mutex_xml = Arc::new(Mutex::new(worksheet_xml));
        self.worksheets
            .insert(sheet_path.clone(), arc_mutex_xml.clone());

        // workbook.xmlにシートを追加
        if let Some(sheets_tag) = self
            .workbook
            .elements
            .first_mut()
            .and_then(|wb| wb.children.iter_mut().find(|c| c.name == "sheets"))
        {
            let mut sheet_element = XmlElement::new("sheet");
            sheet_element
                .attributes
                .insert("name".to_string(), title.clone());
            sheet_element
                .attributes
                .insert("sheetId".to_string(), next_sheet_id.to_string());
            sheet_element
                .attributes
                .insert("r:id".to_string(), next_rid.clone());

            if index < sheets_tag.children.len() {
                sheets_tag.children.insert(index, sheet_element);
            } else {
                sheets_tag.children.push(sheet_element);
            }
        }

        // workbook.xml.relsにリレーションシップを追加
        if let Some(rels) = self
            .rels
            .get_mut("xl/_rels/workbook.xml.rels")
            .and_then(|r| r.elements.first_mut())
        {
            let mut rel_element = XmlElement::new("Relationship");
            rel_element.attributes.insert("Id".to_string(), next_rid);
            rel_element.attributes.insert(
                "Type".to_string(),
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
                    .to_string(),
            );
            rel_element.attributes.insert(
                "Target".to_string(),
                format!("worksheets/sheet{next_sheet_id}.xml"),
            );
            rels.children.push(rel_element);
        }

        Sheet::new(
            title,
            arc_mutex_xml,
            self.shared_strings.clone(),
            self.styles.clone(),
        )
    }

    /// 指定されたパスにワークブックのコピーを作成します。
    pub fn copy(&self, path: &str) {
        let new_file = OpenOptions::new()
            .write(true)
            .create(true)
            .truncate(true)
            .open(path)
            .unwrap();
        let mut zip_writer = ZipWriter::new(new_file);
        let options = FileOptions::default().compression_method(zip::CompressionMethod::Stored);
        let xmls = self.merge_xmls();

        if self.path.is_empty() {
            for (file_name, xml) in &xmls {
                zip_writer.start_file(file_name, options).unwrap();
                zip_writer.write_all(&xml.to_buf()).unwrap();
            }
        } else {
            let file = File::open(&self.path).unwrap();
            let mut archive = ZipArchive::new(file).unwrap();
            self.write_to_zip(&mut archive, &xmls, &mut zip_writer, &options);
        }

        zip_writer.finish().unwrap();
    }
}

impl Book {
    /// ワークブックを元のパスに保存します。
    pub fn save(&self) {
        let file = File::open(&self.path).unwrap();
        let mut archive = ZipArchive::new(file).unwrap();
        let mut buffer = Cursor::new(Vec::new());
        let mut zip_writer = ZipWriter::new(&mut buffer);
        let options = FileOptions::default().compression_method(zip::CompressionMethod::Stored);
        let xmls = self.merge_xmls();

        self.write_to_zip(&mut archive, &xmls, &mut zip_writer, &options);
    }

    /// すべてのXMLコンポーネントを保存用に単一のHashMapにマージします。
    fn merge_xmls(&self) -> HashMap<String, Xml> {
        let mut xmls = self.rels.clone();
        xmls.insert(WORKBOOK_FILENAME.to_string(), self.workbook.clone());
        xmls.insert(
            STYLES_FILENAME.to_string(),
            self.styles.lock().unwrap().clone(),
        );
        xmls.insert(
            SHARED_STRINGS_FILENAME.to_string(),
            self.shared_strings.lock().unwrap().clone(),
        );
        xmls.extend(self.drawings.clone());
        xmls.extend(self.tables.clone());
        xmls.extend(self.pivot_tables.clone());
        xmls.extend(self.pivot_caches.clone());
        xmls.extend(self.sheet_rels.clone());
        for (key, arc_mutex_xml) in &self.worksheets {
            xmls.insert(key.clone(), arc_mutex_xml.lock().unwrap().clone());
        }
        xmls.extend(self.themes.clone());
        xmls
    }

    /// すべてのワークブックデータをzipアーカイブに書き込みます。
    fn write_to_zip<W: Write + std::io::Seek>(
        &self,
        archive: &mut ZipArchive<File>,
        xmls: &HashMap<String, Xml>,
        zip_writer: &mut ZipWriter<W>,
        options: &FileOptions,
    ) {
        // 元のアーカイブからXML以外のファイルをコピー
        for filename in archive
            .file_names()
            .map(|s| s.to_string())
            .collect::<Vec<_>>()
        {
            if !xmls.contains_key(&filename)
                && Some(filename.as_str())
                    != self.vba_project.as_ref().map(|_| VBA_PROJECT_FILENAME)
            {
                let mut file = archive.by_name(&filename).unwrap();
                let mut contents = Vec::new();
                file.read_to_end(&mut contents).unwrap();
                zip_writer.start_file(&filename, *options).unwrap();
                zip_writer.write_all(&contents).unwrap();
            }
        }

        // すべてのXMLファイルを書き込む
        for (file_name, xml) in xmls {
            zip_writer.start_file(file_name, *options).unwrap();
            zip_writer.write_all(&xml.to_buf()).unwrap();
        }

        // VBAプロジェクトが存在する場合は書き込む
        if let Some(vba_project) = &self.vba_project {
            zip_writer
                .start_file(VBA_PROJECT_FILENAME, *options)
                .unwrap();
            zip_writer.write_all(vba_project).unwrap();
        }
    }

    /// ワークブックXMLからすべての`<sheet>`タグを取得します。
    fn sheet_tags(&self) -> Vec<XmlElement> {
        self.workbook
            .elements
            .first()
            .and_then(|wb| wb.children.iter().find(|c| c.name == "sheets"))
            .map_or(Vec::new(), |s| s.children.clone())
    }

    /// ワークブックのリレーションシップXMLからすべての`<Relationship>`タグを取得します。
    fn get_relationships(&self) -> Vec<XmlElement> {
        self.rels
            .get("xl/_rels/workbook.xml.rels")
            .and_then(|r| r.elements.first())
            .map_or(Vec::new(), |rel| rel.children.clone())
    }

    /// シート名とその対応するファイルパスのマップを取得します。
    fn get_sheet_paths(&self) -> HashMap<String, String> {
        let sheet_tags = self.sheet_tags();
        let relationships = self.get_relationships();
        let sheet_paths: HashMap<String, String> = relationships
            .into_iter()
            .map(|x| {
                (
                    x.attributes.get("Id").unwrap().clone(),
                    x.attributes.get("Target").unwrap().clone(),
                )
            })
            .collect();

        sheet_tags
            .into_iter()
            .filter_map(|sheet_tag| {
                let id = sheet_tag.attributes.get("r:id")?;
                let sheet_path = sheet_paths.get(id)?;
                let trimmed_path = sheet_path
                    .trim_start_matches("/xl/")
                    .trim_start_matches("xl/");
                let name = sheet_tag.attributes.get("name")?.clone();
                Some((name, format!("xl/{trimmed_path}")))
            })
            .collect()
    }

    /// 名前で`Sheet`インスタンスを取得します。
    fn get_sheet_by_name(&self, name: &str) -> Option<Sheet> {
        let sheet_path = self.get_sheet_paths().get(name)?.clone();
        let xml = self.worksheets.get(&sheet_path)?;
        Some(Sheet::new(
            name.to_string(),
            xml.clone(),
            self.shared_strings.clone(),
            self.styles.clone(),
        ))
    }

    /// 新しい空のワークブックを作成します。
    fn new_workbook() -> Self {
        let mut rels = HashMap::new();
        let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#;
        rels.insert(
            "xl/_rels/workbook.xml.rels".to_string(),
            Xml::new(workbook_rels),
        );

        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets></sheets></workbook>"#;
        let styles_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>"#;
        let shared_strings_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"></sst>"#;

        Self {
            path: "".to_string(),
            rels,
            drawings: HashMap::new(),
            tables: HashMap::new(),
            pivot_tables: HashMap::new(),
            pivot_caches: HashMap::new(),
            themes: HashMap::new(),
            worksheets: HashMap::new(),
            sheet_rels: HashMap::new(),
            shared_strings: Arc::new(Mutex::new(Xml::new(shared_strings_xml))),
            styles: Arc::new(Mutex::new(Xml::new(styles_xml))),
            workbook: Xml::new(workbook_xml),
            vba_project: None,
        }
    }

    /// ファイルパスからワークブックを読み込みます。
    fn from_path(path: &str) -> Self {
        let file = File::open(path).unwrap_or_else(|_| panic!("File not found: {path}"));
        let mut archive = ZipArchive::new(file).unwrap();
        let mut book = Self::new_workbook();
        book.path = path.to_string();

        for i in 0..archive.len() {
            let mut file = archive.by_index(i).unwrap();
            let name = file.name().to_string();

            if name.ends_with(XML_SUFFIX) || name.ends_with(XML_RELS_SUFFIX) {
                let mut contents = String::new();
                file.read_to_string(&mut contents).unwrap();
                let xml = Xml::new(&contents);
                book.dispatch_xml_file(name, xml);
            } else if name == VBA_PROJECT_FILENAME {
                let mut contents = Vec::new();
                file.read_to_end(&mut contents).unwrap();
                book.vba_project = Some(contents);
            }
        }
        book
    }

    /// 解析されたXMLファイルを`Book`構造体の正しいフィールドに振り分けます。
    fn dispatch_xml_file(&mut self, name: String, xml: Xml) {
        if name.starts_with(DRAWINGS_PREFIX) {
            self.drawings.insert(name, xml);
        } else if name.starts_with(TABLES_PREFIX) {
            self.tables.insert(name, xml);
        } else if name.starts_with(PIVOT_TABLES_PREFIX) {
            self.pivot_tables.insert(name, xml);
        } else if name.starts_with(PIVOT_CACHES_PREFIX) {
            self.pivot_caches.insert(name, xml);
        } else if name.starts_with(THEME_PREFIX) {
            self.themes.insert(name, xml);
        } else if name.starts_with(WORKSHEETS_PREFIX) {
            self.worksheets.insert(name, Arc::new(Mutex::new(xml)));
        } else if name == WORKBOOK_FILENAME {
            self.workbook = xml;
        } else if name == STYLES_FILENAME {
            self.styles = Arc::new(Mutex::new(xml));
        } else if name == SHARED_STRINGS_FILENAME {
            self.shared_strings = Arc::new(Mutex::new(xml));
        } else if name.starts_with(WORKBOOK_RELS_PREFIX) {
            self.rels.insert(name, xml);
        } else if name.starts_with(WORKSHEETS_RELS_PREFIX) {
            self.sheet_rels.insert(name, xml);
        }
    }

    /// ワークシートの.relsファイルにリレーションシップを追加します。
    fn add_relationship_to_sheet(
        &mut self,
        sheet_path: &str,
        target: &str,
        rel_type: &str,
        id: usize,
    ) {
        let rels_filename = format!(
            "xl/worksheets/_rels/{}.rels",
            sheet_path.split('/').next_back().unwrap()
        );
        let rels = self.sheet_rels.entry(rels_filename).or_insert_with(|| {
            Xml::new(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#)
        });
        if rels.elements.is_empty() {
            rels.elements.push(XmlElement::new("Relationships"));
        }
        let relationships = &mut rels.elements[0];
        let mut relationship = XmlElement::new("Relationship");
        relationship
            .attributes
            .insert("Id".to_string(), format!("rId{id}"));
        relationship.attributes.insert(
            "Type".to_string(),
            format!(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/{rel_type}"
            ),
        );
        relationship
            .attributes
            .insert("Target".to_string(), target.to_string());
        relationships.children.push(relationship);
    }
}
