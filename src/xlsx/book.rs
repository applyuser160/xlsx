use crate::sheet::Sheet;
use crate::xml::{Xml, XmlElement};
use pyo3::prelude::*;
use std::collections::HashMap;
use std::fs::{File, OpenOptions};
use std::io::{Read, Write};
use std::sync::{Arc, Mutex};
use zip::result::ZipError;
use zip::write::FileOptions;
use zip::{ZipArchive, ZipWriter};

// --- 定数定義 ---

// XML関連のサフィックス
const XML_SUFFIX: &str = ".xml";
const XML_RELS_SUFFIX: &str = ".xml.rels";

// 主要なファイルパス
const VBA_PROJECT_FILENAME: &str = "xl/vbaProject.bin";
const WORKBOOK_FILENAME: &str = "xl/workbook.xml";
const STYLES_FILENAME: &str = "xl/styles.xml";
const SHARED_STRINGS_FILENAME: &str = "xl/sharedStrings.xml";

// ディレクトリプレフィックス
const WORKBOOK_RELS_PREFIX: &str = "xl/_rels/";
const WORKSHEETS_RELS_PREFIX: &str = "xl/worksheets/_rels/";
const DRAWINGS_PREFIX: &str = "xl/drawings/";
const THEME_PREFIX: &str = "xl/theme/";
const WORKSHEETS_PREFIX: &str = "xl/worksheets/";
const TABLES_PREFIX: &str = "xl/tables/";
const PIVOT_TABLES_PREFIX: &str = "xl/pivotTables/";
const PIVOT_CACHES_PREFIX: &str = "xl/pivotCache/";

/// Excelブック全体を表します。
///
/// この構造体は、Excelファイル（.xlsx）の読み込み、操作、保存の機能を提供します。
/// シートの追加、削除、取得や、テーブルの追加などが可能です。
#[pyclass]
pub struct Book {
    /// Excelファイルのパス。新規作成の場合は空文字列。
    #[pyo3(get, set)]
    pub path: String,

    /// `xl/_rels/` 以下に存在するリレーションシップファイル。
    pub rels: HashMap<String, Xml>,

    /// `xl/drawings/` 以下に存在する図形ファイル。
    pub drawings: HashMap<String, Xml>,

    /// `xl/tables/` 以下に存在するテーブルファイル。
    pub tables: HashMap<String, Xml>,

    /// `xl/pivotTables/` 以下に存在するピボットテーブルファイル。
    pub pivot_tables: HashMap<String, Xml>,

    /// `xl/pivotCache/` 以下に存在するピボットキャッシュファイル。
    pub pivot_caches: HashMap<String, Xml>,

    /// `xl/theme/` 以下に存在するテーマファイル。
    pub themes: HashMap<String, Xml>,

    /// `xl/worksheets/` 以下に存在するワークシートファイル。
    /// 複数箇所から共有されるため、`Arc<Mutex<>>` でラップされています。
    pub worksheets: HashMap<String, Arc<Mutex<Xml>>>,

    /// `xl/worksheets/_rels/` 以下に存在するシートごとのリレーションシップファイル。
    pub sheet_rels: HashMap<String, Xml>,

    /// `xl/sharedStrings.xml`: 共有文字列テーブル。
    pub shared_strings: Arc<Mutex<Xml>>,

    /// `xl/styles.xml`: スタイル情報。
    pub styles: Arc<Mutex<Xml>>,

    /// `xl/workbook.xml`: ブックの構造を定義する主要なXML。
    pub workbook: Xml,

    /// `xl/vbaProject.bin`: VBAマクロプロジェクト（存在する場合）。
    pub vba_project: Option<Vec<u8>>,
}

#[pymethods]
impl Book {
    /// 新しい `Book` インスタンスを作成します。
    ///
    /// - `path` が空の場合: 新しい空のブックを作成します。
    /// - `path` が指定された場合: 既存の.xlsxファイルを読み込みます。
    #[new]
    #[pyo3(signature = (path = ""))]
    pub fn new(path: &str) -> Self {
        if path.is_empty() {
            Self::new_empty()
        } else {
            Self::from_file(path)
        }
    }

    /// シート名のリストを取得します。
    #[getter]
    pub fn sheetnames(&self) -> Vec<String> {
        self.sheet_tags()
            .iter()
            .filter_map(|x| x.attributes.get("name").cloned())
            .collect()
    }

    /// シートをイテレートするためのシート名リストを返します。
    /// Pythonの `for sheet_name in book:` のように使えます。
    pub fn __iter__(&self) -> Vec<String> {
        self.sheetnames()
    }

    /// 指定された名前のシートが存在するかどうかを確認します。
    /// Pythonの `sheet_name in book` のように使えます。
    pub fn __contains__(&self, key: &str) -> bool {
        self.sheetnames().iter().any(|name| name == key)
    }

    /// 指定された名前のシートを取得します。
    /// Pythonの `book[sheet_name]` のように使えます。
    pub fn __getitem__(&self, key: &str) -> PyResult<Sheet> {
        self.get_sheet_by_name(key).ok_or_else(|| {
            PyErr::new::<pyo3::exceptions::PyKeyError, _>(format!("No sheet named '{key}'"))
        })
    }

    /// 指定された名前のシートを削除します。
    /// Pythonの `del book[sheet_name]` のように使えます。
    pub fn __delitem__(&mut self, key: &str) -> PyResult<()> {
        let sheet = self.get_sheet_by_name(key).ok_or_else(|| {
            PyErr::new::<pyo3::exceptions::PyKeyError, _>(format!("No sheet named '{key}'"))
        })?;
        self.remove(&sheet);
        Ok(())
    }

    /// シートを指定して、そのインデックス（0から始まる位置）を取得します。
    pub fn index(&self, sheet: &Sheet) -> PyResult<usize> {
        self.sheetnames()
            .iter()
            .position(|name| name == &sheet.name)
            .ok_or_else(|| {
                PyErr::new::<pyo3::exceptions::PyValueError, _>(format!(
                    "Sheet '{}' not found in workbook",
                    sheet.name
                ))
            })
    }

    /// 指定されたシートをブックから削除します。
    pub fn remove(&mut self, sheet: &Sheet) {
        let sheet_paths = self.get_sheet_paths();
        let sheet_path = match sheet_paths.get(&sheet.name) {
            Some(path) => path,
            None => {
                // 存在しないシートの場合は何もしない
                return;
            }
        };

        // 1. ワークシートXMLを削除
        self.worksheets.remove(sheet_path);

        // 2. workbook.xmlから<sheet>タグを削除し、関連するr:idを取得
        let rid_to_remove = self.remove_sheet_tag_from_workbook(&sheet.name);

        // 3. workbook.xml.relsから対応する<Relationship>タグを削除
        if let Some(rid) = rid_to_remove {
            self.remove_relationship_by_id("xl/_rels/workbook.xml.rels", &rid);
        }

        // TODO: シートに関連するrelsファイル（例: xl/worksheets/_rels/sheet1.xml.rels）も削除する
    }

    /// 新しいシートを作成し、指定された位置に挿入します。
    pub fn create_sheet(&mut self, title: &str, index: Option<usize>) -> PyResult<Sheet> {
        if self.__contains__(title) {
            return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(format!(
                "Sheet with title '{title}' already exists"
            )));
        }

        // 新しいシートIDとrIdを決定
        let next_sheet_id = (self.sheet_tags().len() + 1) as u32;
        let next_rid = format!("rId{}", self.get_relationships().len() + 1);

        // 新しいワークシートXMLを作成
        let sheet_path = format!("xl/worksheets/sheet{next_sheet_id}.xml");
        let worksheet_xml = self.create_empty_worksheet_xml();
        let arc_mutex_xml = Arc::new(Mutex::new(worksheet_xml));
        self.worksheets
            .insert(sheet_path.clone(), arc_mutex_xml.clone());

        // workbook.xmlに新しい<sheet>タグを追加
        self.add_sheet_tag_to_workbook(title, next_sheet_id, &next_rid, index);

        // workbook.xml.relsに新しい<Relationship>タグを追加
        self.add_relationship_to_workbook_rels(&next_rid, &sheet_path);

        // 新しいSheetオブジェクトを返却
        Ok(Sheet::new(
            title.to_string(),
            arc_mutex_xml,
            self.shared_strings.clone(),
            self.styles.clone(),
        ))
    }

    /// ブックを別のパスにコピーします。
    pub fn copy(&self, path: &str) -> PyResult<()> {
        let new_file = OpenOptions::new()
            .write(true)
            .create(true)
            .truncate(true)
            .open(path)
            .map_err(|e| PyErr::new::<pyo3::exceptions::PyIOError, _>(e.to_string()))?;
        let mut zip_writer = ZipWriter::new(new_file);
        let options = FileOptions::default().compression_method(zip::CompressionMethod::Stored);

        let xmls = self.collect_all_xmls();

        if self.path.is_empty() {
            // 新規ブックの場合、すべてのXMLを直接書き込む
            for (file_name, xml) in &xmls {
                zip_writer
                    .start_file(file_name, options)
                    .map_err(|e: ZipError| {
                        PyErr::new::<pyo3::exceptions::PyIOError, _>(e.to_string())
                    })?;
                zip_writer
                    .write_all(&xml.to_buf())
                    .map_err(|e| PyErr::new::<pyo3::exceptions::PyIOError, _>(e.to_string()))?;
            }
        } else {
            // 既存ブックの場合、元のファイルを読み込みながら変更を適用
            let original_file = File::open(&self.path)
                .map_err(|e| PyErr::new::<pyo3::exceptions::PyIOError, _>(e.to_string()))?;
            let mut archive = ZipArchive::new(original_file)
                .map_err(|e| PyErr::new::<pyo3::exceptions::PyIOError, _>(e.to_string()))?;
            self.write_zip_archive(&mut archive, &xmls, &mut zip_writer, options)
                .map_err(|e| PyErr::new::<pyo3::exceptions::PyIOError, _>(e.to_string()))?;
        }

        zip_writer
            .finish()
            .map_err(|e| PyErr::new::<pyo3::exceptions::PyIOError, _>(e.to_string()))?;
        Ok(())
    }

    /// 現在の変更を元のファイルに保存します。
    pub fn save(&self) -> PyResult<()> {
        if self.path.is_empty() {
            return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(
                "Cannot save a new workbook without a path. Use copy() instead.",
            ));
        }
        self.copy(&self.path)
    }
}

// --- Bookの内部実装 ---
impl Book {
    /// 新しい空のブックを作成します。
    fn new_empty() -> Self {
        // ... (実装は元のコードと同様)
        let mut rels: HashMap<String, Xml> = HashMap::new();
        let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"#;
        rels.insert(
            "xl/_rels/workbook.xml.rels".to_string(),
            Xml::new(&workbook_rels.to_string()),
        );

        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
</sheets>
</workbook>"#;

        let styles_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"#;

        Book {
                path: "".to_string(),
                rels,
                drawings: HashMap::new(),
                tables: HashMap::new(),
                pivot_tables: HashMap::new(),
                pivot_caches: HashMap::new(),
                themes: HashMap::new(),
                worksheets: HashMap::new(),
                sheet_rels: HashMap::new(),
                shared_strings: Arc::new(Mutex::new(Xml::new(
                    &r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"></sst>"#.to_string()))),
                styles: Arc::new(Mutex::new(Xml::new(&styles_xml.to_string()))),
                workbook: Xml::new(&workbook_xml.to_string()),
                vba_project: None,
            }
    }

    /// .xlsxファイルを読み込んで `Book` を作成します。
    fn from_file(path: &str) -> Self {
        let file = File::open(path).unwrap_or_else(|_| panic!("File not found: {path}"));
        let mut archive = ZipArchive::new(file).unwrap();
        let mut book = Book::new_empty();
        book.path = path.to_string();

        for i in 0..archive.len() {
            let mut file = archive.by_index(i).unwrap();
            let name = file.name().to_string();

            // ... (読み込みロジックは元のコードと同様)
            if name.ends_with(XML_SUFFIX) {
                let mut contents: String = String::new();
                file.read_to_string(&mut contents).unwrap();
                let xml = Xml::new(&contents);

                if name.starts_with(DRAWINGS_PREFIX) {
                    book.drawings.insert(name, xml);
                } else if name.starts_with(TABLES_PREFIX) {
                    book.tables.insert(name, xml);
                } else if name.starts_with(PIVOT_TABLES_PREFIX) {
                    book.pivot_tables.insert(name, xml);
                } else if name.starts_with(PIVOT_CACHES_PREFIX) {
                    book.pivot_caches.insert(name, xml);
                } else if name.starts_with(THEME_PREFIX) {
                    book.themes.insert(name, xml);
                } else if name.starts_with(WORKSHEETS_PREFIX) {
                    book.worksheets.insert(name, Arc::new(Mutex::new(xml)));
                } else if name == WORKBOOK_FILENAME {
                    book.workbook = xml;
                } else if name == STYLES_FILENAME {
                    book.styles = Arc::new(Mutex::new(xml));
                } else if name == SHARED_STRINGS_FILENAME {
                    book.shared_strings = Arc::new(Mutex::new(xml));
                }
            } else if name.ends_with(XML_RELS_SUFFIX) {
                if name.starts_with(WORKBOOK_RELS_PREFIX) {
                    let mut contents: String = String::new();
                    file.read_to_string(&mut contents).unwrap();
                    book.rels.insert(name, Xml::new(&contents));
                } else if name.starts_with(WORKSHEETS_RELS_PREFIX) {
                    let mut contents: String = String::new();
                    file.read_to_string(&mut contents).unwrap();
                    book.sheet_rels.insert(name, Xml::new(&contents));
                }
            } else if name == VBA_PROJECT_FILENAME {
                let mut contents: Vec<u8> = Vec::new();
                file.read_to_end(&mut contents).unwrap();
                book.vba_project = Some(contents);
            }
        }
        book
    }

    /// ブック内のすべてのXMLデータをHashMapに収集します。
    fn collect_all_xmls(&self) -> HashMap<String, Xml> {
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
        xmls.extend(self.themes.clone());

        for (key, arc_mutex_xml) in &self.worksheets {
            xmls.insert(key.clone(), arc_mutex_xml.lock().unwrap().clone());
        }
        xmls
    }

    /// ZIPアーカイブにファイルと変更を書き込みます。
    fn write_zip_archive<W: Write + std::io::Seek>(
        &self,
        archive: &mut ZipArchive<File>,
        xmls: &HashMap<String, Xml>,
        zip_writer: &mut ZipWriter<W>,
        options: FileOptions,
    ) -> Result<(), ZipError> {
        let file_names: Vec<String> = archive.file_names().map(|s| s.to_string()).collect();

        // 変更されていないファイルをコピー
        for filename in file_names {
            if !xmls.contains_key(&filename)
                && Some(filename.as_str())
                    != self.vba_project.as_ref().map(|_| VBA_PROJECT_FILENAME)
            {
                let mut file = archive.by_name(&filename)?;
                let mut contents = Vec::new();
                file.read_to_end(&mut contents)?;
                zip_writer.start_file(filename, options)?;
                zip_writer.write_all(&contents)?;
            }
        }

        // 変更されたXMLを書き込み
        for (file_name, xml) in xmls {
            zip_writer.start_file(file_name, options)?;
            zip_writer.write_all(&xml.to_buf())?;
        }

        // VBAプロジェクトを書き込み
        if let Some(vba_project) = &self.vba_project {
            zip_writer.start_file(VBA_PROJECT_FILENAME, options)?;
            zip_writer.write_all(vba_project)?;
        }

        Ok(())
    }

    /// `xl/workbook.xml` 内の `<sheet>` タグのリストを取得します。
    fn sheet_tags(&self) -> Vec<XmlElement> {
        self.workbook
            .elements
            .first()
            .and_then(|wb| wb.children.iter().find(|el| el.name == "sheets"))
            .map_or(Vec::new(), |sheets| sheets.children.clone())
    }

    /// `xl/_rels/workbook.xml.rels` 内の `<Relationship>` タグのリストを取得します。
    fn get_relationships(&self) -> Vec<XmlElement> {
        self.rels
            .get("xl/_rels/workbook.xml.rels")
            .and_then(|rels| rels.elements.first())
            .map_or(Vec::new(), |rels| rels.children.clone())
    }

    /// シート名とパスのマッピングを取得します。
    fn get_sheet_paths(&self) -> HashMap<String, String> {
        let relationships = self.get_relationships();
        let sheet_paths_by_rid: HashMap<String, String> = relationships
            .iter()
            .filter(|r| {
                r.attributes
                    .get("Type")
                    .is_some_and(|t| t.ends_with("/worksheet"))
            })
            .filter_map(|r| {
                let id = r.attributes.get("Id")?.clone();
                let target = format!("xl/{}", r.attributes.get("Target")?);
                Some((id, target))
            })
            .collect();

        self.sheet_tags()
            .iter()
            .filter_map(|s| {
                let name = s.attributes.get("name")?.clone();
                let rid = s.attributes.get("r:id")?;
                let path = sheet_paths_by_rid.get(rid)?.clone();
                Some((name, path))
            })
            .collect()
    }

    /// シート名から `Sheet` オブジェクトを取得します。
    fn get_sheet_by_name(&self, name: &str) -> Option<Sheet> {
        let sheet_path = self.get_sheet_paths().get(name)?.clone();
        let xml = self.worksheets.get(&sheet_path)?.clone();
        Some(Sheet::new(
            name.to_string(),
            xml,
            self.shared_strings.clone(),
            self.styles.clone(),
        ))
    }

    // --- シート操作のヘルパー関数 ---

    /// `workbook.xml` から `<sheet>` タグを削除し、その `r:id` を返します。
    fn remove_sheet_tag_from_workbook(&mut self, sheet_name: &str) -> Option<String> {
        let sheets_tag = self
            .workbook
            .elements
            .first_mut()?
            .children
            .iter_mut()
            .find(|el| el.name == "sheets")?;

        let position = sheets_tag
            .children
            .iter()
            .position(|s| s.attributes.get("name").is_some_and(|n| n == sheet_name));

        if let Some(pos) = position {
            let removed_sheet = sheets_tag.children.remove(pos);
            return removed_sheet.attributes.get("r:id").cloned();
        }
        None
    }

    /// 指定されたrelsファイルからIDで指定された`<Relationship>`を削除します。
    fn remove_relationship_by_id(&mut self, rels_path: &str, rid: &str) {
        if let Some(rels) = self.rels.get_mut(rels_path) {
            if let Some(relationships) = rels.elements.first_mut() {
                relationships
                    .children
                    .retain(|r| r.attributes.get("Id").is_none_or(|id| id != rid));
            }
        }
    }

    /// 空のワークシートXMLを生成します。
    fn create_empty_worksheet_xml(&self) -> Xml {
        Xml::new(
            &r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                <sheetData/>
            </worksheet>"#
                .to_string(),
        )
    }

    /// `workbook.xml` の `<sheets>` に新しい `<sheet>` を追加します。
    fn add_sheet_tag_to_workbook(
        &mut self,
        title: &str,
        sheet_id: u32,
        rid: &str,
        index: Option<usize>,
    ) {
        if let Some(sheets_tag) = self
            .workbook
            .elements
            .first_mut()
            .and_then(|wb| wb.children.iter_mut().find(|el| el.name == "sheets"))
        {
            let mut sheet_element = XmlElement::new("sheet");
            sheet_element
                .attributes
                .insert("name".to_string(), title.to_string());
            sheet_element
                .attributes
                .insert("sheetId".to_string(), sheet_id.to_string());
            sheet_element
                .attributes
                .insert("r:id".to_string(), rid.to_string());

            match index {
                Some(i) if i < sheets_tag.children.len() => {
                    sheets_tag.children.insert(i, sheet_element)
                }
                _ => sheets_tag.children.push(sheet_element),
            }
        }
    }

    /// `workbook.xml.rels` に新しいリレーションシップを追加します。
    fn add_relationship_to_workbook_rels(&mut self, rid: &str, target: &str) {
        if let Some(rels) = self
            .rels
            .get_mut("xl/_rels/workbook.xml.rels")
            .and_then(|r| r.elements.first_mut())
        {
            let mut rel_element = XmlElement::new("Relationship");
            rel_element
                .attributes
                .insert("Id".to_string(), rid.to_string());
            rel_element.attributes.insert(
                "Type".to_string(),
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
                    .to_string(),
            );
            rel_element.attributes.insert(
                "Target".to_string(),
                target.strip_prefix("xl/").unwrap_or(target).to_string(),
            );
            rels.children.push(rel_element);
        }
    }
}
