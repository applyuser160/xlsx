use std::collections::HashMap;
use std::fs::{File, OpenOptions};
use std::io::{BufReader, BufWriter, Cursor, Read, Write};
use std::sync::{Arc, Mutex, MutexGuard};
use zip::write::FileOptions;
use zip::{ZipArchive, ZipWriter};

use pyo3::prelude::*;

use crate::sheet::Sheet;
use crate::xml::{Xml, XmlElement};

/// XMLファイルのサフィックス
const XML_SUFFIX: &str = ".xml";
/// リレーションシップファイルのサフィックス
const XML_RELS_SUFFIX: &str = ".xml.rels";
/// VBAプロジェクトのファイル名
const VBA_PROJECT_FILENAME: &str = "xl/vbaProject.bin";

/// ワークブックXMLのファイル名
const WORKBOOK_FILENAME: &str = "xl/workbook.xml";
/// スタイルXMLのファイル名
const STYLES_FILENAME: &str = "xl/styles.xml";
/// 共有文字列XMLのファイル名
const SHARED_STRINGS_FILENAME: &str = "xl/sharedStrings.xml";

/// ワークブックリレーションシップのプレフィックス
const WORKBOOK_RELS_PREFIX: &str = "xl/_rels/";
/// ワークシートリレーションシップのプレフィックス
const WORKSHEETS_RELS_PREFIX: &str = "xl/worksheets/_rels/";
/// 図形のプレフィックス
const DRAWINGS_PREFIX: &str = "xl/drawings/";
/// テーマのプレフィックス
const THEME_PREFIX: &str = "xl/theme/";
/// ワークシートのプレフィックス
const WORKSHEETS_PREFIX: &str = "xl/worksheets/";
/// テーブルのプレフィックス
const TABLES_PREFIX: &str = "xl/tables/";
/// ピボットテーブルのプレフィックス
const PIVOT_TABLES_PREFIX: &str = "xl/pivotTables/";
/// ピボットキャッシュのプレフィックス
const PIVOT_CACHES_PREFIX: &str = "xl/pivotCache/";

/// Excelワークブック
#[pyclass]
pub struct Book {
    /// Excelファイルへのパス
    #[pyo3(get, set)]
    pub path: String,

    /// `xl/_rels/` 内のXMLファイル
    pub rels: HashMap<String, Xml>,

    /// `xl/drawings/` 内のXMLファイル
    pub drawings: HashMap<String, Xml>,

    /// `xl/tables/` 内のXMLファイル
    pub tables: HashMap<String, Xml>,

    /// `xl/pivotTables/` 内のXMLファイル
    pub pivot_tables: HashMap<String, Xml>,

    /// `xl/pivotCache/` 内のXMLファイル
    pub pivot_caches: HashMap<String, Xml>,

    /// `xl/theme/` 内のXMLファイル
    pub themes: HashMap<String, Xml>,

    /// `xl/worksheets/` 内のXMLファイル
    pub worksheets: HashMap<String, Arc<Mutex<Xml>>>,

    /// `xl/worksheets/_rels/` 内のXMLファイル
    pub sheet_rels: HashMap<String, Xml>,

    /// `xl/sharedStrings.xml` ファイル
    pub shared_strings: Arc<Mutex<Xml>>,
    pub shared_strings_map: Arc<Mutex<HashMap<String, usize>>>,

    /// `xl/styles.xml` ファイル
    pub styles: Arc<Mutex<Xml>>,

    /// `workbook.xml` ファイル
    pub workbook: Xml,

    /// `vbaProject.bin` ファイル
    pub vba_project: Option<Vec<u8>>,
}

#[pymethods]
impl Book {
    /// 新しい `Book` インスタンスの作成
    ///
    /// パスが指定されている場合は、ファイルからワークブックを読み込み
    /// それ以外の場合は、新しいワークブックを作成
    #[new]
    #[pyo3(signature = (path = ""))]
    pub fn new(path: &str) -> Self {
        if path.is_empty() {
            Self::new_empty_workbook()
        } else {
            Self::from_file(path)
        }
    }

    /// ワークブック内の全シート名の取得
    #[getter]
    pub fn sheetnames(&self) -> Vec<String> {
        self.sheet_tags()
            .iter()
            .filter_map(|x| x.attributes.get("name").cloned())
            .collect()
    }

    /// シート名のイテレータ
    pub fn __iter__(&self) -> Vec<String> {
        self.sheetnames()
    }

    /// 指定された名前のシートがワークブックに存在するかどうかの確認
    pub fn __contains__(&self, key: String) -> bool {
        self.sheetnames().contains(&key)
    }

    /// 名前によるシートの取得
    pub fn __getitem__(&self, key: String) -> Sheet {
        self.get_sheet_by_name(&key)
            .unwrap_or_else(|| panic!("No sheet named '{key}'"))
    }

    /// ワークシートへのテーブルの追加
    pub fn add_table(&mut self, sheet_name: String, name: String, table_ref: String) {
        let table_id: usize = self.tables.len() + 1;
        let table_filename: String = format!("xl/tables/table{table_id}.xml");

        // テーブルXMLの作成
        self.create_table_xml(&name, &table_ref, table_id, &table_filename);

        // ワークシートへのテーブルパーツの追加
        self.add_table_parts_to_worksheet(&sheet_name, table_id);

        // ワークシートのリレーションシップへのリレーションシップの追加
        self.add_table_relationship(&sheet_name, table_id);
    }

    /// 名前によるシートの削除
    pub fn __delitem__(&mut self, key: String) {
        if let Some(sheet) = self.get_sheet_by_name(&key) {
            self.remove(&sheet);
        } else {
            panic!("No sheet named '{key}'");
        }
    }

    /// シートのインデックス取得
    pub fn index(&self, sheet: &Sheet) -> usize {
        self.sheetnames()
            .iter()
            .position(|x| x == &sheet.name)
            .unwrap_or_else(|| panic!("No sheet named '{}'", &sheet.name))
    }

    /// ワークブックからのシートの削除
    pub fn remove(&mut self, sheet: &Sheet) {
        let sheet_paths: HashMap<String, String> = self.get_sheet_paths();
        let sheet_path: &String = match sheet_paths.get(&sheet.name) {
            Some(path) => path,
            None => panic!("No sheet named '{}'", sheet.name),
        };

        if self.worksheets.remove(sheet_path).is_none() {
            panic!("No sheet named '{}'", sheet.name);
        }

        let rid_to_remove: Option<String> = self
            .workbook
            .elements
            .first_mut()
            .and_then(|wb| wb.children.iter_mut().find(|x| x.name == "sheets"))
            .and_then(|sheets_tag| {
                let rid = sheets_tag
                    .children
                    .iter()
                    .find(|s| s.attributes.get("name").as_ref() == Some(&&sheet.name))
                    .and_then(|s| s.attributes.get("r:id").cloned());
                sheets_tag
                    .children
                    .retain(|s| s.attributes.get("name").as_ref() != Some(&&sheet.name));
                rid
            });

        if let Some(rid) = rid_to_remove {
            if let Some(relationships_tag) = self
                .rels
                .get_mut("xl/_rels/workbook.xml.rels")
                .and_then(|rels| rels.elements.first_mut())
            {
                relationships_tag
                    .children
                    .retain(|r| r.attributes.get("Id").as_ref() != Some(&&rid))
            }
        }
    }

    /// ワークブックへの新しいシートの作成
    pub fn create_sheet(&mut self, title: String, index: usize) -> Sheet {
        let next_sheet_id: usize = self.sheet_tags().len() + 1;
        let next_rid: String = format!("rId{}", self.get_relationships().len() + 1);
        let sheet_path: String = format!("xl/worksheets/sheet{next_sheet_id}.xml");

        let worksheet_xml: Xml = Xml::new(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                <sheetData/>
            </worksheet>"#,
        )
        .expect("Failed to create new worksheet XML");

        let arc_mutex_xml: Arc<Mutex<Xml>> = Arc::new(Mutex::new(worksheet_xml));
        self.worksheets
            .insert(sheet_path.clone(), arc_mutex_xml.clone());

        self.add_sheet_to_workbook_xml(&title, next_sheet_id, &next_rid, index);
        self.add_sheet_relationship(&next_rid, next_sheet_id);

        Sheet::new(
            title,
            arc_mutex_xml,
            self.shared_strings.clone(),
            self.shared_strings_map.clone(),
            self.styles.clone(),
        )
    }

    /// 指定されたパスへのワークブックのコピー作成
    pub fn copy(&self, path: &str) {
        let new_file: File = OpenOptions::new()
            .write(true)
            .create(true)
            .truncate(true)
            .open(path)
            .expect("Failed to create new file");
        let writer: BufWriter<File> = BufWriter::new(new_file);
        let mut zip_writer: ZipWriter<BufWriter<File>> = ZipWriter::new(writer);
        let options: FileOptions =
            FileOptions::default().compression_method(zip::CompressionMethod::Stored);

        if self.path.is_empty() {
            self.write_to_archive::<_, std::io::Cursor<Vec<u8>>>(None, &mut zip_writer, &options);
        } else {
            let file: File = File::open(&self.path).expect("Failed to open original file");
            let reader: BufReader<File> = BufReader::new(file);
            let mut archive: ZipArchive<BufReader<File>> =
                ZipArchive::new(reader).expect("Failed to open zip archive");
            self.write_to_archive(Some(&mut archive), &mut zip_writer, &options);
        }

        zip_writer.finish().expect("Failed to finish zip writing");
    }
}

use crate::xml::XmlError;

trait ToXml {
    fn to_buf(&self) -> Result<Vec<u8>, XmlError>;
}

impl ToXml for Xml {
    fn to_buf(&self) -> Result<Vec<u8>, XmlError> {
        self.to_buf()
    }
}

impl ToXml for &Xml {
    fn to_buf(&self) -> Result<Vec<u8>, XmlError> {
        (*self).to_buf()
    }
}

impl ToXml for Arc<Mutex<Xml>> {
    fn to_buf(&self) -> Result<Vec<u8>, XmlError> {
        self.lock().unwrap().to_buf()
    }
}

impl ToXml for &Arc<Mutex<Xml>> {
    fn to_buf(&self) -> Result<Vec<u8>, XmlError> {
        self.lock().unwrap().to_buf()
    }
}

impl Book {
    /// 新しい空のワークブックの作成
    fn new_empty_workbook() -> Self {
        let mut rels: HashMap<String, Xml> = HashMap::with_capacity(1);
        let workbook_rels: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"#;
        rels.insert(
            "xl/_rels/workbook.xml.rels".to_string(),
            Xml::new(workbook_rels).expect("Failed to create workbook rels"),
        );

        let workbook_xml: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
</sheets>
</workbook>"#;

        let styles_xml: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"></sst>"#,
            ).expect("Failed to create shared strings"))),
            shared_strings_map: Arc::new(Mutex::new(HashMap::new())),
            styles: Arc::new(Mutex::new(Xml::new(styles_xml).expect("Failed to create styles"))),
            workbook: Xml::new(workbook_xml).expect("Failed to create workbook"),
            vba_project: None,
        }
    }

    /// ファイルからのワークブックの読み込み
    fn from_file(path: &str) -> Self {
        let file: File = File::open(path).unwrap_or_else(|_| panic!("File not found: {path}"));
        let reader: BufReader<File> = BufReader::new(file);
        let mut archive: ZipArchive<BufReader<File>> =
            ZipArchive::new(reader).expect("Failed to open zip archive");

        let mut book: Book = Self::new_empty_workbook();
        book.path = path.to_string();

        for i in 0..archive.len() {
            if let Ok(mut file) = archive.by_index(i) {
                let name: String = file.name().to_string();

                if name.ends_with(XML_SUFFIX) || name.ends_with(XML_RELS_SUFFIX) {
                    let mut contents: String = String::with_capacity(file.size() as usize);
                    if file.read_to_string(&mut contents).is_ok() {
                        if let Ok(xml) = Xml::new(&contents) {
                            match name.as_str() {
                                s if s.starts_with(DRAWINGS_PREFIX) => {
                                    book.drawings.insert(name, xml);
                                }
                                s if s.starts_with(TABLES_PREFIX) => {
                                    book.tables.insert(name, xml);
                                }
                                s if s.starts_with(PIVOT_TABLES_PREFIX) => {
                                    book.pivot_tables.insert(name, xml);
                                }
                                s if s.starts_with(PIVOT_CACHES_PREFIX) => {
                                    book.pivot_caches.insert(name, xml);
                                }
                                s if s.starts_with(THEME_PREFIX) => {
                                    book.themes.insert(name, xml);
                                }
                                s if s.starts_with(WORKSHEETS_PREFIX) => {
                                    book.worksheets.insert(name, Arc::new(Mutex::new(xml)));
                                }
                                s if s.starts_with(WORKBOOK_RELS_PREFIX) => {
                                    book.rels.insert(name, xml);
                                }
                                s if s.starts_with(WORKSHEETS_RELS_PREFIX) => {
                                    book.sheet_rels.insert(name, xml);
                                }
                                WORKBOOK_FILENAME => book.workbook = xml,
                                STYLES_FILENAME => book.styles = Arc::new(Mutex::new(xml)),
                                SHARED_STRINGS_FILENAME => {
                                    let mut map = HashMap::new();
                                    if !xml.elements.is_empty() {
                                        for (i, si) in xml.elements[0].children.iter().enumerate() {
                                            let s = si.get_element("t").get_text().to_string();
                                            map.insert(s, i);
                                        }
                                    }
                                    book.shared_strings = Arc::new(Mutex::new(xml));
                                    book.shared_strings_map = Arc::new(Mutex::new(map));
                                }
                                _ => {}
                            }
                        }
                    }
                } else if name == VBA_PROJECT_FILENAME {
                    let mut contents: Vec<u8> = Vec::new();
                    if file.read_to_end(&mut contents).is_ok() {
                        book.vba_project = Some(contents);
                    }
                }
            }
        }
        book
    }

    /// ワークブックの元のファイルパスへの保存
    pub fn save(&self) {
        let file: File = File::open(&self.path)
            .unwrap_or_else(|_| panic!("Failed to open file for saving: {}", &self.path));
        let reader: BufReader<File> = BufReader::new(file);
        let mut archive: ZipArchive<BufReader<File>> =
            ZipArchive::new(reader).expect("Failed to open zip archive for saving");

        let buffer: Cursor<Vec<u8>> = Cursor::new(Vec::new());
        let mut zip_writer: ZipWriter<Cursor<Vec<u8>>> = ZipWriter::new(buffer);
        let options: FileOptions =
            FileOptions::default().compression_method(zip::CompressionMethod::Stored);

        self.write_to_archive(Some(&mut archive), &mut zip_writer, &options);
    }

    /// ワークブックのzipアーカイブへの書き込み
    fn write_to_archive<W: Write + std::io::Seek, R: Read + std::io::Seek>(
        &self,
        archive: Option<&mut ZipArchive<R>>,
        zip_writer: &mut ZipWriter<W>,
        options: &FileOptions,
    ) {
        // 全XMLファイルへの参照を一つのVecにまとめる
        let mut xmls_with_paths: Vec<(&String, Box<dyn ToXml>)> = Vec::new();
        let workbook_filename_str: String = WORKBOOK_FILENAME.to_string();
        let styles_filename_str: String = STYLES_FILENAME.to_string();
        let shared_strings_filename_str: String = SHARED_STRINGS_FILENAME.to_string();

        xmls_with_paths.push((&workbook_filename_str, Box::new(&self.workbook)));
        xmls_with_paths.push((&styles_filename_str, Box::new(&self.styles)));
        xmls_with_paths.push((&shared_strings_filename_str, Box::new(&self.shared_strings)));

        self.rels
            .iter()
            .for_each(|(k, v)| xmls_with_paths.push((k, Box::new(v))));
        self.drawings
            .iter()
            .for_each(|(k, v)| xmls_with_paths.push((k, Box::new(v))));
        self.tables
            .iter()
            .for_each(|(k, v)| xmls_with_paths.push((k, Box::new(v))));
        self.pivot_tables
            .iter()
            .for_each(|(k, v)| xmls_with_paths.push((k, Box::new(v))));
        self.pivot_caches
            .iter()
            .for_each(|(k, v)| xmls_with_paths.push((k, Box::new(v))));
        self.sheet_rels
            .iter()
            .for_each(|(k, v)| xmls_with_paths.push((k, Box::new(v))));
        self.worksheets
            .iter()
            .for_each(|(k, v)| xmls_with_paths.push((k, Box::new(v))));
        self.themes
            .iter()
            .for_each(|(k, v)| xmls_with_paths.push((k, Box::new(v))));

        if let Some(archive) = archive {
            let file_names: Vec<String> = archive.file_names().map(|s| s.to_string()).collect();
            for filename in file_names {
                if !xmls_with_paths
                    .iter()
                    .any(|(path, _)| path.as_str() == filename)
                    && Some(filename.as_str())
                        != self.vba_project.as_ref().map(|_| VBA_PROJECT_FILENAME)
                {
                    let mut file = archive.by_name(&filename).unwrap_or_else(|_| {
                        panic!("Failed to get file by name {} from zip", &filename)
                    });
                    let mut contents: Vec<u8> = Vec::new();
                    file.read_to_end(&mut contents).unwrap_or_else(|_| {
                        panic!("Failed to read file content from {}", &filename)
                    });
                    zip_writer
                        .start_file(filename.clone(), *options)
                        .unwrap_or_else(|_| {
                            panic!("Failed to start file {} in new zip", &filename)
                        });
                    zip_writer.write_all(&contents).unwrap_or_else(|_| {
                        panic!("Failed to write file content to {}", &filename)
                    });
                }
            }
        }

        for (file_name, xml) in xmls_with_paths {
            zip_writer
                .start_file(file_name, *options)
                .unwrap_or_else(|_| panic!("Failed to start file {file_name} in new zip"));
            let buf: Vec<u8> = xml
                .to_buf()
                .unwrap_or_else(|_| panic!("Failed to convert XML to buffer for {file_name}"));
            zip_writer
                .write_all(&buf)
                .unwrap_or_else(|_| panic!("Failed to write XML to {file_name}"));
        }

        if let Some(vba_project) = &self.vba_project {
            zip_writer
                .start_file(VBA_PROJECT_FILENAME, *options)
                .expect("Failed to start VBA project file in zip");
            zip_writer
                .write_all(vba_project)
                .expect("Failed to write VBA project to zip");
        }
    }

    /// `xl/workbook.xml` からのシートタグの取得
    pub fn sheet_tags(&self) -> &[XmlElement] {
        self.workbook
            .elements
            .first()
            .and_then(|wb| wb.children.iter().find(|c| c.name == "sheets"))
            .map_or(&[], |sheets| &sheets.children)
    }

    /// `xl/workbook.xml.rels` からのリレーションシップのリスト取得
    pub fn get_relationships(&self) -> &[XmlElement] {
        self.rels
            .get("xl/_rels/workbook.xml.rels")
            .and_then(|rels| rels.elements.first())
            .map_or(&[], |r| &r.children)
    }

    /// シート名とそのパスのマップ取得
    pub fn get_sheet_paths(&self) -> HashMap<String, String> {
        let relationships: &[XmlElement] = self.get_relationships();
        let sheet_paths: HashMap<String, String> = relationships
            .iter()
            .filter_map(|rel| {
                let id: String = rel.attributes.get("Id")?.clone();
                let target: String = rel.attributes.get("Target")?.clone();
                Some((id, target))
            })
            .collect();

        self.sheet_tags()
            .iter()
            .filter_map(|tag| {
                let name: String = tag.attributes.get("name")?.clone();
                let r_id: &String = tag.attributes.get("r:id")?;
                let path: &String = sheet_paths.get(r_id)?;
                let trimmed_path: &str = path.trim_start_matches("/xl/").trim_start_matches("xl/");
                Some((name, format!("xl/{trimmed_path}")))
            })
            .collect()
    }

    /// 名前によるシートの取得
    pub fn get_sheet_by_name(&self, name: &str) -> Option<Sheet> {
        let sheet_paths: HashMap<String, String> = self.get_sheet_paths();
        sheet_paths.get(name).and_then(|sheet_path| {
            self.worksheets.get(sheet_path).map(|xml| {
                Sheet::new(
                    name.to_string(),
                    xml.clone(),
                    self.shared_strings.clone(),
                    self.shared_strings_map.clone(),
                    self.styles.clone(),
                )
            })
        })
    }

    /// テーブルXMLの作成
    fn create_table_xml(
        &mut self,
        name: &str,
        table_ref: &str,
        table_id: usize,
        table_filename: &str,
    ) {
        let table_xml_str: String = format!(
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
        );
        let new_table_xml: Xml =
            Xml::new(&table_xml_str).expect("Failed to create new table XML from string");
        self.tables
            .insert(table_filename.to_string(), new_table_xml);
    }

    /// ワークシートへのテーブルパーツの追加
    fn add_table_parts_to_worksheet(&mut self, sheet_name: &str, table_id: usize) {
        let sheet_path: String = self
            .get_sheet_paths()
            .get(sheet_name)
            .unwrap_or_else(|| panic!("Sheet {sheet_name} not found"))
            .clone();
        if let Some(sheet_xml_mutex) = self.worksheets.get_mut(&sheet_path) {
            let mut sheet_xml: MutexGuard<Xml> = sheet_xml_mutex
                .lock()
                .unwrap_or_else(|_| panic!("Failed to lock sheet xml for {sheet_name}"));
            if let Some(worksheet) = sheet_xml.elements.get_mut(0) {
                let mut attributes: HashMap<String, String> = HashMap::with_capacity(1);
                attributes.insert("count".to_string(), "1".to_string());

                let mut table_part_attributes: HashMap<String, String> = HashMap::with_capacity(1);
                table_part_attributes.insert("r:id".to_string(), format!("rId{table_id}"));

                let table_part: XmlElement = XmlElement {
                    name: "tablePart".to_string(),
                    attributes: table_part_attributes,
                    ..Default::default()
                };

                worksheet.children.push(XmlElement {
                    name: "tableParts".to_string(),
                    attributes,
                    children: vec![table_part],
                    text: None,
                });
            }
        }
    }

    /// ワークシートのリレーションシップへのテーブルリレーションシップの追加
    fn add_table_relationship(&mut self, sheet_name: &str, table_id: usize) {
        let sheet_paths: HashMap<String, String> = self.get_sheet_paths();
        let sheet_path: &String = sheet_paths
            .get(sheet_name)
            .unwrap_or_else(|| panic!("Sheet path not found for sheet {sheet_name}"));
        let rels_filename: String = format!(
            "xl/worksheets/_rels/{}.rels",
            sheet_path
                .split('/')
                .next_back()
                .expect("Could not extract filename from sheet path")
        );
        let rels: &mut Xml = self.sheet_rels.entry(rels_filename).or_insert_with(|| {
            Xml::new(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"#,
            )
            .expect("Failed to create new relationships XML")
        });

        if rels.elements.is_empty() {
            rels.elements.push(XmlElement {
                name: "Relationships".to_string(),
                ..Default::default()
            });
        }
        if let Some(relationships) = rels.elements.get_mut(0) {
            let mut attributes: HashMap<String, String> = HashMap::with_capacity(3);
            attributes.insert("Id".to_string(), format!("rId{table_id}"));
            attributes.insert(
                "Type".to_string(),
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"
                    .to_string(),
            );
            attributes.insert(
                "Target".to_string(),
                format!("../tables/table{table_id}.xml"),
            );

            relationships.children.push(XmlElement {
                name: "Relationship".to_string(),
                attributes,
                ..Default::default()
            });
        }
    }

    /// workbook.xml へのシートの追加
    fn add_sheet_to_workbook_xml(
        &mut self,
        title: &str,
        sheet_id: usize,
        r_id: &str,
        index: usize,
    ) {
        if let Some(sheets_tag) = self
            .workbook
            .elements
            .first_mut()
            .and_then(|wb| wb.children.iter_mut().find(|x| x.name == "sheets"))
        {
            let mut sheet_element: XmlElement = XmlElement {
                name: "sheet".to_string(),
                attributes: HashMap::with_capacity(3),
                ..Default::default()
            };
            sheet_element
                .attributes
                .insert("name".to_string(), title.to_string());
            sheet_element
                .attributes
                .insert("sheetId".to_string(), sheet_id.to_string());
            sheet_element
                .attributes
                .insert("r:id".to_string(), r_id.to_string());

            if index < sheets_tag.children.len() {
                sheets_tag.children.insert(index, sheet_element);
            } else {
                sheets_tag.children.push(sheet_element);
            }
        }
    }

    /// ワークブックのリレーションシップへのシートリレーションシップの追加
    fn add_sheet_relationship(&mut self, r_id: &str, sheet_id: usize) {
        if let Some(relationships_tag) = self
            .rels
            .get_mut("xl/_rels/workbook.xml.rels")
            .and_then(|rels| rels.elements.first_mut())
        {
            let mut relationship_element: XmlElement = XmlElement {
                name: "Relationship".to_string(),
                attributes: HashMap::with_capacity(3),
                ..Default::default()
            };
            relationship_element
                .attributes
                .insert("Id".to_string(), r_id.to_string());
            relationship_element.attributes.insert(
                "Type".to_string(),
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
                    .to_string(),
            );
            relationship_element.attributes.insert(
                "Target".to_string(),
                format!("worksheets/sheet{sheet_id}.xml"),
            );
            relationships_tag.children.push(relationship_element);
        }
    }
}
