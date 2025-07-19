use std::collections::HashMap;
use std::fs::{File, OpenOptions};
use std::io::{Cursor, Read, Write};
use std::sync::{Arc, Mutex};
use zip::write::FileOptions;
use zip::{ZipArchive, ZipWriter};

use pyo3::prelude::*;

use crate::xlsx::sheet::Sheet;
use crate::xlsx::xml::{Xml, XmlElement};

const XML_SUFFIX: &str = ".xml";
const XML_RELS_SUFFIX: &str = ".xml.rels";

const WORKBOOK_FILENAME: &str = "xl/workbook.xml";
const STYLES_FILENAME: &str = "xl/styles.xml";
const SHARED_STRINGS_FILENAME: &str = "xl/sharedStrings.xml";

const WORKBOOK_RELS_PREFIX: &str = "xl/_rels/";
const DRAWINGS_PREFIX: &str = "xl/drawings/";
const THEME_PREFIX: &str = "xl/theme/";
const WORKSHEETS_PREFIX: &str = "xl/worksheets/";

const INIT_EXCEL_FILENAME: &str = "/src/xlsx/init.xlsx";

#[pyclass]
pub struct Book {
    #[pyo3(get, set)]
    pub path: String,

    /// `xl/_rels/`以下に存在するファイル
    pub rels: HashMap<String, Xml>,

    /// `xl/drawings/`以下に存在するファイル
    pub drawings: HashMap<String, Xml>,

    /// `xl/theme/`以下に存在するファイル
    pub themes: HashMap<String, Xml>,

    /// `xl/worksheets/`以下に存在するファイル
    pub worksheets: HashMap<String, Arc<Mutex<Xml>>>,

    /// `xl/sharedStrings.xml`
    pub shared_strings: Xml,

    /// `xl/styles.xml`
    pub styles: Xml,

    /// `workbook.xml`
    pub workbook: Xml,
}

#[pymethods]
impl Book {
    #[new]
    #[pyo3(signature = (path=INIT_EXCEL_FILENAME.to_string()))]
    pub fn new(path: String) -> Self {
        let file_result: Result<File, std::io::Error> = File::open(&path);
        if file_result.is_err() {
            panic!("File not found: {}", path);
        }
        let file = file_result.unwrap();
        let mut archive: ZipArchive<File> = ZipArchive::new(file).unwrap();

        let mut rels: HashMap<String, Xml> = HashMap::new();
        let mut drawings: HashMap<String, Xml> = HashMap::new();
        let mut themes: HashMap<String, Xml> = HashMap::new();
        let mut worksheets: HashMap<String, Arc<Mutex<Xml>>> = HashMap::new();
        let mut shared_strings: Xml = Xml::new(&String::new());
        let mut styles: Xml = Xml::new(&String::new());
        let mut workbook: Xml = Xml::new(&String::new());

        for i in 0..archive.len() {
            let mut file: zip::read::ZipFile<'_> = archive.by_index(i).unwrap();
            let name: String = file.name().to_string();

            if name.ends_with(XML_SUFFIX) {
                let mut contents: String = String::new();
                file.read_to_string(&mut contents).unwrap();
                let xml = Xml::new(&contents);

                if name.starts_with(DRAWINGS_PREFIX) {
                    drawings.insert(name, xml);
                } else if name.starts_with(THEME_PREFIX) {
                    themes.insert(name, xml);
                } else if name.starts_with(WORKSHEETS_PREFIX) {
                    worksheets.insert(name, Arc::new(Mutex::new(xml)));
                } else if name == WORKBOOK_FILENAME {
                    workbook = xml;
                } else if name == STYLES_FILENAME {
                    styles = xml;
                } else if name == SHARED_STRINGS_FILENAME {
                    shared_strings = xml;
                }
            } else if name.ends_with(XML_RELS_SUFFIX) && name.starts_with(WORKBOOK_RELS_PREFIX) {
                let mut contents: String = String::new();
                file.read_to_string(&mut contents).unwrap();
                rels.insert(name, Xml::new(&contents));
            }
        }

        Book {
            path,
            rels,
            drawings,
            themes,
            worksheets,
            shared_strings,
            styles,
            workbook,
        }
    }

    #[getter]
    pub fn sheetnames(&self) -> Vec<String> {
        self.sheet_tags()
            .iter()
            .map(|x| x.attributes.get("name").unwrap().clone())
            .collect()
    }

    pub fn __contains__(&self, key: String) -> bool {
        self.sheetnames().contains(&key)
    }

    pub fn create_sheet(&mut self, title: String, index: usize) -> Sheet {
        // 次のシートIDとrIdを取得
        let sheet_tags: Vec<XmlElement> = self.sheet_tags();
        let next_sheet_id: usize = sheet_tags.len() + 1;
        let next_rid: String = format!("rId{}", self.get_relationships().len() + 1);

        // シートパスを作成
        let sheet_path: String = format!("xl/worksheets/sheet{}.xml", next_sheet_id);

        // 空のワークシートXMLを作成
        let worksheet_xml: Xml = Xml::new(
            &r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                <sheetData/>
            </worksheet>"#
                .to_string(),
        );

        // ワークシートをコレクションに追加
        let arc_mutex_xml: Arc<Mutex<Xml>> = Arc::new(Mutex::new(worksheet_xml));
        self.worksheets
            .insert(sheet_path.clone(), arc_mutex_xml.clone());

        // workbook.xmlを更新して新しいシートを含める
        if let Some(workbook_tag) = self.workbook.elements.first_mut() {
            if let Some(sheets_tag) = workbook_tag
                .children
                .iter_mut()
                .find(|x| x.name == "sheets")
            {
                // 新しいシート要素を作成
                let mut sheet_element: XmlElement = XmlElement {
                    name: "sheet".to_string(),
                    attributes: HashMap::new(),
                    children: Vec::new(),
                    text: None,
                };

                // 属性を追加
                sheet_element
                    .attributes
                    .insert("name".to_string(), title.clone());
                sheet_element
                    .attributes
                    .insert("sheetId".to_string(), next_sheet_id.to_string());
                sheet_element
                    .attributes
                    .insert("r:id".to_string(), next_rid.clone());

                // 指定されたインデックスに挿入、またはリストの最後に追加
                if index < sheets_tag.children.len() {
                    sheets_tag.children.insert(index, sheet_element);
                } else {
                    sheets_tag.children.push(sheet_element);
                }
            }
        }

        // workbook.xml.relsを更新して関係を含める
        if let Some(workbook_rels) = self.rels.get_mut("xl/_rels/workbook.xml.rels") {
            if let Some(relationships_tag) = workbook_rels.elements.first_mut() {
                // 新しい関係要素を作成
                let mut relationship_element: XmlElement = XmlElement {
                    name: "Relationship".to_string(),
                    attributes: HashMap::new(),
                    children: Vec::new(),
                    text: None,
                };

                // 属性を追加
                relationship_element
                    .attributes
                    .insert("Id".to_string(), next_rid);
                relationship_element.attributes.insert(
                    "Type".to_string(),
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
                        .to_string(),
                );
                relationship_element.attributes.insert(
                    "Target".to_string(),
                    format!("worksheets/sheet{}.xml", next_sheet_id),
                );

                // 関係を追加
                relationships_tag.children.push(relationship_element);
            }
        }

        // Sheetオブジェクトを作成して返す
        Sheet::new(title, arc_mutex_xml)
    }
}

impl Book {
    pub fn save(&self) {
        let file: File = File::open(&self.path).unwrap();
        let mut archive: ZipArchive<File> = ZipArchive::new(file).unwrap();

        let mut buffer: Cursor<Vec<u8>> = Cursor::new(Vec::new());
        let mut zip_writer: ZipWriter<&mut Cursor<Vec<u8>>> = ZipWriter::new(&mut buffer);
        let options: FileOptions =
            FileOptions::default().compression_method(zip::CompressionMethod::Stored);

        // 構造体内のxmlを集合
        let xmls: HashMap<String, Xml> = self.merge_xmls();

        Book::write_file(&mut archive, &xmls, &mut zip_writer, &options);
    }

    pub fn copy(&self, path: &str) {
        let file: File = File::open(&self.path).unwrap();
        let mut archive: ZipArchive<File> = ZipArchive::new(file).unwrap();

        // 新しいファイルを作成
        let new_file: File = OpenOptions::new()
            .write(true)
            .create(true)
            .truncate(true)
            .open(path)
            .unwrap();

        let mut zip_writer: ZipWriter<File> = ZipWriter::new(new_file);
        let options: FileOptions =
            FileOptions::default().compression_method(zip::CompressionMethod::Stored);

        // 構造体内のxmlを集合
        let xmls: HashMap<String, Xml> = self.merge_xmls();

        Book::write_file(&mut archive, &xmls, &mut zip_writer, &options);

        // ファイルを閉じる（ZipWriterのdropで自動的に行われるが、明示的にfinishを呼ぶ）
        zip_writer.finish().unwrap();
    }

    /// 構造体内のxmlを集合
    pub fn merge_xmls(&self) -> HashMap<String, Xml> {
        let mut xmls: HashMap<String, Xml> = self.rels.clone();
        xmls.insert(WORKBOOK_FILENAME.to_string(), self.workbook.clone());
        xmls.insert(STYLES_FILENAME.to_string(), self.styles.clone());
        xmls.insert(
            SHARED_STRINGS_FILENAME.to_string(),
            self.shared_strings.clone(),
        );
        xmls.extend(self.drawings.clone());

        // Arc<Mutex<Xml>>からXmlを取得
        for (key, arc_mutex_xml) in &self.worksheets {
            let xml: Xml = arc_mutex_xml.lock().unwrap().clone();
            xmls.insert(key.clone(), xml);
        }

        xmls.extend(self.themes.clone());
        xmls
    }

    /// ファイルへの書き込み
    pub fn write_file<W: Write + std::io::Seek>(
        archive: &mut ZipArchive<File>,
        xmls: &HashMap<String, Xml>,
        zip_writer: &mut ZipWriter<W>,
        options: &FileOptions,
    ) {
        let file_names: Vec<String> = archive.file_names().map(|s| s.to_string()).collect();
        for filename in file_names {
            if !xmls.contains_key(&filename) {
                let mut file: zip::read::ZipFile<'_> = archive.by_name(&filename).unwrap();
                let mut contents: Vec<u8> = Vec::new();
                file.read_to_end(&mut contents).unwrap();
                zip_writer.start_file(&filename, *options).unwrap();
                zip_writer.write_all(&contents).unwrap();
            }
        }

        for (file_name, xml) in xmls {
            zip_writer.start_file(file_name, *options).unwrap();
            zip_writer.write_all(&xml.to_buf()).unwrap();
        }
    }

    /// xl/workbook.xml内のsheetタグを取得する
    pub fn sheet_tags(&self) -> Vec<XmlElement> {
        if let Some(workbook_tag) = self.workbook.elements.first() {
            if let Some(sheets_tag) = workbook_tag.children.iter().find(|&x| x.name == *"sheets") {
                return sheets_tag.children.clone();
            }
        }
        Vec::new()
    }

    // xl/workbook.xml.rels内のRelationshipタグ一覧を取得する
    pub fn get_relationships(&self) -> Vec<XmlElement> {
        if let Some(workbook_xml_rels) = self.rels.get("xl/_rels/workbook.xml.rels") {
            if let Some(workbook_tag) = workbook_xml_rels.elements.first() {
                return workbook_tag.children.clone();
            }
        }
        Vec::new()
    }

    // sheet_path一覧を取得する
    pub fn get_sheet_paths(&self) -> HashMap<String, String> {
        let mut result: HashMap<String, String> = HashMap::new();
        let sheet_tags: Vec<XmlElement> = self.sheet_tags();
        let relationships: Vec<XmlElement> = self.get_relationships().clone();
        let sheet_paths: HashMap<String, String> = relationships
            .into_iter()
            .map(|x: XmlElement| {
                (
                    x.attributes.get("Id").unwrap().clone(),
                    x.attributes.get("Target").unwrap().clone(),
                )
            })
            .collect();
        for sheet_tag in sheet_tags {
            let id: &str = sheet_tag.attributes.get("r:id").unwrap().as_str();
            let sheet_path: &String = sheet_paths.get(id).unwrap();
            result.insert(
                sheet_tag.attributes.get("name").unwrap().clone(),
                format!("xl/{}", sheet_path),
            );
        }
        result
    }

    // シート名からシートを取得する
    pub fn get_sheet_by_name(&self, name: &str) -> Option<Sheet> {
        let sheet_paths: HashMap<String, String> = self.get_sheet_paths();
        if let Some(sheet_path) = sheet_paths.get(name) {
            if let Some(xml) = self.worksheets.get(sheet_path) {
                return Some(Sheet::new(name.to_string(), xml.clone()));
            }
        }
        None
    }
}
