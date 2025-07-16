use std::collections::HashMap;
use std::fs::{File, OpenOptions};
use std::io::{Cursor, Read, Write};
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
    pub worksheets: HashMap<String, Xml>,

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
        let file = File::open(&path).unwrap();
        let mut archive: ZipArchive<File> = ZipArchive::new(file).unwrap();

        let mut rels: HashMap<String, Xml> = HashMap::new();
        let mut drawings: HashMap<String, Xml> = HashMap::new();
        let mut themes: HashMap<String, Xml> = HashMap::new();
        let mut worksheets: HashMap<String, Xml> = HashMap::new();
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
                    worksheets.insert(name, xml);
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
        self.sheet_tags().iter().map(|x| x.name.clone()).collect()
    }

    pub fn __contains__(&self, key: String) -> bool {
        self.sheetnames().contains(&key)
    }

    // pub fn __repr__(&self) -> String {
    //     format!("<Book path='{}'>", self.path)
    // }

    // pub fn get_value(&self, sheet: String, address: String) -> String {
    //     let worksheet = self.get_sheet_by_name(&sheet);
    //     return match worksheet {
    //         Some(ws) => ws.get_value(address),
    //         None => "Sheet not found".to_string(),
    //     };
    // }
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
    fn merge_xmls(&self) -> HashMap<String, Xml> {
        let mut xmls: HashMap<String, Xml> = self.rels.clone();
        xmls.insert(WORKBOOK_FILENAME.to_string(), self.workbook.clone());
        xmls.insert(STYLES_FILENAME.to_string(), self.styles.clone());
        xmls.insert(
            SHARED_STRINGS_FILENAME.to_string(),
            self.shared_strings.clone(),
        );
        xmls.extend(self.drawings.clone());
        xmls.extend(self.worksheets.clone());
        xmls.extend(self.themes.clone());
        xmls
    }

    /// ファイルへの書き込み
    fn write_file<W: Write + std::io::Seek>(
        archive: &mut ZipArchive<File>,
        xmls: &HashMap<String, Xml>,
        zip_writer: &mut ZipWriter<W>,
        options: &FileOptions,
    ) {
        let file_names: Vec<String> = archive.file_names().map(|s| s.to_string()).collect();
        for filename in file_names {
            if !xmls.contains_key(&filename) {
                let mut file = archive.by_name(&filename).unwrap();
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
    fn sheet_tags(&self) -> Vec<XmlElement> {
        if let Some(workbook_tag) = self.workbook.elements.first() {
            if let Some(sheets_tag) = workbook_tag.children.iter().find(|&x| x.name == *"sheet") {
                return sheets_tag.children.clone();
            }
        }
        Vec::new()
    }

    // xl/workbook.xml.rels内のRelationshipタグ一覧を取得する
    fn get_relationships(&self) -> Vec<XmlElement> {
        if let Some(workbook_xml_rels) = self.rels.get("xl/workbook.xml.rels") {
            if let Some(workbook_tag) = workbook_xml_rels.elements.first() {
                return workbook_tag.children.clone();
            }
        }
        Vec::new()
    }

    // sheet_path一覧を取得する
    fn get_sheet_paths(&self) -> HashMap<String, String> {
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
            let id: &str = sheet_tag.attributes.get("Id").unwrap().as_str();
            let sheet_path: &String = sheet_paths.get(id).unwrap();
            result.insert(
                sheet_tag.attributes.get("name").unwrap().clone(),
                sheet_path.clone(),
            );
        }
        result
    }

    // sheet一覧を取得する
    fn get_sheets(&self) -> Vec<Sheet> {
        let mut result: Vec<Sheet> = Vec::new();
        let sheet_paths: HashMap<String, String> = self.get_sheet_paths();
        for (_id, path) in sheet_paths {
            let xml: Xml = self.worksheets.get(&path).unwrap().clone();
            let sheet: Sheet = Sheet::new(path, xml);
            result.push(sheet);
        }
        result
    }

    // pub fn get_sheet_by_name(&self, name: &String) -> Option<&Worksheet> {
    //     self.value.get_sheet_by_name(name)
    // }

    // pub fn get_sheet_by_index(&self, index: &usize) -> Option<&Worksheet> {
    //     self.value.get_sheet(index)
    // }
}

impl Drop for Book {
    fn drop(&mut self) {
        self.save();
    }
}
