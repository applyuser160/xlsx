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
const VBA_PROJECT_FILENAME: &str = "xl/vbaProject.bin";

const WORKBOOK_FILENAME: &str = "xl/workbook.xml";
const STYLES_FILENAME: &str = "xl/styles.xml";
const SHARED_STRINGS_FILENAME: &str = "xl/sharedStrings.xml";

const WORKBOOK_RELS_PREFIX: &str = "xl/_rels/";
const WORKSHEETS_RELS_PREFIX: &str = "xl/worksheets/_rels/";
const DRAWINGS_PREFIX: &str = "xl/drawings/";
const THEME_PREFIX: &str = "xl/theme/";
const WORKSHEETS_PREFIX: &str = "xl/worksheets/";
const TABLES_PREFIX: &str = "xl/tables/";
const PIVOT_TABLES_PREFIX: &str = "xl/pivotTables/";
const PIVOT_CACHES_PREFIX: &str = "xl/pivotCache/";

const INIT_EXCEL_FILENAME: &str = "/src/xlsx/init.xlsx";

#[pyclass]
pub struct Book {
    #[pyo3(get, set)]
    pub path: String,

    /// `xl/_rels/`以下に存在するファイル
    pub rels: HashMap<String, Xml>,

    /// `xl/drawings/`以下に存在するファイル
    pub drawings: HashMap<String, Xml>,

    /// `xl/tables/`以下に存在するファイル
    pub tables: HashMap<String, Xml>,

    /// `xl/pivotTables/`以下に存在するファイル
    pub pivot_tables: HashMap<String, Xml>,

    /// `xl/pivotCache/`以下に存在するファイル
    pub pivot_caches: HashMap<String, Xml>,

    /// `xl/theme/`以下に存在するファイル
    pub themes: HashMap<String, Xml>,

    /// `xl/worksheets/`以下に存在するファイル
    pub worksheets: HashMap<String, Arc<Mutex<Xml>>>,

    /// `xl/worksheets/_rels/`以下に存在するファイル
    pub sheet_rels: HashMap<String, Xml>,

    /// `xl/sharedStrings.xml`
    pub shared_strings: Arc<Mutex<Xml>>,

    /// `xl/styles.xml`
    pub styles: Arc<Mutex<Xml>>,

    /// `workbook.xml`
    pub workbook: Xml,

    /// `vbaProject.bin`
    pub vba_project: Option<Vec<u8>>,
}

#[pymethods]
impl Book {
    #[new]
    #[pyo3(signature = (path = ""))]
    pub fn new(path: &str) -> Self {
        if path.is_empty() {
            // New workbook
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
        } else {
            let file_result: Result<File, std::io::Error> = File::open(path);
        if file_result.is_err() {
            panic!("File not found: {path}");
        }
        let file = file_result.unwrap();
        let mut archive: ZipArchive<File> = ZipArchive::new(file).unwrap();

        let mut rels: HashMap<String, Xml> = HashMap::new();
        let mut drawings: HashMap<String, Xml> = HashMap::new();
        let mut tables: HashMap<String, Xml> = HashMap::new();
        let mut pivot_tables: HashMap<String, Xml> = HashMap::new();
        let mut pivot_caches: HashMap<String, Xml> = HashMap::new();
        let mut themes: HashMap<String, Xml> = HashMap::new();
        let mut worksheets: HashMap<String, Arc<Mutex<Xml>>> = HashMap::new();
        let mut sheet_rels: HashMap<String, Xml> = HashMap::new();
        let mut shared_strings: Arc<Mutex<Xml>> = Arc::new(Mutex::new(Xml::new(&String::new())));
        let mut styles: Arc<Mutex<Xml>> = Arc::new(Mutex::new(Xml::new(&String::new())));
        let mut workbook: Xml = Xml::new(&String::new());
        let mut vba_project: Option<Vec<u8>> = None;

        for i in 0..archive.len() {
            let mut file: zip::read::ZipFile<'_> = archive.by_index(i).unwrap();
            let name: String = file.name().to_string();

            if name.ends_with(XML_SUFFIX) {
                let mut contents: String = String::new();
                file.read_to_string(&mut contents).unwrap();
                let xml = Xml::new(&contents);

                if name.starts_with(DRAWINGS_PREFIX) {
                    drawings.insert(name, xml);
                } else if name.starts_with(TABLES_PREFIX) {
                    tables.insert(name, xml);
                } else if name.starts_with(PIVOT_TABLES_PREFIX) {
                    pivot_tables.insert(name, xml);
                } else if name.starts_with(PIVOT_CACHES_PREFIX) {
                    pivot_caches.insert(name, xml);
                } else if name.starts_with(THEME_PREFIX) {
                    themes.insert(name, xml);
                } else if name.starts_with(WORKSHEETS_PREFIX) {
                    worksheets.insert(name, Arc::new(Mutex::new(xml)));
                } else if name == WORKBOOK_FILENAME {
                    workbook = xml;
                } else if name == STYLES_FILENAME {
                    styles = Arc::new(Mutex::new(xml));
                } else if name == SHARED_STRINGS_FILENAME {
                    shared_strings = Arc::new(Mutex::new(xml));
                }
            } else if name.ends_with(XML_RELS_SUFFIX) {
                if name.starts_with(WORKBOOK_RELS_PREFIX) {
                    let mut contents: String = String::new();
                    file.read_to_string(&mut contents).unwrap();
                    rels.insert(name, Xml::new(&contents));
                } else if name.starts_with(WORKSHEETS_RELS_PREFIX) {
                    let mut contents: String = String::new();
                    file.read_to_string(&mut contents).unwrap();
                    sheet_rels.insert(name, Xml::new(&contents));
                }
            } else if name == VBA_PROJECT_FILENAME {
                let mut contents: Vec<u8> = Vec::new();
                file.read_to_end(&mut contents).unwrap();
                vba_project = Some(contents);
            }
        }

        Book {
            path: path.to_string(),
            rels,
            drawings,
            tables,
            pivot_tables,
            pivot_caches,
            themes,
            worksheets,
            sheet_rels,
            shared_strings,
            styles,
            workbook,
            vba_project,
        }
    }
}

    #[getter]
    pub fn sheetnames(&self) -> Vec<String> {
        self.sheet_tags()
            .iter()
            .map(|x| x.attributes.get("name").unwrap().clone())
            .collect()
    }

    pub fn __iter__(&self) -> Vec<String> {
        self.sheetnames()
    }

    pub fn __contains__(&self, key: String) -> bool {
        self.sheetnames().contains(&key)
    }

    pub fn __getitem__(&self, key: String) -> Sheet {
        if let Some(sheet) = self.get_sheet_by_name(key.as_str()) {
            return sheet;
        }
        panic!("No sheet named '{key}'");
    }

    pub fn add_table(&mut self, sheet_name: String, name: String, table_ref: String) {
        let table_id = self.tables.len() + 1;
        let table_filename = format!("xl/tables/table{}.xml", table_id);

        let new_table_xml = Xml::new(&format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="{}" name="{}" displayName="{}" ref="{}" totalsRowShown="0">
    <autoFilter ref="{}"/>
    <tableColumns count="3">
        <tableColumn id="1" name="Column1"/>
        <tableColumn id="2" name="Column2"/>
        <tableColumn id="3" name="Column3"/>
    </tableColumns>
    <tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
</table>"#,
            table_id, name, name, table_ref, table_ref
        ));
        self.tables.insert(table_filename.clone(), new_table_xml);

        let sheet_path = self.get_sheet_paths().get(&sheet_name).unwrap().clone();
        let sheet_xml = self.worksheets.get_mut(&sheet_path).unwrap();
        let mut sheet_xml_locked = sheet_xml.lock().unwrap();
        let worksheet = &mut sheet_xml_locked.elements[0];
        worksheet.children.push(XmlElement {
            name: "tableParts".to_string(),
            attributes: {
                let mut map = HashMap::new();
                map.insert("count".to_string(), "1".to_string());
                map
            },
            children: vec![
                XmlElement {
                    name: "tablePart".to_string(),
                    attributes: {
                        let mut map = HashMap::new();
                        map.insert("r:id".to_string(), format!("rId{}", table_id));
                        map
                    },
                    children: Vec::new(),
                    text: None,
                }
            ],
            text: None,
        });

        let rels_filename = format!("xl/worksheets/_rels/{}.rels", sheet_path.split('/').last().unwrap());
        let rels = self.sheet_rels.entry(rels_filename).or_insert_with(|| Xml::new(&r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"#.to_string()));

        if rels.elements.is_empty() {
            rels.elements.push(XmlElement {
                name: "Relationships".to_string(),
                attributes: HashMap::new(),
                children: Vec::new(),
                text: None,
            });
        }
        let relationships = &mut rels.elements[0];
        relationships.children.push(XmlElement {
            name: "Relationship".to_string(),
            attributes: {
                let mut map = HashMap::new();
                map.insert("Id".to_string(), format!("rId{}", table_id));
                map.insert("Type".to_string(), "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table".to_string());
                map.insert("Target".to_string(), format!("../tables/table{}.xml", table_id));
                map
            },
            children: Vec::new(),
            text: None,
        });
    }

    pub fn __delitem__(&mut self, key: String) {
        if let Some(sheet) = self.get_sheet_by_name(key.as_str()) {
            self.remove(&sheet);
            return;
        }
        panic!("No sheet named '{key}'");
    }

    pub fn index(&self, sheet: &Sheet) -> usize {
        let sheet_name = &sheet.name;
        let sheet_names = self.sheetnames();
        if let Some(sheet_index) = sheet_names.iter().position(|x| x == sheet_name) {
            return sheet_index;
        }
        panic!("No sheet named '{sheet_name}'");
    }

    pub fn remove(&mut self, sheet: &Sheet) {
        let sheet_paths = self.get_sheet_paths();
        if let Some(sheet_path) = sheet_paths.get(&sheet.name) {
            if self.worksheets.contains_key(sheet_path) {
                self.worksheets.remove(sheet_path);

                // workbook.xmlからsheetタグを削除し、r:idを取得
                let mut rid_to_remove = String::new();
                if let Some(workbook_tag) = self.workbook.elements.first_mut() {
                    if let Some(sheets_tag) = workbook_tag
                        .children
                        .iter_mut()
                        .find(|x| x.name == "sheets")
                    {
                        if let Some(sheet_element) = sheets_tag
                            .children
                            .iter()
                            .find(|s| s.attributes.get("name") == Some(&sheet.name))
                        {
                            if let Some(rid) = sheet_element.attributes.get("r:id") {
                                rid_to_remove = rid.clone();
                            }
                        }
                        sheets_tag
                            .children
                            .retain(|s| s.attributes.get("name") != Some(&sheet.name));
                    }
                }

                // workbook.xml.relsからRelationshipタグを削除
                if !rid_to_remove.is_empty() {
                    if let Some(workbook_rels) = self.rels.get_mut("xl/_rels/workbook.xml.rels") {
                        if let Some(relationships_tag) = workbook_rels.elements.first_mut() {
                            relationships_tag
                                .children
                                .retain(|r| r.attributes.get("Id") != Some(&rid_to_remove));
                        }
                    }
                }
                return;
            }
        }
        panic!("No sheet named '{}'", sheet.name);
    }

    pub fn create_sheet(&mut self, title: String, index: usize) -> Sheet {
        // 次のシートIDとrIdを取得
        let sheet_tags: Vec<XmlElement> = self.sheet_tags();
        let next_sheet_id: usize = sheet_tags.len() + 1;
        let next_rid: String = format!("rId{}", self.get_relationships().len() + 1);

        // シートパスを作成
        let sheet_path: String = format!("xl/worksheets/sheet{next_sheet_id}.xml");

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
                    format!("worksheets/sheet{next_sheet_id}.xml"),
                );

                // 関係を追加
                relationships_tag.children.push(relationship_element);
            }
        }

        // Sheetオブジェクトを作成して返す
        Sheet::new(
            title,
            arc_mutex_xml,
            self.shared_strings.clone(),
            self.styles.clone(),
        )
    }

    pub fn copy(&self, path: &str) {
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
        let xmls: HashMap<String, Xml> = self.merge_xmls();

        if self.path.is_empty() {
            for (file_name, xml) in &xmls {
                zip_writer.start_file(file_name, options).unwrap();
                zip_writer.write_all(&xml.to_buf()).unwrap();
            }
        } else {
            let file = File::open(&self.path).unwrap();
            let mut archive = ZipArchive::new(file).unwrap();
            self.write_file(&mut archive, &xmls, &mut zip_writer, &options);
        }

        zip_writer.finish().unwrap();
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

        self.write_file(&mut archive, &xmls, &mut zip_writer, &options);
    }

    /// 構造体内のxmlを集合
    pub fn merge_xmls(&self) -> HashMap<String, Xml> {
        let mut xmls: HashMap<String, Xml> = self.rels.clone();
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
        &self,
        archive: &mut ZipArchive<File>,
        xmls: &HashMap<String, Xml>,
        zip_writer: &mut ZipWriter<W>,
        options: &FileOptions,
    ) {
        let file_names: Vec<String> = archive.file_names().map(|s| s.to_string()).collect();
        for filename in file_names {
            if !xmls.contains_key(&filename) && Some(filename.as_str()) != self.vba_project.as_ref().map(|_| VBA_PROJECT_FILENAME) {
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

        if let Some(vba_project) = &self.vba_project {
            zip_writer.start_file(VBA_PROJECT_FILENAME, *options).unwrap();
            zip_writer.write_all(vba_project).unwrap();
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
            let trimmed_path = sheet_path
                .trim_start_matches("/xl/")
                .trim_start_matches("xl/");
            result.insert(
                sheet_tag.attributes.get("name").unwrap().clone(),
                format!("xl/{trimmed_path}"),
            );
        }
        result
    }

    // シート名からシートを取得する
    pub fn get_sheet_by_name(&self, name: &str) -> Option<Sheet> {
        let sheet_paths: HashMap<String, String> = self.get_sheet_paths();
        if let Some(sheet_path) = sheet_paths.get(name) {
            if let Some(xml) = self.worksheets.get(sheet_path) {
                return Some(Sheet::new(
                    name.to_string(),
                    xml.clone(),
                    self.shared_strings.clone(),
                    self.styles.clone(),
                ));
            }
        }
        None
    }
}
