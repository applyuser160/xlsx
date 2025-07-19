use pyo3::prelude::*;
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};
use std::collections::HashMap;
use std::fs::OpenOptions;
use std::io::{BufRead, BufWriter, Write};

/// Xmlオブジェクト
#[pyclass]
#[derive(Debug, Clone)]
pub struct Xml {
    // /// ファイルパス
    // pub path: String,
    /// XML宣言
    pub decl: HashMap<String, String>,
    /// タグリスト
    pub elements: Vec<XmlElement>,
}

#[pyclass]
#[derive(Debug, Clone)]
/// Xmlタグ
pub struct XmlElement {
    /// タグ名
    pub name: String,

    /// 属性
    pub attributes: HashMap<String, String>,

    /// 子要素
    pub children: Vec<XmlElement>,

    /// テキスト内容
    pub text: Option<String>,
}

impl Xml {
    pub fn new(contents: &String) -> Self {
        let mut reader = Reader::from_str(contents);
        let mut buf: Vec<u8> = Vec::new();
        let mut elements: Vec<XmlElement> = Vec::new();
        let mut decl: HashMap<String, String> = HashMap::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let root: XmlElement = Xml::parse_element(&mut reader, e);
                    elements.push(root);
                    break;
                }
                Ok(Event::Decl(ref e)) => {
                    decl = Xml::parse_decl_element(e);
                }
                Ok(Event::Eof) => break,
                Err(e) => panic!("エラー: {:?}", e),
                _ => {}
            }
            buf.clear();
        }
        Self { decl, elements }
    }

    pub fn save_file(&self, path: &str) {
        let file = OpenOptions::new()
            .write(true)
            .create(true)
            .truncate(true)
            .open(path)
            .unwrap();
        let writer = BufWriter::new(file);
        let mut xml_writer = Writer::new(writer);
        Xml::write_decl(&mut xml_writer, &self.decl);
        for element in &self.elements {
            Xml::write_element(&mut xml_writer, element);
        }
    }

    pub fn to_buf(&self) -> Vec<u8> {
        let mut buffer = Vec::new();
        let mut xml_writer = Writer::new(&mut buffer);

        // XML宣言を書き込み
        Xml::write_decl(&mut xml_writer, &self.decl);

        // 各要素を書き込み
        for element in &self.elements {
            Xml::write_element(&mut xml_writer, element);
        }

        buffer
    }

    /// Xmlタグの解析(通常タグ)
    fn parse_element<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &quick_xml::events::BytesStart,
    ) -> XmlElement {
        // タグ名およびタグ要素を取得
        let name: String = Xml::get_name(start_tag);
        let attributes: HashMap<String, String> = Xml::get_attributes(start_tag);

        let mut children: Vec<XmlElement> = Vec::new();
        let mut text: Option<String> = None;
        let mut buf: Vec<u8> = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                // 開始タグ: <name>
                Ok(Event::Start(e)) => {
                    let child: XmlElement = Xml::parse_element(reader, &e);
                    children.push(child);
                }

                // タグではない要素: <v>text</v>の`text`が該当
                Ok(Event::Text(e)) => {
                    let content: String = e.decode().unwrap().to_string();
                    if !content.trim().is_empty() {
                        text = Some(content);
                    }
                }

                // 終了タグ: </name>
                Ok(Event::End(e)) => {
                    if e.name() == start_tag.name() {
                        break;
                    }
                }

                // 空要素(<name />のようなタグ)
                Ok(Event::Empty(e)) => {
                    let child: XmlElement = Xml::parse_empty_element(&e);
                    children.push(child);
                }

                Ok(Event::Eof) => break,
                _ => {}
            }

            buf.clear();
        }

        XmlElement {
            name,
            attributes,
            children,
            text,
        }
    }

    /// Xmlタグの解析(空タグ)
    fn parse_empty_element(start_tag: &quick_xml::events::BytesStart) -> XmlElement {
        // タグ名およびタグ要素を取得
        let name: String = Xml::get_name(start_tag);
        let attributes: HashMap<String, String> = Xml::get_attributes(start_tag);

        XmlElement {
            name,
            attributes,
            children: Vec::new(),
            text: None,
        }
    }

    fn parse_decl_element(decl: &quick_xml::events::BytesDecl) -> HashMap<String, String> {
        let mut map = HashMap::new();

        if let Ok(version) = decl.version() {
            map.insert(
                "version".to_string(),
                String::from_utf8_lossy(&version).into_owned(),
            );
        }

        if let Some(encoding) = decl.encoding() {
            if let Ok(encoding) = encoding {
                map.insert(
                    "encoding".to_string(),
                    String::from_utf8_lossy(encoding.as_ref()).to_string(),
                );
            }
        }

        if let Some(standalone) = decl.standalone() {
            if let Ok(standalone) = standalone {
                map.insert(
                    "standalone".to_string(),
                    String::from_utf8_lossy(standalone.as_ref()).to_string(),
                );
            }
        }

        map
    }

    /// Xmlタグの書き込み
    fn write_element<W: Write>(writer: &mut Writer<W>, element: &XmlElement) {
        let mut start = BytesStart::new(element.name.as_str());

        for (k, v) in &element.attributes {
            start.push_attribute((k.as_str(), v.as_str()));
        }

        if element.children.is_empty() && element.text.is_none() {
            writer.write_event(Event::Empty(start)).unwrap();
            return;
        }

        writer.write_event(Event::Start(start)).unwrap();

        if let Some(ref text) = element.text {
            writer
                .write_event(Event::Text(BytesText::new(text)))
                .unwrap();
        }

        for child in &element.children {
            Xml::write_element(writer, child);
        }

        writer
            .write_event(Event::End(BytesEnd::new(element.name.as_str())))
            .unwrap();
    }

    /// declの書き込み
    fn write_decl<W: Write>(writer: &mut Writer<W>, decl_hash_map: &HashMap<String, String>) {
        // declの作成
        let version = decl_hash_map.get("version").map(|e| e.as_str());
        let encoding = decl_hash_map.get("encoding").map(|e| e.as_str());
        let standalone = decl_hash_map.get("standalone").map(|s| s.as_str());
        let decl = BytesDecl::new(version.unwrap_or(""), encoding, standalone);

        // ファイル書き込み
        writer.write_event(Event::Decl(decl)).unwrap();
    }

    /// タグ名を取得
    fn get_name(start_tag: &quick_xml::events::BytesStart) -> String {
        String::from_utf8_lossy(start_tag.name().as_ref()).to_string()
    }

    /// タグ属性を取得
    fn get_attributes(start_tag: &quick_xml::events::BytesStart) -> HashMap<String, String> {
        let mut attributes: HashMap<String, String> = HashMap::new();
        for attr in start_tag.attributes().flatten() {
            let key: String = std::str::from_utf8(attr.key.as_ref()).unwrap().to_string();
            let value: String = attr.unescape_value().unwrap().to_string();
            attributes.insert(key, value);
        }
        attributes
    }
}
