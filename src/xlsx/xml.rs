use pyo3::prelude::*;
use quick_xml::encoding::EncodingError;
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};
use std::collections::HashMap;
use std::fs::{File, OpenOptions};
use std::io::{BufRead, BufWriter, Write};
use thiserror::Error;

/// XML操作中のエラー
#[derive(Error, Debug)]
pub enum XmlError {
    #[error("XML parsing error: {0}")]
    Parse(#[from] quick_xml::Error),
    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),
    #[error("UTF-8 conversion error: {0}")]
    Utf8(#[from] std::string::FromUtf8Error),
    #[error("Attribute conversion error: {0}")]
    Attr(#[from] quick_xml::events::attributes::AttrError),
    #[error("Encoding error: {0}")]
    Encoding(#[from] EncodingError),
}

type Result<T> = std::result::Result<T, XmlError>;

/// XMLファイル
#[pyclass]
#[derive(Debug, Clone)]
pub struct Xml {
    /// XML宣言
    pub decl: HashMap<String, String>,
    /// XMLファイル内のルート要素のリスト
    pub elements: Vec<XmlElement>,
}

#[pymethods]
impl XmlElement {
    /// 与えられたタグ名での新しい `XmlElement` の作成
    #[new]
    pub fn new(name: &str) -> Self {
        XmlElement {
            name: name.to_string(),
            ..Default::default()
        }
    }
}

impl XmlElement {
    pub fn get_element(&self, path: &str) -> &XmlElement {
        let mut current_element = self;
        for tag in path.split('>') {
            current_element = current_element
                .children
                .iter()
                .find(|c| c.name == tag)
                .unwrap();
        }
        current_element
    }

    pub fn get_elements(&self, path: &str) -> Vec<&XmlElement> {
        let mut current_elements = vec![self];
        for tag in path.split('>') {
            current_elements = current_elements
                .iter()
                .flat_map(|e| e.children.iter().filter(|c| c.name == tag))
                .collect();
        }
        current_elements
    }

    pub fn get_element_mut(&mut self, path: &str) -> &mut XmlElement {
        let mut current_element = self;
        for tag in path.split('>') {
            current_element = current_element
                .children
                .iter_mut()
                .find(|c| c.name == tag)
                .unwrap();
        }
        current_element
    }

    pub fn get_attribute(&self, key: &str) -> Option<&String> {
        self.attributes.get(key)
    }

    pub fn get_text(&self) -> &str {
        self.text.as_deref().unwrap_or_default()
    }

    pub fn push_str(&mut self, content: &str) {
        // This is a bit of a hack to append raw XML content.
        // It assumes the content is well-formed.
        // A proper implementation would parse the string into XmlElement objects.
        if self.text.is_none() {
            self.text = Some(String::new());
        }
        self.text.as_mut().unwrap().push_str(content);
    }
}

impl Xml {
    /// タグ名による子要素への可変参照の取得
    ///
    /// 子要素が存在しない場合は作成
    pub fn get_mut_or_create_child_by_tag(&mut self, tag_name: &str) -> &mut XmlElement {
        let style_sheet: &mut XmlElement = match self.elements.first_mut() {
            Some(element) => element,
            None => panic!("No elements in XML"),
        };
        // 子要素の検索
        if let Some(pos) = style_sheet.children.iter().position(|c| c.name == tag_name) {
            &mut style_sheet.children[pos]
        } else {
            // 新規作成して返却
            let new_element: XmlElement = XmlElement::new(tag_name);
            style_sheet.children.push(new_element);
            match style_sheet.children.last_mut() {
                Some(element) => element,
                None => panic!("Failed to add new element"),
            }
        }
    }
}

/// XML要素
#[pyclass]
#[derive(Debug, Clone, Default, PartialEq)]
pub struct XmlElement {
    /// 要素のタグ名
    pub name: String,
    /// 要素の属性
    pub attributes: HashMap<String, String>,
    /// 要素の子要素
    pub children: Vec<XmlElement>,
    /// 要素のテキストコンテンツ
    pub text: Option<String>,
}

impl Xml {
    /// 文字列からの新しい `Xml` インスタンスの作成
    pub fn new(contents: &str) -> Result<Self> {
        let mut reader: Reader<&[u8]> = Reader::from_str(contents);
        let mut buf: Vec<u8> = Vec::new();
        let mut elements: Vec<XmlElement> = Vec::new();
        let mut decl: HashMap<String, String> = HashMap::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let root: XmlElement = Self::parse_element(&mut reader, e)?;
                    elements.push(root);
                    break;
                }
                Ok(Event::Decl(ref e)) => {
                    decl = Self::parse_decl_element(e);
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }
        Ok(Self { decl, elements })
    }

    /// `Xml` 構造体のファイルへの保存
    pub fn save_file(&self, path: &str) -> Result<()> {
        let file: File = OpenOptions::new()
            .write(true)
            .create(true)
            .truncate(true)
            .open(path)?;
        let writer: BufWriter<File> = BufWriter::new(file);
        let mut xml_writer: Writer<BufWriter<File>> = Writer::new(writer);
        Self::write_decl(&mut xml_writer, &self.decl)?;
        for element in &self.elements {
            Self::write_element(&mut xml_writer, element)?;
        }
        Ok(())
    }

    /// `Xml` 構造体のバイトベクターへの変換
    pub fn to_buf(&self) -> Result<Vec<u8>> {
        let mut buffer: Vec<u8> = Vec::new();
        let mut xml_writer: Writer<&mut Vec<u8>> = Writer::new(&mut buffer);
        Self::write_decl(&mut xml_writer, &self.decl)?;
        for element in &self.elements {
            Self::write_element(&mut xml_writer, element)?;
        }
        Ok(buffer)
    }

    /// 通常のXML要素の解析
    fn parse_element<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
    ) -> Result<XmlElement> {
        let name: String = Self::get_name(start_tag)?;
        let attributes: HashMap<String, String> = Self::get_attributes(start_tag)?;
        let mut children: Vec<XmlElement> = Vec::new();
        let mut text: Option<String> = None;
        let mut buf: Vec<u8> = Vec::new();

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) => children.push(Self::parse_element(reader, &e)?),
                Event::Text(e) => {
                    let content: String = e.decode()?.to_string();
                    if !content.trim().is_empty() {
                        text = Some(content);
                    }
                }
                Event::End(e) if e.name() == start_tag.name() => break,
                Event::Empty(e) => children.push(Self::parse_empty_element(&e)?),
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        Ok(XmlElement {
            name,
            attributes,
            children,
            text,
        })
    }

    /// 空のXML要素の解析
    fn parse_empty_element(start_tag: &BytesStart) -> Result<XmlElement> {
        Ok(XmlElement {
            name: Self::get_name(start_tag)?,
            attributes: Self::get_attributes(start_tag)?,
            ..Default::default()
        })
    }

    /// XML宣言要素の解析
    fn parse_decl_element(decl: &BytesDecl) -> HashMap<String, String> {
        let mut map: HashMap<String, String> = HashMap::with_capacity(3);
        if let Ok(version) = decl.version() {
            map.insert(
                "version".to_string(),
                String::from_utf8_lossy(&version).into_owned(),
            );
        }
        if let Some(Ok(encoding)) = decl.encoding() {
            map.insert(
                "encoding".to_string(),
                String::from_utf8_lossy(encoding.as_ref()).to_string(),
            );
        }
        if let Some(Ok(standalone)) = decl.standalone() {
            map.insert(
                "standalone".to_string(),
                String::from_utf8_lossy(standalone.as_ref()).to_string(),
            );
        }
        map
    }

    /// XML要素のライターへの書き込み
    fn write_element<W: Write>(writer: &mut Writer<W>, element: &XmlElement) -> Result<()> {
        let mut start: BytesStart = BytesStart::new(&element.name);
        for (k, v) in &element.attributes {
            start.push_attribute((k.as_str(), v.as_str()));
        }

        if element.children.is_empty() && element.text.is_none() {
            writer.write_event(Event::Empty(start))?;
        } else {
            writer.write_event(Event::Start(start))?;
            if let Some(ref text) = element.text {
                writer.write_event(Event::Text(BytesText::new(text)))?;
            }
            for child in &element.children {
                Self::write_element(writer, child)?;
            }
            writer.write_event(Event::End(BytesEnd::new(&element.name)))?;
        }
        Ok(())
    }

    /// XML宣言のライターへの書き込み
    fn write_decl<W: Write>(
        writer: &mut Writer<W>,
        decl_hash_map: &HashMap<String, String>,
    ) -> Result<()> {
        let version: Option<&str> = decl_hash_map.get("version").map(|e| e.as_str());
        let encoding: Option<&str> = decl_hash_map.get("encoding").map(|e| e.as_str());
        let standalone: Option<&str> = decl_hash_map.get("standalone").map(|s| s.as_str());
        let decl: BytesDecl = BytesDecl::new(version.unwrap_or("1.0"), encoding, standalone);
        writer.write_event(Event::Decl(decl))?;
        Ok(())
    }

    /// `BytesStart` イベントからのタグ名の取得
    fn get_name(start_tag: &BytesStart) -> Result<String> {
        Ok(String::from_utf8(start_tag.name().as_ref().to_vec())?)
    }

    /// `BytesStart` イベントからの属性の取得
    fn get_attributes(start_tag: &BytesStart) -> Result<HashMap<String, String>> {
        start_tag
            .attributes()
            .map(|attr_result| {
                let attr: quick_xml::events::attributes::Attribute<'_> = attr_result?;
                let key: String = String::from_utf8(attr.key.as_ref().to_vec())?;
                let value: String = attr.unescape_value()?.to_string();
                Ok((key, value))
            })
            .collect()
    }
}
