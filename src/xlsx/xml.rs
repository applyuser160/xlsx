use pyo3::prelude::*;
use quick_xml::encoding::EncodingError;
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};
use std::collections::HashMap;
use std::fs::OpenOptions;
use std::io::{BufRead, BufWriter, Write};
use thiserror::Error;

/// XML操作中に発生する可能性のあるエラー
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

/// XMLファイルを表す構造体
///
/// この構造体はXML宣言とルート要素のリストを保持
#[pyclass]
#[derive(Debug, Clone)]
pub struct Xml {
    /// XML宣言
    ///
    /// キーと値のペアは宣言の属性を表す
    pub decl: HashMap<String, String>,
    /// XMLファイル内のルート要素のリスト
    pub elements: Vec<XmlElement>,
}
#[pymethods]
impl XmlElement {
    /// 与えられたタグ名で新しい `XmlElement` を作成
    #[new]
    pub fn new(name: &str) -> Self {
        XmlElement {
            name: name.to_string(),
            attributes: HashMap::new(),
            children: Vec::new(),
            text: None,
        }
    }
}

impl Xml {
    /// タグ名で子要素への可変参照を取得
    ///
    /// 子要素が存在しない場合は作成
    pub fn get_mut_or_create_child_by_tag(&mut self, tag_name: &str) -> &mut XmlElement {
        let style_sheet = self.elements.first_mut().unwrap();
        let position = style_sheet.children.iter().position(|c| c.name == tag_name);
        match position {
            Some(pos) => &mut style_sheet.children[pos],
            None => {
                let new_element = XmlElement::new(tag_name);
                style_sheet.children.push(new_element);
                style_sheet.children.last_mut().unwrap()
            }
        }
    }
}

/// XML要素を表す構造体
///
/// この構造体はタグ名、属性、子要素、テキストコンテンツを保持
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
    /// 文字列から新しい `Xml` インスタンスを作成
    ///
    /// この関数はXML文字列を解析し、 `Xml` 構造体を構築
    pub fn new(contents: &str) -> Result<Self> {
        let mut reader = Reader::from_str(contents);
        let mut buf = Vec::new();
        let mut elements = Vec::new();
        let mut decl = HashMap::new();

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(ref e) => {
                    let root = Self::parse_element(&mut reader, e)?;
                    elements.push(root);
                    break;
                }
                Event::Decl(ref e) => {
                    decl = Self::parse_decl_element(e);
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }
        Ok(Self { decl, elements })
    }

    /// `Xml` 構造体をファイルに保存
    pub fn save_file(&self, path: &str) -> Result<()> {
        let file = OpenOptions::new()
            .write(true)
            .create(true)
            .truncate(true)
            .open(path)?;
        let writer = BufWriter::new(file);
        let mut xml_writer = Writer::new(writer);
        Self::write_decl(&mut xml_writer, &self.decl)?;
        for element in &self.elements {
            Self::write_element(&mut xml_writer, element)?;
        }
        Ok(())
    }

    /// `Xml` 構造体をバイトベクターに変換
    pub fn to_buf(&self) -> Result<Vec<u8>> {
        let mut buffer = Vec::new();
        let mut xml_writer = Writer::new(&mut buffer);
        Self::write_decl(&mut xml_writer, &self.decl)?;
        for element in &self.elements {
            Self::write_element(&mut xml_writer, element)?;
        }
        Ok(buffer)
    }

    /// 通常のXML要素を解析
    fn parse_element<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &quick_xml::events::BytesStart,
    ) -> Result<XmlElement> {
        let name = Self::get_name(start_tag)?;
        let attributes = Self::get_attributes(start_tag)?;
        let mut children = Vec::new();
        let mut text = None;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) => children.push(Self::parse_element(reader, &e)?),
                Event::Text(e) => {
                    let content = e.decode()?.to_string();
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

    /// 空のXML要素を解析
    fn parse_empty_element(start_tag: &quick_xml::events::BytesStart) -> Result<XmlElement> {
        Ok(XmlElement {
            name: Self::get_name(start_tag)?,
            attributes: Self::get_attributes(start_tag)?,
            ..Default::default()
        })
    }

    /// XML宣言要素を解析
    fn parse_decl_element(decl: &quick_xml::events::BytesDecl) -> HashMap<String, String> {
        let mut map = HashMap::new();
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

    /// XML要素をライターに書き込み
    fn write_element<W: Write>(writer: &mut Writer<W>, element: &XmlElement) -> Result<()> {
        let mut start = BytesStart::new(&element.name);
        for (k, v) in &element.attributes {
            start.push_attribute((k.as_str(), v.as_str()));
        }

        if element.children.is_empty() && element.text.is_none() {
            writer.write_event(Event::Empty(start))?;
            return Ok(());
        }

        writer.write_event(Event::Start(start))?;
        if let Some(ref text) = element.text {
            writer.write_event(Event::Text(BytesText::new(text)))?;
        }

        for child in &element.children {
            Self::write_element(writer, child)?;
        }

        writer.write_event(Event::End(BytesEnd::new(&element.name)))?;
        Ok(())
    }

    /// XML宣言をライターに書き込み
    fn write_decl<W: Write>(
        writer: &mut Writer<W>,
        decl_hash_map: &HashMap<String, String>,
    ) -> Result<()> {
        let version = decl_hash_map.get("version").map(|e| e.as_str());
        let encoding = decl_hash_map.get("encoding").map(|e| e.as_str());
        let standalone = decl_hash_map.get("standalone").map(|s| s.as_str());
        let decl = BytesDecl::new(version.unwrap_or(""), encoding, standalone);
        writer.write_event(Event::Decl(decl))?;
        Ok(())
    }

    /// `BytesStart` イベントからタグ名を取得
    fn get_name(start_tag: &quick_xml::events::BytesStart) -> Result<String> {
        Ok(String::from_utf8(start_tag.name().as_ref().to_vec())?)
    }

    /// `BytesStart` イベントから属性を取得
    fn get_attributes(
        start_tag: &quick_xml::events::BytesStart,
    ) -> Result<HashMap<String, String>> {
        start_tag
            .attributes()
            .map(|attr| {
                let attr = attr?;
                let key = String::from_utf8(attr.key.as_ref().to_vec())?;
                let value = attr.unescape_value()?.to_string();
                Ok((key, value))
            })
            .collect()
    }
}
