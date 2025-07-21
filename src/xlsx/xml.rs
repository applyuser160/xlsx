use pyo3::prelude::*;
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};
use std::collections::HashMap;
use std::fs::OpenOptions;
use std::io::{BufRead, BufWriter, Write};

/// Represents an XML document.
///
/// This struct holds the XML declaration and a list of root elements.
#[pyclass]
#[derive(Debug, Clone)]
pub struct Xml {
    /// The XML declaration attributes, such as version and encoding.
    pub decl: HashMap<String, String>,
    /// The root elements of the XML document.
    pub elements: Vec<XmlElement>,
}

/// Represents an element in an XML document.
///
/// This struct holds the tag name, attributes, child elements, and optional text content.
#[pyclass]
#[derive(Debug, Clone)]
pub struct XmlElement {
    /// The name of the XML tag.
    pub name: String,
    /// The attributes of the XML tag.
    pub attributes: HashMap<String, String>,
    /// The child elements of the XML tag.
    pub children: Vec<XmlElement>,
    /// The text content of the XML tag, if any.
    pub text: Option<String>,
}

#[pymethods]
impl XmlElement {
    /// Creates a new `XmlElement` with the given tag name.
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
    /// Creates a new `Xml` instance by parsing the given XML content string.
    pub fn new(contents: &String) -> Self {
        let mut reader = Reader::from_str(contents);
        let mut buf: Vec<u8> = Vec::new();
        let mut elements: Vec<XmlElement> = Vec::new();
        let mut decl: HashMap<String, String> = HashMap::new();

        // XML declaration parsing function
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

        // Main parsing loop
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let root = XmlElement::parse_element(&mut reader, e);
                    elements.push(root);
                    break;
                }
                Ok(Event::Decl(ref e)) => {
                    decl = parse_decl_element(e);
                }
                Ok(Event::Eof) => break,
                Err(e) => panic!("Error during XML parsing: {e:?}"),
                _ => {}
            }
            buf.clear();
        }
        Self { decl, elements }
    }

    /// Saves the XML document to the specified file path.
    pub fn save_file(&self, path: &str) {
        let file = OpenOptions::new()
            .write(true)
            .create(true)
            .truncate(true)
            .open(path)
            .unwrap();
        let writer = BufWriter::new(file);
        let mut xml_writer = Writer::new(writer);

        // Write declaration
        Self::write_decl(&mut xml_writer, &self.decl);

        // Write elements
        for element in &self.elements {
            XmlElement::write_element(&mut xml_writer, element);
        }
    }

    /// Converts the XML document to a byte buffer.
    pub fn to_buf(&self) -> Vec<u8> {
        let mut buffer = Vec::new();
        let mut xml_writer = Writer::new(&mut buffer);

        // Write declaration
        Self::write_decl(&mut xml_writer, &self.decl);

        // Write elements
        for element in &self.elements {
            XmlElement::write_element(&mut xml_writer, element);
        }
        buffer
    }

    /// Finds a mutable child element by tag name, or creates it if it doesn't exist.
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

    /// Writes the XML declaration to the writer.
    fn write_decl<W: Write>(writer: &mut Writer<W>, decl_hash_map: &HashMap<String, String>) {
        let version = decl_hash_map.get("version").map(|e| e.as_str());
        let encoding = decl_hash_map.get("encoding").map(|e| e.as_str());
        let standalone = decl_hash_map.get("standalone").map(|s| s.as_str());
        let decl = BytesDecl::new(version.unwrap_or("1.0"), encoding, standalone);
        writer.write_event(Event::Decl(decl)).unwrap();
    }
}

impl XmlElement {
    /// Parses an XML element from the reader.
    fn parse_element<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &quick_xml::events::BytesStart,
    ) -> XmlElement {
        let name = Self::get_name(start_tag);
        let attributes = Self::get_attributes(start_tag);
        let mut children = Vec::new();
        let mut text = None;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    children.push(Self::parse_element(reader, &e));
                }
                Ok(Event::Text(e)) => {
                    let content = e.decode().unwrap().to_string();
                    if !content.trim().is_empty() {
                        text = Some(content);
                    }
                }
                Ok(Event::End(e)) => {
                    if e.name() == start_tag.name() {
                        break;
                    }
                }
                Ok(Event::Empty(e)) => {
                    children.push(Self::parse_empty_element(&e));
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

    /// Parses an empty XML element.
    fn parse_empty_element(start_tag: &quick_xml::events::BytesStart) -> XmlElement {
        let name = Self::get_name(start_tag);
        let attributes = Self::get_attributes(start_tag);
        XmlElement {
            name,
            attributes,
            children: Vec::new(),
            text: None,
        }
    }

    /// Writes the XML element to the writer.
    fn write_element<W: Write>(writer: &mut Writer<W>, element: &XmlElement) {
        let mut start = BytesStart::new(&element.name);
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
            Self::write_element(writer, child);
        }

        writer
            .write_event(Event::End(BytesEnd::new(&element.name)))
            .unwrap();
    }

    /// Extracts the tag name from a `BytesStart` event.
    fn get_name(start_tag: &quick_xml::events::BytesStart) -> String {
        String::from_utf8_lossy(start_tag.name().as_ref()).to_string()
    }

    /// Extracts the attributes from a `BytesStart` event.
    fn get_attributes(start_tag: &quick_xml::events::BytesStart) -> HashMap<String, String> {
        start_tag
            .attributes()
            .flatten()
            .map(|attr| {
                let key = std::str::from_utf8(attr.key.as_ref()).unwrap().to_string();
                let value = attr.unescape_value().unwrap().to_string();
                (key, value)
            })
            .collect()
    }
}
