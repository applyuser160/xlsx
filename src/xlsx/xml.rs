use pyo3::prelude::*;
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};
use std::collections::HashMap;
use std::fs::OpenOptions;
use std::io::{BufRead, BufWriter, Write};

/// A struct that represents an XML file.
///
/// This struct holds the XML declaration and a list of root elements.
#[pyclass]
#[derive(Debug, Clone)]
pub struct Xml {
    /// The XML declaration, e.g., `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`.
    ///
    /// The key-value pairs represent the attributes of the declaration.
    pub decl: HashMap<String, String>,
    /// A list of root elements in the XML file.
    pub elements: Vec<XmlElement>,
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
    /// Gets a mutable reference to a child element by its tag name.
    ///
    /// If the child element does not exist, it will be created.
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

/// A struct that represents an XML element.
///
/// This struct holds the tag name, attributes, child elements, and text content of an XML element.
#[pyclass]
#[derive(Debug, Clone)]
pub struct XmlElement {
    /// The tag name of the element.
    pub name: String,

    /// The attributes of the element.
    pub attributes: HashMap<String, String>,

    /// The child elements of the element.
    pub children: Vec<XmlElement>,

    /// The text content of the element.
    pub text: Option<String>,
}

impl Xml {
    /// Creates a new `Xml` instance from a string.
    ///
    /// This function parses the XML string and builds the `Xml` struct.
    pub fn new(contents: &str) -> Self {
        let mut reader = Reader::from_str(contents);
        let mut buf: Vec<u8> = Vec::new();
        let mut elements: Vec<XmlElement> = Vec::new();
        let mut decl: HashMap<String, String> = HashMap::new();

        // Parse the XML string
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    // Parse the root element
                    let root: XmlElement = Xml::parse_element(&mut reader, e);
                    elements.push(root);
                    break;
                }
                Ok(Event::Decl(ref e)) => {
                    // Parse the XML declaration
                    decl = Xml::parse_decl_element(e);
                }
                Ok(Event::Eof) => break,
                Err(e) => panic!("Error at position {}: {:?}", reader.buffer_position(), e),
                _ => {}
            }
            buf.clear();
        }
        Self { decl, elements }
    }

    /// Saves the `Xml` struct to a file.
    pub fn save_file(&self, path: &str) {
        let file = OpenOptions::new()
            .write(true)
            .create(true)
            .truncate(true)
            .open(path)
            .unwrap();
        let writer = BufWriter::new(file);
        let mut xml_writer = Writer::new(writer);
        // Write the XML declaration
        Xml::write_decl(&mut xml_writer, &self.decl);
        // Write the elements
        for element in &self.elements {
            Xml::write_element(&mut xml_writer, element);
        }
    }

    /// Converts the `Xml` struct to a byte vector.
    pub fn to_buf(&self) -> Vec<u8> {
        let mut buffer = Vec::new();
        let mut xml_writer = Writer::new(&mut buffer);

        // Write the XML declaration
        Xml::write_decl(&mut xml_writer, &self.decl);

        // Write the elements
        for element in &self.elements {
            Xml::write_element(&mut xml_writer, element);
        }

        buffer
    }

    /// Parses a regular XML element.
    fn parse_element<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &quick_xml::events::BytesStart,
    ) -> XmlElement {
        // Get the tag name and attributes
        let name: String = Xml::get_name(start_tag);
        let attributes: HashMap<String, String> = Xml::get_attributes(start_tag);

        let mut children: Vec<XmlElement> = Vec::new();
        let mut text: Option<String> = None;
        let mut buf: Vec<u8> = Vec::new();

        // Parse the children and text
        loop {
            match reader.read_event_into(&mut buf) {
                // Start tag: <name>
                Ok(Event::Start(e)) => {
                    let child: XmlElement = Xml::parse_element(reader, &e);
                    children.push(child);
                }

                // Text content: <v>text</v>
                Ok(Event::Text(e)) => {
                    let content: String = e.decode().unwrap().to_string();
                    if !content.trim().is_empty() {
                        text = Some(content);
                    }
                }

                // End tag: </name>
                Ok(Event::End(e)) => {
                    if e.name() == start_tag.name() {
                        break;
                    }
                }

                // Empty element: <name />
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

    /// Parses an empty XML element.
    fn parse_empty_element(start_tag: &quick_xml::events::BytesStart) -> XmlElement {
        // Get the tag name and attributes
        let name: String = Xml::get_name(start_tag);
        let attributes: HashMap<String, String> = Xml::get_attributes(start_tag);

        XmlElement {
            name,
            attributes,
            children: Vec::new(),
            text: None,
        }
    }

    /// Parses the XML declaration element.
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

    /// Writes an XML element to the writer.
    fn write_element<W: Write>(writer: &mut Writer<W>, element: &XmlElement) {
        let mut start = BytesStart::new(element.name.as_str());

        // Add attributes
        for (k, v) in &element.attributes {
            start.push_attribute((k.as_str(), v.as_str()));
        }

        // Write empty tag if there are no children and no text
        if element.children.is_empty() && element.text.is_none() {
            writer.write_event(Event::Empty(start)).unwrap();
            return;
        }

        // Write start tag
        writer.write_event(Event::Start(start)).unwrap();

        // Write text content
        if let Some(ref text) = element.text {
            writer
                .write_event(Event::Text(BytesText::new(text)))
                .unwrap();
        }

        // Write child elements
        for child in &element.children {
            Xml::write_element(writer, child);
        }

        // Write end tag
        writer
            .write_event(Event::End(BytesEnd::new(element.name.as_str())))
            .unwrap();
    }

    /// Writes the XML declaration to the writer.
    fn write_decl<W: Write>(writer: &mut Writer<W>, decl_hash_map: &HashMap<String, String>) {
        // Create the declaration
        let version = decl_hash_map.get("version").map(|e| e.as_str());
        let encoding = decl_hash_map.get("encoding").map(|e| e.as_str());
        let standalone = decl_hash_map.get("standalone").map(|s| s.as_str());
        let decl = BytesDecl::new(version.unwrap_or(""), encoding, standalone);

        // Write the declaration to the file
        writer.write_event(Event::Decl(decl)).unwrap();
    }

    /// Gets the tag name from a `BytesStart` event.
    fn get_name(start_tag: &quick_xml::events::BytesStart) -> String {
        String::from_utf8_lossy(start_tag.name().as_ref()).to_string()
    }

    /// Gets the attributes from a `BytesStart` event.
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
