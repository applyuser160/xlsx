use pyo3::prelude::*;
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};
use std::collections::HashMap;
use std::fs::OpenOptions;
use std::io::{BufRead, BufWriter, Write};

/// XMLドキュメントの宣言とルート要素の表現
///
/// XML宣言（例： `<?xml version="1.0" ...>`）と
/// ルート要素のリストを含む、解析済みXMLデータの保持
#[pyclass]
#[derive(Debug, Clone)]
pub struct Xml {
    /// XML宣言の属性（"version"や"encoding"など）を格納するマップ
    pub decl: HashMap<String, String>,
    /// XMLドキュメントのルート要素を表す`XmlElement`構造体のベクター
    pub elements: Vec<XmlElement>,
}

/// XMLドキュメント内の単一要素の表現
///
/// 名前、属性、子要素、およびオプションのテキストコンテンツの保持
/// XMLドキュメントのツリー表現の構築に使用
#[pyclass]
#[derive(Debug, Clone)]
pub struct XmlElement {
    /// XMLタグ名
    pub name: String,
    /// 要素の属性マップ（キー：属性名、値：属性値）
    pub attributes: HashMap<String, String>,
    /// ネストされた要素を表す子`XmlElement`のベクター
    pub children: Vec<XmlElement>,
    /// オプションのテキストコンテンツ
    pub text: Option<String>,
}

#[pymethods]
impl XmlElement {
    /// 指定された名前での新しい`XmlElement`の作成
    ///
    /// # 引数
    ///
    /// * `name` - XML要素の名前
    ///
    /// # 戻り値
    ///
    /// 空の属性、子要素、テキストコンテンツなしの新しい`XmlElement`インスタンス
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
    /// タグ名による子要素への可変参照の検索（存在しない場合は作成）
    ///
    /// `Xml`構造体に少なくとも1つのルート要素が存在することの前提
    ///
    /// # 引数
    ///
    /// * `tag_name` - 検索または作成する子要素の名前
    ///
    /// # 戻り値
    ///
    /// 見つかった、または新しく作成された`XmlElement`への可変参照
    ///
    /// # パニック
    ///
    /// `Xml`構造体にルート要素がない場合にパニック
    pub fn get_mut_or_create_child_by_tag(&mut self, tag_name: &str) -> &mut XmlElement {
        let style_sheet = self
            .elements
            .first_mut()
            .expect("少なくとも1つのルート要素が必要");
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

impl Xml {
    /// XMLコンテンツの文字列解析による新しい`Xml`インスタンスの作成
    ///
    /// # 引数
    ///
    /// * `contents` - 解析対象のXMLコンテンツを含む文字列スライス
    ///
    /// # 戻り値
    ///
    /// 新しい`Xml`インスタンス
    pub fn new(contents: &str) -> Self {
        let mut reader = Reader::from_str(contents);
        let mut buf = Vec::new();
        let mut elements = Vec::new();
        let mut decl = HashMap::new();

        // XMLイベントの処理
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let root = Self::parse_element(&mut reader, e);
                    elements.push(root);
                    break; // ルート要素は1つと仮定
                }
                Ok(Event::Decl(ref e)) => {
                    decl = Self::parse_decl_element(e);
                }
                Ok(Event::Eof) => break,
                Err(e) => panic!("位置 {} でエラー: {:?}", reader.buffer_position(), e),
                _ => {}
            }
            buf.clear();
        }

        Self { decl, elements }
    }

    /// XMLコンテンツのファイル保存
    ///
    /// 既存ファイルは上書き
    ///
    /// # 引数
    ///
    /// * `path` - 保存先のファイルパス
    pub fn save_file(&self, path: &str) {
        let file = OpenOptions::new()
            .write(true)
            .create(true)
            .truncate(true)
            .open(path)
            .unwrap();
        let writer = BufWriter::new(file);
        let mut xml_writer = Writer::new(writer);
        Self::write_decl(&mut xml_writer, &self.decl);
        for element in &self.elements {
            Self::write_element(&mut xml_writer, element);
        }
    }

    /// `Xml`インスタンスのバイトベクターへのシリアライズ
    ///
    /// # 戻り値
    ///
    /// シリアライズされたXMLコンテンツを含む`Vec<u8>`
    pub fn to_buf(&self) -> Vec<u8> {
        let mut buffer = Vec::new();
        let mut xml_writer = Writer::new(&mut buffer);

        Self::write_decl(&mut xml_writer, &self.decl);

        for element in &self.elements {
            Self::write_element(&mut xml_writer, element);
        }

        buffer
    }

    /// 標準的なXML要素（子要素・テキストコンテンツを含む）の解析
    fn parse_element<R: BufRead>(reader: &mut Reader<R>, start_tag: &BytesStart) -> XmlElement {
        let name = Self::get_name(start_tag);
        let attributes = Self::get_attributes(start_tag);
        let mut children = Vec::new();
        let mut text = None;
        let mut buf = Vec::new();

        // 子要素の再帰的解析
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => children.push(Self::parse_element(reader, &e)),
                Ok(Event::Text(e)) => {
                    if let Ok(content) = e.decode() {
                        if !content.trim().is_empty() {
                            text = Some(content.to_string());
                        }
                    }
                }
                Ok(Event::End(e)) if e.name() == start_tag.name() => break,
                Ok(Event::Empty(e)) => children.push(Self::parse_empty_element(&e)),
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

    /// 空のXML要素（例： `<tag/>`）の解析
    fn parse_empty_element(start_tag: &BytesStart) -> XmlElement {
        XmlElement {
            name: Self::get_name(start_tag),
            attributes: Self::get_attributes(start_tag),
            children: Vec::new(),
            text: None,
        }
    }

    /// XML宣言の属性解析
    fn parse_decl_element(decl: &BytesDecl) -> HashMap<String, String> {
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

    /// `XmlElement`のXMLライターへの書き込み
    fn write_element<W: Write>(writer: &mut Writer<W>, element: &XmlElement) {
        let mut start = BytesStart::new(&element.name);
        for (k, v) in &element.attributes {
            start.push_attribute((k.as_str(), v.as_str()));
        }

        if element.children.is_empty() && element.text.is_none() {
            writer.write_event(Event::Empty(start)).unwrap();
        } else {
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
    }

    /// XML宣言のXMLライターへの書き込み
    fn write_decl<W: Write>(writer: &mut Writer<W>, decl_hash_map: &HashMap<String, String>) {
        let version = decl_hash_map.get("version").map_or("", |s| s.as_str());
        let encoding = decl_hash_map.get("encoding").map(|s| s.as_str());
        let standalone = decl_hash_map.get("standalone").map(|s| s.as_str());

        if !version.is_empty() {
            let decl = BytesDecl::new(version, encoding, standalone);
            writer.write_event(Event::Decl(decl)).unwrap();
        }
    }

    /// `BytesStart`イベントからのタグ名取得
    fn get_name(start_tag: &BytesStart) -> String {
        String::from_utf8_lossy(start_tag.name().as_ref()).to_string()
    }

    /// `BytesStart`イベントからの属性取得（`HashMap`へ）
    fn get_attributes(start_tag: &BytesStart) -> HashMap<String, String> {
        start_tag
            .attributes()
            .flatten()
            .filter_map(|attr| {
                let key = std::str::from_utf8(attr.key.as_ref()).ok()?.to_string();
                let value = attr.unescape_value().ok()?.to_string();
                Some((key, value))
            })
            .collect()
    }
}
