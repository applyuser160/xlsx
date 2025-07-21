#[cfg(test)]
mod tests {
    use std::fs::{self, File};
    use std::io::Read;
    use std::path::Path;

    use crate::xml::Xml;

    #[test]
    fn test_xml_read() {
        // 観点: xmlファイルが読み取れること

        // Act
        let mut file: File = File::open("data/sheet1.xml").unwrap();
        let mut buf = String::new();
        file.read_to_string(&mut buf).unwrap();
        let xml: Xml = Xml::new(&buf).unwrap();

        // Assert

        // タグ
        assert_eq!(xml.elements.len(), 1);

        // decl
        assert_eq!(xml.decl.get("version").unwrap().as_str(), "1.0");
    }

    #[test]
    fn test_xml_write_file() {
        // 観点: xmlファイルが作成されること

        // Arrange

        // ファイルが存在しないことを確認
        if Path::new("data/sheet2.xml").exists() {
            let _ = fs::remove_file("data/sheet2.xml");
        }
        assert!(!Path::new("data/sheet2.xml").exists());

        // Act
        let mut file: File = File::open("data/sheet1.xml").unwrap();
        let mut buf = String::new();
        file.read_to_string(&mut buf).unwrap();
        let xml: Xml = Xml::new(&buf).unwrap();
        xml.save_file("data/sheet2.xml").unwrap();

        // Assert

        // ファイルが作成されること
        assert!(Path::new("data/sheet2.xml").exists());
    }

    #[test]
    fn test_xml_to_buf() {
        // 観点: xml文字列が作成されること

        // Act
        let mut file: File = File::open("data/sheet1.xml").unwrap();
        let mut buf = String::new();
        file.read_to_string(&mut buf).unwrap();
        let xml: Xml = Xml::new(&buf).unwrap();
        let buf = xml.to_buf().unwrap();

        // Assert
        assert!(!buf.is_empty());
    }
}
