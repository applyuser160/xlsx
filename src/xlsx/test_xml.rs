#[cfg(test)]
mod tests {
    use std::fs;
    use std::path::Path;

    use crate::xlsx::xml::Xml;

    #[test]
    fn test_xml_read() {
        // 観点: xmlファイルが読み取れること

        // Act
        let xml: Xml = Xml::new("data/sheet1.xml");

        // Assert

        // path
        assert_eq!(xml.path, "data/sheet1.xml");

        // タグ
        assert_eq!(xml.elements.len(), 1);

        // decl
        assert_eq!(xml.decl.get("version").unwrap().as_str(), "1.0");
    }

    #[test]
    fn test_xml_write() {
        // 観点: xmlファイルが作成されること

        // Arrange

        // ファイルが存在しないことを確認
        if Path::new("data/sheet2.xml").exists() {
            let _ = fs::remove_file("data/sheet2.xml");
        }
        assert!(!Path::new("data/sheet2.xml").exists());

        // Act
        let xml: Xml = Xml::new("data/sheet1.xml");
        xml.save(Some("data/sheet2.xml"));

        // Assert

        // ファイルが作成されること
        assert!(Path::new("data/sheet2.xml").exists());
    }
}
