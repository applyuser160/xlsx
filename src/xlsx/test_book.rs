#[cfg(test)]
mod tests {
    use std::{fs, path::Path};

    use crate::xlsx::book::Book;

    #[test]
    fn test_new_book() {
        // 観点: Excelファイルの読み取り

        // Act
        let book = Book::new("data/sample.xlsx".to_string());

        // Assert
        let xml = book.worksheets.get("xl/worksheets/sheet1.xml").unwrap();
        let xml_guard = xml.lock().unwrap();
        assert_eq!(xml_guard.decl.get("version").unwrap(), "1.0");
        assert_eq!(xml_guard.decl.get("encoding").unwrap(), "UTF-8");
        assert_eq!(xml_guard.decl.get("standalone").unwrap(), "yes");
    }

    #[test]
    fn test_save_book() {
        // 観点: Excelファイルの書き込み

        // Arrange

        // ファイルが存在しないことを確認
        if Path::new("data/sample2.xlsx").exists() {
            let _ = fs::remove_file("data/sample2.xlsx");
        }
        assert!(!Path::new("data/sample2.xlsx").exists());

        // Act
        let book = Book::new("data/sample.xlsx".to_string());
        let xml = book.worksheets.get("xl/worksheets/sheet1.xml").unwrap();
        let mut xml_guard = xml.lock().unwrap();
        let version = xml_guard.decl.get_mut("version").unwrap();
        *version = "1.0".to_string();
        drop(xml_guard); // ロックを解放
        book.save();

        // Assert
        let book = Book::new("data/sample.xlsx".to_string());
        let xml = book.worksheets.get("xl/worksheets/sheet1.xml").unwrap();
        let xml_guard = xml.lock().unwrap();
        assert_eq!(xml_guard.decl.get("version").unwrap(), "1.0");
        assert_eq!(xml_guard.decl.get("encoding").unwrap(), "UTF-8");
        assert_eq!(xml_guard.decl.get("standalone").unwrap(), "yes");
    }

    #[test]
    fn test_copy_book() {
        // 観点: Excelファイルの名前をつけて保存

        // Arrange

        // ファイルが存在しないことを確認
        if Path::new("data/sample2.xlsx").exists() {
            let _ = fs::remove_file("data/sample2.xlsx");
        }
        assert!(!Path::new("data/sample2.xlsx").exists());

        // Act
        let book = Book::new("data/sample.xlsx".to_string());
        let xml = book.worksheets.get("xl/worksheets/sheet1.xml").unwrap();
        let mut xml_guard = xml.lock().unwrap();
        let version = xml_guard.decl.get_mut("version").unwrap();
        *version = "2.0".to_string();
        drop(xml_guard); // ロックを解放
        book.copy("./data/sample2.xlsx");

        // Assert
        let book = Book::new("./data/sample2.xlsx".to_string());
        let xml = book.worksheets.get("xl/worksheets/sheet1.xml").unwrap();
        let xml_guard = xml.lock().unwrap();
        assert_eq!(xml_guard.decl.get("version").unwrap(), "2.0");
        assert_eq!(xml_guard.decl.get("encoding").unwrap(), "UTF-8");
        assert_eq!(xml_guard.decl.get("standalone").unwrap(), "yes");
    }

    #[test]
    fn test_sheetnames() {
        // 観点: シート名一覧の取得

        // Act
        let book = Book::new("data/sample.xlsx".to_string());
        let sheetnames = book.sheetnames();

        // Assert
        assert!(!sheetnames.is_empty());
        assert!(sheetnames.contains(&"シート1".to_string()));
    }

    #[test]
    fn test_contains__() {
        // 観点: シート名の存在確認

        // Act
        let book = Book::new("data/sample.xlsx".to_string());

        // Assert
        assert!(book.__contains__("シート1".to_string()));
        assert!(!book.__contains__("存在しないシート".to_string()));
    }

    #[test]
    fn test_create_sheet() {
        // 観点: 新規シートの作成

        // Arrange
        let mut book = Book::new("data/sample.xlsx".to_string());
        let sheet_count_before = book.sheetnames().len();

        // Act
        let sheet = book.create_sheet("TestSheet".to_string(), sheet_count_before);

        // Assert
        assert_eq!(sheet.name, "TestSheet");
        assert_eq!(book.sheetnames().len(), sheet_count_before + 1);
        assert!(book.__contains__("TestSheet".to_string()));
    }

    #[test]
    fn test_merge_xmls() {
        // 観点: XMLの結合

        // Act
        let book = Book::new("data/sample.xlsx".to_string());
        let xmls = book.merge_xmls();

        // Assert
        assert!(xmls.contains_key("xl/workbook.xml"));
        assert!(xmls.contains_key("xl/styles.xml"));
        assert!(xmls.contains_key("xl/sharedStrings.xml"));

        // worksheetsのキーが含まれていることを確認
        for key in book.worksheets.keys() {
            assert!(xmls.contains_key(key));
        }
    }

    #[test]
    fn test_write_file() {
        // 観点: ファイルへの書き込み
        // 注: write_fileはprivateメソッドなので、間接的にテスト

        // Arrange
        if Path::new("data/write_test.xlsx").exists() {
            let _ = fs::remove_file("data/write_test.xlsx");
        }

        // Act
        let book = Book::new("data/sample.xlsx".to_string());
        book.copy("data/write_test.xlsx");

        // Assert
        assert!(Path::new("data/write_test.xlsx").exists());

        // 書き込まれたファイルが読み込み可能であることを確認
        let new_book = Book::new("data/write_test.xlsx".to_string());
        assert!(!new_book.worksheets.is_empty());
    }

    #[test]
    fn test_sheet_tags() {
        // 観点: シートタグの取得

        // Act
        let book = Book::new("data/sample.xlsx".to_string());
        let sheet_tags = book.sheet_tags();

        // Assert
        assert!(!sheet_tags.is_empty());

        // シートタグに必要な属性があることを確認
        let first_sheet = &sheet_tags[0];
        assert!(first_sheet.attributes.contains_key("name"));
        assert!(first_sheet.attributes.contains_key("sheetId"));
        assert!(first_sheet.attributes.contains_key("r:id"));
    }

    #[test]
    fn test_relationships() {
        // 観点: リレーションシップの取得

        // Act
        let book = Book::new("data/sample.xlsx".to_string());
        let relationships = book.get_relationships();

        // Assert
        assert!(!relationships.is_empty());

        // リレーションシップに必要な属性があることを確認
        let first_rel = &relationships[0];
        assert!(first_rel.attributes.contains_key("Id"));
        assert!(first_rel.attributes.contains_key("Type"));
        assert!(first_rel.attributes.contains_key("Target"));
    }

    #[test]
    fn test_sheet_paths() {
        // 観点: シートパスの取得

        // Act
        let book = Book::new("data/sample.xlsx".to_string());
        let sheet_paths = book.get_sheet_paths();

        // Assert
        assert!(!sheet_paths.is_empty());

        // Sheet1のパスが存在することを確認
        assert!(sheet_paths.contains_key("シート1"));

        // パスの形式が正しいことを確認
        for path in sheet_paths.values() {
            assert!(path.starts_with("xl/worksheets/"));
            assert!(path.ends_with(".xml"));
        }
    }
}
