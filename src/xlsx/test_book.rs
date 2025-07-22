#[cfg(test)]
mod tests {
    use crate::book::Book;
    use std::{fs, path::Path};

    fn setup_book(test_name: &str) -> Book {
        let original_path = "data/sample.xlsx";
        let test_path = format!("data/test_book_{test_name}.xlsx");
        if Path::new(&test_path).exists() {
            let _ = fs::remove_file(&test_path);
        }
        fs::copy(original_path, &test_path).unwrap();
        Book::new(&test_path)
    }

    fn cleanup(book: Book) {
        let _ = fs::remove_file(book.path);
    }

    #[test]
    fn test_new_book() {
        // 観点: Excelファイルの読み取り

        // Act
        let book = Book::new("data/sample.xlsx");

        // Assert
        let xml = book.worksheets.get("xl/worksheets/sheet1.xml").unwrap();
        let xml_guard = xml.lock().unwrap();
        assert_eq!(xml_guard.decl.get("version").unwrap(), "1.0");
        assert_eq!(xml_guard.decl.get("encoding").unwrap(), "UTF-8");
        assert_eq!(xml_guard.decl.get("standalone").unwrap(), "yes");
    }

    #[test]
    fn test_copy_book() {
        // 観点: Excelファイルの名前をつけて保存
        let book = setup_book("copy_book");
        let copy_path = format!("{}.copy.xlsx", book.path);

        // Act
        let xml = book.worksheets.get("xl/worksheets/sheet1.xml").unwrap();
        let mut xml_guard = xml.lock().unwrap();
        let version = xml_guard.decl.get_mut("version").unwrap();
        *version = "2.0".to_string();
        drop(xml_guard); // ロックを解放
        book.copy(&copy_path);

        // Assert
        let book_copied = Book::new(&copy_path);
        let xml_copied = book_copied
            .worksheets
            .get("xl/worksheets/sheet1.xml")
            .unwrap();
        let xml_guard_copied = xml_copied.lock().unwrap();
        assert_eq!(xml_guard_copied.decl.get("version").unwrap(), "2.0");

        cleanup(book);
        let _ = fs::remove_file(copy_path);
    }

    #[test]
    fn test_sheetnames() {
        // 観点: シート名一覧の取得

        // Act
        let book = Book::new("data/sample.xlsx");
        let sheetnames = book.sheetnames();

        // Assert
        assert!(!sheetnames.is_empty());
        assert!(sheetnames.contains(&"シート1".to_string()));
    }

    #[test]
    fn test_contains__() {
        // 観点: シート名の存在確認

        // Act
        let book = Book::new("data/sample.xlsx");

        // Assert
        assert!(book.__contains__("シート1".to_string()));
        assert!(!book.__contains__("存在しないシート".to_string()));
    }

    #[test]
    fn test_create_sheet() {
        // 観点: 新規シートの作成

        // Arrange
        let mut book = setup_book("create_sheet");
        let sheet_count_before = book.sheetnames().len();

        // Act
        let sheet = book.create_sheet("TestSheet".to_string(), sheet_count_before);

        // Assert
        assert_eq!(sheet.name, "TestSheet");
        assert_eq!(book.sheetnames().len(), sheet_count_before + 1);
        assert!(book.__contains__("TestSheet".to_string()));
        cleanup(book);
    }

    #[test]
    fn test_write_file_indirectly() {
        // 観点: ファイルへの書き込み（copy経由での間接テスト）
        let book = setup_book("write_file");
        let copy_path = format!("{}.copy.xlsx", book.path);

        // Act
        book.copy(&copy_path);

        // Assert
        assert!(Path::new(&copy_path).exists());

        cleanup(book);
        let _ = fs::remove_file(copy_path);
    }

    #[test]
    fn test_sheet_tags() {
        // 観点: シートタグの取得

        // Act
        let book = Book::new("data/sample.xlsx");
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
        let book = Book::new("data/sample.xlsx");
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
        let book = Book::new("data/sample.xlsx");
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

    #[test]
    fn test_delete_sheet() {
        // 観点: シートを削除できるか
        let mut book = setup_book("delete_sheet");
        let sheet_count_before = book.sheetnames().len();
        assert!(book.__contains__("シート1".to_string()));

        // Act
        let sheet_to_delete = book.__getitem__("シート1".to_string());
        book.__delitem__(sheet_to_delete.name.clone());

        // Assert
        assert_eq!(book.sheetnames().len(), sheet_count_before - 1);
        assert!(!book.__contains__("シート1".to_string()));

        cleanup(book);
    }

    #[test]
    fn test_sheet_index() {
        // 観点: シートのインデックスを取得できるか
        let book = setup_book("sheet_index");
        let sheet = book.__getitem__("シート1".to_string());

        // Act
        let index = book.index(&sheet);

        // Assert
        assert_eq!(index, 0);

        cleanup(book);
    }

    #[test]
    fn test_create_sheet_with_index() {
        // 観点: 指定したインデックスにシートを作成できるか
        let mut book = setup_book("create_with_index");

        // Act
        let new_sheet = book.create_sheet("NewSheetAt0".to_string(), 0);

        // Assert
        let sheetnames = book.sheetnames();
        assert_eq!(sheetnames.len(), 2);
        assert_eq!(sheetnames[0], "NewSheetAt0");
        assert_eq!(sheetnames[1], "シート1");
        assert_eq!(new_sheet.name, "NewSheetAt0");

        cleanup(book);
    }

    #[test]
    fn test_add_table() {
        // 観点: テーブルを追加できるか
        let mut book = setup_book("add_table");

        // Act
        book.add_table(
            "シート1".to_string(),
            "Table1".to_string(),
            "A1:C5".to_string(),
        );

        // Assert
        assert!(book.tables.contains_key("xl/tables/table1.xml"));
        let sheet = book.get_sheet_by_name("シート1").unwrap();
        let sheet_xml_arc = sheet.get_xml();
        let sheet_xml = sheet_xml_arc.lock().unwrap();
        let table_parts = sheet_xml.elements[0]
            .children
            .iter()
            .find(|e| e.name == "tableParts")
            .unwrap();
        assert_eq!(table_parts.attributes.get("count").unwrap(), "1");

        cleanup(book);
    }
}
