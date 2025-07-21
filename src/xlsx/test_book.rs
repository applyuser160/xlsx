#[cfg(test)]
mod tests {
    use crate::xlsx::book::Book;
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
    fn test_active_sheet_logic() {
        // 観点: アクティブシートのロジック確認
        let mut book = setup_book("active_sheet_logic");
        assert_eq!(book.active_sheet_index, 0);

        // Act
        book.create_sheet("NewSheet".to_string(), 1);
        book.active_sheet_index = 1;
        book.update_active_tab();

        // Assert
        let workbook_tag = book.workbook.elements.first().unwrap();
        let book_views = workbook_tag.children.iter().find(|c| c.name == "bookViews").unwrap();
        let workbook_view = book_views.children.iter().find(|c| c.name == "workbookView").unwrap();
        assert_eq!(workbook_view.attributes.get("activeTab").unwrap(), "1");

        cleanup(book);
    }

    #[test]
    fn test_named_range() {
        // 観点: 名前付き範囲の作成と削除
        let mut book = setup_book("named_range");
        assert!(book.defined_names.is_empty());

        // Act (作成)
        book.create_named_range("TestRange".to_string(), "シート1!$A$1".to_string(), None);

        // Assert (作成)
        assert_eq!(book.defined_names.len(), 1);
        let named_range = &book.defined_names[0];
        assert_eq!(named_range.attributes.get("name").unwrap(), "TestRange");
        assert_eq!(named_range.text.as_ref().unwrap(), "シート1!$A$1");

        // XMLの確認 (作成)
        let workbook_tag = book.workbook.elements.first().unwrap();
        let defined_names_tag = workbook_tag.children.iter().find(|c| c.name == "definedNames").unwrap();
        assert_eq!(defined_names_tag.children.len(), 1);

        // Act (削除)
        book.delete_named_range("TestRange".to_string());

        // Assert (削除)
        assert!(book.defined_names.is_empty());
        let workbook_tag_after_delete = book.workbook.elements.first().unwrap();
        let defined_names_tag_after_delete = workbook_tag_after_delete.children.iter().find(|c| c.name == "definedNames").unwrap();
        assert!(defined_names_tag_after_delete.children.is_empty());


        cleanup(book);
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
    fn test_merge_xmls() {
        // 観点: XMLの結合

        // Act
        let book = Book::new("data/sample.xlsx");
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
    fn test_set_print_area() {
        // 観点: 印刷範囲を設定できるか
        let mut book = setup_book("set_print_area");

        // Act
        book.set_print_area("シート1", "A1:B10");

        // Assert
        let workbook = &book.workbook.elements[0];
        let defined_names = workbook.children.iter().find(|e| e.name == "definedNames").unwrap();
        let defined_name = defined_names.children.iter().find(|dn| dn.attributes.get("name").unwrap() == "_xlnm.Print_Area").unwrap();
        assert_eq!(defined_name.text.as_ref().unwrap(), "'シート1'!A1:B10");
        assert_eq!(defined_name.attributes.get("localSheetId").unwrap(), "0");

        cleanup(book);
    }

    #[test]
    fn test_copy_worksheet() {
        // 観点: シートをコピーできるか
        let mut book = setup_book("copy_worksheet");

        // Act
        let copied_sheet = book.copy_worksheet("シート1", "シート1 コピー");

        // Assert
        assert_eq!(copied_sheet.name, "シート1 コピー");
        assert!(book.__contains__("シート1 コピー".to_string()));

        let original_sheet = book.__getitem__("シート1".to_string());
        let original_xml = original_sheet.xml.lock().unwrap();
        let copied_xml = copied_sheet.xml.lock().unwrap();

        assert_eq!(original_xml.elements.len(), copied_xml.elements.len());
        for i in 0..original_xml.elements.len() {
            assert_eq!(original_xml.elements[i].name, copied_xml.elements[i].name);
            assert_eq!(original_xml.elements[i].attributes, copied_xml.elements[i].attributes);
            assert_eq!(original_xml.elements[i].text, copied_xml.elements[i].text);
        }

        cleanup(book);
    }
}
