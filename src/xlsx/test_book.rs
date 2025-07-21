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
        let book = Book::new("data/sample.xlsx");
        let xml = book.worksheets.get("xl/worksheets/sheet1.xml").unwrap();
        let xml_guard = xml.lock().unwrap();
        assert_eq!(xml_guard.decl.get("version").unwrap(), "1.0");
        assert_eq!(xml_guard.decl.get("encoding").unwrap(), "UTF-8");
        assert_eq!(xml_guard.decl.get("standalone").unwrap(), "yes");
    }

    #[test]
    fn test_copy_book() {
        let book = setup_book("copy_book");
        let copy_path = format!("{}.copy.xlsx", book.path);
        let xml = book.worksheets.get("xl/worksheets/sheet1.xml").unwrap();
        let mut xml_guard = xml.lock().unwrap();
        let version = xml_guard.decl.get_mut("version").unwrap();
        *version = "2.0".to_string();
        drop(xml_guard);
        book.copy(&copy_path);
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
        let book = Book::new("data/sample.xlsx");
        let sheetnames = book.sheetnames();
        assert!(!sheetnames.is_empty());
        assert!(sheetnames.contains(&"シート1".to_string()));
    }

    #[test]
    fn test_contains__() {
        let book = Book::new("data/sample.xlsx");
        assert!(book.__contains__("シート1".to_string()));
        assert!(!book.__contains__("存在しないシート".to_string()));
    }

    #[test]
    fn test_create_sheet() {
        let mut book = setup_book("create_sheet");
        let sheet_count_before = book.sheetnames().len();
        let sheet = book.create_sheet("TestSheet".to_string(), sheet_count_before);
        assert_eq!(sheet.name, "TestSheet");
        assert_eq!(book.sheetnames().len(), sheet_count_before + 1);
        assert!(book.__contains__("TestSheet".to_string()));
        cleanup(book);
    }

    #[test]
    fn test_write_file_indirectly() {
        let book = setup_book("write_file");
        let copy_path = format!("{}.copy.xlsx", book.path);
        book.copy(&copy_path);
        assert!(Path::new(&copy_path).exists());
        cleanup(book);
        let _ = fs::remove_file(copy_path);
    }

    #[test]
    fn test_delete_sheet() {
        let mut book = setup_book("delete_sheet");
        let sheet_count_before = book.sheetnames().len();
        assert!(book.__contains__("シート1".to_string()));
        let sheet_to_delete = book.__getitem__("シート1".to_string()).unwrap();
        book.__delitem__(sheet_to_delete.name.clone()).unwrap();
        assert_eq!(book.sheetnames().len(), sheet_count_before - 1);
        assert!(!book.__contains__("シート1".to_string()));
        cleanup(book);
    }

    #[test]
    fn test_sheet_index() {
        let book = setup_book("sheet_index");
        let sheet = book.__getitem__("シート1".to_string()).unwrap();
        let index = book.index(&sheet).unwrap();
        assert_eq!(index, 0);
        cleanup(book);
    }

    #[test]
    fn test_create_sheet_with_index() {
        let mut book = setup_book("create_with_index");
        let new_sheet = book.create_sheet("NewSheetAt0".to_string(), 0);
        let sheetnames = book.sheetnames();
        assert_eq!(sheetnames.len(), 2);
        assert_eq!(sheetnames[0], "NewSheetAt0");
        assert_eq!(sheetnames[1], "シート1");
        assert_eq!(new_sheet.name, "NewSheetAt0");
        cleanup(book);
    }

    #[test]
    fn test_add_table() {
        let mut book = setup_book("add_table");
        book.add_table(
            "シート1".to_string(),
            "Table1".to_string(),
            "A1:C5".to_string(),
        );
        assert!(book.tables.contains_key("xl/tables/table1.xml"));
        {
            let sheet_xml_arc = book.worksheets.get("xl/worksheets/sheet1.xml").unwrap();
            let sheet_xml = sheet_xml_arc.lock().unwrap();
            let table_parts = sheet_xml.elements[0]
                .children
                .iter()
                .find(|e| e.name == "tableParts")
                .unwrap();
            assert_eq!(table_parts.attributes.get("count").unwrap(), "1");
        }
        cleanup(book);
    }
}
