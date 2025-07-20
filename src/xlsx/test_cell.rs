#[cfg(test)]
mod tests {
    use crate::xlsx::book::Book;
    use std::fs;

    fn setup_book(test_name: &str) -> Book {
        // テスト用のExcelファイルをコピーして使用
        let original_path = "data/sample.xlsx";
        let test_path = format!("data/test_cell_{test_name}.xlsx");
        if std::path::Path::new(&test_path).exists() {
            let _ = fs::remove_file(&test_path);
        }
        fs::copy(original_path, &test_path).unwrap();
        Book::new(test_path)
    }

    #[test]
    fn test_get_numeric_value() {
        // 観点: 数値セルの値が正しく読み取れるか
        let book = setup_book("get_numeric");
        let sheet = book.__getitem__("シート1".to_string());

        // Act
        let cell = sheet.__getitem__("A1");

        // Assert
        assert_eq!(cell.value().unwrap(), "1.0");
        let _ = fs::remove_file(&book.path);
    }

    #[test]
    fn test_get_non_existent_cell_value() {
        // 観点: 存在しないセルの値を読み取るとNoneが返るか
        let book = setup_book("get_non_existent");
        let sheet = book.__getitem__("シート1".to_string());

        // Act
        let cell = sheet.__getitem__("Z99");

        // Assert
        assert!(cell.value().is_none());
        let _ = fs::remove_file(&book.path);
    }

    #[test]
    fn test_set_numeric_value() {
        // 観点: 数値セルの値を書き換えることができるか
        let book = setup_book("set_numeric");
        let sheet = book.__getitem__("シート1".to_string());
        let copy_path = format!("{}.copy.xlsx", book.path);

        // Act
        let mut cell = sheet.__getitem__("A1");
        cell.set_value("999".to_string());
        book.copy(&copy_path);

        // Assert
        let book_reloaded = Book::new(copy_path.clone());
        let sheet_reloaded = book_reloaded.__getitem__("シート1".to_string());
        let cell_reloaded = sheet_reloaded.__getitem__("A1");
        assert_eq!(cell_reloaded.value().unwrap(), "999");

        let _ = fs::remove_file(&book.path);
        let _ = fs::remove_file(copy_path);
    }

    // TODO: `inlineStr`と`sharedStrings`のテストを追加
    // これには、それらの形式を持つテスト用のxlsxファイルが必要
}
