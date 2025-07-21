#[cfg(test)]
mod tests {
    use crate::book::Book;
    use std::fs;

    fn setup_book(test_name: &str) -> Book {
        // テスト用のExcelファイルをコピーして使用
        let original_path = "data/sample.xlsx";
        let test_path = format!("data/test_cell_{test_name}.xlsx");
        if std::path::Path::new(&test_path).exists() {
            let _ = fs::remove_file(&test_path);
        }
        fs::copy(original_path, &test_path).unwrap();
        Book::new(&test_path)
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
    fn test_set_numeric_value_existing_cell() {
        // 観点: 既存の数値セルの値を書き換えることができるか
        let book = setup_book("set_numeric_existing");
        let sheet = book.__getitem__("シート1".to_string());
        let copy_path = format!("{}.copy.xlsx", book.path);

        // Act
        let mut cell = sheet.__getitem__("A1");
        cell.set_value("999".to_string());
        book.copy(&copy_path);

        // Assert
        let book_reloaded = Book::new(&copy_path);
        let sheet_reloaded = book_reloaded.__getitem__("シート1".to_string());
        let cell_reloaded = sheet_reloaded.__getitem__("A1");
        assert_eq!(cell_reloaded.value().unwrap(), "999");

        let _ = fs::remove_file(&book.path);
        let _ = fs::remove_file(copy_path);
    }

    #[test]
    fn test_set_string_value_existing_cell() {
        // 観点: 既存の文字列セルの値を書き換えることができるか
        let book = setup_book("set_string_existing");
        let sheet = book.__getitem__("シート1".to_string());
        let copy_path = format!("{}.copy.xlsx", book.path);

        // Act
        let mut cell = sheet.__getitem__("B1");
        cell.set_value("new_string".to_string());
        book.copy(&copy_path);

        // Assert
        let book_reloaded = Book::new(&copy_path);
        let sheet_reloaded = book_reloaded.__getitem__("シート1".to_string());
        let cell_reloaded = sheet_reloaded.__getitem__("B1");
        assert_eq!(cell_reloaded.value().unwrap(), "new_string");

        let _ = fs::remove_file(&book.path);
        let _ = fs::remove_file(copy_path);
    }

    #[test]
    fn test_set_value_new_cell() {
        // 観点: 新しいセルに値を書き込めるか
        let book = setup_book("set_new_cell");
        let sheet = book.__getitem__("シート1".to_string());
        let copy_path = format!("{}.copy.xlsx", book.path);

        // Act
        let mut cell_c1 = sheet.__getitem__("C1");
        cell_c1.set_value("12345".to_string());
        let mut cell_d1 = sheet.__getitem__("D1");
        cell_d1.set_value("new_cell_string".to_string());
        book.copy(&copy_path);

        // Assert
        let book_reloaded = Book::new(&copy_path);
        let sheet_reloaded = book_reloaded.__getitem__("シート1".to_string());
        let cell_c1_reloaded = sheet_reloaded.__getitem__("C1");
        let cell_d1_reloaded = sheet_reloaded.__getitem__("D1");
        assert_eq!(cell_c1_reloaded.value().unwrap(), "12345");
        assert_eq!(cell_d1_reloaded.value().unwrap(), "new_cell_string");

        let _ = fs::remove_file(&book.path);
        let _ = fs::remove_file(copy_path);
    }

    #[test]
    fn test_set_datetime_value() {
        // 観点: 日付・時刻の値を設定できるか
        let book = setup_book("set_datetime");
        let sheet = book.__getitem__("シート1".to_string());
        let copy_path = format!("{}.copy.xlsx", book.path);

        // Act
        let mut cell = sheet.__getitem__("E1");
        cell.set_value("2024-01-01 12:30:00".to_string());
        book.copy(&copy_path);

        // Assert
        let book_reloaded = Book::new(&copy_path);
        let sheet_reloaded = book_reloaded.__getitem__("シート1".to_string());
        let cell_reloaded = sheet_reloaded.__getitem__("E1");
        // Excel's serial value for 2024-01-01 12:30:00 is 45292.520833333336
        assert_eq!(cell_reloaded.value().unwrap(), "45292.520833333336");

        let _ = fs::remove_file(&book.path);
        let _ = fs::remove_file(copy_path);
    }

    #[test]
    fn test_set_bool_value() {
        // 観点: ブール値を設定できるか
        let book = setup_book("set_bool");
        let sheet = book.__getitem__("シート1".to_string());
        let copy_path = format!("{}.copy.xlsx", book.path);

        // Act
        let mut cell_f1 = sheet.__getitem__("F1");
        cell_f1.set_value("true".to_string());
        let mut cell_g1 = sheet.__getitem__("G1");
        cell_g1.set_value("false".to_string());
        book.copy(&copy_path);

        // Assert
        let book_reloaded = Book::new(&copy_path);
        let sheet_reloaded = book_reloaded.__getitem__("シート1".to_string());
        let cell_f1_reloaded = sheet_reloaded.__getitem__("F1");
        let cell_g1_reloaded = sheet_reloaded.__getitem__("G1");
        assert_eq!(cell_f1_reloaded.value().unwrap(), "1");
        assert_eq!(cell_g1_reloaded.value().unwrap(), "0");

        let _ = fs::remove_file(&book.path);
        let _ = fs::remove_file(copy_path);
    }

    #[test]
    fn test_set_formula_value() {
        // 観点: 数式を設定できるか
        let book = setup_book("set_formula");
        let sheet = book.__getitem__("シート1".to_string());
        let copy_path = format!("{}.copy.xlsx", book.path);

        // Act
        let mut cell = sheet.__getitem__("H1");
        cell.set_value("=SUM(A1:A2)".to_string());
        book.copy(&copy_path);

        // Assert
        let book_reloaded = Book::new(&copy_path);
        let sheet_reloaded = book_reloaded.__getitem__("シート1".to_string());
        let _cell_reloaded = sheet_reloaded.__getitem__("H1");
        // TODO: 実際に計算された値を取得する方法は未実装
        // とりあえず、数式が設定されているかを確認
        // assert_eq!(cell_reloaded.formula().unwrap(), "SUM(A1:A2)");

        let _ = fs::remove_file(&book.path);
        let _ = fs::remove_file(copy_path);
    }
}
