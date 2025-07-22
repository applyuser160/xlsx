#[cfg(test)]
mod tests {
    use crate::book::Book;

    #[test]
    fn test_getitem() {
        // 観点: セルをA1表記で取得できるか
        let book = Book::new("data/sample.xlsx");
        let sheet = book.__getitem__("シート1".to_string());

        // Act
        let cell = sheet.__getitem__("A1");

        // Assert
        assert_eq!(cell.value().unwrap(), "1.0");
    }

    #[test]
    fn test_cell() {
        // 観点: セルを行・列で取得できるか
        let book = Book::new("data/sample.xlsx");
        let sheet = book.__getitem__("シート1".to_string());

        // Act
        let cell = sheet.cell(1, 1);

        // Assert
        assert_eq!(cell.value().unwrap(), "1.0");
    }

    #[test]
    fn test_append_and_iter_rows() {
        // 観点: データの追記と読み込みが正常に行えるか
        let mut book = Book::new("data/sample.xlsx");
        let sheet = book.create_sheet("test_sheet".to_string(), 1);
        let data = vec![
            "Test1".to_string(),
            "Test2".to_string(),
            "Test3".to_string(),
        ];
        sheet.append(data.clone());

        // Act
        let rows = sheet.iter_rows(true).unwrap();

        // Assert
        assert_eq!(rows.len(), 1);
        assert_eq!(rows[0], data);
    }

    #[test]
    fn test_append_empty_and_special_chars() {
        // 観点: 空の行や特殊文字を含む行を扱えるか
        let mut book = Book::new("data/sample.xlsx");
        let sheet = book.create_sheet("test_sheet_2".to_string(), 2);
        let empty_data: Vec<String> = vec![];
        let special_chars_data = vec!["<&>".to_string(), "\"".to_string(), "'".to_string()];
        sheet.append(empty_data.clone());
        sheet.append(special_chars_data.clone());

        // Act
        let rows = sheet.iter_rows(true).unwrap();

        // Assert
        assert_eq!(rows.len(), 2);
        assert_eq!(rows[0], empty_data);
        assert_eq!(rows[1], special_chars_data);
    }
}
