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
    fn test_append() {
        // 観点: 行を追加できるか
        let book = Book::new("data/sample.xlsx");
        let sheet = book.__getitem__("シート1".to_string());
        let new_row = vec!["foo".to_string(), "bar".to_string()];

        // Act
        sheet.append(new_row);

        // Assert
        let sheet_row_len = sheet.iter_rows().len();
        let binding = sheet.iter_rows();
        let appended_row = binding.get(sheet_row_len - 1).unwrap();
        assert_eq!(appended_row[0], "foo".to_string());
        assert_eq!(appended_row[1], "bar".to_string());
    }

    #[test]
    fn test_iter_rows() {
        // 観点: 行をイテレートできるか
        let book = Book::new("data/sample.xlsx");
        let sheet = book.__getitem__("シート1".to_string());
        let rows = sheet.iter_rows();
        let first_row = rows.first().unwrap();
        assert_eq!(first_row[0], "1.0".to_string());
        assert_eq!(first_row[1], "3.0".to_string());
        let second_row = rows.get(1).unwrap();
        assert_eq!(second_row[0], "2.0".to_string());
        assert_eq!(second_row[1], "4.0".to_string());
    }
}
