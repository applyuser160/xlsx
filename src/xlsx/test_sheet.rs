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
        let appended_row = sheet.iter_rows(true).last().unwrap();
        assert_eq!(appended_row[0].value(), Some("foo".to_string()));
        assert_eq!(appended_row[1].value(), Some("bar".to_string()));
    }

    #[test]
    fn test_iter_rows() {
        // 観点: 行をイテレートできるか
        let book = Book::new("data/sample.xlsx");
        let sheet = book.__getitem__("シート1".to_string());
        let mut rows = sheet.iter_rows(true);
        let first_row = rows.next().unwrap();
        assert_eq!(first_row[0].value(), Some("1.0".to_string()));
        assert_eq!(first_row[1].value(), Some("3.0".to_string()));
        let second_row = rows.next().unwrap();
        assert_eq!(second_row[0].value(), Some("2.0".to_string()));
        assert_eq!(second_row[1].value(), Some("4.0".to_string()));
    }
}
