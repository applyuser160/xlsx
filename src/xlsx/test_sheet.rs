#[cfg(test)]
mod tests {
    use crate::book::Book;

    #[test]
    fn test_getitem() {
        // 観点: シート名でのセル取得
        let book = Book::new("data/sample.xlsx");
        let sheet = book.__getitem__("シート1").unwrap();

        // Act
        let cell = sheet.__getitem__("A1");

        // Assert
        assert_eq!(cell.value().unwrap(), "1.0");
    }

    #[test]
    fn test_cell() {
        // 観点: 行・列番号でのセル取得
        let book = Book::new("data/sample.xlsx");
        let sheet = book.__getitem__("シート1").unwrap();

        // Act
        let cell = sheet.cell(1, 1);

        // Assert
        assert_eq!(cell.value().unwrap(), "1.0");
    }
}
