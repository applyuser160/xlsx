#[cfg(test)]
mod tests {
    use crate::book::Book;

    #[test]
    fn test_getitem() {
        // 観点: セルをA1表記で取得できるか
        let book: Book = Book::new("data/sample.xlsx");
        let sheet = book.__getitem__("シート1".to_string());

        // Act
        let cell = sheet.__getitem__("A1");

        // Assert
        assert_eq!(cell.value().unwrap(), "1.0");
    }

    #[test]
    fn test_cell() {
        // 観点: セルを行・列で取得できるか
        let book: Book = Book::new("data/sample.xlsx");
        let sheet = book.__getitem__("シート1".to_string());

        // Act
        let cell = sheet.cell(1, 1);

        // Assert
        assert_eq!(cell.value().unwrap(), "1.0");
    }

    #[test]
    fn test_append() {
        // 観点: 行を追加できるか
        // Setup
        let book: Book = Book::new("data/sample.xlsx");
        let sheet = book.__getitem__("シート1".to_string());
        let row_data: Vec<String> = vec!["A".to_string(), "B".to_string(), "C".to_string()];

        // Act
        sheet.append(row_data);

        // Assert
        let xml = sheet.get_xml();
        let xml = xml.lock().unwrap();
        let worksheet = &xml.elements[0];
        let sheet_data = worksheet.get_element("sheetData");
        let rows = sheet_data.get_elements("row");
        let last_row = rows.last().unwrap();
        assert_eq!(last_row.get_attribute("r").unwrap(), "4");
    }

    #[test]
    fn test_iter_rows() {
        // 観点: 行をイテレートできるか
        // Setup
        let book: Book = Book::new("data/sample.xlsx");
        let sheet = book.__getitem__("シート1".to_string());
        let expected: Vec<Vec<String>> = vec![
            vec!["1.0".to_string(), "2.0".to_string(), "3.0".to_string()],
            vec!["4.0".to_string(), "5.0".to_string(), "6.0".to_string()],
            vec![
                "7.0".to_string(),
                "8.0".to_string(),
                "9.0".to_string(),
                "10.0".to_string(),
            ],
        ];

        // Act
        let rows = sheet.iter_rows(true).unwrap();

        // Assert
        assert_eq!(rows, expected);
    }
}
