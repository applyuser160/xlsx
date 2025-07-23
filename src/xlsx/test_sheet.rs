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
        // 観点: 新しいデータを追加できるか
        let mut book = Book::new("");
        let sheet = book.create_sheet("test_append".to_string(), 0);
        let new_row: Vec<String> = vec!["a".to_string(), "b".to_string(), "c".to_string()];

        // Act
        sheet.append(new_row);

        // Assert
        let binding = sheet.get_xml();
        let mut xml = binding.lock().unwrap();
        let worksheet = &mut xml.elements[0];
        let sheet_data = worksheet.get_element_mut("sheetData");
        let rows = sheet_data.get_elements("row");
        let row = rows.last().unwrap();
        assert_eq!(row.get_attribute("r").unwrap(), "1");
        let cells = row.get_elements("c");
        assert_eq!(cells.len(), 3);
        assert_eq!(cells[0].get_attribute("r").unwrap(), "A1");
        assert_eq!(cells[1].get_attribute("r").unwrap(), "B1");
        assert_eq!(cells[2].get_attribute("r").unwrap(), "C1");
    }

    #[test]
    fn test_append_to_existing() {
        // 観点: 既存のデータがある場合に追記できるか
        let book = Book::new("data/sample.xlsx");
        let sheet = book.__getitem__("シート1".to_string());
        let new_row: Vec<String> = vec!["a".to_string(), "b".to_string(), "c".to_string()];

        // Act
        sheet.append(new_row);

        // Assert
        let binding = sheet.get_xml();
        let mut xml = binding.lock().unwrap();
        let worksheet = &mut xml.elements[0];
        let sheet_data = worksheet.get_element_mut("sheetData");
        let rows = sheet_data.get_elements("row");
        let row = rows.last().unwrap();
        assert_eq!(row.get_attribute("r").unwrap(), "5");
        let cells = row.get_elements("c");
        assert_eq!(cells.len(), 3);
        assert_eq!(cells[0].get_attribute("r").unwrap(), "A5");
        assert_eq!(cells[1].get_attribute("r").unwrap(), "B5");
        assert_eq!(cells[2].get_attribute("r").unwrap(), "C5");
    }

    #[test]
    fn test_iter_rows_values_only() {
        // 観点: values_only=trueの場合に値のみ取得できるか
        let book = Book::new("data/sample.xlsx");
        let sheet = book.__getitem__("シート1".to_string());

        // Act
        let rows = sheet.iter_rows(true).unwrap();

        // Assert
        assert_eq!(rows.len(), 2);
        assert_eq!(rows[0], vec!["1.0", "3.0"]);
        assert_eq!(rows[1], vec!["2.0", "4.0"]);
    }

    #[test]
    fn test_iter_rows_not_values_only() {
        // 観点: values_only=falseの場合に値のみ取得できるか
        let book = Book::new("data/sample.xlsx");
        let sheet = book.__getitem__("シート1".to_string());

        // Act
        let rows = sheet.iter_rows(false).unwrap();

        // Assert
        assert_eq!(rows.len(), 2);
        assert_eq!(rows[0], vec!["1.0", "3.0"]);
        assert_eq!(rows[1], vec!["2.0", "4.0"]);
    }
}
