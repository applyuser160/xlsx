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
        let book: Book = Book::new("data/sample.xlsx");
        let sheet = book.__getitem__("シート1".to_string());
        let new_row: Vec<String> = vec!["d".to_string(), "e".to_string(), "f".to_string()];

        // Act
        sheet.append(new_row);

        // Assert
        let xml = sheet.get_xml();
        let xml = xml.lock().unwrap();
        let worksheet = &xml.elements[0];
        let sheet_data = worksheet.get_element("sheetData");
        let rows = sheet_data.get_elements("row");
        let last_row = rows.last().unwrap();
        let last_row_num: usize = last_row.get_attribute("r").unwrap().parse().unwrap();
        assert_eq!(last_row_num, 4);
        let cells = last_row.get_elements("c");
        assert_eq!(cells.len(), 3);
        assert_eq!(cells[0].get_element("is>t").get_text(), "d");
        assert_eq!(cells[1].get_element("is>t").get_text(), "e");
        assert_eq!(cells[2].get_element("is>t").get_text(), "f");
    }

    #[test]
    fn test_iter_rows() {
        // 観点: 行をイテレートできるか
        let book: Book = Book::new("data/sample.xlsx");
        let sheet = book.__getitem__("シート1".to_string());

        // Act
        let rows = sheet.iter_rows(false).unwrap();

        // Assert
        assert_eq!(rows.len(), 3);
        assert_eq!(rows[0], vec!["1.0", "2.0", "3.0"]);
        assert_eq!(rows[1], vec!["4.0", "5.0", "6.0"]);
        assert_eq!(rows[2], vec!["7.0", "8.0", "9.0"]);
    }
}
