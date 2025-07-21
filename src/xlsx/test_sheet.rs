#[cfg(test)]
mod tests {
    use crate::xlsx::book::Book;

    #[test]
    fn test_getitem() {
        // 観点: セルをA1表記で取得できるか
        let book = Book::new("data/sample.xlsx".to_string());
        let sheet = book.__getitem__("シート1".to_string());

        // Act
        let cell = sheet.__getitem__("A1");

        // Assert
        assert_eq!(cell.value().unwrap(), "1.0");
    }

    #[test]
    fn test_cell() {
        // 観点: セルを行・列で取得できるか
        let book = Book::new("data/sample.xlsx".to_string());
        let sheet = book.__getitem__("シート1".to_string());

        // Act
        let cell = sheet.cell(1, 1);

        // Assert
        assert_eq!(cell.value().unwrap(), "1.0");
    }

    #[test]
    fn test_insert_rows() {
        // 観点: 行を挿入できるか
        let book = Book::new("data/sample.xlsx".to_string());
        let sheet = book.__getitem__("シート1".to_string());
        let cell_before = sheet.cell(2, 1);
        assert_eq!(cell_before.value().unwrap(), "2.0");

        // Act
        sheet.insert_rows(2, 1);

        // Assert
        let cell_after = sheet.cell(3, 1);
        assert_eq!(cell_after.value().unwrap(), "2.0");
    }

    #[test]
    fn test_delete_rows() {
        // 観点: 行を削除できるか
        let book = Book::new("data/sample.xlsx".to_string());
        let sheet = book.__getitem__("シート1".to_string());
        let cell_before = sheet.cell(2, 1);
        assert_eq!(cell_before.value().unwrap(), "2.0");

        // Act
        sheet.delete_rows(1, 1);

        // Assert
        let cell_after = sheet.cell(1, 1);
        assert_eq!(cell_after.value().unwrap(), "2.0");
    }

    #[test]
    fn test_insert_cols() {
        // 観点: 列を挿入できるか
        let book = Book::new("data/sample.xlsx".to_string());
        let sheet = book.__getitem__("シート1".to_string());
        let cell_before = sheet.cell(1, 2);
        assert_eq!(cell_before.value().unwrap(), "3.0");

        // Act
        sheet.insert_cols(1, 1);

        // Assert
        let cell_after = sheet.cell(1, 3);
        assert_eq!(cell_after.value().unwrap(), "3.0");
    }

    #[test]
    fn test_delete_cols() {
        // 観点: 列を削除できるか
        let book = Book::new("data/sample.xlsx".to_string());
        let sheet = book.__getitem__("シート1".to_string());
        let cell_before = sheet.cell(1, 2);
        assert_eq!(cell_before.value().unwrap(), "3.0");

        // Act
        sheet.delete_cols(1, 1);

        // Assert
        let cell_after = sheet.cell(1, 1);
        assert_eq!(cell_after.value().unwrap(), "3.0");
    }

    #[test]
    fn test_set_row_height() {
        // 観点: 行の高さを設定できるか
        let book = Book::new("data/sample.xlsx".to_string());
        let sheet = book.__getitem__("シート1".to_string());

        // Act
        sheet.set_row_height(1, 30.0);

        // Assert
        let xml = sheet.xml.lock().unwrap();
        let sheet_data = xml.elements.iter().find(|e| e.name == "worksheet").unwrap().children.iter().find(|e| e.name == "sheetData").unwrap();
        let row = sheet_data.children.iter().find(|r| r.attributes.get("r").unwrap() == "1").unwrap();
        assert_eq!(row.attributes.get("ht").unwrap(), "30");
        assert_eq!(row.attributes.get("customHeight").unwrap(), "1");
    }

    #[test]
    fn test_set_column_width() {
        // 観点: 列の幅を設定できるか
        let mut book = Book::new("data/sample.xlsx".to_string());
        let sheet = book.__getitem__("シート1".to_string());

        // Act
        sheet.set_column_width(1, 20.0);

        // Assert
        let xml = sheet.xml.lock().unwrap();
        let worksheet = xml.elements.iter().find(|e| e.name == "worksheet").unwrap();
        let cols = worksheet.children.iter().find(|e| e.name == "cols").unwrap();
        let col = cols.children.iter().find(|c| c.attributes.get("min").unwrap() == "1").unwrap();
        assert_eq!(col.attributes.get("width").unwrap(), "20");
        assert_eq!(col.attributes.get("customWidth").unwrap(), "1");
    }

    #[test]
    fn test_merge_cells() {
        // 観点: セルをマージできるか
        let book = Book::new("data/sample.xlsx".to_string());
        let sheet = book.__getitem__("シート1".to_string());

        // Act
        sheet.merge_cells("A1:B2");

        // Assert
        let xml = sheet.xml.lock().unwrap();
        let worksheet = xml.elements.iter().find(|e| e.name == "worksheet").unwrap();
        let merge_cells = worksheet.children.iter().find(|e| e.name == "mergeCells").unwrap();
        let merge_cell = merge_cells.children.iter().find(|mc| mc.attributes.get("ref").unwrap() == "A1:B2").unwrap();
        assert!(merge_cell.attributes.get("ref").is_some());
    }

    #[test]
    fn test_unmerge_cells() {
        // 観点: セルのマージを解除できるか
        let book = Book::new("data/sample.xlsx".to_string());
        let sheet = book.__getitem__("シート1".to_string());
        sheet.merge_cells("A1:B2");

        // Act
        sheet.unmerge_cells("A1:B2");

        // Assert
        let xml = sheet.xml.lock().unwrap();
        let worksheet = xml.elements.iter().find(|e| e.name == "worksheet").unwrap();
        let merge_cells = worksheet.children.iter().find(|e| e.name == "mergeCells").unwrap();
        assert!(merge_cells.children.is_empty());
    }

    #[test]
    fn test_freeze_panes() {
        // 観点: ウィンドウ枠を固定できるか
        let book = Book::new("data/sample.xlsx".to_string());
        let sheet = book.__getitem__("シート1".to_string());

        // Act
        sheet.freeze_panes("B2");

        // Assert
        let xml = sheet.xml.lock().unwrap();
        let worksheet = xml.elements.iter().find(|e| e.name == "worksheet").unwrap();
        let sheet_views = worksheet.children.iter().find(|e| e.name == "sheetViews").unwrap();
        let sheet_view = sheet_views.children.iter().find(|e| e.name == "sheetView").unwrap();
        let pane = sheet_view.children.iter().find(|e| e.name == "pane").unwrap();
        assert_eq!(pane.attributes.get("topLeftCell").unwrap(), "B2");
        assert_eq!(pane.attributes.get("xSplit").unwrap(), "1");
        assert_eq!(pane.attributes.get("ySplit").unwrap(), "1");
    }
}
