#[cfg(test)]
mod tests {
    use std::{fs, path::Path};

    use crate::xlsx::book::Book;

    #[test]
    fn test_new_book() {
        // 観点: Excelファイルの読み取り

        // Act
        let book = Book::new("data/sample.xlsx".to_string());

        // Assert
        let xml = book.worksheets.get("xl/worksheets/sheet1.xml").unwrap();
        assert_eq!(xml.decl.get("version").unwrap(), "1.0");
        assert_eq!(xml.decl.get("encoding").unwrap(), "UTF-8");
        assert_eq!(xml.decl.get("standalone").unwrap(), "yes");
    }

    #[test]
    fn test_save_book() {
        // 観点: Excelファイルの書き込み

        // Arrange

        // ファイルが存在しないことを確認
        if Path::new("data/sample2.xlsx").exists() {
            let _ = fs::remove_file("data/sample2.xlsx");
        }
        assert!(!Path::new("data/sample2.xlsx").exists());

        // Act
        let mut book = Book::new("data/sample.xlsx".to_string());
        let xml = book.worksheets.get_mut("xl/worksheets/sheet1.xml").unwrap();
        let version = xml.decl.get_mut("version").unwrap();
        *version = "1.0".to_string();
        book.save();

        // Assert
        let mut book = Book::new("data/sample.xlsx".to_string());
        let xml = book.worksheets.get_mut("xl/worksheets/sheet1.xml").unwrap();
        assert_eq!(xml.decl.get("version").unwrap(), "1.0");
        assert_eq!(xml.decl.get("encoding").unwrap(), "UTF-8");
        assert_eq!(xml.decl.get("standalone").unwrap(), "yes");
    }

    #[test]
    fn test_copy_book() {
        // 観点: Excelファイルの名前をつけて保存

        // Arrange

        // ファイルが存在しないことを確認
        if Path::new("data/sample2.xlsx").exists() {
            let _ = fs::remove_file("data/sample2.xlsx");
        }
        assert!(!Path::new("data/sample2.xlsx").exists());

        // Act
        let mut book = Book::new("data/sample.xlsx".to_string());
        let xml = book.worksheets.get_mut("xl/worksheets/sheet1.xml").unwrap();
        let version = xml.decl.get_mut("version").unwrap();
        *version = "2.0".to_string();
        book.copy("data/sample2.xlsx");

        // Assert
        let mut book = Book::new("data/sample2.xlsx".to_string());
        let xml = book.worksheets.get_mut("xl/worksheets/sheet1.xml").unwrap();
        assert_eq!(xml.decl.get("version").unwrap(), "2.0");
        assert_eq!(xml.decl.get("encoding").unwrap(), "UTF-8");
        assert_eq!(xml.decl.get("standalone").unwrap(), "yes");
    }
}
