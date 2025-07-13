use std::collections::HashMap;
use std::fs::{File, OpenOptions};
use std::io::{Cursor, Read, Write};
use zip::write::FileOptions;
use zip::{ZipArchive, ZipWriter};

use pyo3::prelude::*;

use crate::xlsx::xml::Xml;

#[pyclass]
pub struct Book {
    #[pyo3(get, set)]
    pub path: String,
    pub xmls: HashMap<String, Xml>,
}

#[pymethods]
impl Book {
    #[new]
    pub fn new(path: String) -> Self {
        let file: File = File::open(&path).unwrap();
        let mut archive: ZipArchive<File> = ZipArchive::new(file).unwrap();
        let mut xmls: HashMap<String, Xml> = HashMap::new();

        for i in 0..archive.len() {
            let mut file: zip::read::ZipFile<'_> = archive.by_index(i).unwrap();
            let name: String = file.name().to_string();

            if name.ends_with(".xml") {
                let mut contents: String = String::new();
                file.read_to_string(&mut contents).unwrap();

                // xmlsの追加
                xmls.insert(name, Xml::new(&contents));
            }
        }

        Book { path, xmls }
    }

    // pub fn __repr__(&self) -> String {
    //     format!("<Book path='{}'>", self.path)
    // }

    // pub fn get_value(&self, sheet: String, address: String) -> String {
    //     let worksheet = self.get_sheet_by_name(&sheet);
    //     return match worksheet {
    //         Some(ws) => ws.get_value(address),
    //         None => "Sheet not found".to_string(),
    //     };
    // }
}

impl Book {
    pub fn save(&self) {
        let file: File = File::open(&self.path).unwrap();
        let mut archive: ZipArchive<File> = ZipArchive::new(file).unwrap();

        let mut buffer: Cursor<Vec<u8>> = Cursor::new(Vec::new());
        let mut zip_writer: ZipWriter<&mut Cursor<Vec<u8>>> = ZipWriter::new(&mut buffer);
        let options: FileOptions =
            FileOptions::default().compression_method(zip::CompressionMethod::Stored);

        for i in 0..archive.len() {
            let mut file: zip::read::ZipFile<'_> = archive.by_index(i).unwrap();
            let name: String = file.name().to_string();

            if let Some(new_contents) = self.xmls.get(&name) {
                // 差し替え対象
                zip_writer.start_file(name, options).unwrap();
                zip_writer.write_all(&new_contents.to_buf()).unwrap();
            } else {
                // その他 → そのままコピー
                let mut contents: Vec<u8> = Vec::new();
                file.read_to_end(&mut contents).unwrap();
                zip_writer.start_file(name, options).unwrap();
                zip_writer.write_all(&contents).unwrap();
            }
        }
    }

    pub fn copy(&self, path: &str) {
        let file: File = File::open(&self.path).unwrap();
        let mut archive: ZipArchive<File> = ZipArchive::new(file).unwrap();

        // 新しいファイルを作成
        let new_file: File = OpenOptions::new()
            .write(true)
            .create(true)
            .truncate(true)
            .open(path)
            .unwrap();

        let mut zip_writer: ZipWriter<File> = ZipWriter::new(new_file);
        let options: FileOptions =
            FileOptions::default().compression_method(zip::CompressionMethod::Stored);

        for i in 0..archive.len() {
            let mut file: zip::read::ZipFile<'_> = archive.by_index(i).unwrap();
            let name: String = file.name().to_string();

            if let Some(new_contents) = self.xmls.get(&name) {
                // 差し替え対象のXMLファイル
                zip_writer.start_file(name, options).unwrap();
                zip_writer.write_all(&new_contents.to_buf()).unwrap();
            } else {
                // その他のファイル → そのままコピー
                let mut contents: Vec<u8> = Vec::new();
                file.read_to_end(&mut contents).unwrap();
                zip_writer.start_file(name, options).unwrap();
                zip_writer.write_all(&contents).unwrap();
            }
        }

        // ファイルを閉じる（ZipWriterのdropで自動的に行われるが、明示的にfinishを呼ぶ）
        zip_writer.finish().unwrap();
    }

    // pub fn get_sheet_by_name(&self, name: &String) -> Option<&Worksheet> {
    //     self.value.get_sheet_by_name(name)
    // }

    // pub fn get_sheet_by_index(&self, index: &usize) -> Option<&Worksheet> {
    //     self.value.get_sheet(index)
    // }
}

impl Drop for Book {
    fn drop(&mut self) {
        self.save();
    }
}
