use std::sync::{Arc, Mutex};

use pyo3::prelude::*;

use crate::xlsx::cell::Cell;
use crate::xlsx::xml::Xml;

#[pyclass]
#[derive(Clone)]
pub struct Sheet {
    #[pyo3(get)]
    pub name: String,
    pub xml: Arc<Mutex<Xml>>,
    shared_strings: Arc<Mutex<Xml>>,
}

#[pymethods]
impl Sheet {
    // pub fn __repr__(&self) -> String {
    //     format!("<Sheet name='{}'>", self.name)
    // }

    pub fn __getitem__(&self, key: &str) -> Cell {
        Cell::new(
            self.xml.clone(),
            self.shared_strings.clone(),
            key.to_string(),
        )
    }

    #[pyo3(signature = (row, column))]
    pub fn cell(&self, row: usize, column: usize) -> Cell {
        let address = Self::coordinate_to_string(row, column);
        Cell::new(self.xml.clone(), self.shared_strings.clone(), address)
    }

    pub fn insert_rows(&self, from_row_idx: usize, num_rows: usize) {
        let mut xml = self.xml.lock().unwrap();
        let sheet_data = xml
            .elements
            .iter_mut()
            .find(|e| e.name == "worksheet")
            .unwrap()
            .children
            .iter_mut()
            .find(|e| e.name == "sheetData")
            .unwrap();

        for row in sheet_data.children.iter_mut() {
            let row_idx: usize = row.attributes.get("r").unwrap().parse().unwrap();
            if row_idx >= from_row_idx {
                let new_row_idx = row_idx + num_rows;
                row.attributes
                    .insert("r".to_string(), new_row_idx.to_string());
                for cell in row.children.iter_mut() {
                    let cell_address = cell.attributes.get("r").unwrap();
                    let (col_str, _) = Self::split_address(cell_address);
                    let new_address = format!("{}{}", col_str, new_row_idx);
                    cell.attributes
                        .insert("r".to_string(), new_address.to_string());
                }
            }
        }
    }

    pub fn delete_rows(&self, from_row_idx: usize, num_rows: usize) {
        let mut xml = self.xml.lock().unwrap();
        let sheet_data = xml
            .elements
            .iter_mut()
            .find(|e| e.name == "worksheet")
            .unwrap()
            .children
            .iter_mut()
            .find(|e| e.name == "sheetData")
            .unwrap();

        sheet_data
            .children
            .retain(|row| {
                let row_idx: usize = row.attributes.get("r").unwrap().parse().unwrap();
                row_idx < from_row_idx || row_idx >= from_row_idx + num_rows
            });

        for row in sheet_data.children.iter_mut() {
            let row_idx: usize = row.attributes.get("r").unwrap().parse().unwrap();
            if row_idx >= from_row_idx + num_rows {
                let new_row_idx = row_idx - num_rows;
                row.attributes
                    .insert("r".to_string(), new_row_idx.to_string());
                for cell in row.children.iter_mut() {
                    let cell_address = cell.attributes.get("r").unwrap();
                    let (col_str, _) = Self::split_address(cell_address);
                    let new_address = format!("{}{}", col_str, new_row_idx);
                    cell.attributes
                        .insert("r".to_string(), new_address.to_string());
                }
            }
        }
    }

    pub fn insert_cols(&self, from_col_idx: usize, num_cols: usize) {
        let mut xml = self.xml.lock().unwrap();
        let sheet_data = xml
            .elements
            .iter_mut()
            .find(|e| e.name == "worksheet")
            .unwrap()
            .children
            .iter_mut()
            .find(|e| e.name == "sheetData")
            .unwrap();

        for row in sheet_data.children.iter_mut() {
            for cell in row.children.iter_mut() {
                let cell_address = cell.attributes.get("r").unwrap();
                let (col_str, row_idx) = Self::split_address(cell_address);
                let col_idx = Self::col_str_to_num(&col_str);
                if col_idx >= from_col_idx {
                    let new_col_idx = col_idx + num_cols;
                    let new_col_str = Self::col_num_to_str(new_col_idx);
                    let new_address = format!("{}{}", new_col_str, row_idx);
                    cell.attributes
                        .insert("r".to_string(), new_address.to_string());
                }
            }
        }
    }

    pub fn delete_cols(&self, from_col_idx: usize, num_cols: usize) {
        let mut xml = self.xml.lock().unwrap();
        let sheet_data = xml
            .elements
            .iter_mut()
            .find(|e| e.name == "worksheet")
            .unwrap()
            .children
            .iter_mut()
            .find(|e| e.name == "sheetData")
            .unwrap();

        for row in sheet_data.children.iter_mut() {
            row.children.retain(|cell| {
                let cell_address = cell.attributes.get("r").unwrap();
                let (col_str, _) = Self::split_address(cell_address);
                let col_idx = Self::col_str_to_num(&col_str);
                col_idx < from_col_idx || col_idx >= from_col_idx + num_cols
            });

            for cell in row.children.iter_mut() {
                let cell_address = cell.attributes.get("r").unwrap();
                let (col_str, row_idx) = Self::split_address(cell_address);
                let col_idx = Self::col_str_to_num(&col_str);
                if col_idx >= from_col_idx + num_cols {
                    let new_col_idx = col_idx - num_cols;
                    let new_col_str = Self::col_num_to_str(new_col_idx);
                    let new_address = format!("{}{}", new_col_str, row_idx);
                    cell.attributes.insert("r".to_string(), new_address.to_string());
                }
            }
        }
    }

    pub fn set_row_height(&self, row_idx: usize, height: f64) {
        let mut xml = self.xml.lock().unwrap();
        let sheet_data = xml
            .elements
            .iter_mut()
            .find(|e| e.name == "worksheet")
            .unwrap()
            .children
            .iter_mut()
            .find(|e| e.name == "sheetData")
            .unwrap();

        let row = sheet_data
            .children
            .iter_mut()
            .find(|r| r.attributes.get("r").unwrap().parse::<usize>().unwrap() == row_idx);

        if let Some(row) = row {
            row.attributes
                .insert("ht".to_string(), height.to_string());
            row.attributes.insert("customHeight".to_string(), "1".to_string());
        }
    }

    pub fn set_column_width(&self, col_idx: usize, width: f64) {
        let mut xml = self.xml.lock().unwrap();
        let worksheet = xml
            .elements
            .iter_mut()
            .find(|e| e.name == "worksheet")
            .unwrap();

        let cols = worksheet.children.iter_mut().find(|e| e.name == "cols");

        if cols.is_none() {
            let new_cols = crate::xlsx::xml::XmlElement::new("cols");
            worksheet.children.insert(0, new_cols);
        }

        let cols = worksheet.children.iter_mut().find(|e| e.name == "cols").unwrap();
        let col = cols.children.iter_mut().find(|c| {
            let min: usize = c.attributes.get("min").unwrap().parse().unwrap();
            let max: usize = c.attributes.get("max").unwrap().parse().unwrap();
            min <= col_idx && col_idx <= max
        });

        if let Some(col) = col {
            col.attributes
                .insert("width".to_string(), width.to_string());
        } else {
            let mut new_col = crate::xlsx::xml::XmlElement::new("col");
            new_col.attributes.insert("min".to_string(), col_idx.to_string());
            new_col.attributes.insert("max".to_string(), col_idx.to_string());
            new_col.attributes.insert("width".to_string(), width.to_string());
            new_col.attributes.insert("customWidth".to_string(), "1".to_string());
            cols.children.push(new_col);
        }
    }

    pub fn merge_cells(&self, range: &str) {
        let mut xml = self.xml.lock().unwrap();
        let worksheet = xml
            .elements
            .iter_mut()
            .find(|e| e.name == "worksheet")
            .unwrap();

        let merge_cells = worksheet
            .children
            .iter_mut()
            .find(|e| e.name == "mergeCells");

        if merge_cells.is_none() {
            let new_merge_cells = crate::xlsx::xml::XmlElement::new("mergeCells");
            worksheet.children.push(new_merge_cells);
        }

        let merge_cells = worksheet
            .children
            .iter_mut()
            .find(|e| e.name == "mergeCells")
            .unwrap();

        let mut new_merge_cell = crate::xlsx::xml::XmlElement::new("mergeCell");
        new_merge_cell
            .attributes
            .insert("ref".to_string(), range.to_string());
        merge_cells.children.push(new_merge_cell);
    }

    pub fn unmerge_cells(&self, range: &str) {
        let mut xml = self.xml.lock().unwrap();
        let worksheet = xml
            .elements
            .iter_mut()
            .find(|e| e.name == "worksheet")
            .unwrap();

        let merge_cells = worksheet
            .children
            .iter_mut()
            .find(|e| e.name == "mergeCells");

        if let Some(merge_cells) = merge_cells {
            merge_cells
                .children
                .retain(|mc| mc.attributes.get("ref").unwrap() != range);
        }
    }

    pub fn freeze_panes(&self, cell_address: &str) {
        let mut xml = self.xml.lock().unwrap();
        let worksheet = xml
            .elements
            .iter_mut()
            .find(|e| e.name == "worksheet")
            .unwrap();

        let sheet_views = worksheet
            .children
            .iter_mut()
            .find(|e| e.name == "sheetViews")
            .unwrap();

        let sheet_view = sheet_views
            .children
            .iter_mut()
            .find(|e| e.name == "sheetView")
            .unwrap();

        let pane = sheet_view.children.iter_mut().find(|e| e.name == "pane");

        if pane.is_none() {
            let new_pane = crate::xlsx::xml::XmlElement::new("pane");
            sheet_view.children.push(new_pane);
        }

        let (col_str, row_idx) = Self::split_address(cell_address);
        let col_idx = Self::col_str_to_num(&col_str);

        let pane = sheet_view.children.iter_mut().find(|e| e.name == "pane").unwrap();
        if row_idx > 1 {
            pane.attributes.insert("ySplit".to_string(), (row_idx - 1).to_string());
        }
        if col_idx > 1 {
            pane.attributes.insert("xSplit".to_string(), (col_idx - 1).to_string());
        }
        pane.attributes.insert("topLeftCell".to_string(), cell_address.to_string());
        pane.attributes.insert("activePane".to_string(), "bottomRight".to_string());
        pane.attributes.insert("state".to_string(), "frozen".to_string());
    }
}

impl Sheet {
    fn split_address(address: &str) -> (String, usize) {
        let mut col_str = String::new();
        let mut row_str = String::new();
        for c in address.chars() {
            if c.is_alphabetic() {
                col_str.push(c);
            } else {
                row_str.push(c);
            }
        }
        (col_str, row_str.parse().unwrap())
    }

    fn col_str_to_num(col_str: &str) -> usize {
        let mut col_num = 0;
        for c in col_str.chars() {
            col_num = col_num * 26 + (c as usize - 'A' as usize + 1);
        }
        col_num
    }

    fn col_num_to_str(col_num: usize) -> String {
        let mut col_str = String::new();
        let mut col_num = col_num;
        while col_num > 0 {
            let remainder = (col_num - 1) % 26;
            col_str.insert(0, (b'A' + remainder as u8) as char);
            col_num = (col_num - 1) / 26;
        }
        col_str
    }

    fn coordinate_to_string(row: usize, col: usize) -> String {
        let col_str = Self::col_num_to_str(col);
        format!("{col_str}{row}")
    }

    pub fn new(name: String, xml: Arc<Mutex<Xml>>, shared_strings: Arc<Mutex<Xml>>) -> Self {
        Sheet {
            name,
            xml,
            shared_strings,
        }
    }

    // pub fn get_name(&self) -> &str {
    //     &self.name
    // }
}
