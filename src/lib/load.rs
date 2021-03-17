use calamine::{open_workbook_auto, DataType, Reader, Sheets};

pub struct Load {
    workbook: Sheets,
    sheet_name: String,
}

impl Load {
    pub fn new(filename: &str) -> Load {
        println!("Parse: {}", filename);
        let sheet_name: String;
        let workbook = match open_workbook_auto(filename) {
            Ok(wk) => {
                /* Search headers in source files */
                sheet_name = match wk.sheet_names().first() {
                    Some(s) => s.clone(),
                    None => panic!("unable to get sheet names"),
                };
                wk
            }
            Err(error) => panic!("Error while parsing file: {:?}", error),
        };

        println!("Sheets: {}", sheet_name);
        Load {
            workbook,
            sheet_name,
        }
    }

    pub fn read(&mut self) -> Vec<Vec<String>> {
        let mut data: Vec<Vec<String>> = Vec::new();
        match self.workbook.worksheet_range(self.sheet_name.as_str()) {
            Some(Ok(range)) => {
                let (rw, cl) = range.get_size();
                for row in 0..rw {
                    let mut element: Vec<String> = Vec::new();
                    for column in 0..cl {
                        if let Some(DataType::String(s)) = range.get((row, column)) {
                            element.push(s.clone());
                        }
                    }
                    data.push(element);
                }
            }
            _ => panic!("peggio.."),
        }
        data
    }
}
