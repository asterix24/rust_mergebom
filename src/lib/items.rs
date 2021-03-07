use calamine::{open_workbook_auto, DataType, Reader, Sheets};
use lazy_static::lazy_static;
use regex::Regex;

use super::utils::{convert_comment_to_value, detect_measure_unit, guess_category};

#[derive(Default, Debug, Clone)]
struct ExtraCol {
    label: String,
    value: String,
}

#[derive(Default, Clone)]
pub struct Item {
    unique_id: String,
    category: String,
    base_exp: (f32, i32),
    fmt_value: String,
    measure_unit: String,
    designator: Vec<String>,
    comment: String,
    footprint: String,
    description: String,
    extra: Vec<ExtraCol>,
}

#[derive(Default, Debug, Clone)]
pub struct HeaderMap {
    key: String,
    index: usize,
    order: usize,
}
pub struct DataParser {
    workbook: Sheets,
}

impl DataParser {
    pub fn new(filename: &str) -> DataParser {
        println!("Parse: {}", filename);
        let workbook = match open_workbook_auto(filename) {
            Ok(wk) => wk,
            Err(error) => panic!("Error while parsing file: {:?}", error),
        };

        DataParser { workbook: workbook }
    }

    pub fn parse(mut self, header_map: &Vec<HeaderMap>) -> Vec<Item> {
        let data = self.parse_xlsx(&header_map);
        self.sets(data)
    }

    pub fn headers(&mut self) -> Vec<HeaderMap> {
        let mut header_map: Vec<HeaderMap> = Vec::new();
        lazy_static! {
            static ref RE: Regex = Regex::new(r"[NOTE|CODE]\s(.*)").unwrap();
        }

        /* Search headers in source files */
        if let Some(Ok(range)) = self.workbook.worksheet_range("Sheet1") {
            let (rw, cl) = range.get_size();
            let extra_col: usize = 0;
            for row in 0..rw {
                for column in 0..cl {
                    if let Some(DataType::String(s)) = range.get((row, column)) {
                        match s.to_lowercase().as_str() {
                            "designator" => header_map.push(HeaderMap {
                                key: String::from("Designator"),
                                index: column,
                                order: 0,
                            }),
                            "comment" => header_map.push(HeaderMap {
                                key: String::from("Comment"),
                                index: column,
                                order: 1,
                            }),
                            "footprint" => header_map.push(HeaderMap {
                                key: String::from("Footprint"),
                                index: column,
                                order: 2,
                            }),
                            "description" => header_map.push(HeaderMap {
                                key: String::from("Description"),
                                index: column,
                                order: 3,
                            }),
                            "mounttechnology" | "mount_technology" => header_map.push(HeaderMap {
                                key: String::from("Mount Tecnology"),
                                index: column,
                                order: 4,
                            }),
                            _ => {
                                // println!("{:?}", s);
                                match RE.captures(s.as_ref()) {
                                    Some(cc) => {
                                        if let Some(m) =
                                            cc.get(1).map_or(None, |m| Some(m.as_str()))
                                        {
                                            header_map.push(HeaderMap {
                                                key: format!("Alt. {:}", m),
                                                index: column,
                                                order: 4 + extra_col,
                                            })
                                        }
                                    }
                                    _ => (),
                                }
                            }
                        }
                    }
                }
            }
        }
        println!("Trovato: {:?}", header_map);
        header_map
    }

    pub fn parse_xlsx(&mut self, header_map: &Vec<HeaderMap>) -> Vec<Item> {
        let mut items: Vec<Item> = Vec::new();

        /* Find data in source with column position find above */
        if let Some(Ok(range)) = self.workbook.worksheet_range("Sheet1") {
            let (rw, _) = range.get_size();
            for row in 0..rw {
                let mut template = Item::default();
                let mut skip_row = false;
                for header_label in header_map {
                    // this row contain a header, so we should skip it.
                    match range.get((row, header_label.index)) {
                        Some(DataType::String(value)) => {
                            match header_label.key.to_lowercase().as_str() {
                                "designator" => {
                                    if *value == header_label.key || value.is_empty() {
                                        println!("skip: [{:?}]", value);
                                        skip_row = true;
                                        continue;
                                    }
                                    template.designator = value
                                        .split(",")
                                        .map(|m| m.trim().to_string())
                                        .collect::<Vec<_>>();

                                    let des = template.designator.first().unwrap();
                                    template.category = guess_category(des.trim());
                                    template.measure_unit = detect_measure_unit(des.trim());
                                }
                                "comment" => {
                                    template.comment = value.clone();
                                    template.base_exp = convert_comment_to_value(value);
                                }
                                "description" => {
                                    template.description = value.clone();
                                }
                                "footprint" => {
                                    template.footprint = value.clone();
                                }
                                "layer" | "mounttechnoloy" | "mounting_technoloy" => {
                                    template.extra.push(ExtraCol {
                                        label: header_label.key.clone(),
                                        value: value.to_uppercase(),
                                    });
                                }
                                _ => {
                                    template.extra.push(ExtraCol {
                                        label: header_label.key.clone(),
                                        value: value.clone(),
                                    });
                                }
                            }
                        }
                        Some(DataType::Int(value)) => {
                            template.extra.push(ExtraCol {
                                label: header_label.key.clone(),
                                value: value.to_string(),
                            });
                        }

                        Some(DataType::Float(value)) => {
                            template.extra.push(ExtraCol {
                                label: header_label.key.clone(),
                                value: value.to_string(),
                            });
                        }
                        Some(DataType::Empty) => (), //println!("Empty cell.. skip"),
                        _ => println!(
                            "Invalid data type..[{:?}]",
                            range.get((row, header_label.index))
                        ),
                    }
                }
                if !skip_row {
                    let mut ext_str: String = String::new();
                    for ext in template.extra.clone() {
                        ext_str = format!("{}{}", ext_str, ext.value);
                    }

                    let key = match template.category.as_str() {
                        "connectors" | "diode" => {
                            format!("{}{}{}", template.footprint, template.description, ext_str)
                        }
                        _ => format!(
                            "{}{}{}{}",
                            template.comment, template.footprint, template.description, ext_str
                        ),
                    };
                    template.unique_id = key;
                    items.push(template.clone());
                }
            }
        }
        return items;
    }

    pub fn sets(self, data: Vec<Item>) -> Vec<Item> {
        let mut items: Vec<Item> = Vec::new();
        for row in data {
            match items.iter().position(|m| m.unique_id == row.unique_id) {
                Some(cc) => {
                    let mut des: Vec<String> = Vec::new();
                    des.append(&mut items[cc].designator.clone());
                    des.append(&mut row.designator.clone());
                    items[cc].designator = des;
                    println!("{:?}", items[cc].designator);
                }
                _ => items.push(Item { ..row.clone() }),
            }
        }
        return items;
    }
}

pub fn headers_to_str(header_map: &Vec<HeaderMap>) -> Vec<String> {
    let mut hdr: Vec<String> = Vec::new();
    header_map.clone().sort_by_key(|x| x.order);
    for i in header_map.clone() {
        hdr.push(i.key);
    }
    hdr
}

pub fn categories(data: &Vec<Item>) -> Vec<String> {
    let mut cat: Vec<String> = Vec::new();
    for c in data {
        if !cat.contains(&c.category) {
            cat.push(c.category.clone());
        }
    }
    cat
}
pub fn dump(data: &Vec<Item>) {
    for i in data {
        println!("{:?}", i);
    }
}

impl PartialEq for Item {
    fn eq(&self, other: &Self) -> bool {
        self.unique_id == other.unique_id
    }
}

impl std::fmt::Debug for Item {
    fn fmt(&self, f: &mut std::fmt::Formatter) -> std::fmt::Result {
        write!(
            f,
            "Item:\n\tunique_id: {}\n\tcategory: {}\n\tbase_exp: {:?}\n\tfmt_value: {}\n\tmeasure_unit: {}\n\tdesignator: {:?}\n\tcomment:{}\n\tfootprint:{}\n\tdescription:{}\n\textra: {:?}",
            self.unique_id,
            self.category,
            self.base_exp,
            self.fmt_value,
            self.measure_unit,
            self.designator,
            self.comment,
            self.footprint,
            self.description,
            self.extra
        )
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn test_header_map() {
        let header_map_check: Vec<HeaderMap> = vec![
            HeaderMap {
                key: String::from("quantity"),
                index: 0,
            },
            HeaderMap {
                key: String::from("designator"),
                index: 1,
            },
            HeaderMap {
                key: String::from("comment"),
                index: 2,
            },
            HeaderMap {
                key: String::from("footprint"),
                index: 3,
            },
            HeaderMap {
                key: String::from("description"),
                index: 4,
            },
            HeaderMap {
                key: String::from("mounttechnology"),
                index: 8,
            },
            HeaderMap {
                key: String::from("123"),
                index: 10,
            },
            HeaderMap {
                key: String::from("farnell"),
                index: 11,
            },
            HeaderMap {
                key: String::from("mouser"),
                index: 12,
            },
            HeaderMap {
                key: String::from("description"),
                index: 13,
            },
            HeaderMap {
                key: String::from("digikey"),
                index: 14,
            },
        ];
        let data: DataParser = DataParser::new("test_data/bom0.xlsx");
        let hdr_map: Vec<HeaderMap> = data.headers();
        assert_eq!(hdr_map.len(), header_map_check.len());
        for (n, i) in hdr_map.iter().enumerate() {
            assert_eq!(i.key, header_map_check.get(n).unwrap().key);
            assert_eq!(i.index, header_map_check.get(n).unwrap().index);
        }
    }

    #[test]
    fn test_extra_col() {}
    #[test]
    fn test_merge() {
        let len_check = vec![2, 2, 2, 2, 4, 1, 2];
        let data: DataParser = DataParser::new("test_data/bom_merge.xlsx");
        let items = data.collect();
        assert_eq!(len_check.len(), items.len());
        for (n, c) in items.iter().enumerate() {
            assert_eq!(c.designator.len(), len_check[n]);
            println!("{:?}", c);
        }
    }
}
