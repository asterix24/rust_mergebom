use calamine::{open_workbook_auto, DataType, Reader, Sheets};
use lazy_static::lazy_static;
use regex::Regex;
use std::collections::{hash_map::Entry, HashMap};

use super::utils::{convert_comment_to_value, detect_measure_unit, guess_category};

#[derive(Default, Debug, Clone)]
struct ExtraCol {
    label: String,
    value: String,
}

#[derive(Default, Debug, Clone)]
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
}

pub struct DataParser {
    header_map: Vec<HeaderMap>,
    workbook: Sheets,
}

impl DataParser {
    pub fn new(filename: &str) -> DataParser {
        println!("Parse: {}", filename);
        let workbook = match open_workbook_auto(filename) {
            Ok(wk) => wk,
            Err(error) => panic!("Error while parsing file: {:?}", error),
        };

        DataParser {
            header_map: Vec::new(),
            workbook: workbook,
        }
    }
    pub fn headers(mut self) -> Vec<HeaderMap> {
        self.find_header();
        self.header_map.clone()
    }
    pub fn xlsx(mut self) -> Vec<Item> {
        self.find_header();
        let data = self.parse_xlsx();
        self.sets(&data)
    }
    pub fn find_header(&mut self) {
        lazy_static! {
            static ref RE: Regex = Regex::new(r"[NOTE|CODE]\s(.*)").unwrap();
        }

        /* Search headers in source files */
        if let Some(Ok(range)) = self.workbook.worksheet_range("Sheet1") {
            let (rw, cl) = range.get_size();
            for row in 0..rw {
                for column in 0..cl {
                    if let Some(DataType::String(s)) = range.get((row, column)) {
                        match s.to_lowercase().as_str() {
                            "quantity" | "designator" | "comment" | "footprint" | "description"
                            | "mounttechnology" | "mount_technology" => {
                                self.header_map.push(HeaderMap {
                                    key: String::from(s.to_lowercase()),
                                    index: column,
                                })
                            }
                            _ => {
                                // println!("{:?}", s);
                                match RE.captures(s.as_ref()) {
                                    Some(cc) => {
                                        if let Some(m) =
                                            cc.get(1).map_or(None, |m| Some(m.as_str()))
                                        {
                                            self.header_map.push(HeaderMap {
                                                key: format!("{:}", m),
                                                index: column,
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
        println!("Trovato: {:?}", self.header_map);
    }

    pub fn parse_xlsx(&mut self) -> Vec<Item> {
        let mut items: Vec<Item> = Vec::new();

        /* Find data in source with column position find above */
        if let Some(Ok(range)) = self.workbook.worksheet_range("Sheet1") {
            let (rw, _) = range.get_size();
            for row in 0..rw {
                let mut template = Item::default();
                let mut skip_row = false;
                for header_label in &self.header_map {
                    // this row contain a header, so we should skip it.
                    if let Some(DataType::String(value)) = range.get((row, header_label.index)) {
                        match header_label.key.to_lowercase().as_str() {
                            "designator" => {
                                if value.to_lowercase() == header_label.key || value.is_empty() {
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

    pub fn sets(&mut self, data: &Vec<Item>) -> Vec<Item> {
        let mut items: Vec<Item> = Vec::new();
        let mut item_sets: HashMap<String, Vec<Item>> = HashMap::new();
        for row in data {
            match item_sets.entry(row.unique_id.clone()) {
                Entry::Occupied(mut o) => {
                    o.get_mut().push(row.clone());
                }
                Entry::Vacant(v) => {
                    v.insert(vec![row.clone(); 1]);
                }
            };
        }

        for (k, v) in item_sets {
            let template: Item = Item::default();
            println!("{}", k);
            for d in v {
                println!("\t{:?} {:?} {:?}", d.category, d.designator, d.extra);
            }
            items.push(template)
        }
        return items;
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
}
