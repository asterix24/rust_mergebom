use calamine::{open_workbook_auto, DataType, Reader, Sheets};
use lazy_static::lazy_static;
use regex::Regex;

use std::fmt;

use super::utils::{convert_comment_to_value, detect_measure_unit, guess_category};

#[derive(Debug, Clone)]
pub struct Stats {
    pub label: Category,
    pub value: usize,
}
#[derive(Debug, Clone)]
pub struct ExtraCol {
    pub label: Header,
    pub value: String,
}
#[derive(Debug, PartialEq, PartialOrd, Clone, Eq, Ord)]
pub enum Category {
    Connectors,
    Mechanicals,
    Fuses,
    Resistors,
    Capacitors,
    Diode,
    Inductors,
    Transistor,
    Transformes,
    Cristal,
    IC,
    IVALID,
}

#[derive(Debug, PartialEq, PartialOrd, Clone, Eq, Ord, Copy)]
pub enum Header {
    Quantity,
    Designator,
    Comment,
    Footprint,
    Description,
    MountTecnology,
    Layer,
    Extra,
}

#[derive(Debug)]
pub struct Item {
    unique_id: String,
    pub category: Category,
    pub base_exp: (f32, i32),
    pub measure_unit: String,
    pub designator: Vec<String>,
    pub comment: String,
    pub footprint: String,
    pub description: String,
    pub layer: Vec<String>,
    pub extra: Vec<ExtraCol>,
}

#[derive(PartialEq, Eq, PartialOrd, Ord)]
pub struct HeaderMap {
    pub key: Header,
    pub label: String,
    pub index: usize,
}

impl fmt::Debug for HeaderMap {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        write!(f, "{:?}-> {} {}\n", self.key, self.label, self.index)
    }
}

pub struct DataParser {
    workbook: Sheets,
    sheet_name: String,
}

impl DataParser {
    pub fn new(filename: &str) -> DataParser {
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
        DataParser {
            workbook,
            sheet_name,
        }
    }

    pub fn parse(mut self, header_map: &Vec<HeaderMap>) -> Vec<Item> {
        let data = self.parse_xlsx(&header_map);
        self.sets(data)
    }

    pub fn headers(&mut self) -> Vec<HeaderMap> {
        let mut header_map: Vec<HeaderMap> = Vec::new();
        lazy_static! {
            static ref RE_NOTE: Regex = Regex::new(r"NOTE\s(.*)").unwrap();
            static ref RE_CODE: Regex = Regex::new(r"CODE\s(.*)").unwrap();
        }

        match self.workbook.worksheet_range(self.sheet_name.as_str()) {
            Some(Ok(range)) => {
                let (rw, cl) = range.get_size();
                let mut header_found = false;
                for row in 0..rw {
                    if header_found {
                        break;
                    }
                    for column in 0..cl {
                        if let Some(DataType::String(s)) = range.get((row, column)) {
                            match s.to_lowercase().as_str() {
                                "designator" => {
                                    header_found = true;
                                    header_map.push(HeaderMap {
                                        key: Header::Designator,
                                        label: String::from("Designator"),
                                        index: column,
                                    })
                                }
                                "comment" => header_map.push(HeaderMap {
                                    key: Header::Comment,
                                    label: String::from("Comment"),
                                    index: column,
                                }),
                                "footprint" => header_map.push(HeaderMap {
                                    key: Header::Footprint,
                                    label: String::from("Footprint"),
                                    index: column,
                                }),
                                "description" => header_map.push(HeaderMap {
                                    key: Header::Description,
                                    label: String::from("Description"),
                                    index: column,
                                }),
                                "mounttechnology" | "mount_technology" => {
                                    header_map.push(HeaderMap {
                                        key: Header::MountTecnology,
                                        label: String::from("Mount Technology"),
                                        index: column,
                                    })
                                }
                                "layer" => header_map.push(HeaderMap {
                                    key: Header::Layer,
                                    label: String::from("Layer"),
                                    index: column,
                                }),
                                _ => {
                                    match RE_CODE.captures(s.as_ref()) {
                                        Some(cc) => {
                                            if let Some(m) =
                                                cc.get(1).map_or(None, |m| Some(m.as_str()))
                                            {
                                                header_map.push(HeaderMap {
                                                    key: Header::Extra,
                                                    index: column,
                                                    label: format!("Code {:}", m),
                                                })
                                            }
                                        }
                                        _ => (),
                                    }
                                    match RE_NOTE.captures(s.as_ref()) {
                                        Some(cc) => {
                                            if let Some(m) =
                                                cc.get(1).map_or(None, |m| Some(m.as_str()))
                                            {
                                                header_map.push(HeaderMap {
                                                    key: Header::Extra,
                                                    index: column,
                                                    label: format!("Note {:}", m),
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
            _ => panic!("peggio.."),
        }
        header_map.sort_by_key(|m| m.key);
        println!("{:?}", header_map);
        header_map
    }

    pub fn parse_xlsx(&mut self, header_map: &Vec<HeaderMap>) -> Vec<Item> {
        let mut items: Vec<Item> = Vec::new();

        /* Find data in source with column position find above */
        if let Some(Ok(range)) = self.workbook.worksheet_range(self.sheet_name.as_str()) {
            let (rw, _) = range.get_size();
            for row in 0..rw {
                let mut template = Item {
                    unique_id: String::new(),
                    category: Category::IVALID,
                    base_exp: (0.0, 0),
                    measure_unit: String::new(),
                    designator: vec![],
                    comment: String::new(),
                    footprint: String::new(),
                    description: String::new(),
                    layer: vec![],
                    extra: vec![],
                };
                let mut skip_row = false;
                for header_label in header_map {
                    // this row contain a header, so we should skip it.
                    match range.get((row, header_label.index)) {
                        Some(DataType::String(value)) => match header_label.key {
                            Header::Designator => {
                                if value == "Designator" || value.is_empty() {
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
                            Header::Comment => {
                                template.comment = value.clone();
                                template.base_exp = convert_comment_to_value(value);
                            }
                            Header::Description => {
                                template.description = value.clone();
                            }
                            Header::Footprint => {
                                template.footprint = value.clone();
                            }
                            Header::Layer | Header::MountTecnology => {
                                template.layer.push(value.clone());
                            }
                            _ => {
                                template.extra.push(ExtraCol {
                                    label: header_label.key.clone(),
                                    value: value.clone(),
                                });
                            }
                        },
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
                        Some(DataType::Empty) => (),
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

                    let key = match template.category {
                        Category::Connectors => {
                            format!("{}{}{}", template.footprint, template.description, ext_str)
                        }
                        _ => format!(
                            "{}{}{}{}",
                            template.comment, template.footprint, template.description, ext_str
                        ),
                    };
                    template.unique_id = key;
                    items.push(template);
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

                    //TODO: Merge all columns
                }
                _ => items.push(Item { ..row }),
            }
        }
        return items;
    }
}

pub fn categories(data: &Vec<Item>) -> Vec<Category> {
    let mut cat: Vec<Category> = Vec::new();
    for c in data {
        if !cat.contains(&c.category) {
            cat.push(c.category.clone());
        }
    }
    cat.sort();
    cat
}

pub fn stats(data: &Vec<Item>) -> Vec<Stats> {
    data.iter().fold(Vec::<Stats>::new(), |mut acc, i| {
        let mut is_new = true;
        for mut x in acc.iter_mut() {
            if x.label == i.category {
                x.value += 1;
                is_new = false;
            }
        }
        if is_new {
            acc.push(Stats {
                label: i.category.clone(),
                value: 1,
            });
        }
        acc
    })
}

impl PartialEq for Item {
    fn eq(&self, other: &Self) -> bool {
        self.unique_id == other.unique_id
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn test_find_headers() {
        let boms = [
            "test_data/test0.xlsx",
            "test_data/test1.xlsx",
            "test_data/test2.xlsx",
        ];

        let header_map_check = (
            [
                (Header::Designator, "Designator", 3),
                (Header::Comment, "Comment", 1),
                (Header::Footprint, "Footprint", 4),
                (Header::Description, "Description", 2),
            ],
            [
                (Header::Designator, "Designator", 1),
                (Header::Comment, "Comment", 2),
                (Header::Footprint, "Footprint", 3),
                (Header::Description, "Description", 4),
                (Header::Extra, "Code farnell", 5),
                (Header::Extra, "Note produzione", 6),
                (Header::Extra, "Code digikey", 7),
            ],
            [
                (Header::Designator, "Designator", 5),
                (Header::Comment, "Comment", 2),
                (Header::Footprint, "Footprint", 3),
                (Header::Description, "Description", 0),
            ],
        );

        let mut data: DataParser = DataParser::new(boms[0]);
        let hdr_map: Vec<HeaderMap> = data.headers();
        assert_eq!(hdr_map.len(), header_map_check.0.len());
        for (n, i) in hdr_map.iter().enumerate() {
            assert_eq!(i.key, header_map_check.0[n].0);
            assert_eq!(i.label, header_map_check.0[n].1);
            assert_eq!(i.index, header_map_check.0[n].2);
        }

        let mut data: DataParser = DataParser::new(boms[1]);
        let hdr_map: Vec<HeaderMap> = data.headers();
        assert_eq!(hdr_map.len(), header_map_check.1.len());
        for (n, i) in hdr_map.iter().enumerate() {
            assert_eq!(i.key, header_map_check.1[n].0);
            assert_eq!(i.label, header_map_check.1[n].1);
            assert_eq!(i.index, header_map_check.1[n].2);
        }

        let mut data: DataParser = DataParser::new(boms[2]);
        let hdr_map: Vec<HeaderMap> = data.headers();
        assert_eq!(hdr_map.len(), header_map_check.2.len());
        for (n, i) in hdr_map.iter().enumerate() {
            assert_eq!(i.key, header_map_check.2[n].0);
            assert_eq!(i.label, header_map_check.2[n].1);
            assert_eq!(i.index, header_map_check.2[n].2);
        }
    }
}
