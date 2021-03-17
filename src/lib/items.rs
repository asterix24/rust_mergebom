use regex::Regex;

use std::fmt;

use super::load::Load;
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
        write!(f, "{:?}-> {} {}", self.key, self.label, self.index)
    }
}

pub struct DataParser {
    items: Vec<Item>,
    headers: Vec<HeaderMap>,
}

impl DataParser {
    pub fn new(mut load: Load) -> DataParser {
        let mut headers = Vec::new();

        let re_note = Regex::new(r"NOTE\s(.*)").unwrap();
        let re_code = Regex::new(r"CODE\s(.*)").unwrap();

        let mut header_found = false;
        for row in load.read() {
            for (n, col) in row.iter().enumerate() {
                match col.to_lowercase().as_str() {
                    "designator" => {
                        header_found = true;
                        headers.push(HeaderMap {
                            key: Header::Designator,
                            label: String::from("Designator"),
                            index: n,
                        })
                    }
                    "comment" => headers.push(HeaderMap {
                        key: Header::Comment,
                        label: String::from("Comment"),
                        index: n,
                    }),
                    "footprint" => headers.push(HeaderMap {
                        key: Header::Footprint,
                        label: String::from("Footprint"),
                        index: n,
                    }),
                    "description" => headers.push(HeaderMap {
                        key: Header::Description,
                        label: String::from("Description"),
                        index: n,
                    }),
                    "mounttechnology" | "mount_technology" => headers.push(HeaderMap {
                        key: Header::MountTecnology,
                        label: String::from("Mount Technology"),
                        index: n,
                    }),
                    "layer" => headers.push(HeaderMap {
                        key: Header::Layer,
                        label: String::from("Layer"),
                        index: n,
                    }),
                    _ => {
                        if let Some(cc) = re_code.captures(col.as_ref()) {
                            if let Some(m) = cc.get(1).map(|m| m.as_str()) {
                                headers.push(HeaderMap {
                                    key: Header::Extra,
                                    index: n,
                                    label: format!("Code {:}", m),
                                })
                            }
                        }
                        if let Some(cc) = re_note.captures(col.as_ref()) {
                            if let Some(m) = cc.get(1).map(|m| m.as_str()) {
                                headers.push(HeaderMap {
                                    key: Header::Extra,
                                    index: n,
                                    label: format!("Note {:}", m),
                                })
                            }
                        }
                    }
                }
            }

            if header_found {
                break;
            }
        }

        headers.sort_by_key(|m| m.key);
        println!("{:?}", headers);

        let data = Self::parse_data(&mut load, &headers);
        let items = Self::sets(data);

        DataParser { headers, items }
    }

    pub fn headers(&self) -> &[HeaderMap] {
        &self.headers
    }

    pub fn items(&self) -> &[Item] {
        &self.items
    }

    pub fn categories(&self) -> Vec<Category> {
        let mut cat: Vec<Category> = Vec::new();
        for c in &self.items {
            if !cat.contains(&c.category) {
                cat.push(c.category.clone());
            }
        }
        cat.sort();
        cat
    }

    pub fn stats(&self) -> Vec<Stats> {
        self.items.iter().fold(Vec::<Stats>::new(), |mut acc, i| {
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

    fn parse_data(load: &mut Load, headers: &[HeaderMap]) -> Vec<Item> {
        let mut items = Vec::new();

        for row in load.read() {
            /* Find data in source with column position find above */
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
            for header_label in headers {
                match row.get(header_label.index) {
                    Some(value) => match header_label.key {
                        Header::Designator => {
                            // this row contain a header, so we should skip it.
                            if value == "Designator" || value.is_empty() {
                                println!("skip: [{:?}]", value);
                                skip_row = true;
                                continue;
                            }
                            template.designator = value
                                .split(',')
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
                                label: header_label.key,
                                value: value.clone(),
                            });
                        }
                    },
                    _ => println!("Invalid data type..",),
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
        items
    }

    fn sets(data: Vec<Item>) -> Vec<Item> {
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

        items
    }
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

        let mut data: DataParser = DataParser::new(Load::new(boms[0]));
        let hdr_map: Vec<HeaderMap> = data.headers();
        assert_eq!(hdr_map.len(), header_map_check.0.len());
        for (n, i) in hdr_map.iter().enumerate() {
            assert_eq!(i.key, header_map_check.0[n].0);
            assert_eq!(i.label, header_map_check.0[n].1);
            assert_eq!(i.index, header_map_check.0[n].2);
        }

        let mut data: DataParser = DataParser::new(Load::new(boms[1]));
        let hdr_map: Vec<HeaderMap> = data.headers();
        assert_eq!(hdr_map.len(), header_map_check.1.len());
        for (n, i) in hdr_map.iter().enumerate() {
            assert_eq!(i.key, header_map_check.1[n].0);
            assert_eq!(i.label, header_map_check.1[n].1);
            assert_eq!(i.index, header_map_check.1[n].2);
        }

        let mut data: DataParser = DataParser::new(Load::new(boms[2]));
        let hdr_map: Vec<HeaderMap> = data.headers();
        assert_eq!(hdr_map.len(), header_map_check.2.len());
        for (n, i) in hdr_map.iter().enumerate() {
            assert_eq!(i.key, header_map_check.2[n].0);
            assert_eq!(i.label, header_map_check.2[n].1);
            assert_eq!(i.index, header_map_check.2[n].2);
        }
    }
}
