pub mod items;
pub mod utils;

use items::*;
use lazy_static::lazy_static;
use regex::Regex;
use utils::*;

use calamine::{open_workbook, DataType, Reader, Xlsx};

const CATEGORY: [&'static str; 11] = [
    "connectors",
    "mechanicals",
    "fuses",
    "resistors",
    "capacitors",
    "diode",
    "inductors",
    "transistor",
    "transformes",
    "cristal",
    "ic",
];

pub fn find_header(bomfile: &str) -> Vec<HeaderMap> {
    lazy_static! {
        static ref RE: Regex = Regex::new(r"NOTE|CODE (.*)").unwrap();
    }

    println!("Parse: {}", bomfile);
    let mut workbook: Xlsx<_> = match open_workbook(bomfile) {
        Ok(wk) => wk,
        Err(error) => panic!("Error while parsing file: {:?}", error),
    };

    let mut header_map: Vec<HeaderMap> = Vec::new();
    /* Search headers in source files */
    if let Some(Ok(range)) = workbook.worksheet_range("Sheet1") {
        let (rw, cl) = range.get_size();
        for row in 0..rw {
            for column in 0..cl {
                if let Some(DataType::String(s)) = range.get((row, column)) {
                    match s.to_lowercase().as_str() {
                        "quantity" | "designator" | "comment" | "footprint" | "description"
                        | "layer" | "mounttechnoloy" | "mounting_technoloy" => {
                            header_map.push(HeaderMap {
                                key: String::from(s.to_lowercase()),
                                index: column,
                            })
                        }
                        _ => match RE.captures(s.as_ref()) {
                            Some(cc) => match cc.get(1).map_or(None, |m| Some(m.as_str())) {
                                None => (),
                                Some(m) => header_map.push(HeaderMap {
                                    key: format!("{:}", m),
                                    index: column,
                                }),
                            },
                            None => (),
                        },
                    }
                }
            }
        }
    }
    println!("Trovato: {:?}", header_map);
    return header_map;
}

pub fn parse_xlsx(bomfile: &str, heade_map: &Vec<HeaderMap>) -> Vec<Item> {
    let mut items: Vec<Item> = Vec::new();

    println!("Parse: {}", bomfile);
    let mut workbook: Xlsx<_> = match open_workbook(bomfile) {
        Ok(wk) => wk,
        Err(error) => panic!("Error while parsing file: {:?}", error),
    };

    /* Find data in source with column position find above */
    if let Some(Ok(range)) = workbook.worksheet_range("Sheet1") {
        let (rw, _) = range.get_size();
        for row in 0..rw {
            let mut template = Item::default();
            let mut skip_row = false;
            for header_label in heade_map {
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
                                index: header_label.index,
                            });
                        }
                        _ => {
                            template.extra.push(ExtraCol {
                                label: header_label.key.clone(),
                                value: value.clone(),
                                index: header_label.index,
                            });
                        }
                    }
                }
            }
            if !skip_row {
                items.push(template);
            }
        }
    }
    return items;
}

fn report() {
    //     headers: Vec<String>,
    //     data: HashMap<String, Vec<HashMap<String, Vec<Item>>>>,
    //     merge_file: &str,
    // ) {
    //     let workbook = Workbook::new(merge_file);
    //     let default_fmt = workbook.add_format().set_font_size(10.0).set_text_wrap();

    //     let hdr_fmt = workbook
    //         .add_format()
    //         .set_bg_color(FormatColor::Cyan)
    //         .set_bold()
    //         .set_font_size(12.0);
    //     let merge_fmt = workbook
    //         .add_format()
    //         .set_bg_color(FormatColor::Yellow)
    //         .set_bold()
    //         .set_border(FormatBorder::Thin)
    //         .set_align(FormatAlignment::CenterAcross);

    //     let tot_fmt = workbook
    //         .add_format()
    //         .set_bg_color(FormatColor::Lime)
    //         .set_bold()
    //         .set_font_size(12.0);

    //     let mut sheet1 = match workbook.add_worksheet(None) {
    //         Ok(wk) => wk,
    //         _ => panic!("Unable to open wk"),
    //     };

    //     let mut row_curr: u32 = 10;
    //     let mut col: u16 = 0;
    //     for label in HEADERS.iter() {
    //         match sheet1.write_string(row_curr, col, label, Some(&hdr_fmt)) {
    //             Ok(m) => m,
    //             _ => panic!("Error!"),
    //         };
    //         col += 1;
    //     }
    //     row_curr += 1;
    //     for (category, items) in data {
    //         println!(">>>> {:?}", category);
    //         match sheet1.merge_range(
    //             row_curr,
    //             0,
    //             row_curr,
    //             HEADERS.len() as u16,
    //             category.as_str(),
    //             Some(&merge_fmt),
    //         ) {
    //             Ok(m) => m,
    //             _ => panic!("Error!"),
    //         };

    //         row_curr += 1;
    //         //let sorted_item = items.sort_by_key(|k| Ord(pow(k.base_exp[0], k.base_exp[1])));
    //         for unique_row in items {
    //             for (unique_key, parts) in unique_row {
    //                 //println!("{:?} {:?}", unique_key, parts.len());
    //                 let row_elem = match parts.first() {
    //                     Some(r) => r,
    //                     _ => panic!("empty part list"),
    //                 };
    //                 for label in headers.iter() {
    //                     match label.as_str() {
    //                         "quantity" => {
    //                             //println!("{:?}", parts.len());
    //                             match sheet1.write_string(
    //                                 row_curr,
    //                                 0,
    //                                 parts.len().to_string().as_str(),
    //                                 Some(&tot_fmt),
    //                             ) {
    //                                 Ok(m) => m,
    //                                 _ => panic!("Error!"),
    //                             };
    //                         }
    //                         "designator" => {
    //                             let mut ss: Vec<String> = Vec::new();
    //                             for s in parts.iter() {
    //                                 ss.push(s.designator.clone());
    //                             }
    //                             match sheet1.write_string(
    //                                 row_curr,
    //                                 1,
    //                                 ss.join(", ").as_str(),
    //                                 Some(&default_fmt),
    //                             ) {
    //                                 Ok(m) => m,
    //                                 _ => panic!("Error!"),
    //                             };
    //                         }
    //                         "comment" => {
    //                             match sheet1.write_string(
    //                                 row_curr,
    //                                 2,
    //                                 row_elem.comment.as_str(),
    //                                 Some(&default_fmt),
    //                             ) {
    //                                 Ok(m) => m,
    //                                 _ => panic!("Error!"),
    //                             };
    //                         }
    //                         "footprint" => {
    //                             match sheet1.write_string(
    //                                 row_curr,
    //                                 3,
    //                                 row_elem.footprint.as_str(),
    //                                 Some(&default_fmt),
    //                             ) {
    //                                 Ok(m) => m,
    //                                 _ => panic!("Error!"),
    //                             };
    //                         }
    //                         "description" => {
    //                             match sheet1.write_string(
    //                                 row_curr,
    //                                 4,
    //                                 row_elem.description.as_str(),
    //                                 Some(&default_fmt),
    //                             ) {
    //                                 Ok(m) => m,
    //                                 _ => panic!("Error!"),
    //                             };
    //                         }
    //                         _ => println!("Invalid category"),
    //                     }
    //                 }
    //                 row_curr += 1;
    //             }
    //         }
    //     }
    //     match workbook.close() {
    //         Ok(m) => m,
    //         _ => panic!("Error!"),
    //     };
    // }
}

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
                key: String::from("layer"),
                index: 7,
            },
            HeaderMap {
                key: String::from("farnell"),
                index: 11,
            },
            HeaderMap {
                key: String::from("prova"),
                index: 12,
            },
            HeaderMap {
                key: String::from("sdfk"),
                index: 14,
            },
        ];
        let header_map = find_header("test_data/bom0.xlsx");
        assert_eq!(header_map.len(), header_map_check.len());
        for (n, i) in header_map.iter().enumerate() {
            assert_eq!(i.key, header_map_check.get(n).unwrap().key);
            assert_eq!(i.index, header_map_check.get(n).unwrap().index);
        }
    }
}
