use bomerge::{convert_comment_to_value, detect_measure_unit};
use calamine::{open_workbook, DataType, Reader, Xlsx};
use clap::{App, Arg};
use lazy_static::lazy_static;
use regex::Regex;
use std::{collections::hash_map::Entry, u16};
use std::{collections::HashMap, str::FromStr};
use xlsxwriter::*;

#[derive(Default, Debug, Clone)]
pub struct Item {
    category: String,
    base_exp: (f32, i32),
    fmt_value: String,
    measure_unit: String,
    designator: String,
    comment: String,
    footprint: String,
    description: String,
}

#[derive(Default, Debug, Clone)]
struct ItemRow {
    quantity: String,
    designator: Vec<String>,
    comment: String,
    footprint: String,
    description: String,
}

const HEADERS: [&'static str; 5] = [
    "quantity",
    "designator",
    "comment",
    "footprint",
    "description",
];

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

fn guess_category<S: AsRef<str>>(designator: S) -> String {
    lazy_static! {
        static ref RE: Regex = Regex::new(r"^([a-zA-Z_]{1,3})").unwrap();
    }

    match RE.captures(designator.as_ref()) {
        None => String::from("ivalid"),
        Some(cc) => match cc.get(1).map_or("", |m| m.as_str()).as_ref() {
            "J" | "X" | "P" | "SIM" => String::from("connectors"),
            "S" | "SCR" | "SPA" | "BAT" | "BUZ" | "BT" | "B" | "SW" | "MP" | "K" => {
                String::from("mechanicals")
            }
            "F" | "FU" => String::from("fuses"),
            "R" | "RN" | "R_G" => String::from("resistors"),
            "C" | "CAP" => String::from("capacitors"),
            "D" | "DZ" => String::from("diode"),
            "L" => String::from("inductors"),
            "Q" => String::from("transistor"),
            "TR" => String::from("transformes"),
            "Y" => String::from("cristal"),
            "U" => String::from("ic"),
            _ => panic!("Invalid category"),
        },
    }
}

pub fn report(
    headers: Vec<String>,
    data: HashMap<String, Vec<HashMap<String, Vec<Item>>>>,
    merge_file: &str,
) {
    let workbook = Workbook::new(merge_file);
    let default_fmt = workbook.add_format().set_font_size(10.0).set_text_wrap();

    let hdr_fmt = workbook
        .add_format()
        .set_bg_color(FormatColor::Cyan)
        .set_bold()
        .set_font_size(12.0);
    let merge_fmt = workbook
        .add_format()
        .set_bg_color(FormatColor::Yellow)
        .set_bold()
        .set_border(FormatBorder::Thin)
        .set_align(FormatAlignment::CenterAcross);

    let tot_fmt = workbook
        .add_format()
        .set_bg_color(FormatColor::Lime)
        .set_bold()
        .set_font_size(12.0);

    let mut sheet1 = match workbook.add_worksheet(None) {
        Ok(wk) => wk,
        _ => panic!("Unable to open wk"),
    };

    let mut row_curr: u32 = 10;
    let mut col: u16 = 0;
    for label in HEADERS.iter() {
        match sheet1.write_string(row_curr, col, label, Some(&hdr_fmt)) {
            Ok(m) => m,
            _ => panic!("Error!"),
        };
        col += 1;
    }
    row_curr += 1;
    for (category, items) in data {
        println!(">>>> {:?}", category);
        match sheet1.merge_range(
            row_curr,
            0,
            row_curr,
            HEADERS.len() as u16,
            category.as_str(),
            Some(&merge_fmt),
        ) {
            Ok(m) => m,
            _ => panic!("Error!"),
        };

        row_curr += 1;
        //let sorted_item = items.sort_by_key(|k| Ord(pow(k.base_exp[0], k.base_exp[1])));
        for unique_row in items {
            for (unique_key, parts) in unique_row {
                //println!("{:?} {:?}", unique_key, parts.len());
                let row_elem = match parts.first() {
                    Some(r) => r,
                    _ => panic!("empty part list"),
                };
                for label in headers.iter() {
                    match label.as_str() {
                        "quantity" => {
                            //println!("{:?}", parts.len());
                            match sheet1.write_string(
                                row_curr,
                                0,
                                parts.len().to_string().as_str(),
                                Some(&tot_fmt),
                            ) {
                                Ok(m) => m,
                                _ => panic!("Error!"),
                            };
                        }
                        "designator" => {
                            let mut ss: Vec<String> = Vec::new();
                            for s in parts.iter() {
                                ss.push(s.designator.clone());
                            }
                            match sheet1.write_string(
                                row_curr,
                                1,
                                ss.join(", ").as_str(),
                                Some(&default_fmt),
                            ) {
                                Ok(m) => m,
                                _ => panic!("Error!"),
                            };
                        }
                        "comment" => {
                            match sheet1.write_string(
                                row_curr,
                                2,
                                row_elem.comment.as_str(),
                                Some(&default_fmt),
                            ) {
                                Ok(m) => m,
                                _ => panic!("Error!"),
                            };
                        }
                        "footprint" => {
                            match sheet1.write_string(
                                row_curr,
                                3,
                                row_elem.footprint.as_str(),
                                Some(&default_fmt),
                            ) {
                                Ok(m) => m,
                                _ => panic!("Error!"),
                            };
                        }
                        "description" => {
                            match sheet1.write_string(
                                row_curr,
                                4,
                                row_elem.description.as_str(),
                                Some(&default_fmt),
                            ) {
                                Ok(m) => m,
                                _ => panic!("Error!"),
                            };
                        }
                        _ => println!("Invalid category"),
                    }
                }
                row_curr += 1;
            }
        }
    }
    match workbook.close() {
        Ok(m) => m,
        _ => panic!("Error!"),
    };
}

fn main() {
    let matches = App::new("Rust MergeBom")
        .version("0.1.0")
        .author("Daniele Basile <asterix24@gmail.com>")
        .about("Pretty merger and formatter for Bill Of Materials.")
        .arg(
            Arg::with_name("BOMFile")
                .help("BOM to Merge")
                .required(true)
                .min_values(1),
        )
        .get_matches();

    let boms: Vec<&str> = matches.values_of("BOMFile").unwrap().collect();

    let mut header_map = HashMap::new();
    let mut items: Vec<Item> = Vec::new();

    for item in boms {
        println!("Parse: {}", item);
        let mut workbook: Xlsx<_> = match open_workbook(item) {
            Ok(wk) => wk,
            Err(error) => {
                println!("Unable read bom: {}", error);
                continue;
            }
        };

        /* Search headers in source files */
        if let Some(Ok(range)) = workbook.worksheet_range("Sheet1") {
            let (rw, cl) = range.get_size();
            for row in 0..rw {
                for column in 0..cl {
                    let d = range.get((row, column));
                    if let Some(DataType::String(s)) = d {
                        if HEADERS.contains(&s.to_lowercase().as_str()) {
                            println!("Trovato: {:?} >> {}", s, column);
                            header_map.insert(s.to_lowercase(), column);
                        }
                    }
                }
            }
        }

        /* Find data in source with column position find above */
        if let Some(Ok(range)) = workbook.worksheet_range("Sheet1") {
            let (rw, _) = range.get_size();
            for row in 0..rw {
                if let Some(&designator) = header_map.get("designator") {
                    let mut template = Item::default();

                    let mut skip_item = false;
                    for label in HEADERS.iter() {
                        match header_map.get(&label.to_string()) {
                            Some(&col) => {
                                if let Some(DataType::String(value)) = range.get((row, col)) {
                                    if value.to_lowercase() == *label {
                                        println!("skip header..");
                                        skip_item = true;
                                        continue;
                                    }
                                    match *label {
                                        "comment" => template.comment = value.to_string(),
                                        "footprint" => template.footprint = value.to_string(),
                                        "description" => template.description = value.to_string(),
                                        _ => println!("[{:?}] skip..", *label),
                                    }
                                }
                            }
                            _ => {}
                        }
                    }

                    if !skip_item {
                        items.append(
                            &mut range
                                .get((row, designator))
                                .unwrap()
                                .to_string()
                                .split(",")
                                .map(|designator| Item {
                                    designator: designator.trim().to_owned(),
                                    category: guess_category(designator.trim()),
                                    base_exp: convert_comment_to_value(&template.comment),
                                    measure_unit: detect_measure_unit(&template.designator.trim()),
                                    ..template.clone()
                                })
                                .collect::<Vec<_>>(),
                        );
                    }
                }
            }
        }
    }

    let mut grouped_items: HashMap<String, Vec<HashMap<String, Vec<Item>>>> = HashMap::new();
    for category in CATEGORY.iter() {
        let group = items.iter().filter(|n| n.category == *category);

        let mut item_sets: HashMap<String, Vec<Item>> = HashMap::new();
        for row in group {
            let key = match *category {
                "connectors" => format!("{}{}", row.footprint, row.description),
                _ => format!("{}{}{}", row.comment, row.footprint, row.description),
            };

            match item_sets.entry(key) {
                Entry::Occupied(mut o) => {
                    o.get_mut().push(row.clone());
                }
                Entry::Vacant(v) => {
                    v.insert(vec![row.clone(); 1]);
                }
            };
        }
        match grouped_items.entry(String::from(*category)) {
            Entry::Occupied(mut o) => {
                o.get_mut().push(item_sets);
            }
            Entry::Vacant(v) => {
                v.insert(vec![item_sets; 1]);
            }
        };

        header_map.insert(String::from("quantity"), 0);
        report(
            header_map.keys().cloned().collect(),
            grouped_items.clone(),
            "merged_bom.xlsx",
        );
    }
}
