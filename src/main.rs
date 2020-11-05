use bomerge::{convert_comment_to_value, detect_measure_unit};
use calamine::{open_workbook, DataType, Reader, Xlsx};
use clap::{App, Arg};
use lazy_static::lazy_static;
use regex::Regex;
use std::collections::hash_map::Entry;
use std::collections::HashMap;

#[derive(Default, Debug, Clone)]
struct ItemBOM {
    category: String,
    value: f32,
    base_exp: (f32, i32),
    fmt_value: String,
    measure_unit: String,
    designator: String,
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
    let mut items: Vec<ItemBOM> = Vec::new();

    for item in boms {
        println!("Parse: {}", item);
        let mut workbook: Xlsx<_> = match open_workbook(item) {
            Ok(wk) => wk,
            Err(error) => {
                println!("Unable read bom: {}", error);
                continue;
            }
        };

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
        if let Some(Ok(range)) = workbook.worksheet_range("Sheet1") {
            let (rw, _) = range.get_size();
            for row in 0..rw {
                if let Some(&designator) = header_map.get("designator") {
                    let mut template = ItemBOM::default();

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
                                        _ => println!("skip.."),
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
                                .map(|designator| ItemBOM {
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

    // compute_value(&mut items);
    // for i in &items {
    //     println!("{:?}", i);
    // }
    // items.sort_by_key(|i| i.base_exp.1);

    let mut grouped_items: Vec<HashMap<String, Vec<ItemBOM>>> = Vec::new();
    for category in CATEGORY.iter() {
        let group = items.iter().filter(|n| n.category == *category);

        let mut item_sets: HashMap<String, Vec<ItemBOM>> = HashMap::new();
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
        grouped_items.push(item_sets);
        //println!("{} {:?}", category, group.size_hint());
    }

    for v in grouped_items {
        for (_, m) in v {
            let mut d: Vec<String> = Vec::new();
            for n in m.iter() {
                d.push(n.designator.clone());
            }

            println!("{} {:?} {}", m[0].category, m.len(), d.join(", "));
        }
    }
}

fn compute_value(items: &mut Vec<ItemBOM>) {
    for i in items {
        i.value = i.base_exp.0 * 10_f32.powf(i.base_exp.1 as f32);
    }
}
