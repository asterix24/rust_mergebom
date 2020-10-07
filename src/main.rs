use calamine::{open_workbook, DataType, Reader, Xlsx};
use clap::{App, Arg};
use lazy_static::lazy_static;
use num_traits::pow;
use regex::Regex;
use std::collections::HashMap;

#[derive(Default, Debug, Clone)]
struct ItemBOM {
    category: String,
    value: f64,
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

fn detect_measure_unit(comment: &str) -> String {
    lazy_static! {
        static ref RE: Regex = Regex::new(r"^([KkR|C|L|Y])").unwrap();
    }
    match RE.captures(comment.as_ref()) {
        None => String::from("unknow"),
        Some(cc) => match cc.get(1).map_or("", |m| m.as_str()).as_ref() {
            "K" | "k" | "R" => String::from("ohm"),
            "C" => String::from("F"),
            "L" => String::from("H"),
            "Y" => String::from("Hz"),
            _ => String::from("missing"),
        },
    }
}
#[test]
fn test_detect_measure_unit() {
    let test_data = vec![
        ["C123", "F"],
        ["R123", "ohm"],
        ["L232", "H"],
        ["Y123", "Hz"],
        ["Q123", "unknow"],
        ["TR123", "unknow"],
    ];

    for data in test_data.iter() {
        assert_eq!(detect_measure_unit(data[0]), data[1]);
    }
}

fn convert_comment_to_value(comment: &str) -> f64 {
    if comment == "NP" {
        return -1.0;
    }

    let v = comment
        .split(",")
        .map(|item| item.trim())
        .collect::<Vec<_>>();

    let value = match v.get(0) {
        None => return -1.0,
        Some(v) => v,
    };

    lazy_static! {
        static ref VAL: Regex = Regex::new(r"^([0-9.,]*)([GMkKRmunp])([0-9.,]*)").unwrap();
    }

    for cap in VAL.captures_iter(value) {
        println!("{:?}", cap);
    }
    0.0
}

#[test]
fn test_convert_comment_to_value() {
    struct TestData {
        refs: &'static str,
        val: f64,
    }

    let test_data = vec![
        TestData {
            refs: "100nF",
            val: 100e-9,
        },
        TestData {
            refs: "1R0",
            val: 100e-9,
        },
        TestData {
            refs: "1k",
            val: 1e3,
        },
        TestData {
            refs: "2k3",
            val: 2300.0,
        },
        TestData {
            refs: "4mH",
            val: 4e-3,
        },
        TestData {
            refs: "12MHZ",
            val: 12e6,
        },
        TestData {
            refs: "33nohm",
            val: 12e6,
        },
        TestData {
            refs: "100pF",
            val: 12e6,
        },
        TestData {
            refs: "1.1R",
            val: 12e6,
        },
        TestData {
            refs: "32.768kHz",
            val: 12e6,
        },
        TestData {
            refs: "12.134kHz",
            val: 12e6,
        },
        TestData {
            refs: "100uH",
            val: 12e6,
        },
        TestData {
            refs: "5K421",
            val: 12e6,
        },
        TestData {
            refs: "0.33R",
            val: 12e6,
        },
        TestData {
            refs: "2.2uH",
            val: 12e6,
        },
        TestData {
            refs: "0.3",
            val: 12e6,
        },
        TestData {
            refs: "4.7mH inductor",
            val: 12e6,
        },
    ];

    for data in test_data.iter() {
        // assert_eq!(convert_comment_to_value(data.refs), data.val);
        convert_comment_to_value(data.refs);
    }
    assert_eq!(0, 1);
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
                                    value: convert_comment_to_value(&template.comment),
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

    let tt = items.clone();
    for i in items {
        println!("{:?}", i);
    }
    for v in CATEGORY.iter() {
        println!("{}: {}", v, tt.iter().filter(|&n| n.category == *v).count());
    }
}
