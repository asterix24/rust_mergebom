use calamine::{open_workbook, DataType, Reader, Xlsx};
use clap::{App, Arg};
use regex::Regex;
use std::collections::HashMap;

#[derive(Default, Debug, Clone)]
struct ItemBOM {
    category: String,
    value: i64,
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

fn guess_category(s: String) -> String {
    return String::from("cat");
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
                                    //println!("label {:?}: cell ->{:?}", label, value);
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
                                    ..template.clone()
                                })
                                .collect::<Vec<_>>(),
                        );
                    }
                }
            }
        }
    }

    let re = Regex::new(r"^([a-zA-Z_]{1,3})").unwrap();
    for i in items {
        println!("{:?}", i);
        for cap in re.captures_iter(&i.designator) {
            println!("group: {:?}", &cap[1]);
        }
    }
}
