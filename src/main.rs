use calamine::{open_workbook, DataType, Reader, Xlsx};
use clap::{App, Arg};
use std::collections::HashMap;

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
    let mut header: HashMap<&str, Option<usize>> = [
        ("quantity", None),
        ("designator", None),
        ("comment", None),
        ("footprint", None),
        ("description", None),
    ]
    .iter()
    .cloned()
    .collect();

    let headers = vec![
        "quantity",
        "designator",
        "comment",
        "footprint",
        "description",
    ];

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
                    println!("{:?} -> {} {}", d, row, column);

                    if let Some(DataType::String(s)) = d {
                        println!("{}", s);
                        if let Some(val) = header.get_mut(s.to_lowercase().as_str()) {
                            *val = Some(column)
                        }
                    }
                }
            }

            for (key, val) in header.iter() {
                println!("key: {} val: {:?}", key, val);
            }
        }

        if let Some(Ok(range)) = workbook.worksheet_range("Sheet1") {
            let (rw, cl) = range.get_size();
            for row in 0..rw {
                for column in 0..cl {
                    let d = range.get((row, column));
                    println!("{:?} -> {} {}", d, row, column);
                }
            }
        }
    }
}
