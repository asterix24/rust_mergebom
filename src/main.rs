use clap::{App, Arg};
mod lib;
use lib::items::{DataParser, Item};

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

    let bom = matches.values_of("BOMFile").unwrap();
    for i in bom {
        let data: DataParser = DataParser::new(i);
        let items: Vec<Item> = data.xlsx();
        for i in items {
            println!("{:?}", i);
        }
    }
}

