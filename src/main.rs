use clap::{App, Arg};
mod lib;
use lib::items::{categories, dump, DataParser, Item};
use lib::outjob::OutJob;

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
        let v: Vec<Item> = data.collect();
        let c: Vec<String> = categories(&v);
        for i in c {
            println!("-> {}", i);
        }
        dump(&v);
    }

    let out = OutJob::new("merge_bom.xlsx");
    //out.write();
}
