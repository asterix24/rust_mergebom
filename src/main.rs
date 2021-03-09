use clap::{App, Arg};
mod lib;
use lib::items::Category;
use lib::items::{categories, dump, DataParser, HeaderMap, Item};
use lib::outjob::OutJobXlsx;

fn main() {
    let matches = App::new("Rust MergeBom")
        .version("0.1.0")
        .author("Daniele Basile <asterix24@gmail.com>")
        .about("Pretty merger and formatter Bill Of Materials.")
        .arg(
            Arg::with_name("BOMFile")
                .help("BOM to Merge")
                .required(true)
                .min_values(1),
        )
        .get_matches();

    let bom = matches.values_of("BOMFile").unwrap();
    for i in bom {
        let mut data: DataParser = DataParser::new(i);
        let hdr: Vec<HeaderMap> = data.headers();
        println!("-> {:?}", hdr);

        let v: Vec<Item> = data.parse(&hdr);
        let c: Vec<Category> = categories(&v);
        dump(&v);
        for i in c.iter() {
            println!("-> {:?}", i);
        }
        for i in hdr.iter() {
            println!("=> {:?}", i);
        }
        let out = OutJobXlsx::new("merged_bom");
        out.write(&hdr, &v, c.clone());
    }
}
