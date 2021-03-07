use clap::{App, Arg};
mod lib;
use lib::items::{categories, dump, headers_to_str, DataParser, HeaderMap, Item};
use lib::outjob::OutJobXlsx;

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
        let mut data: DataParser = DataParser::new(i);
        let hdr: Vec<HeaderMap> = data.headers();
        println!("-> {:?}", hdr);

        let v: Vec<Item> = data.parse(&hdr);
        let c: Vec<String> = categories(&v);
        let h: Vec<String> = headers_to_str(&hdr);
        dump(&v);
        for i in c.iter() {
            println!("-> {}", i);
        }
        for i in h.iter() {
            println!("=> {}", i);
        }
        let out = OutJobXlsx::new("merged_bom");
        out.write(&h, &v, &c);
    }
}
