use clap::{App, Arg};
mod lib;
use lib::items::Category;
use lib::items::DataParser;
use lib::load::Load;
use lib::outjob::OutJobXlsx;
use lib::ASCII_LOGO;

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

    println!("{}", ASCII_LOGO);

    let bom = matches.values_of("BOMFile").unwrap();
    for i in bom {
        let ld: Load = Load::new(i);
        let data: DataParser = DataParser::new(ld);

        let c: Vec<Category> = data.categories();
        for x in data.stats() {
            println!("->\t{:?} {}", x.label, x.value);
        }
        //dump(&v);
        // for i in c.iter() {
        //     println!("-> {:?}", i);
        // }
        // for i in hdr.iter() {
        //     println!("=> {:?}", i);
        // }
        let out = OutJobXlsx::new("merged_bom");
        out.write(data.headers(), data.items(), c.clone());
    }
}
