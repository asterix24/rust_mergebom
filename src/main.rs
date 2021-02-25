use clap::{App, Arg};
mod lib;

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

    let mut items: Vec<lib::items::Item> = Vec::new();
    for bom in matches.values_of("BOMFile").unwrap() {
        let header_map = lib::find_header(bom);
        items = lib::parse_xlsx(bom, &header_map);
    }

    for i in items {
        println!("{:?}", i);
    }
}

//for item in boms {}

//     let mut grouped_items: HashMap<String, Vec<HashMap<String, Vec<Item>>>> = HashMap::new();
//     for category in CATEGORY.iter() {
//         let group = items.iter().filter(|n| n.category == *category);

//         let mut item_sets: HashMap<String, Vec<Item>> = HashMap::new();
//         for row in group {
//             let key = match *category {
//                 "connectors" => format!("{}{}", row.footprint, row.description),
//                 _ => format!("{}{}{}", row.comment, row.footprint, row.description),
//             };

//             match item_sets.entry(key) {
//                 Entry::Occupied(mut o) => {
//                     o.get_mut().push(row.clone());
//                 }
//                 Entry::Vacant(v) => {
//                     v.insert(vec![row.clone(); 1]);
//                 }
//             };
//         }
//         match grouped_items.entry(String::from(*category)) {
//             Entry::Occupied(mut o) => {
//                 o.get_mut().push(item_sets);
//             }
//             Entry::Vacant(v) => {
//                 v.insert(vec![item_sets; 1]);
//             }
//         };

//         header_map.insert(String::from("quantity"), 0);
//         report(
//             header_map.keys().cloned().collect(),
//             grouped_items.clone(),
//             "merged_bom.xlsx",
//         );
//     }
// }
