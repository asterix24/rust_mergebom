pub mod items;
pub mod utils;

//fn report() {
//     headers: Vec<String>,
//     data: HashMap<String, Vec<HashMap<String, Vec<Item>>>>,
//     merge_file: &str,
// ) {
//     let workbook = Workbook::new(merge_file);
//     let default_fmt = workbook.add_format().set_font_size(10.0).set_text_wrap();

//     let hdr_fmt = workbook
//         .add_format()
//         .set_bg_color(FormatColor::Cyan)
//         .set_bold()
//         .set_font_size(12.0);
//     let merge_fmt = workbook
//         .add_format()
//         .set_bg_color(FormatColor::Yellow)
//         .set_bold()
//         .set_border(FormatBorder::Thin)
//         .set_align(FormatAlignment::CenterAcross);

//     let tot_fmt = workbook
//         .add_format()
//         .set_bg_color(FormatColor::Lime)
//         .set_bold()
//         .set_font_size(12.0);

//     let mut sheet1 = match workbook.add_worksheet(None) {
//         Ok(wk) => wk,
//         _ => panic!("Unable to open wk"),
//     };

//     let mut row_curr: u32 = 10;
//     let mut col: u16 = 0;
//     for label in HEADERS.iter() {
//         match sheet1.write_string(row_curr, col, label, Some(&hdr_fmt)) {
//             Ok(m) => m,
//             _ => panic!("Error!"),
//         };
//         col += 1;
//     }
//     row_curr += 1;
//     for (category, items) in data {
//         println!(">>>> {:?}", category);
//         match sheet1.merge_range(
//             row_curr,
//             0,
//             row_curr,
//             HEADERS.len() as u16,
//             category.as_str(),
//             Some(&merge_fmt),
//         ) {
//             Ok(m) => m,
//             _ => panic!("Error!"),
//         };

//         row_curr += 1;
//         //let sorted_item = items.sort_by_key(|k| Ord(pow(k.base_exp[0], k.base_exp[1])));
//         for unique_row in items {
//             for (unique_key, parts) in unique_row {
//                 //println!("{:?} {:?}", unique_key, parts.len());
//                 let row_elem = match parts.first() {
//                     Some(r) => r,
//                     _ => panic!("empty part list"),
//                 };
//                 for label in headers.iter() {
//                     match label.as_str() {
//                         "quantity" => {
//                             //println!("{:?}", parts.len());
//                             match sheet1.write_string(
//                                 row_curr,
//                                 0,
//                                 parts.len().to_string().as_str(),
//                                 Some(&tot_fmt),
//                             ) {
//                                 Ok(m) => m,
//                                 _ => panic!("Error!"),
//                             };
//                         }
//                         "designator" => {
//                             let mut ss: Vec<String> = Vec::new();
//                             for s in parts.iter() {
//                                 ss.push(s.designator.clone());
//                             }
//                             match sheet1.write_string(
//                                 row_curr,
//                                 1,
//                                 ss.join(", ").as_str(),
//                                 Some(&default_fmt),
//                             ) {
//                                 Ok(m) => m,
//                                 _ => panic!("Error!"),
//                             };
//                         }
//                         "comment" => {
//                             match sheet1.write_string(
//                                 row_curr,
//                                 2,
//                                 row_elem.comment.as_str(),
//                                 Some(&default_fmt),
//                             ) {
//                                 Ok(m) => m,
//                                 _ => panic!("Error!"),
//                             };
//                         }
//                         "footprint" => {
//                             match sheet1.write_string(
//                                 row_curr,
//                                 3,
//                                 row_elem.footprint.as_str(),
//                                 Some(&default_fmt),
//                             ) {
//                                 Ok(m) => m,
//                                 _ => panic!("Error!"),
//                             };
//                         }
//                         "description" => {
//                             match sheet1.write_string(
//                                 row_curr,
//                                 4,
//                                 row_elem.description.as_str(),
//                                 Some(&default_fmt),
//                             ) {
//                                 Ok(m) => m,
//                                 _ => panic!("Error!"),
//                             };
//                         }
//                         _ => println!("Invalid category"),
//                     }
//                 }
//                 row_curr += 1;
//             }
//         }
//     }
//     match workbook.close() {
//         Ok(m) => m,
//         _ => panic!("Error!"),
//     };
// }
//}
