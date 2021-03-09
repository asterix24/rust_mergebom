use super::items::{Category, HeaderMap, Item};
use xlsxwriter::*;

pub struct OutJobXlsx {
    wk: Workbook,
    curr_row: u32,
}

impl OutJobXlsx {
    pub fn new(filename: &str) -> OutJobXlsx {
        OutJobXlsx {
            wk: Workbook::new(format!("{}.xlsx", filename).as_str()),
            curr_row: 0,
        }
    }
    pub fn write(mut self, headers: &Vec<HeaderMap>, data: &Vec<Item>, categories: Vec<Category>) {
        let fmt_defalt = self.wk.add_format().set_font_size(10.0).set_text_wrap();
        let fmt_header = self
            .wk
            .add_format()
            .set_bg_color(FormatColor::Cyan)
            .set_bold()
            .set_font_size(12.0);
        let fmt_category = self
            .wk
            .add_format()
            .set_bg_color(FormatColor::Yellow)
            .set_bold()
            .set_border(FormatBorder::Thin)
            .set_align(FormatAlignment::CenterAcross);
        let fmt_qty = self
            .wk
            .add_format()
            .set_bg_color(FormatColor::Lime)
            .set_bold()
            .set_font_size(12.0);

        let mut sheet = match self.wk.add_worksheet(None) {
            Ok(wk) => wk,
            _ => panic!("Unable to add sheet to open wk"),
        };

        let mut column: u16 = 0;
        for label in headers.iter() {
            match sheet.write_string(
                self.curr_row,
                column,
                format!("{:?}", label).as_str(),
                Some(&fmt_header),
            ) {
                Ok(m) => m,
                _ => panic!("Error!"),
            };
            column += 1;
        }
        self.curr_row += 1;

        match self.wk.close() {
            Ok(m) => m,
            _ => panic!("Error! while close workbook!"),
        };
    }
}

//         let mut row_curr: u32 = 10;
//         let mut col: u16 = 0;
//         for label in HEADERS.iter() {
//             match sheet1.write_string(row_curr, col, label, Some(&hdr_fmt)) {
//                 Ok(m) => m,
//                 _ => panic!("Error!"),
//             };
//             col += 1;
//         }
//         row_curr += 1;
//         for (category, items) in data {
//             println!(">>>> {:?}", category);
//             match sheet1.merge_range(
//                 row_curr,
//                 0,
//                 row_curr,
//                 HEADERS.len() as u16,
//                 category.as_str(),
//                 Some(&merge_fmt),
//             ) {
//                 Ok(m) => m,
//                 _ => panic!("Error!"),
//             };

//             row_curr += 1;
//             //let sorted_item = items.sort_by_key(|k| Ord(pow(k.base_exp[0], k.base_exp[1])));
//             for unique_row in items {
//                 for (unique_key, parts) in unique_row {
//                     //println!("{:?} {:?}", unique_key, parts.len());
//                     let row_elem = match parts.first() {
//                         Some(r) => r,
//                         _ => panic!("empty part list"),
//                     };
//                     for label in headers.iter() {
//                         match label.as_str() {
//                             "quantity" => {
//                                 //println!("{:?}", parts.len());
//                                 match sheet1.write_string(
//                                     row_curr,
//                                     0,
//                                     parts.len().to_string().as_str(),
//                                     Some(&tot_fmt),
//                                 ) {
//                                     Ok(m) => m,
//                                     _ => panic!("Error!"),
//                                 };
//                             }
//                             "designator" => {
//                                 let mut ss: Vec<String> = Vec::new();
//                                 for s in parts.iter() {
//                                     ss.push(s.designator.clone());
//                                 }
//                                 match sheet1.write_string(
//                                     row_curr,
//                                     1,
//                                     ss.join(", ").as_str(),
//                                     Some(&default_fmt),
//                                 ) {
//                                     Ok(m) => m,
//                                     _ => panic!("Error!"),
//                                 };
//                             }
//                             "comment" => {
//                                 match sheet1.write_string(
//                                     row_curr,
//                                     2,
//                                     row_elem.comment.as_str(),
//                                     Some(&default_fmt),
//                                 ) {
//                                     Ok(m) => m,
//                                     _ => panic!("Error!"),
//                                 };
//                             }
//                             "footprint" => {
//                                 match sheet1.write_string(
//                                     row_curr,
//                                     3,
//                                     row_elem.footprint.as_str(),
//                                     Some(&default_fmt),
//                                 ) {
//                                     Ok(m) => m,
//                                     _ => panic!("Error!"),
//                                 };
//                             }
//                             "description" => {
//                                 match sheet1.write_string(
//                                     row_curr,
//                                     4,
//                                     row_elem.description.as_str(),
//                                     Some(&default_fmt),
//                                 ) {
//                                     Ok(m) => m,
//                                     _ => panic!("Error!"),
//                                 };
//                             }
//                             _ => println!("Invalid category"),
//                         }
//                     }
//                     row_curr += 1;
//                 }
//             }
//         }

//     }
//     }
// }
