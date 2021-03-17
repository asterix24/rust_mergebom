use super::items::{Category, Header, HeaderMap, Item};
use super::utils::value_to_eng_notation;
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
    pub fn write(mut self, headers: &[HeaderMap], data: &[Item], categories: Vec<Category>) {
        let fmt_defalt = self
            .wk
            .add_format()
            .set_text_wrap()
            .set_font_size(10.0)
            .set_text_wrap();
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
        sheet
            .write_string(self.curr_row, column, "Qty", Some(&fmt_qty))
            .unwrap();
        column += 1;
        for hdr in headers.iter() {
            sheet
                .write_string(
                    self.curr_row,
                    column,
                    format!("{}", hdr.label).as_str(),
                    Some(&fmt_header),
                )
                .unwrap();
            column += 1;
        }
        self.curr_row += 1;
        for i in categories.iter() {
            // Write Category Header
            sheet
                .merge_range(
                    self.curr_row,
                    0,
                    self.curr_row,
                    headers.len() as u16,
                    format!("{:?}", i).as_str(),
                    Some(&fmt_category),
                )
                .unwrap();
            self.curr_row += 1;
            for item in data.iter().filter(|m| m.category == *i) {
                // Write Qty
                sheet
                    .write_string(
                        self.curr_row,
                        Header::Quantity as u16,
                        item.designator.len().to_string().as_str(),
                        Some(&fmt_qty),
                    )
                    .unwrap();
                // Write designator
                sheet
                    .write_string(
                        self.curr_row,
                        Header::Designator as u16,
                        item.designator.join(", ").as_str(),
                        Some(&fmt_defalt),
                    )
                    .unwrap();
                // Write Comment
                sheet
                    .write_string(
                        self.curr_row,
                        Header::Comment as u16,
                        value_to_eng_notation(
                            item.base_exp.0,
                            item.base_exp.1,
                            item.measure_unit.as_str(),
                        )
                        .as_str(),
                        Some(&fmt_defalt),
                    )
                    .unwrap();
                // Write Footprint
                sheet
                    .write_string(
                        self.curr_row,
                        Header::Footprint as u16,
                        item.footprint.as_str(),
                        Some(&fmt_defalt),
                    )
                    .unwrap();
                // Write Description
                sheet
                    .write_string(
                        self.curr_row,
                        Header::Description as u16,
                        item.description.as_str(),
                        Some(&fmt_defalt),
                    )
                    .unwrap();
                // Write extra column
                for (n, m) in item.extra.iter().enumerate() {
                    sheet
                        .write_string(
                            self.curr_row,
                            m.label as u16 + n as u16,
                            m.value.as_str(),
                            Some(&fmt_defalt),
                        )
                        .unwrap();
                }
                self.curr_row += 1;
            }
        }

        self.wk.close().unwrap();
    }
}
