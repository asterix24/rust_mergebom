#[derive(Default, Debug, Clone)]
pub struct Item {
    pub category: String,
    pub base_exp: (f32, i32),
    pub fmt_value: String,
    pub measure_unit: String,
    pub designator: String,
    pub comment: String,
    pub footprint: String,
    pub description: String,
}

#[derive(Default, Debug, Clone)]
pub struct ItemRow {
    pub quantity: String,
    pub designator: Vec<String>,
    pub comment: String,
    pub footprint: String,
    pub description: String,
}
