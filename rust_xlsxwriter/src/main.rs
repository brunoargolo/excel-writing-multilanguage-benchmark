use std::fs::File;
use std::io::Read;
use flate2::read::GzDecoder;
use serde::Deserialize;
use rust_xlsxwriter::{ExcelDateTime, Format, Workbook, Worksheet};
use std::time::Instant;
use std::env;
use rayon::prelude::*;
use std::sync::{Arc, Mutex};

#[derive(Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct Record {
    #[serde(rename(deserialize = "id"))]
    pub id: String,
    #[serde(rename(deserialize = "myString1"))]
    pub my_string_1: String,
    #[serde(rename(deserialize = "myDate1"))]
    pub my_date_1: String,
    #[serde(rename(deserialize = "myDate2"))]
    pub my_date_2: String,
    #[serde(rename(deserialize = "amount"))]
    pub amount: f64,
    #[serde(rename(deserialize = "myNumericString"))]
    pub my_numeric_string2: Option<String>,
    #[serde(rename(deserialize = "myString2"))]
    pub my_string_2: Option<String>,
}

fn get_content() -> Result<Vec<Record>, Box<dyn std::error::Error>> {
    let file = File::open("../input.json.gzip")?;
    let mut gz = GzDecoder::new(file);
    let mut contents = String::new();
    gz.read_to_string(&mut contents)?;
    
    let records: Vec<Record> = serde_json::from_str(&contents)?;
    Ok(records)
}

fn write_sheet(worksheet: &mut Worksheet, recs: &[Record], decimal_format: &Format, date_format: &Format) {
    worksheet.set_column_width(0, 22).unwrap();

    for (i, rec) in recs.iter().enumerate() {
        let r = u32::try_from(i).unwrap();

        worksheet.write(r, 0, &rec.id).unwrap();
        worksheet.write(r, 1, &rec.my_string_1).unwrap();
        worksheet.write(r, 2, rec.my_numeric_string2.as_deref().unwrap_or("")).unwrap();
        worksheet.write(r, 3, rec.my_string_2.as_deref().unwrap_or("")).unwrap();

        worksheet.write_with_format(r, 4, rec.amount, decimal_format).unwrap();

        let my_date_2 = ExcelDateTime::parse_from_str(&rec.my_date_2).unwrap();
        worksheet.write_with_format(r, 5, &my_date_2, date_format).unwrap();    
        
        let my_date_1 = ExcelDateTime::parse_from_str(&rec.my_date_1).unwrap();
        worksheet.write_with_format(r, 6, &my_date_1, date_format).unwrap();    
    }
}



fn to_excel(recs: Vec<Record>) {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();
    
    // Get the number of sheets from the environment variable
    let n_sheets = match env::var("N_SHEETS") {
        Ok(val) => {
            let n: u8 = val.parse().unwrap_or(1);
            if n > 0 && n < 9 {
                n
            } else {
                println!("Invalid N_SHEETS value. Using default of 1.");
                1
            }
        },
        Err(_) => 1,
    };

    // Create some formats to use in the worksheet.
    let decimal_format = Format::new().set_num_format("0.000");
    let date_format = Format::new().set_num_format("yyyy-mm-dd");

    for sheet_num in 1..=n_sheets {
        // Add a worksheet to the workbook.
        let worksheet = workbook.add_worksheet_with_low_memory();
        // Set the worksheet name
        worksheet.set_name(&format!("Sheet{}", sheet_num)).unwrap();
        // Set the column width for clarity.
        worksheet.set_column_width(0, 22);

        for (i, rec) in recs.iter().enumerate() {
            let r = u32::try_from(i).unwrap();

            worksheet.write(r, 0, &rec.id);
            worksheet.write(r, 1, &format!("{}", &rec.my_string_1));
            worksheet.write(r, 2, rec.my_numeric_string2.as_deref().unwrap_or(""));
            worksheet.write(r, 3, rec.my_string_2.as_deref().unwrap_or(""));

            worksheet.write_with_format(r, 4, rec.amount, &decimal_format);

            let my_date_2 = ExcelDateTime::parse_from_str(&rec.my_date_2).unwrap();
            worksheet.write_with_format(r, 5, &my_date_2, &date_format);    
            
            let my_date_1 = ExcelDateTime::parse_from_str(&rec.my_date_1).unwrap();
            worksheet.write_with_format(r, 6, &my_date_1, &date_format);    
        }
    }
    
    // Save the file to disk.
    workbook.save("demo.xlsx");
}

fn main(){
    let start = Instant::now();
    let records = get_content().unwrap();
    println!("Load Time {} seconds", start.elapsed().as_secs());

    let start = Instant::now();
    to_excel(records);
    println!("Xlsx Write Time {} seconds", start.elapsed().as_secs());
}