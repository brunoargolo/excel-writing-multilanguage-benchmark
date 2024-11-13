use flate2::read::GzDecoder;
use rayon::prelude::*;
use rust_xlsxwriter::{Format, Workbook, Worksheet, XlsxError};
use serde_json::Value;
use std::env;
use std::fs::File;
use std::io::{self, Read};
use std::sync::{Arc, Mutex};
use std::time::Instant;

// Custom error type
#[derive(Debug)]
enum AppError {
    Io(io::Error),
    Json(serde_json::Error),
    Xlsx(XlsxError),
    Other(String),
}

impl From<io::Error> for AppError {
    fn from(err: io::Error) -> Self {
        AppError::Io(err)
    }
}

impl From<serde_json::Error> for AppError {
    fn from(err: serde_json::Error) -> Self {
        AppError::Json(err)
    }
}

impl From<XlsxError> for AppError {
    fn from(err: XlsxError) -> Self {
        AppError::Xlsx(err)
    }
}

fn read_compressed_json_file(filename: &str) -> Result<Vec<Value>, AppError> {
    let file = File::open(filename)?;
    let mut gz = GzDecoder::new(file);
    let mut json_string = String::new();
    gz.read_to_string(&mut json_string)?;
    let records: Vec<Value> = serde_json::from_str(&json_string)?;
    Ok(records)
}

fn write_to_excel(records: &[Value]) -> Result<(), AppError> {
    let workbook = Arc::new(Mutex::new(Workbook::new()));

    // Get N_SHEETS value from environment variable, default to 1 if not set
    let n_sheets = env::var("N_SHEETS")
        .unwrap_or_else(|_| "1".to_string())
        .parse::<usize>()
        .unwrap_or(1);

    // Validate N_SHEETS value
    if n_sheets < 1 || n_sheets > 9 {
        return Err(AppError::Other("N_SHEETS must be between 1 and 9".to_string()));
    }

    let date_format = Format::new().set_num_format("yyyy-mm-dd");
    let amount_format = Format::new().set_num_format("0.000");

    // Create sheets in parallel
    (1..=n_sheets).into_par_iter().try_for_each(|i| {
        create_sheet(&workbook, i, records, &date_format, &amount_format)
    })?;

    // Save the workbook
    workbook.lock().unwrap().save("demo.xlsx")?;
    Ok(())
}

fn create_sheet(
    workbook: &Arc<Mutex<Workbook>>,
    sheet_num: usize,
    records: &[Value],
    date_format: &Format,
    amount_format: &Format,
) -> Result<(), AppError> {
    let mut workbook = workbook.lock().unwrap();
    let worksheet = workbook.add_worksheet_with_low_memory();
    worksheet.set_name(&format!("Sheet{}", sheet_num))?;

    // Define columns
    let columns = [
        ("ID", 22.0),
        ("My String 1", 22.0),
        ("My Numeric String", 22.0),
        ("My String 2", 22.0),
        ("Amount", 15.0),
        ("My Date 1", 15.0),
        ("My Date 2", 15.0),
    ];

    // Write headers and set column widths
    for (col, (header, width)) in columns.iter().enumerate() {
        worksheet.write_string(0, col as u16, *header)?;
        worksheet.set_column_width(col as u16, *width)?;
    }

    // Write data
    for (row, record) in records.iter().enumerate() {
        worksheet.write_string((row + 1) as u32, 0, record["id"].as_str().unwrap_or(""))?;
        worksheet.write_string((row + 1) as u32, 1, record["myString1"].as_str().unwrap_or(""))?;
        worksheet.write_string((row + 1) as u32, 2, record["myNumericString"].as_str().unwrap_or(""))?;
        worksheet.write_string((row + 1) as u32, 3, record["myString2"].as_str().unwrap_or(""))?;
        worksheet.write_number_with_format((row + 1) as u32, 4, record["amount"].as_f64().unwrap_or(0.0), amount_format)?;
        worksheet.write_string_with_format((row + 1) as u32, 5, record["myDate1"].as_str().unwrap_or(""), date_format)?;
        worksheet.write_string_with_format((row + 1) as u32, 6, record["myDate2"].as_str().unwrap_or(""), date_format)?;
    }

    Ok(())
}

fn main() -> Result<(), AppError> {
    let start = Instant::now();
    let records = read_compressed_json_file("../input.json.gzip")?;
    println!("Load Time: {:?}", start.elapsed());

    println!("Retrieved {} records", records.len());

    let start = Instant::now();
    write_to_excel(&records)?;
    println!("Write Time: {:?}", start.elapsed());

    Ok(())
}