use anyhow::{Context, Result};
use calamine::{RangeDeserializerBuilder, Reader, Xlsx, open_workbook};
use csv::Writer;
use rfd::FileDialog;
use std::fs::File;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Open file dialog
    let file_path = FileDialog::new()
        .set_title("Select a Pcl Invoice Excel file")
        .add_filter("Excel files", &["csv", "xls", "xlsx"])
        .pick_file()
        .context("No file selected")?; // Convert Option to Result with context

    println!("Selected file: {:?}", file_path);

    // Open workbook (moved inside the if let block)
    let mut workbook: Xlsx<_> = open_workbook(&file_path)?;
    println!("Workbook opened successfully.");

    let output_path = FileDialog::new()
        .set_title("Select output CSV location")
        .add_filter("CSV files", &["csv"])
        .save_file()
        .context("No output file selected")?;

    let mut wtr = Writer::from_path(&output_path)?;

    // List sheets (moved after workbook is opened)

    let mut start_processing = false;

    for sheet in workbook.sheet_names() {
        if sheet == "WR1" {
            start_processing = true;
        }

        if start_processing {
            println!("Processing sheet: {}", sheet);
            if let Ok(range) = workbook.worksheet_range(&sheet) {
                for (row_idx, row) in range.rows().enumerate() {
                    let mut record = vec![sheet.clone(), row_idx.to_string()];
                    record.extend(row.iter().map(|cell| cell.to_string()));
                    wtr.write_record(&record)?;
                }
            }

            if sheet == "FIXED FEE " {
                break;
            }
        }
    }
    wtr.flush()?;
    Ok(())
}
