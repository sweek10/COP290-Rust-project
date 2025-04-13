// main.rs
mod types;
mod sheet;
mod dependencies;
mod cell;
mod utils;

use std::io::{self, BufRead, Write};
use std::time::Instant;
use crate::types::SHEET;
use crate::sheet::{create_sheet, process_command, display_sheet};
use crate::utils::is_valid_command;
use std::fs::File;
use std::path::Path;
use calamine::{Reader, open_workbook, Xlsx, Range, DataType};

const MAX_ROWS: i32 = 999;
const MAX_COLS: i32 = 18278;

fn load_csv_file(sheet: &mut crate::types::Sheet, filename: &str) -> Result<(), String> {
    // Open the file
    let file = File::open(filename).map_err(|e| format!("Failed to open CSV file: {}", e))?;
    let reader = io::BufReader::new(file);
    
    // Read each line
    let mut row_idx = 0;
    for line in reader.lines() {
        if row_idx >= sheet.rows {
            return Err(format!("CSV file has more rows than the spreadsheet (max: {})", sheet.rows));
        }
        
        let line = line.map_err(|e| format!("Error reading CSV line: {}", e))?;
        let values: Vec<&str> = line.split(',').collect();
        
        // Process each value in the row
        let mut col_idx = 0;
        for value in values {
            if col_idx >= sheet.cols {
                return Err(format!("CSV file has more columns than the spreadsheet (max: {})", sheet.cols));
            }
            
            let value = value.trim();
            
            // Try to parse as number first
            if let Ok(num_value) = value.parse::<i32>() {
                // Set as direct value
                sheet.cells[row_idx as usize][col_idx as usize].value = num_value;
            } else if value.starts_with('=') {
                // It's a formula
                let formula = &value[1..]; // Remove the = prefix
                
                // Create cell reference string for this cell
                let mut cell_ref = String::new();
                crate::utils::encode_column(col_idx, &mut cell_ref);
                cell_ref.push_str(&(row_idx + 1).to_string());
                
                // Apply the formula to the cell
                crate::cell::update_cell(sheet, &cell_ref, formula);
            } else if !value.is_empty() {
                // For non-empty, non-numeric, non-formula values, store as value 0
                // with a comment (this is a limitation as our spreadsheet only supports numbers)
                sheet.cells[row_idx as usize][col_idx as usize].value = 0;
                // Could add a comment field to Cell struct for future enhancement
            }
            
            col_idx += 1;
        }
        
        row_idx += 1;
    }
    
    Ok(())
}

fn load_excel_file(sheet: &mut crate::types::Sheet, filename: &str) -> Result<(), String> {
    // Open the workbook
    let mut workbook: Xlsx<_> = open_workbook(filename)
        .map_err(|e| format!("Failed to open Excel file: {}", e))?;
    
    // Get the first worksheet (we'll only read the first sheet)
    let sheet_names = workbook.sheet_names().to_vec();
    if sheet_names.is_empty() {
        return Err("Excel file doesn't contain any worksheets".to_string());
    }
    
    let worksheet = workbook.worksheet_range(&sheet_names[0])
        .ok_or_else(|| "Failed to get first worksheet".to_string())?
        .map_err(|e| format!("Error accessing worksheet: {}", e))?;
    
    // Process each cell in the worksheet
    let height = worksheet.height() as i32;
    let width = worksheet.width() as i32;
    
    if height > sheet.rows {
        return Err(format!("Excel file has more rows than the spreadsheet (file: {}, max: {})", 
                          height, sheet.rows));
    }
    
    if width > sheet.cols {
        return Err(format!("Excel file has more columns than the spreadsheet (file: {}, max: {})", 
                          width, sheet.cols));
    }
    
    for row_idx in 0..height {
        for col_idx in 0..width {
            match worksheet.get_value(((row_idx as usize).try_into().unwrap(), (col_idx as usize).try_into().unwrap())) {
                Some(DataType::Int(value)) => {
                    sheet.cells[row_idx as usize][col_idx as usize].value = *value as i32;
                },
                Some(DataType::Float(value)) => {
                    // Since our spreadsheet only supports integers, we'll round floats
                    sheet.cells[row_idx as usize][col_idx as usize].value = *value as i32;
                },
                Some(DataType::String(value)) => {
                    if value.starts_with('=') {
                        // It's a formula
                        let formula = &value[1..]; // Remove the = prefix
                        
                        // Create cell reference string for this cell
                        let mut cell_ref = String::new();
                        crate::utils::encode_column(col_idx, &mut cell_ref);
                        cell_ref.push_str(&(row_idx + 1).to_string());
                        
                        // Apply the formula to the cell
                        crate::cell::update_cell(sheet, &cell_ref, formula);
                    } else {
                        // For non-formula strings, store as value 0
                        sheet.cells[row_idx as usize][col_idx as usize].value = 0;
                    }
                },
                Some(DataType::Bool(value)) => {
                    // Convert boolean to 1 (true) or 0 (false)
                    sheet.cells[row_idx as usize][col_idx as usize].value = if *value { 1 } else { 0 };
                },
                Some(DataType::DateTime(_)) => {
                    // For date/time values, store as 0 (could be enhanced in the future)
                    sheet.cells[row_idx as usize][col_idx as usize].value = 0;
                },
                Some(DataType::Error(_)) => {
                    // For error values, store as 0
                    sheet.cells[row_idx as usize][col_idx as usize].value = 0;
                },
                None => {
                    // Empty cell, do nothing
                }
                _ => {
                    // Other types, set to 0
                    sheet.cells[row_idx as usize][col_idx as usize].value = 0;
                }
            }
        }
    }
    
    Ok(())
}

fn main() {
    let args: Vec<String> = std::env::args().collect();
    let mut extension_enabled = false;
    let mut row_col_args = Vec::new();
    let mut input_file = None;

    // Parse command-line arguments
    let mut i = 1;
    while i < args.len() {
        if args[i] == "--extension" {
            extension_enabled = true;
            i += 1;
        } else {
            row_col_args.push(args[i].clone());
            i += 1;
        }
    }

    // Check if the last argument is a data file (when using extension mode)
    if extension_enabled && row_col_args.len() == 3 {
        let potential_file = &row_col_args[2];
        if Path::new(potential_file).exists() {
            input_file = Some(potential_file.clone());
            row_col_args.pop(); // Remove the filename from row_col_args
        }
    }

    if row_col_args.len() != 2 {
        println!("Usage: {} [--extension] <rows> <columns> [input_file.csv|xlsx]", args[0]);
        println!("Note: File loading is only available with --extension flag");
        return;
    }

    let rows: i32 = row_col_args[0].parse().unwrap_or(0);
    let cols: i32 = row_col_args[1].parse().unwrap_or(0);

    if rows < 1 || rows > MAX_ROWS || cols < 1 || cols > MAX_COLS {
        println!("Invalid dimensions. Rows: 1-{}, Columns: 1-{}", MAX_ROWS, MAX_COLS);
        return;
    }

    let mut sheet_guard = SHEET.lock().unwrap();
    *sheet_guard = create_sheet(rows, cols, extension_enabled);
    
    // If an input file was specified and extension is enabled, load it
    if extension_enabled {
        if let Some(filename) = input_file {
            if let Some(ref mut sheet) = *sheet_guard {
                // Check file extension to determine how to load it
                let extension = Path::new(&filename)
                    .extension()
                    .and_then(|ext| ext.to_str())
                    .unwrap_or("");
                
                let result = match extension.to_lowercase().as_str() {
                    "csv" => load_csv_file(sheet, &filename),
                    "xlsx" => load_excel_file(sheet, &filename),
                    _ => Err(format!("Unsupported file format: {}", extension))
                };
                
                match result {
                    Ok(_) => println!("Successfully loaded file: {}", filename),
                    Err(e) => println!("Error loading file: {}", e)
                }
            }
        }
    }
    
    drop(sheet_guard);

    let mut elapsed_time = 0.0;
    let mut is_valid = true;
    let stdin = io::stdin();

    loop {
        {
            let sheet_guard = SHEET.lock().unwrap();
            if let Some(ref sheet) = *sheet_guard {
                display_sheet(sheet);
            }
        }

        print!("[{:.1}] {}> ", elapsed_time, if is_valid {
            if SHEET.lock().unwrap().as_ref().unwrap().circular_dependency_detected { "(err)" } else { "(ok)" }
        } else { "(err)" });
        io::stdout().flush().unwrap();

        let mut command = String::new();
        if stdin.lock().read_line(&mut command).is_err() {
            break;
        }
        let command = command.trim();

        if command == "q" {
            break;
        }
        is_valid = is_valid_command(&mut SHEET.lock().unwrap().as_mut().unwrap(), command);
        let start = Instant::now();
        {
            let mut sheet_guard = SHEET.lock().unwrap();
            if let Some(ref mut sheet) = *sheet_guard {
                process_command(sheet, command);
            }
        }
        elapsed_time = start.elapsed().as_secs_f64();
    }
}
