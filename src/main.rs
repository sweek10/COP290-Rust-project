mod types;
mod sheet;
mod dependencies;
mod cell;
mod utils;

use std::io::{self, BufRead, Write};
use std::time::Instant;
use std::path::Path;
use std::fs::File;
use axum::{
    routing::{get, post},
    Router, Json,
    http::StatusCode,
    response::{IntoResponse, Html, Response},
};
use tokio::net::TcpListener;
use serde::{Serialize, Deserialize};
use crate::types::{SHEET, Sheet};
use crate::sheet::{create_sheet, process_command, display_sheet};
use crate::utils::is_valid_command;
use calamine::{Reader, open_workbook, Xlsx, DataType};

const MAX_ROWS: i32 = 999;
const MAX_COLS: i32 = 18278;

#[derive(Serialize, Deserialize)]
struct CellResponse {
    row: i32,
    col: i32,
    value: i32,
    formula: Option<String>,
    is_error: bool,
    is_bold: bool,
    is_italic: bool,
    is_underline: bool,
}

#[derive(Deserialize)]
struct CommandRequest {
    command: String,
}

async fn run_web_interface(rows: i32, cols: i32, extension_enabled: bool, input_file: Option<String>) {
    // Initialize the sheet
    let mut sheet_guard = SHEET.lock().unwrap();
    *sheet_guard = create_sheet(rows, cols, extension_enabled);

    // Load input file if provided
    if let Some(filename) = input_file {
        if let Some(ref mut sheet) = *sheet_guard {
            let extension = Path::new(&filename)
                .extension()
                .and_then(|ext| ext.to_str())
                .unwrap_or("");
            let result = match extension.to_lowercase().as_str() {
                "csv" => load_csv_file(sheet, &filename),
                "xlsx" => load_excel_file(sheet, &filename),
                _ => Err(format!("Unsupported file format: {}", extension)),
            };
            match result {
                Ok(_) => println!("Successfully loaded file: {}", filename),
                Err(e) => println!("Error loading file: {}", e),
            }
        }
    }
    drop(sheet_guard);

    // Define the web server routes
    let app = Router::new()
        .route("/", get(root))
        .route("/api/sheet", get(get_sheet))
        .route("/api/command", post(execute_command));

    // Start the server
    let addr = "127.0.0.1:3000";
    println!("Web interface running at http://{}", addr);
    let listener = TcpListener::bind(addr).await.unwrap();
    axum::serve(listener, app).await.unwrap();
}

// Serve the main HTML page
async fn root() -> impl IntoResponse {
    Html(include_str!("../web/index.html"))
}

// Get the current state of the sheet
async fn get_sheet() -> Response {
    let sheet_guard = SHEET.lock().unwrap();
    if let Some(sheet) = &*sheet_guard {
        let mut cells = Vec::new();
        for i in 0..sheet.rows.min(100) { // Limit for performance
            for j in 0..sheet.cols.min(100) {
                let cell = &sheet.cells[i as usize][j as usize];
                cells.push(CellResponse {
                    row: i,
                    col: j,
                    value: cell.value,
                    formula: cell.formula.clone(),
                    is_error: cell.is_error,
                    is_bold: cell.is_bold,
                    is_italic: cell.is_italic,
                    is_underline: cell.is_underline,
                });
            }
        }
        Json(cells).into_response()
    } else {
        (StatusCode::INTERNAL_SERVER_ERROR, "Sheet not initialized").into_response()
    }
}

// Execute a command
async fn execute_command(Json(payload): Json<CommandRequest>) -> impl IntoResponse {
    let mut sheet_guard = SHEET.lock().unwrap();
    if let Some(ref mut sheet) = *sheet_guard {
        if is_valid_command(sheet, &payload.command) {
            process_command(sheet, &payload.command);
            (StatusCode::OK, "Command executed")
        } else {
            (StatusCode::BAD_REQUEST, "Invalid command")
        }
    } else {
        (StatusCode::INTERNAL_SERVER_ERROR, "Sheet not initialized")
    }
}

// Existing load_csv_file and load_excel_file functions remain unchanged
fn load_csv_file(sheet: &mut Sheet, filename: &str) -> Result<(), String> {
    let file = File::open(filename).map_err(|e| format!("Failed to open CSV file: {}", e))?;
    let reader = io::BufReader::new(file);
    let mut row_idx = 0;
    for line in reader.lines() {
        if row_idx >= sheet.rows {
            return Err(format!("CSV file has more rows than the spreadsheet (max: {})", sheet.rows));
        }
        let line = line.map_err(|e| format!("Error reading CSV line: {}", e))?;
        let values: Vec<&str> = line.split(',').collect();
        let mut col_idx = 0;
        for value in values {
            if col_idx >= sheet.cols {
                return Err(format!("CSV file has more columns than the spreadsheet (max: {})", sheet.cols));
            }
            let value = value.trim();
            if let Ok(num_value) = value.parse::<i32>() {
                sheet.cells[row_idx as usize][col_idx as usize].value = num_value;
            } else if value.starts_with('=') {
                let formula = &value[1..];
                let mut cell_ref = String::new();
                crate::utils::encode_column(col_idx, &mut cell_ref);
                cell_ref.push_str(&(row_idx + 1).to_string());
                crate::cell::update_cell(sheet, &cell_ref, formula);
            } else if !value.is_empty() {
                sheet.cells[row_idx as usize][col_idx as usize].value = 0;
            }
            col_idx += 1;
        }
        row_idx += 1;
    }
    Ok(())
}

fn load_excel_file(sheet: &mut Sheet, filename: &str) -> Result<(), String> {
    let mut workbook: Xlsx<_> = open_workbook(filename)
        .map_err(|e| format!("Failed to open Excel file: {}", e))?;
    let sheet_names = workbook.sheet_names().to_vec();
    if sheet_names.is_empty() {
        return Err("Excel file doesn't contain any worksheets".to_string());
    }
    let worksheet = workbook.worksheet_range(&sheet_names[0])
        .ok_or_else(|| "Failed to get first worksheet".to_string())?
        .map_err(|e| format!("Error accessing worksheet: {}", e))?;
    let height = worksheet.height() as i32;
    let width = worksheet.width() as i32;
    if height > sheet.rows {
        return Err(format!("Excel file has more rows than the spreadsheet (file: {}, max: {})", height, sheet.rows));
    }
    if width > sheet.cols {
        return Err(format!("Excel file has more columns than the spreadsheet (file: {}, max: {})", width, sheet.cols));
    }
    for row_idx in 0..height {
        for col_idx in 0..width {
            match worksheet.get_value(((row_idx as usize).try_into().unwrap(), (col_idx as usize).try_into().unwrap())) {
                Some(DataType::Int(value)) => {
                    sheet.cells[row_idx as usize][col_idx as usize].value = *value as i32;
                },
                Some(DataType::Float(value)) => {
                    sheet.cells[row_idx as usize][col_idx as usize].value = *value as i32;
                },
                Some(DataType::String(value)) => {
                    if value.starts_with('=') {
                        let formula = &value[1..];
                        let mut cell_ref = String::new();
                        crate::utils::encode_column(col_idx, &mut cell_ref);
                        cell_ref.push_str(&(row_idx + 1).to_string());
                        crate::cell::update_cell(sheet, &cell_ref, formula);
                    } else {
                        sheet.cells[row_idx as usize][col_idx as usize].value = 0;
                    }
                },
                Some(DataType::Bool(value)) => {
                    sheet.cells[row_idx as usize][col_idx as usize].value = if *value { 1 } else { 0 };
                },
                Some(DataType::DateTime(_)) => {
                    sheet.cells[row_idx as usize][col_idx as usize].value = 0;
                },
                Some(DataType::Error(_)) => {
                    sheet.cells[row_idx as usize][col_idx as usize].value = 0;
                },
                None => {},
                _ => {
                    sheet.cells[row_idx as usize][col_idx as usize].value = 0;
                }
            }
        }
    }
    Ok(())
}

// Terminal interface logic
fn run_terminal_interface(rows: i32, cols: i32, extension_enabled: bool, input_file: Option<String>) {
    let mut sheet_guard = SHEET.lock().unwrap();
    *sheet_guard = create_sheet(rows, cols, extension_enabled);
    
    if extension_enabled {
        if let Some(filename) = input_file {
            if let Some(ref mut sheet) = *sheet_guard {
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

#[tokio::main]
async fn main() {
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
            row_col_args.pop();
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

    if extension_enabled {
        run_web_interface(rows, cols, extension_enabled, input_file).await;
    } else {
        run_terminal_interface(rows, cols, extension_enabled, input_file);
    }
}
