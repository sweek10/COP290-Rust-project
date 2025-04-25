mod cell;
mod dependencies;
mod sheet;
#[cfg(test)]
mod tests;
mod types;
mod utils;

use crate::sheet::{create_sheet, display_sheet, process_command};
use crate::types::{Sheet, SHEET};
use crate::utils::{encode_column, is_valid_command};
use calamine::{open_workbook, Reader, Xlsx};
use rocket::{form::Form, get, post, response::Redirect};
use rocket_dyn_templates::Template;
use serde_json::json;
use std::fs::File;
use std::io::{self, BufRead, Write};
use std::path::Path;
use std::time::Instant;

const MAX_ROWS: i32 = 999;
const MAX_COLS: i32 = 18278;
const DISPLAY_SIZE: i32 = 10;

#[derive(rocket::form::FromForm)]
struct CommandForm {
    command: String,
}

/// Renders the main web interface for the spreadsheet.
///
/// This function generates the HTML template for the spreadsheet's web interface, displaying
/// a portion of the sheet based on the current view. It includes cell values, styling, and
/// any error messages.
///
/// # Arguments
/// * `message` - An optional message to display (e.g., error or success message).
///
/// # Returns
/// A `Template` containing the rendered HTML.
///
/// # Example
/// ```text
/// // Access the web interface at http://localhost:8000
/// // The interface shows a 10x10 grid of cells with values and styling
/// ```
#[cfg(not(tarpaulin_include))]
#[get("/?<message>")]
fn index(message: Option<String>) -> Template {
    let sheet = SHEET.lock().unwrap();
    let sheet = sheet.as_ref().unwrap();
    let view_row = sheet.view_row;
    let view_col = sheet.view_col;
    let rows = (view_row..(view_row + DISPLAY_SIZE).min(sheet.rows)).collect::<Vec<_>>();
    let columns = (view_col..(view_col + DISPLAY_SIZE).min(sheet.cols))
        .map(|col| {
            let mut col_str = String::new();
            encode_column(col, &mut col_str);
            col_str
        })
        .collect::<Vec<_>>();

    let rows_data = rows
        .iter()
        .map(|&row| {
            let cells = (view_col..(view_col + DISPLAY_SIZE).min(sheet.cols))
                .map(|col| {
                    let cell = &sheet.cells[row as usize][col as usize];
                    let value = if cell.is_error && !cell.has_circular {
                        "err".to_string()
                    } else {
                        cell.value.to_string()
                    };
                    let classes = {
                        let mut c = Vec::new();
                        if cell.is_bold {
                            c.push("bold");
                        }
                        if cell.is_italic {
                            c.push("italic");
                        }
                        if cell.is_underline {
                            c.push("underline");
                        }
                        c.join(" ")
                    };
                    json!({
                        "value": value,
                        "classes": classes,
                    })
                })
                .collect::<Vec<_>>();
            json!({
                "number": (row + 1).to_string(),
                "cells": cells,
            })
        })
        .collect::<Vec<_>>();

    Template::render(
        "index",
        json!({
            "columns": columns,
            "rows": rows_data,
            "circular_detected": sheet.circular_dependency_detected,
            "message": message,
        }),
    )
}

/// Processes a command submitted via the web interface.
///
/// This function handles commands submitted through the web form, processes them using
/// `process_command`, and redirects to the main page with an optional message.
///
/// # Arguments
/// * `form` - The form data containing the command string.
///
/// # Returns
/// A `Redirect` to the main page, possibly with a message.
///
/// # Example
/// ```text
/// // POST request to /command with form data "command=A1=5"
/// // Redirects to main page after setting A1 to 5
/// ```
#[cfg(not(tarpaulin_include))]
#[post("/command", data = "<form>")]
fn command(form: Form<CommandForm>) -> Redirect {
    let command = form.command.clone();
    let message = {
        let mut sheet = SHEET.lock().unwrap();
        if let Some(ref mut sheet) = *sheet {
            process_command(sheet, &command)
        } else {
            None
        }
    };
    if let Some(msg) = message {
        let encoded_msg = urlencoding::encode(&msg);
        Redirect::to(format!("/?message={}", encoded_msg))
    } else {
        Redirect::to("/")
    }
}

/// Scrolls the spreadsheet view in the specified direction.
///
/// This function processes a scroll command (w, a, s, d) submitted via a POST request
/// and updates the spreadsheet's view accordingly.
///
/// # Arguments
/// * `direction` - The direction to scroll (w, a, s, d).
///
/// # Returns
/// A `Redirect` to the main page.
///
/// # Example
/// ```text
/// // POST request to /scroll/s
/// // Scrolls the view down by 10 rows
/// ```
#[cfg(not(tarpaulin_include))]
#[post("/scroll/<direction>")]
fn scroll(direction: String) -> Redirect {
    let direction = direction.chars().next().unwrap_or(' ');
    if !['w', 'a', 's', 'd'].contains(&direction) {
        return Redirect::to("/");
    }
    let mut sheet = SHEET.lock().unwrap();
    if let Some(ref mut sheet) = *sheet {
        crate::sheet::scroll_sheet(sheet, direction);
    }
    Redirect::to("/")
}

/// Loads a CSV file into the spreadsheet.
///
/// This function reads a CSV file and populates the spreadsheet with its values.
/// Formulas (starting with '=') are processed after all values are loaded.
///
/// # Arguments
/// * `sheet` - A mutable reference to the spreadsheet.
/// * `filename` - The path to the CSV file.
///
/// # Returns
/// A `Result<(), String>` indicating success or an error message.
///
/// # Example
/// ```
/// let mut sheet = create_sheet(10, 10, true).unwrap();
/// // Assume "data.csv" contains "5,=A1+1"
/// load_csv_file(&mut sheet, "data.csv").unwrap();
/// assert_eq!(sheet.cells[0][0].value, 5);
/// assert_eq!(sheet.cells[0][1].value, 6);
/// ```
#[cfg(not(tarpaulin_include))]
fn load_csv_file(sheet: &mut Sheet, filename: &str) -> Result<(), String> {
    let file = File::open(filename).map_err(|e| format!("Failed to open CSV file: {}", e))?;
    let reader = io::BufReader::new(file);

    let mut formulas = Vec::new();
    for (row_idx_usize, line) in reader.lines().enumerate() {
        let row_idx = row_idx_usize as i32;
        if row_idx >= sheet.rows {
            return Err(format!(
                "CSV file has more rows than the spreadsheet (max: {})",
                sheet.rows
            ));
        }

        let line = line.map_err(|e| format!("Error reading CSV line: {}", e))?;
        let values: Vec<&str> = line.split(',').collect();

        for (col_idx_usize, value) in values.into_iter().enumerate() {
            let col_idx = col_idx_usize as i32;
            if col_idx >= sheet.cols {
                return Err(format!(
                    "CSV file has more columns than the spreadsheet (max: {})",
                    sheet.cols
                ));
            }

            let value = value.trim();
            if let Ok(num_value) = value.parse::<i32>() {
                sheet.cells[row_idx as usize][col_idx as usize].value = num_value;
            } else if let Some(stripped) = value.strip_prefix('=') {
                let formula = stripped.to_string();
                formulas.push((row_idx, col_idx, formula));
            } else if !value.is_empty() {
                sheet.cells[row_idx as usize][col_idx as usize].value = 0;
            }
        }
    }

    for (row, col, formula) in formulas {
        crate::cell::update_cell(sheet, row, col, &formula);
    }
    Ok(())
}

/// Loads an Excel file into the spreadsheet.
///
/// This function reads an Excel (.xlsx) file and populates the spreadsheet with its values.
/// It handles various data types (int, float, string, bool) and processes formulas.
///
/// # Arguments
/// * `sheet` - A mutable reference to the spreadsheet.
/// * `filename` - The path to the Excel file.
///
/// # Returns
/// A `Result<(), String>` indicating success or an error message.
///
/// # Example
/// ```
/// let mut sheet = create_sheet(10, 10, true).unwrap();
/// // Assume "data.xlsx" contains a sheet with "10" in A1 and "=A1+5" in B1
/// load_excel_file(&mut sheet, "data.xlsx").unwrap();
/// assert_eq!(sheet.cells[0][0].value, 10);
/// assert_eq!(sheet.cells[0][1].value, 15);
/// ```
#[cfg(not(tarpaulin_include))]
fn load_excel_file(sheet: &mut Sheet, filename: &str) -> Result<(), String> {
    let mut workbook: Xlsx<_> =
        open_workbook(filename).map_err(|e| format!("Failed to open Excel file: {}", e))?;

    let sheet_names = workbook.sheet_names().to_vec();
    if sheet_names.is_empty() {
        return Err("Excel file doesn't contain any worksheets".to_string());
    }

    let worksheet = workbook
        .worksheet_range(&sheet_names[0])
        .ok_or_else(|| "Failed to get first worksheet".to_string())?
        .map_err(|e| format!("Error accessing worksheet: {}", e))?;

    let height = worksheet.height() as i32;
    let width = worksheet.width() as i32;

    if height > sheet.rows {
        return Err(format!(
            "Excel file has more rows than the spreadsheet (file: {}, max: {})",
            height, sheet.rows
        ));
    }

    if width > sheet.cols {
        return Err(format!(
            "Excel file has more columns than the spreadsheet (file: {}, max: {})",
            width, sheet.cols
        ));
    }

    for row_idx in 0..height {
        for col_idx in 0..width {
            match worksheet.get_value((
                (row_idx as usize).try_into().unwrap(),
                (col_idx as usize).try_into().unwrap(),
            )) {
                Some(calamine::DataType::Int(value)) => {
                    sheet.cells[row_idx as usize][col_idx as usize].value = *value as i32;
                }
                Some(calamine::DataType::Float(value)) => {
                    sheet.cells[row_idx as usize][col_idx as usize].value = *value as i32;
                }
                Some(calamine::DataType::String(value)) => {
                    if let Some(stripped) = value.strip_prefix('=') {
                        let formula = &stripped;
                        crate::cell::update_cell(sheet, row_idx, col_idx, formula);
                    } else {
                        sheet.cells[row_idx as usize][col_idx as usize].value = 0;
                    }
                }
                Some(calamine::DataType::Bool(value)) => {
                    sheet.cells[row_idx as usize][col_idx as usize].value =
                        if *value { 1 } else { 0 };
                }
                _ => {
                    sheet.cells[row_idx as usize][col_idx as usize].value = 0;
                }
            }
        }
    }
    Ok(())
}

/// The main entry point for the spreadsheet application.
///
/// This function initializes the spreadsheet, processes command-line arguments, and either
/// starts a web server (if extensions are enabled) or runs a terminal-based interface.
/// It also handles loading input files if provided.
///
/// # Returns
/// A `Result<(), rocket::Error>` indicating whether the application ran successfully.
///
/// # Example
/// ```text
/// // Run with: cargo run -- 10 10
/// // Creates a 10x10 spreadsheet and starts the terminal interface
/// // Or with: cargo run --extension 10 10 data.csv
/// // Loads data.csv and starts the web server
/// ```
#[cfg(not(tarpaulin_include))]
#[rocket::main]
async fn main() -> Result<(), rocket::Error> {
    let args: Vec<String> = std::env::args().collect();
    let mut extension_enabled = false;
    let mut row_col_args = Vec::new();
    let mut input_file = None;

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

    if extension_enabled && row_col_args.len() == 3 {
        let potential_file = &row_col_args[2];
        if Path::new(potential_file).exists() {
            input_file = Some(potential_file.clone());
            row_col_args.pop();
        }
    }

    if row_col_args.len() != 2 {
        println!(
            "Usage: {} [--extension] <rows> <columns> [input_file.csv|xlsx]",
            args[0]
        );
        println!("Note: File loading is only available with --extension flag");
        return Ok(());
    }

    let rows: i32 = row_col_args[0].parse().unwrap_or(0);
    let cols: i32 = row_col_args[1].parse().unwrap_or(0);

    if !(1..=MAX_ROWS).contains(&rows) || !(1..=MAX_COLS).contains(&cols) {
        println!(
            "Invalid dimensions. Rows: 1-{}, Columns: 1-{}",
            MAX_ROWS, MAX_COLS
        );
        return Ok(());
    }

    {
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
                        _ => Err(format!("Unsupported file format: {}", extension)),
                    };

                    match result {
                        Ok(_) => println!("Successfully loaded file: {}", filename),
                        Err(e) => println!("Error loading file: {}", e),
                    }
                }
            }
        }
    }

    if extension_enabled {
        rocket::build()
            .configure(rocket::Config::figment()
                .merge(("port", 1835)))  // Configure to use port 1835 because port 80 is privileged and can only be accessed by using sudo.
            .mount("/", rocket::routes![index, command, scroll])
            .attach(Template::fairing())
            .launch()
            .await?;
    } else {
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

            print!(
                "[{:.1}] {}> ",
                elapsed_time,
                if is_valid {
                    if SHEET
                        .lock()
                        .unwrap()
                        .as_ref()
                        .unwrap()
                        .circular_dependency_detected
                    {
                        "(err)"
                    } else {
                        "(ok)"
                    }
                } else {
                    "(err)"
                }
            );
            io::stdout().flush().unwrap();

            let mut command = String::new();
            if stdin.lock().read_line(&mut command).is_err() {
                break;
            }
            let command = command.trim();

            if command == "q" {
                break;
            }
            is_valid = is_valid_command(SHEET.lock().unwrap().as_mut().unwrap(), command);
            let start = Instant::now();
            let message = {
                let mut sheet_guard = SHEET.lock().unwrap();
                if let Some(ref mut sheet) = *sheet_guard {
                    process_command(sheet, command)
                } else {
                    None
                }
            };
            if let Some(msg) = message {
                println!("{}", msg);
            }
            elapsed_time = start.elapsed().as_secs_f64();
        }
    }
    Ok(())
}
