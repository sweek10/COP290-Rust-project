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

const MAX_ROWS: i32 = 999;
const MAX_COLS: i32 = 18278;

fn main() {
    let args: Vec<String> = std::env::args().collect();
    if args.len() != 3 {
        println!("Usage: {} <rows> <columns>", args[0]);
        return;
    }

    let rows: i32 = args[1].parse().unwrap_or(0);
    let cols: i32 = args[2].parse().unwrap_or(0);

    if rows < 1 || rows > MAX_ROWS || cols < 1 || cols > MAX_COLS {
        println!("Invalid dimensions. Rows: 1-{}, Columns: 1-{}", MAX_ROWS, MAX_COLS);
        return;
    }

    let mut sheet_guard = SHEET.lock().unwrap();
    *sheet_guard = create_sheet(rows, cols);
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
    //    println!("Command received: {}", command);
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
