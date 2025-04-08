// sheet.rs
use std::io::{self, Write};
use crate::types::{Sheet, Cell};
// use crate::types::SHEET;
use crate::utils::{encode_column, parse_cell_reference};
use crate::cell::update_cell;

const DISPLAY_SIZE: i32 = 10;
const MAX_CELL_REF_LEN: usize = 10;

pub fn create_sheet(rows: i32, cols: i32) -> Option<Sheet> {
    let mut cells = Vec::with_capacity(rows as usize);
    for _ in 0..rows {
        let row: Vec<Cell> = vec![Cell {
            value: 0,
            formula: None,
            is_formula: false,
            is_error: false,
            dependencies: None,
            dependents: None,
            has_circular: false,
        }; cols as usize];
        cells.push(row);
    }

    Some(Sheet {
        cells,
        rows,
        cols,
        view_row: 0,
        view_col: 0,
        output_enabled: true,
        circular_dependency_detected: false,
    })
}

pub fn scroll_sheet(sheet: &mut Sheet, direction: char) {
    match direction {
        'w' => {
            if sheet.view_row > 0 && sheet.view_row - 10 >= 0 {
                sheet.view_row -= DISPLAY_SIZE;
            } else if sheet.view_row >= 0 {
                sheet.view_row = 0;
            }
        }
        's' => {
            if sheet.view_row + DISPLAY_SIZE < sheet.rows && sheet.view_row + 20 <= sheet.rows {
                sheet.view_row += DISPLAY_SIZE;
            } else if sheet.view_row + DISPLAY_SIZE < sheet.rows && sheet.view_row + 10 <= sheet.rows && sheet.view_row + 20 > sheet.rows {
                sheet.view_row += sheet.rows - sheet.view_row - 10;
            }
        }
        'a' => {
            if sheet.view_col - DISPLAY_SIZE >= 0 && sheet.view_col - 10 >= 0 {
                sheet.view_col -= DISPLAY_SIZE;
            } else if sheet.view_col >= 0 {
                sheet.view_col = 0;
            }
        }
        'd' => {
            if sheet.view_col + DISPLAY_SIZE < sheet.cols && sheet.view_col + 20 <= sheet.cols {
                sheet.view_col += DISPLAY_SIZE;
            } else if sheet.view_col + DISPLAY_SIZE < sheet.cols && sheet.view_col + 10 <= sheet.cols && sheet.view_col + 20 > sheet.cols {
                sheet.view_col += sheet.cols - sheet.view_col - 10;
            }
        }
        _ => {}
    }
}

pub fn scroll_to_cell(sheet: &mut Sheet, cell_ref: &str) {
    if let Some((row, col)) = parse_cell_reference(sheet,cell_ref) {
        sheet.view_row = row;
        sheet.view_col = col;
    } else {
        println!("Invalid cell reference for scroll");
    }
}

pub fn process_command(sheet: &mut Sheet, command: &str) {
    if command.is_empty() {
        return;
    }
   

    if command.len() == 1 {
        match command.chars().next().unwrap() {
            'w' => scroll_sheet(sheet, 'w'),
            'a' => scroll_sheet(sheet, 'a'),
            's' => scroll_sheet(sheet, 's'),
            'd' => scroll_sheet(sheet, 'd'),
            'q' => std::process::exit(0),
            _ => {}
        }
        return;
    }

    if command == "disable_output" {
        sheet.output_enabled = false;
        return;
    }
    if command == "enable_output" {
        sheet.output_enabled = true;
        return;
    }

    if command.starts_with("scroll_to ") {
        scroll_to_cell(sheet, &command[10..]);
        return;
    }

     if let Some((cell_ref, formula)) = command.split_once('=') {
        let cell_ref = cell_ref.trim();
        let formula = formula.trim();
        // println!("Processing: cell_ref={}, formula={}", cell_ref, formula);
       
        update_cell(sheet,cell_ref, formula);
        // println!("cell is updated");
    } else {
        // println!("Invalid command format");
    }
}

pub fn display_sheet(sheet: &Sheet) {
    if !sheet.output_enabled {
        return;
    }

    let mut max_widths = vec![0; DISPLAY_SIZE as usize];
    for j in sheet.view_col..(sheet.view_col + DISPLAY_SIZE).min(sheet.cols) {
        let mut col_header = String::new();
        encode_column(j, &mut col_header);
        max_widths[(j - sheet.view_col) as usize] = col_header.len();
    }

    for i in sheet.view_row..(sheet.view_row + DISPLAY_SIZE).min(sheet.rows) {
        for j in sheet.view_col..(sheet.view_col + DISPLAY_SIZE).min(sheet.cols) {
            let cell = &sheet.cells[i as usize][j as usize];
            let width = if cell.is_error && !cell.has_circular {
                3
            } else {
                let val = cell.value;
                if val == 0 { 1 } else { val.to_string().len() }
            };
            max_widths[(j - sheet.view_col) as usize] = max_widths[(j - sheet.view_col) as usize].max(width);
        }
    }

    print!("     ");
    for j in sheet.view_col..(sheet.view_col + DISPLAY_SIZE).min(sheet.cols) {
        let mut col_header = String::new();
        encode_column(j, &mut col_header);
        print!("{:width$} ", col_header, width = max_widths[(j - sheet.view_col) as usize]);
    }
    println!();

    for i in sheet.view_row..(sheet.view_row + DISPLAY_SIZE).min(sheet.rows) {
        print!("{:4} ", i + 1);
        for j in sheet.view_col..(sheet.view_col + DISPLAY_SIZE).min(sheet.cols) {
            let cell = &sheet.cells[i as usize][j as usize];
            let width = max_widths[(j - sheet.view_col) as usize];
            if cell.is_error && !cell.has_circular {
                print!("{:width$} ", "err", width = width);
            } else {
                print!("{:width$} ", cell.value, width = width);
            }
        }
        println!();
    }
    io::stdout().flush().unwrap();
}