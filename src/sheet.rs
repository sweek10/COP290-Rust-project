// sheet.rs
use std::io::{self, Write};
use crate::types::{Sheet, Cell,PatternType};
use crate::utils::{encode_column, parse_cell_reference, parse_range, detect_pattern};

use crate::cell::update_cell;

const DISPLAY_SIZE: i32 = 10;
const MAX_CELL_REF_LEN: usize = 10;

pub fn create_sheet(rows: i32, cols: i32, extension_enabled: bool) -> Option<Sheet> {
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
            is_bold: false,
            is_italic: false,
            is_underline:false,
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
        extension_enabled,
        command_history: Vec::with_capacity(10),
        command_position: 0,
        max_history_size: 10,
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

// Add to sheet.rs
pub fn add_to_history(sheet: &mut Sheet, command: &str) {
    // Don't record navigation commands, undo, redo, or display commands
    if command.len() == 1 && "wasd".contains(command) || 
       command == "undo" || 
       command == "redo" ||
       command == "disable_output" || 
       command == "enable_output" ||
       command.starts_with("scroll_to ") ||
       command.contains("AUTOFILL") {
        return;
    }

    // If we're not at the end of the history, truncate it
    if sheet.command_position < sheet.command_history.len() {
        sheet.command_history.truncate(sheet.command_position);
    }

    // Add the new command
    sheet.command_history.push(command.to_string());
    
    // Limit history size
    if sheet.command_history.len() > sheet.max_history_size {
        sheet.command_history.remove(0);
    } else {
        sheet.command_position += 1;
    }
}

pub fn undo(sheet: &mut Sheet) -> bool {
    if sheet.command_position == 0 || sheet.command_history.is_empty() {
        return false; // Nothing to undo
    }

    sheet.command_position -= 1;
    rebuild_sheet_state(sheet);
    true
}

pub fn redo(sheet: &mut Sheet) -> bool {
    if sheet.command_position >= sheet.command_history.len() {
        return false; // Nothing to redo
    }

    sheet.command_position += 1;
    rebuild_sheet_state(sheet);
    true
}

fn rebuild_sheet_state(sheet: &mut Sheet) {
    // Save current view position
    let view_row = sheet.view_row;
    let view_col = sheet.view_col;
    let output_enabled = sheet.output_enabled;
    
    // Clear all cells
    for row in &mut sheet.cells {
        for cell in row {
            *cell = Cell::new();
        }
    }
    
    // Reset circular dependency flag
    sheet.circular_dependency_detected = false;
    
    // Replay commands up to current position
    let commands_to_replay: Vec<String> = sheet.command_history[0..sheet.command_position].to_vec();
    
    for cmd in commands_to_replay {
        if let Some((cell_ref, formula)) = cmd.split_once('=') {
            let cell_ref = cell_ref.trim();
            let formula = formula.trim();
            update_cell(sheet, cell_ref, formula);
        }
    }
    
    // Restore view position
    sheet.view_row = view_row;
    sheet.view_col = view_col;
    sheet.output_enabled = output_enabled;
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
    
    // Handle undo/redo commands if extension is enabled
    if sheet.extension_enabled {
        if command == "undo" {
            undo(sheet);
            return;
        }
        if command == "redo" {
            redo(sheet);
            return;
        }

        if command.starts_with("FORMULA ") {
            let cell_ref = command[8..].trim();
            if let Some((row, col)) = parse_cell_reference(sheet, cell_ref) {
                let cell = &sheet.cells[row as usize][col as usize];
                if let Some(formula) = &cell.formula {
                    println!("Formula in cell {}: {}", cell_ref, formula);
                } else {
                    println!("No formula stored in cell {}", cell_ref);
                }
            } else {
                println!("Invalid cell reference: {}", cell_ref);
            }
            return;
        }
    }

    if command.starts_with("scroll_to ") {
        scroll_to_cell(sheet, &command[10..]);
        return;
    }
    
    if let Some((cell_ref, formula)) = command.split_once('=') {
        let cell_ref = cell_ref.trim();
        let formula = formula.trim();
        
        // For cell updates, store command in history if extension is enabled
        if sheet.extension_enabled {
            add_to_history(sheet, command);
        }

        if (sheet.extension_enabled) {
        if let Some((func_name, args)) = formula.split_once('(') {
            if let Some(range_arg) = args.strip_suffix(')') {
                let range_arg = range_arg.trim();
                if func_name.trim().to_uppercase() == "AUTOFILL" {
                    if let Some((start_row, start_col, end_row, end_col)) = parse_range(sheet, range_arg) {
                        // Ensure the range is a single column or row for simplicity
                        if start_col != end_col && start_row != end_row {
                            return; // Only support single column or row for now
                        }

                        if start_col == end_col { // Vertical autofill
                            let mut values = Vec::new();
                            for i in (0.max(start_row - 5)..start_row).rev() { // Fixed j to i
                                values.push(sheet.cells[i as usize][start_col as usize].value);
                            }
                            if values.is_empty() {
                                return;
                            }
                            match detect_pattern(sheet, start_row, start_col) {
                                PatternType::Constant(value) => {
                                    for i in start_row..=end_row {
                                        let cell = &mut sheet.cells[i as usize][start_col as usize];
                                        cell.value = value;
                                        cell.formula = None;
                                        cell.is_formula = false;
                                        cell.is_error = false;
                                    }
                                }
                                PatternType::Arithmetic(_initial, diff) => {
                                    let last_value = values[0]; // Last known value (A2=2)
                                    for i in start_row..=end_row {
                                        let offset = i - (start_row -1); // Offset from the last known cell (A2)
                                        let new_value = last_value - diff * offset;
                                        let cell = &mut sheet.cells[i as usize][start_col as usize];
                                        cell.value = new_value;
                                        cell.formula = None;
                                        cell.is_formula = false;
                                        cell.is_error = false;
                                    }
                                }
                                PatternType::Fibonacci(mut penult, mut last) => {
                                    for i in start_row..=end_row {
                                        let new_value = penult + last;
                                        let cell = &mut sheet.cells[i as usize][start_col as usize];
                                        cell.value = new_value;
                                        cell.formula = None;
                                        cell.is_formula = false;
                                        cell.is_error = false;
                                        penult = last;
                                        last = new_value;
                                    }
                                }
                                PatternType::Unknown => {}
                            }
                        }else if start_row == end_row { // Horizontal autofill
                            let mut values = Vec::new();
    for j in (0.max(start_col - 5)..start_col).rev() {
        values.push(sheet.cells[start_row as usize][j as usize].value);
    }
    if values.is_empty() {
        return; // No values to work with
    }
                            match detect_pattern(sheet, start_row, start_col) {
                                PatternType::Constant(value) => {
                                    for j in start_col..=end_col {
                                        let cell = &mut sheet.cells[start_row as usize][j as usize];
                                        cell.value = value;
                                        cell.formula = None;
                                        cell.is_formula = false;
                                        cell.is_error = false;
                                    }
                                }
                                PatternType::Arithmetic(initial, diff) => {
                                    for j in start_col..=end_col {
                                        let offset = j - (start_col - values.len() as i32);
                                        let new_value = initial + diff * offset;
                                        let cell = &mut sheet.cells[start_row as usize][j as usize];
                                        cell.value = new_value;
                                        cell.formula = None;
                                        cell.is_formula = false;
                                        cell.is_error = false;
                                    }
                                }
                                PatternType::Fibonacci(mut penult, mut last) => {
                                    for j in start_col..=end_col {
                                        let new_value = penult + last;
                                        let cell = &mut sheet.cells[start_row as usize][j as usize];
                                        cell.value = new_value;
                                        cell.formula = None;
                                        cell.is_formula = false;
                                        cell.is_error = false;
                                        penult = last;
                                        last = new_value;
                                    }
                                }
                                PatternType::Unknown => {} // Do nothing if no pattern detected
                            }
                        }
                        return;
                    }
                } else if let Some(cell_arg) = args.strip_suffix(')') {
                    let cell_arg = cell_arg.trim();
                    if let Some((row, col)) = parse_cell_reference(sheet, cell_arg) {
                        match func_name.trim().to_uppercase().as_str() {
                            "BOLD" => {
                                sheet.cells[row as usize][col as usize].is_bold = true;
                                return;
                            }
                            "ITALIC" => {
                                sheet.cells[row as usize][col as usize].is_italic = true;
                                return;
                            }
                            "UNDERLINE" => {
                                sheet.cells[row as usize][col as usize].is_underline = true;
                                return;
                            }
                            _ => {}
                        }
                    }
                }
            }
        }
    }

        update_cell(sheet, cell_ref, formula);
    
 
 } else {println!{"invalid command format"};}
}

// ... (display_sheet unchanged) ...
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
        print!("{:4} ", i + 1); // Row number
        for j in sheet.view_col..(sheet.view_col + DISPLAY_SIZE).min(sheet.cols) {
            let cell = &sheet.cells[i as usize][j as usize];
            let width = max_widths[(j - sheet.view_col) as usize];
            let value_str = if cell.is_error && !cell.has_circular {
                "err".to_string()
            } else {
                cell.value.to_string()
            };

            // Apply ANSI formatting
            let mut formatted = String::new();
            if cell.is_bold {
                formatted.push_str("\x1b[1m"); // Bold
            }
            if cell.is_italic {
                formatted.push_str("\x1b[3m"); // Italic
            }
            if cell.is_underline {
                formatted.push_str("\x1b[4m"); // Underline
            }
            formatted.push_str(&value_str);
            if cell.is_bold || cell.is_italic || cell.is_underline {
                formatted.push_str("\x1b[0m"); // Reset formatting
            }

            // Print with proper padding (right-aligned)
            print!("{:>width$} ", formatted, width = width);
        }
        println!();
    }
    io::stdout().flush().unwrap();
}
