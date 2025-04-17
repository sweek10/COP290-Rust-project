// sheet.rs
use std::io::{self, Write};
use crate::types::{Sheet, Cell,PatternType};
use crate::utils::{encode_column, parse_cell_reference, parse_range, detect_pattern};
use crate::types::GraphType;
use crate::types::{Clipboard, CLIPBOARD};
use crate::utils::factorial;
use crate::utils::triangular;
use crate::cell::update_cell;

const DISPLAY_SIZE: i32 = 10;
// const MAX_CELL_REF_LEN: usize = 10;

pub fn create_sheet(rows: i32, cols: i32, extension_enabled: bool) -> Option<Sheet> {
    let mut cells = Vec::with_capacity(rows as usize);
    for _ in 0..rows {
        let row: Vec<Cell> = vec![Cell {
            value: 0,
            formula: None,
            is_formula: false,
            is_error: false,
            dependencies: Vec::new(),
            dependents: Vec::new(),
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

        if command.starts_with("ROWDEL ") {
            let row_str = &command[7..].trim();
            if let Ok(row) = row_str.parse::<i32>() {
                if row >= 1 && row <= sheet.rows {
                    // Delete all contents of the row (set to initial state)
                    for col in 0..sheet.cols {
                        let cell = &mut sheet.cells[(row-1) as usize][col as usize];
                        cell.value = 0;
                        cell.formula = None;
                        cell.is_formula = false;
                        cell.is_error = false;
                        cell.dependencies = Vec::new();
                        cell.is_bold = false;
                        cell.is_italic = false;
                        cell.is_underline = false;
                    }
                    add_to_history(sheet, command);
                }
            }
            return;
        }
        
        if command.starts_with("COLDEL ") {
            let col_str = &command[7..].trim();
            if !col_str.is_empty() && col_str.chars().all(|c| c.is_ascii_alphabetic()) {
                if let Some((_, col)) = parse_cell_reference(sheet, &format!("{}1", col_str)) {
                    // Delete all contents of the column (set to initial state)
                    for row in 0..sheet.rows {
                        let cell = &mut sheet.cells[row as usize][col as usize];
                        cell.value = 0;
                        cell.formula = None;
                        cell.is_formula = false;
                        cell.is_error = false;
                        cell.dependencies = Vec::new();
                        cell.is_bold = false;
                        cell.is_italic = false;
                        cell.is_underline = false;
                    }
                    add_to_history(sheet, command);
                }
            }
            return;
        }
    

    

    if command.starts_with("COPY") {
        let range = if command.starts_with("COPY ") {
            &command[5..]
        } else {
            &command[4..]
        };
        if copy_range(sheet, range) {
            if sheet.output_enabled {
                println!("Copied to clipboard");
            }
        } else {
            println!("Invalid range for copy");
        }
        return;
    }
    
    if command.starts_with("CUT") {
        let range = if command.starts_with("CUT ") {
            &command[4..]
        } else {
            &command[3..]
        };
        if cut_range(sheet, range) {
            if sheet.output_enabled {
                println!("Cut to clipboard");
            }
        } else {
            println!("Invalid range for cut");
        }
        return;
    }
    
    if command.starts_with("PASTE") {
        let cell_ref = if command.starts_with("PASTE ") {
            &command[6..]
        } else {
            &command[5..]
        };
        if paste_range(sheet, cell_ref) {
            if sheet.output_enabled {
                println!("Pasted from clipboard");
                display_sheet(sheet);
            }
        } else {
            println!("Nothing to paste or invalid target");
        }
        return;
    }

    if command.starts_with("GRAPH ") {
        let parts: Vec<&str> = command.split_whitespace().collect();
        if parts.len() == 3 {
            let graph_type = match parts[1].to_uppercase().as_str() {
                "(BAR)" => GraphType::Bar,
                "(SCATTER)" => GraphType::Scatter,
                _ => {
                    println!("Invalid graph type. Use BAR or LINE");
                    return;
                }
            };
            display_graph(sheet, graph_type, parts[2]);
        } else {
            println!("Usage: GRAPH <type> <range> (e.g., GRAPH BAR A1:A10)");
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

        if sheet.extension_enabled {
        if let Some((func_name, args)) = formula.split_once('(') {
            if let Some(range_arg) = args.strip_suffix(')') {
                let range_arg = range_arg.trim();

                if func_name.trim().to_uppercase() == "SORTA" || func_name.trim().to_uppercase() == "SORTD" {
                    if let Some((start_row, start_col, end_row, end_col)) = parse_range(sheet, range_arg) {
                        // Check if it's a single column or single row
                        if start_col == end_col { // Single column
                            // Collect values from the column
                            let mut values: Vec<(i32, i32)> = Vec::new();
                            for i in start_row..=end_row {
                                values.push((i, sheet.cells[i as usize][start_col as usize].value));
                            }
                            
                            // Sort the values
                            if func_name.trim().to_uppercase() == "SORTA" {
                                values.sort_by(|a, b| a.1.cmp(&b.1));
                            } else {
                                values.sort_by(|a, b| b.1.cmp(&a.1));
                            }
                            
                            // Create a map of original values for storing
                            let mut original_values: Vec<(i32, Option<String>, bool, bool, bool, bool, bool)> = Vec::new();
                            for i in start_row..=end_row {
                                let cell = &sheet.cells[i as usize][start_col as usize];
                                original_values.push((
                                    cell.value,
                                    cell.formula.clone(),
                                    cell.is_formula,
                                    cell.is_error,
                                    cell.is_bold,
                                    cell.is_italic,
                                    cell.is_underline
                                ));
                            }
                            
                            // Update cells with sorted values
                            for (idx, (orig_row, value)) in values.iter().enumerate() {
                                let i = start_row + idx as i32;
                                let orig_idx = (orig_row - start_row) as usize;
                                
                                let cell = &mut sheet.cells[i as usize][start_col as usize];
                                cell.value = *value;
                                cell.formula = original_values[orig_idx].1.clone();
                                cell.is_formula = original_values[orig_idx].2;
                                cell.is_error = original_values[orig_idx].3;
                                cell.is_bold = original_values[orig_idx].4;
                                cell.is_italic = original_values[orig_idx].5;
                                cell.is_underline = original_values[orig_idx].6;
                            }
                        } else if start_row == end_row { // Single row
                            // Collect values from the row
                            let mut values: Vec<(i32, i32)> = Vec::new();
                            for j in start_col..=end_col {
                                values.push((j, sheet.cells[start_row as usize][j as usize].value));
                            }
                            
                            // Sort the values
                            if func_name.trim().to_uppercase() == "SORTA" {
                                values.sort_by(|a, b| a.1.cmp(&b.1));
                            } else {
                                values.sort_by(|a, b| b.1.cmp(&a.1));
                            }
                            
                            // Create a map of original values and properties
                            let mut original_values: Vec<(i32, Option<String>, bool, bool, bool, bool, bool)> = Vec::new();
                            for j in start_col..=end_col {
                                let cell = &sheet.cells[start_row as usize][j as usize];
                                original_values.push((
                                    cell.value,
                                    cell.formula.clone(),
                                    cell.is_formula,
                                    cell.is_error,
                                    cell.is_bold,
                                    cell.is_italic,
                                    cell.is_underline
                                ));
                            }
                            
                            // Update cells with sorted values
                            for (idx, (orig_col, value)) in values.iter().enumerate() {
                                let j = start_col + idx as i32;
                                let orig_idx = (orig_col - start_col) as usize;
                                
                                let cell = &mut sheet.cells[start_row as usize][j as usize];
                                cell.value = *value;
                                cell.formula = original_values[orig_idx].1.clone();
                                cell.is_formula = original_values[orig_idx].2;
                                cell.is_error = original_values[orig_idx].3;
                                cell.is_bold = original_values[orig_idx].4;
                                cell.is_italic = original_values[orig_idx].5;
                                cell.is_underline = original_values[orig_idx].6;
                            }
                        } else {
                            // For 2D ranges, we can collect all values, sort them, and reassign
                            // in row-major or column-major order
                            let mut all_values: Vec<i32> = Vec::new();
                            for i in start_row..=end_row {
                                for j in start_col..=end_col {
                                    all_values.push(sheet.cells[i as usize][j as usize].value);
                                }
                            }
                            
                            // Sort all values
                            if func_name.trim().to_uppercase() == "SORTA" {
                                all_values.sort();
                            } else {
                                all_values.sort_by(|a, b| b.cmp(a));
                            }
                            
                            // Update cells with sorted values (row-major order)
                            let mut idx = 0;
                            for i in start_row..=end_row {
                                for j in start_col..=end_col {
                                    if idx < all_values.len() {
                                        let cell = &mut sheet.cells[i as usize][j as usize];
                                        cell.value = all_values[idx];
                                        // Clear formula for sorted cells
                                        cell.formula = None;
                                        cell.is_formula = false;
                                        cell.is_error = false;
                                        idx += 1;
                                    }
                                }
                            }
                        }
                        return;
                    }
                }
                
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
                            println!("Vertical values: {:?}", values);
                            if values.is_empty() {
                                println!("no value");
                                return;
                            }
                            let pattern = detect_pattern(sheet, start_row, start_col, end_row, end_col);
                             println!("Detected pattern: {:?}", pattern); // Debug
                            match pattern {
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
                                PatternType::Geometric(_initial, ratio) => {
                                    let last_value = values[0];
                                    for i in start_row..=end_row {
                                        let offset = i - (start_row-1);
                                        let new_value = (last_value as f64 * ratio.powi(offset)).round() as i32;
                                        let cell = &mut sheet.cells[i as usize][start_col as usize];
                                        cell.value = new_value;
                                        cell.formula = None;
                                        cell.is_formula = false;
                                        cell.is_error = false;
                                    }
                                }
                                PatternType::Factorial(_last_value, mut next_index) => {
                                    for i in start_row..=end_row {
                                        let new_value = factorial(next_index);
                                        let cell = &mut sheet.cells[i as usize][start_col as usize];
                                        cell.value = new_value;
                                        cell.formula = None;
                                        cell.is_formula = false;
                                        cell.is_error = false;
                                        next_index += 1;
                                    }
                                }
                                PatternType::Triangular(_last_value, mut next_index) => {
                                    for i in start_row..=end_row {
                                        let new_value = triangular(next_index);
                                        let cell = &mut sheet.cells[i as usize][start_col as usize];
                                        cell.value = new_value;
                                        cell.formula = None;
                                        cell.is_formula = false;
                                        cell.is_error = false;
                                        next_index += 1;
                                    }
                                }
                                PatternType::Unknown => {}
                            }
                        }else if start_row == end_row { // Horizontal autofill
                            let mut values = Vec::new();
    for j in (0.max(start_col - 5)..start_col).rev() {
        values.push(sheet.cells[start_row as usize][j as usize].value);
    }
    // println!("Horizontal values: {:?}", values);
    if values.is_empty() {
        // println!("No values for pattern detection");
        return;
    }
    let pattern = detect_pattern(sheet, start_row, start_col, end_row, end_col);
    // println!("Detected pattern: {:?}", pattern);
                            match pattern {
                                PatternType::Constant(value) => {
                                    for j in start_col..=end_col {
                                        let cell = &mut sheet.cells[start_row as usize][j as usize];
                                        cell.value = value;
                                        cell.formula = None;
                                        cell.is_formula = false;
                                        cell.is_error = false;
                                    }
                                }
                                PatternType::Arithmetic(_, diff) => {
                                    let last_value = values[0]; 
                                    for j in start_col..=end_col {
                                        let offset = j - (start_col-1);
                                        let new_value = last_value - diff * (offset); // Fix: Add diff, adjust offset
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
                                PatternType::Geometric(_, ratio) => {
                                    let last_value = values[0];
                                    for j in start_col..=end_col { // Fix: Use j instead of i
                                        let offset = j - (start_col - 1);
                                        let new_value = (last_value as f64 * ratio.powi(offset)).round() as i32;
                                        let cell = &mut sheet.cells[start_row as usize][j as usize];
                                        cell.value = new_value;
                                        cell.formula = None;
                                        cell.is_formula = false;
                                        cell.is_error = false;
                                    }
                                }
                                PatternType::Factorial(_last_value, mut next_index) => {
                                    for j in start_col..=end_col {
                                        let new_value = factorial(next_index);
                                        let cell = &mut sheet.cells[start_row as usize][j as usize];
                                        cell.value = new_value;
                                        cell.formula = None;
                                        cell.is_formula = false;
                                        cell.is_error = false;
                                        next_index += 1;
                                    }
                                }
                                PatternType::Triangular(_last_value, mut next_index) => {
                                    for j in start_col..=end_col {
                                        let new_value = triangular(next_index);
                                        let cell = &mut sheet.cells[start_row as usize][j as usize];
                                        cell.value = new_value;
                                        cell.formula = None;
                                        cell.is_formula = false;
                                        cell.is_error = false;
                                        next_index += 1;
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

pub fn display_graph(sheet: &mut Sheet, graph_type: GraphType, range: &str) {
    if let Some((start_row, start_col, end_row, end_col)) = parse_range(sheet, range) {
        // Collect values from the range
        let mut values = Vec::new();
        let mut labels = Vec::new();
        
        for i in start_row..=end_row {
            for j in start_col..=end_col {
                let cell = &sheet.cells[i as usize][j as usize];
                values.push(cell.value);
                
                // Create label (e.g., "A1")
                let mut label = String::new();
                encode_column(j, &mut label);
                label.push_str(&(i + 1).to_string());
                labels.push(label);
            }
        }
        
        // Find max value for y-axis scale
        let max_val = *values.iter().filter(|&&v| v > 0).max().unwrap_or(&10);
        let max_label_width = labels.iter().map(|l| l.len()).max().unwrap_or(2);
        let column_width = max_label_width.max(3) + 1; // Ensure minimum width of 3 for bars

        match graph_type {
            GraphType::Bar => {
                println!("\nBar Graph for range {}:", range);
                
                // Print vertical axis and bars
                for value in (1..=max_val).rev() {
                    print!("{:2} |", value);
                    
                    for &cell_value in &values {
                        if cell_value >= value {
                            print!("{:^width$}", "█", width = column_width);
                        } else {
                            print!("{:^width$}", " ", width = column_width);
                        }
                    }
                    println!();
                }
                
                // Print bottom separator
                print!("---+");
                for _ in &values {
                    print!("{}", "-".repeat(column_width));
                }
                println!();
                
                // Print x-axis labels
                print!("   |");
                for label in &labels {
                    print!("{:^width$}", label, width = column_width);
                }
                println!("\n");
            },
            
            GraphType::Scatter => {
                println!("\nScatter Plot for range {}:", range);
                
                // Print vertical axis and points
                for value in (1..=max_val).rev() {
                    print!("{:2} |", value);
                    
                    for &cell_value in &values {
                        let center = column_width / 2;
                        
                        if cell_value == value {
                            // Print point centered in column
                            print!("{:width$}", " ", width = center);
                            print!("*");
                            print!("{:width$}", " ", width = column_width - center - 1);
                        } else {
                            // Empty space
                            print!("{:^width$}", " ", width = column_width);
                        }
                    }
                    println!();
                }
                
                // Print bottom separator
                print!("---+");
                for _ in &values {
                    print!("{}", "-".repeat(column_width));
                }
                println!();
                
                // Print x-axis labels
                print!("   |");
                for label in &labels {
                    print!("{:^width$}", label, width = column_width);
                }
                println!("\n");
            }
        }
    } else {
        println!("Invalid range specified for graph");
    }
}

impl Sheet {
    pub fn get_cell_range(&self, start_row: i32, start_col: i32, end_row: i32, end_col: i32) -> Vec<Vec<Cell>> {
        let mut range = Vec::new();
        for i in start_row..=end_row {
            let mut row = Vec::new();
            for j in start_col..=end_col {
                row.push(self.cells[i as usize][j as usize].clone());
            }
            range.push(row);
        }
        range
    }

    pub fn set_cell_range(&mut self, start_row: i32, start_col: i32, values: &[Vec<Cell>]) {
        for (i, row) in values.iter().enumerate() {
            for (j, cell) in row.iter().enumerate() {
                let target_row = start_row + i as i32;
                let target_col = start_col + j as i32;
                if target_row < self.rows && target_col < self.cols {
                    self.cells[target_row as usize][target_col as usize] = cell.clone();
                    // Recalculate if it's a formula cell
                    if let Some(formula) = &cell.formula {
                        let mut cell_ref = String::new();
                        encode_column(target_col, &mut cell_ref);
                        cell_ref.push_str(&(target_row + 1).to_string());
                        update_cell(self, &cell_ref, formula);
                    }
                }
            }
        }
    }
}

pub fn copy_range(sheet: &mut Sheet, range: &str) -> bool {
    if let Some((start_row, start_col, end_row, end_col)) = parse_range(sheet, range) {
        let contents = sheet.get_cell_range(start_row, start_col, end_row, end_col);
        *CLIPBOARD.lock().unwrap() = Some(Clipboard {
            contents,
            is_cut: false,
            source_range: None,
        });
        true
    } else {
        false
    }
}

pub fn cut_range(sheet: &mut Sheet, range: &str) -> bool {
    if let Some((start_row, start_col, end_row, end_col)) = parse_range(sheet, range) {
        let contents = sheet.get_cell_range(start_row, start_col, end_row, end_col);
        
        // Clear the source cells
        for i in start_row..=end_row {
            for j in start_col..=end_col {
                let cell = &mut sheet.cells[i as usize][j as usize];
                *cell = Cell::new(); // Reset to empty cell
            }
        }
        
        *CLIPBOARD.lock().unwrap() = Some(Clipboard {
            contents,
            is_cut: true,
            source_range: Some((start_row, start_col, end_row, end_col)),
        });
        true
    } else {
        false
    }
}

pub fn paste_range(sheet: &mut Sheet, cell_ref: &str) -> bool {
    // Create a scope for the clipboard lock to ensure it's released before display_sheet
    let success = {
        let mut clipboard = CLIPBOARD.lock().unwrap();
        if let Some(clipboard_data) = &*clipboard {
            if let Some((start_row, start_col)) = parse_cell_reference(sheet, cell_ref) {
                sheet.set_cell_range(start_row, start_col, &clipboard_data.contents);
                
                // If this was a cut operation, clear the clipboard after paste
                if clipboard_data.is_cut {
                    *clipboard = None;
                }
                true
            } else {
                false
            }
        } else {
            false
        }
    }; 
    
    if success {
        display_sheet(sheet);
    }
    
    success
}
