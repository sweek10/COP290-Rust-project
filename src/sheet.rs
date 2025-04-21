use std::io::{self, Write};
use crate::types::{Sheet, Cell, PatternType, GraphType, Clipboard, CLIPBOARD, DependencyType};
use crate::utils::{encode_column, parse_cell_reference, parse_range, detect_pattern, factorial, triangular};
use crate::cell::update_cell;
use std::collections::HashMap;
use crate::dependencies::remove_dependency;

const DISPLAY_SIZE: i32 = 10;


pub fn create_sheet(rows: i32, cols: i32, extension_enabled: bool) -> Option<Sheet> {
    let mut cells = Vec::with_capacity(rows as usize);
    for _ in 0..rows {
        let row: Vec<Cell> = vec![Cell::new(); cols as usize];
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
        dependency_graph: HashMap::new(),
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

pub fn scroll_to_cell(sheet: &mut Sheet, row: i32, col: i32) {
    if row >= 0 && row < sheet.rows && col >= 0 && col < sheet.cols {
        sheet.view_row = row;
        sheet.view_col = col;
    } else {
        println!("Invalid cell coordinates for scroll");
    }
}

pub fn add_to_history(sheet: &mut Sheet, command: &str) {
    if command.len() == 1 && "wasd".contains(command) || 
       command == "undo" || 
       command == "redo" ||
       command == "disable_output" || 
       command == "enable_output" ||
       command.starts_with("scroll_to ") ||
       command.contains("AUTOFILL") {
        return;
    }

    if sheet.command_position < sheet.command_history.len() {
        sheet.command_history.truncate(sheet.command_position);
    }

    sheet.command_history.push(command.to_string());
    if sheet.command_history.len() > sheet.max_history_size {
        sheet.command_history.remove(0);
    } else {
        sheet.command_position += 1;
    }
}

pub fn undo(sheet: &mut Sheet) -> bool {
    if sheet.command_position == 0 || sheet.command_history.is_empty() {
        return false;
    }

    sheet.command_position -= 1;
    rebuild_sheet_state(sheet);
    true
}

pub fn redo(sheet: &mut Sheet) -> bool {
    if sheet.command_position >= sheet.command_history.len() {
        return false;
    }

    sheet.command_position += 1;
    rebuild_sheet_state(sheet);
    true
}

fn rebuild_sheet_state(sheet: &mut Sheet) {
    let view_row = sheet.view_row;
    let view_col = sheet.view_col;
    let output_enabled = sheet.output_enabled;
    
    for row in &mut sheet.cells {
        for cell in row {
            *cell = Cell::new();
        }
    }
    
    sheet.circular_dependency_detected = false;
    
    let commands_to_replay: Vec<String> = sheet.command_history[0..sheet.command_position].to_vec();
    
    for cmd in commands_to_replay {
        if let Some((cell_ref, formula)) = cmd.split_once('=') {
            let cell_ref = cell_ref.trim();
            if let Some((row, col)) = parse_cell_reference(sheet, cell_ref) {
                update_cell(sheet, row, col, formula.trim());
            }
        }
    }
    
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

        // Update ROWDEL
    if command.starts_with("ROWDEL ") {
        let row_str = &command[7..].trim();
        if let Ok(row) = row_str.parse::<i32>() {
            if row >= 1 && row <= sheet.rows {
                for col in 0..sheet.cols {
                    let cell = &mut sheet.cells[(row-1) as usize][col as usize];
                    cell.value = 0;
                    cell.formula = None;
                    cell.is_formula = false;
                    cell.is_error = false;
                    cell.is_bold = false;
                    cell.is_italic = false;
                    cell.is_underline = false;
                    // Remove this cell's dependencies and dependents
                    if let Some(cell_deps) = sheet.dependency_graph.remove(&((row-1), col)) {
                        for dep in cell_deps.dependencies {
                            match dep {
                                DependencyType::Single { row: r, col: c } => {
                                    remove_dependency(sheet, r, c, row-1, col, true);
                                }
                                DependencyType::Range { start_row, start_col, end_row, end_col } => {
                                    for i in start_row..=end_row {
                                        for j in start_col..=end_col {
                                            remove_dependency(sheet, i, j, row-1, col, true);
                                        }
                                    }
                                }
                            }
                        }
                        for dep in cell_deps.dependents {
                            match dep {
                                DependencyType::Single { row: r, col: c } => {
                                    remove_dependency(sheet, r, c, row-1, col, false);
                                }
                                DependencyType::Range { .. } => {} // Dependents are Single
                            }
                        }
                    }
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
                    for row in 0..sheet.rows {
                        let cell = &mut sheet.cells[row as usize][col as usize];
                        cell.value = 0;
                        cell.formula = None;
                        cell.is_formula = false;
                        cell.is_error = false;
                        cell.is_bold = false;
                        cell.is_italic = false;
                        cell.is_underline = false;
                        // Remove this cell's dependencies and dependents
                        if let Some(cell_deps) = sheet.dependency_graph.remove(&(row, col)) {
                            for dep in cell_deps.dependencies {
                                match dep {
                                    DependencyType::Single { row: r, col: c } => {
                                        remove_dependency(sheet, r, c, row, col, true);
                                    }
                                    DependencyType::Range { start_row, start_col, end_row, end_col } => {
                                        for i in start_row..=end_row {
                                            for j in start_col..=end_col {
                                                remove_dependency(sheet, i, j, row, col, true);
                                            }
                                        }
                                    }
                                }
                            }
                            for dep in cell_deps.dependents {
                                match dep {
                                    DependencyType::Single { row: r, col: c } => {
                                        remove_dependency(sheet, r, c, row, col, false);
                                    }
                                    DependencyType::Range { .. } => {}
                                }
                            }
                        }
                    }
                    add_to_history(sheet, command);
                }
            }
            return;
        }
        if command.starts_with("COPY ") {
            let range = &command[5..];
            if let Some((start_row, start_col, end_row, end_col)) = parse_range(sheet, range) {
                if copy_range(sheet, start_row, start_col, end_row, end_col) {
                    if sheet.output_enabled {
                        println!("Copied to clipboard");
                    }
                } else {
                    println!("Invalid range for copy");
                }
            }
            return;
        }

        if command.starts_with("CUT ") {
            let range = &command[4..];
            if let Some((start_row, start_col, end_row, end_col)) = parse_range(sheet, range) {
                if cut_range(sheet, start_row, start_col, end_row, end_col) {
                    if sheet.output_enabled {
                        println!("Cut to clipboard");
                    }
                } else {
                    println!("Invalid range for cut");
                }
            }
            return;
        }

        if command.starts_with("PASTE ") {
            let cell_ref = &command[6..];
            if let Some((row, col)) = parse_cell_reference(sheet, cell_ref) {
                if paste_range(sheet, row, col) {
                    if sheet.output_enabled {
                        println!("Pasted from clipboard");
                        display_sheet(sheet);
                    }
                } else {
                    println!("Nothing to paste or invalid target");
                }
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
                if let Some((start_row, start_col, end_row, end_col)) = parse_range(sheet, parts[2]) {
                    display_graph(sheet, graph_type, start_row, start_col, end_row, end_col);
                } else {
                    println!("Invalid range for graph");
                }
            } else {
                println!("Usage: GRAPH <type> <range> (e.g., GRAPH BAR A1:A10)");
            }
            return;
        }
    }

    if command.starts_with("scroll_to ") {
        let cell_ref = &command[10..];
        if let Some((row, col)) = parse_cell_reference(sheet, cell_ref) {
            scroll_to_cell(sheet, row, col);
        } else {
            println!("Invalid cell reference for scroll");
        }
        return;
    }

    if let Some((cell_ref, formula)) = command.split_once('=') {
        let cell_ref = cell_ref.trim();
        let formula = formula.trim();
        if let Some((row, col)) = parse_cell_reference(sheet, cell_ref) {
            if sheet.extension_enabled {
                add_to_history(sheet, command);
            }
            if let Some((func_name, args)) = formula.split_once('(') {
                if let Some(range_arg) = args.strip_suffix(')') {
                    let range_arg = range_arg.trim();
                    if func_name.trim().to_uppercase() == "SORTA" || func_name.trim().to_uppercase() == "SORTD" {
                        if let Some((start_row, start_col, end_row, end_col)) = parse_range(sheet, range_arg) {
                            if start_col == end_col {
                                let mut values: Vec<(i32, i32)> = Vec::new();
                                for i in start_row..=end_row {
                                    values.push((i, sheet.cells[i as usize][start_col as usize].value));
                                }
                                if func_name.trim().to_uppercase() == "SORTA" {
                                    values.sort_by(|a, b| a.1.cmp(&b.1));
                                } else {
                                    values.sort_by(|a, b| b.1.cmp(&a.1));
                                }
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
                            } else if start_row == end_row {
                                let mut values: Vec<(i32, i32)> = Vec::new();
                                for j in start_col..=end_col {
                                    values.push((j, sheet.cells[start_row as usize][j as usize].value));
                                }
                                if func_name.trim().to_uppercase() == "SORTA" {
                                    values.sort_by(|a, b| a.1.cmp(&b.1));
                                } else {
                                    values.sort_by(|a, b| b.1.cmp(&a.1));
                                }
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
                                let mut all_values: Vec<i32> = Vec::new();
                                for i in start_row..=end_row {
                                    for j in start_col..=end_col {
                                        all_values.push(sheet.cells[i as usize][j as usize].value);
                                    }
                                }
                                if func_name.trim().to_uppercase() == "SORTA" {
                                    all_values.sort();
                                } else {
                                    all_values.sort_by(|a, b| b.cmp(a));
                                }
                                let mut idx = 0;
                                for i in start_row..=end_row {
                                    for j in start_col..=end_col {
                                        if idx < all_values.len() {
                                            let cell = &mut sheet.cells[i as usize][j as usize];
                                            cell.value = all_values[idx];
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
                    } else if func_name.trim().to_uppercase() == "AUTOFILL" {
                        if let Some((start_row, start_col, end_row, end_col)) = parse_range(sheet, range_arg) {
                            if start_col != end_col && start_row != end_row {
                                return;
                            }
                            if start_col == end_col {
                                let mut values = Vec::new();
                                for i in (0.max(start_row - 5)..start_row).rev() {
                                    values.push(sheet.cells[i as usize][start_col as usize].value);
                                }
                                if values.is_empty() {
                                    return;
                                }
                                let pattern = detect_pattern(sheet, start_row, start_col, end_row, end_col);
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
                                        let last_value = values[0];
                                        for i in start_row..=end_row {
                                            let offset = i - (start_row - 1);
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
                                            let offset = i - (start_row - 1);
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
                            } else if start_row == end_row {
                                let mut values = Vec::new();
                                for j in (0.max(start_col - 5)..start_col).rev() {
                                    values.push(sheet.cells[start_row as usize][j as usize].value);
                                }
                                if values.is_empty() {
                                    return;
                                }
                                let pattern = detect_pattern(sheet, start_row, start_col, end_row, end_col);
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
                                    PatternType::Arithmetic(_initial, diff) => {
                                        let last_value = values[0];
                                        for j in start_col..=end_col {
                                            let offset = j - (start_col - 1);
                                            let new_value = last_value - diff * offset;
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
                                    PatternType::Geometric(_initial, ratio) => {
                                        let last_value = values[0];
                                        for j in start_col..=end_col {
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
                                    PatternType::Unknown => {}
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
            update_cell(sheet, row, col, formula);
        } else {
            println!("Invalid cell reference");
        }
    } else {
        println!("Invalid command format");
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
            let value_str = if cell.is_error && !cell.has_circular {
                "err".to_string()
            } else {
                cell.value.to_string()
            };

            let mut formatted = String::new();
            if cell.is_bold {
                formatted.push_str("\x1b[1m");
            }
            if cell.is_italic {
                formatted.push_str("\x1b[3m");
            }
            if cell.is_underline {
                formatted.push_str("\x1b[4m");
            }
            formatted.push_str(&value_str);
            if cell.is_bold || cell.is_italic || cell.is_underline {
                formatted.push_str("\x1b[0m");
            }

            print!("{:>width$} ", formatted, width = width);
        }
        println!();
    }
    io::stdout().flush().unwrap();
}

pub fn display_graph(sheet: &mut Sheet, graph_type: GraphType, start_row: i32, start_col: i32, end_row: i32, end_col: i32) {
    let mut values = Vec::new();
    let mut labels = Vec::new();
    
    for i in start_row..=end_row {
        for j in start_col..=end_col {
            let cell = &sheet.cells[i as usize][j as usize];
            values.push(cell.value);
            let mut label = String::new();
            encode_column(j, &mut label);
            label.push_str(&(i + 1).to_string());
            labels.push(label);
        }
    }
    
    let max_val = *values.iter().filter(|&&v| v > 0).max().unwrap_or(&10);
    let max_label_width = labels.iter().map(|l| l.len()).max().unwrap_or(2);
    let column_width = max_label_width.max(3) + 1;

    match graph_type {
        GraphType::Bar => {
            println!("\nBar Graph for range:");
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
            print!("---+");
            for _ in &values {
                print!("{}", "-".repeat(column_width));
            }
            println!();
            print!("   |");
            for label in &labels {
                print!("{:^width$}", label, width = column_width);
            }
            println!("\n");
        },
        GraphType::Scatter => {
            println!("\nScatter Plot for range:");
            for value in (1..=max_val).rev() {
                print!("{:2} |", value);
                for &cell_value in &values {
                    let center = column_width / 2;
                    if cell_value == value {
                        print!("{:width$}", " ", width = center);
                        print!("*");
                        print!("{:width$}", " ", width = column_width - center - 1);
                    } else {
                        print!("{:^width$}", " ", width = column_width);
                    }
                }
                println!();
            }
            print!("---+");
            for _ in &values {
                print!("{}", "-".repeat(column_width));
            }
            println!();
            print!("   |");
            for label in &labels {
                print!("{:^width$}", label, width = column_width);
            }
            println!("\n");
        }
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
                    if let Some(formula) = &cell.formula {
                        let mut cell_ref = String::new();
                        encode_column(target_col, &mut cell_ref);
                        cell_ref.push_str(&(target_row + 1).to_string());
                        update_cell(self, target_row, target_col, formula);
                    }
                }
            }
        }
    }
}

pub fn copy_range(sheet: &mut Sheet, start_row: i32, start_col: i32, end_row: i32, end_col: i32) -> bool {
    if start_row >= 0 && start_col >= 0 && end_row < sheet.rows && end_col < sheet.cols && start_row <= end_row && start_col <= end_col {
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

pub fn cut_range(sheet: &mut Sheet, start_row: i32, start_col: i32, end_row: i32, end_col: i32) -> bool {
    if start_row >= 0 && start_col >= 0 && end_row < sheet.rows && end_col < sheet.cols && start_row <= end_row && start_col <= end_col {
        let contents = sheet.get_cell_range(start_row, start_col, end_row, end_col);
        for i in start_row..=end_row {
            for j in start_col..=end_col {
                let cell = &mut sheet.cells[i as usize][j as usize];
                *cell = Cell::new();
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

pub fn paste_range(sheet: &mut Sheet, start_row: i32, start_col: i32) -> bool {
    let success = {
        let mut clipboard = CLIPBOARD.lock().unwrap();
        if let Some(clipboard_data) = &*clipboard {
            if start_row >= 0 && start_col >= 0 && start_row < sheet.rows && start_col < sheet.cols {
                sheet.set_cell_range(start_row, start_col, &clipboard_data.contents);
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
