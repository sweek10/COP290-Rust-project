use crate::cell::update_cell;
use crate::dependencies::remove_dependency;
use crate::types::{
    Cell, Clipboard, DependencyType, GraphType, PatternType, Sheet, SheetState, CLIPBOARD,
};
use crate::utils::{
    detect_pattern, encode_column, factorial, parse_cell_reference, parse_range, triangular,
};
use std::collections::HashMap;
use std::io::{self, Write};

const DISPLAY_SIZE: i32 = 10;
type CellAttributes = (i32, Option<String>, bool, bool, bool, bool, bool);

/// Creates a new spreadsheet with the specified dimensions.
///
/// This function initializes a spreadsheet with the given number of rows and columns,
/// creating a grid of empty cells. It also sets whether extensions are enabled.
///
/// # Arguments
/// * `rows` - The number of rows in the spreadsheet.
/// * `cols` - The number of columns in the spreadsheet.
/// * `extension_enabled` - Whether advanced features are enabled.
///
/// # Returns
/// An `Option<Sheet>` containing the initialized spreadsheet.
///
/// # Example
/// ```
/// let sheet = create_sheet(5, 5, false).unwrap();
/// assert_eq!(sheet.rows, 5);
/// assert_eq!(sheet.cols, 5);
/// ```
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
        // command_history: Vec::with_capacity(10),
        // command_position: 0,
        max_history_size: 10,
        dependency_graph: HashMap::new(),
        undo_stack: Vec::new(),
        redo_stack: Vec::new(),
    })
}

/// Scrolls the spreadsheet view in the specified direction.
///
/// This function adjusts the viewable portion of the spreadsheet based on the direction
/// (w: up, s: down, a: left, d: right), ensuring the view stays within bounds.
///
/// # Arguments
/// * `sheet` - A mutable reference to the spreadsheet.
/// * `direction` - The direction to scroll (w, a, s, d).
///
/// # Example
/// ```
/// let mut sheet = create_sheet(20, 20, false).unwrap();
/// scroll_sheet(&mut sheet, 's');
/// assert_eq!(sheet.view_row, 10);
/// ```
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
            } else if sheet.view_row + DISPLAY_SIZE < sheet.rows
                && sheet.view_row + 10 <= sheet.rows
                && sheet.view_row + 20 > sheet.rows
            {
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
            } else if sheet.view_col + DISPLAY_SIZE < sheet.cols
                && sheet.view_col + 10 <= sheet.cols
                && sheet.view_col + 20 > sheet.cols
            {
                sheet.view_col += sheet.cols - sheet.view_col - 10;
            }
        }
        _ => {}
    }
}

/// Scrolls the spreadsheet view to a specific cell.
///
/// # Description
/// Sets the spreadsheet's view to center on the specified cell coordinates (row, col).
/// If the coordinates are invalid (outside the spreadsheet bounds), it prints an error message and does not change the view.
///
/// # Arguments
/// * `sheet` - A mutable reference to the spreadsheet.
/// * `row` - The target row index (0-based).
/// * `col` - The target column index (0-based).
///
/// # Example
/// ```
/// use spreadsheet::{create_sheet, scroll_to_cell};
/// let mut sheet = create_sheet(20, 20, false).unwrap();
/// scroll_to_cell(&mut sheet, 15, 15);
/// assert_eq!(sheet.view_row, 15);
/// assert_eq!(sheet.view_col, 15);
/// scroll_to_cell(&mut sheet, 25, 25); // Invalid coordinates
/// // Prints "Invalid cell coordinates for scroll"
/// ```
pub fn scroll_to_cell(sheet: &mut Sheet, row: i32, col: i32) {
    if row >= 0 && row < sheet.rows && col >= 0 && col < sheet.cols {
        sheet.view_row = row;
        sheet.view_col = col;
    } else {
        println!("Invalid cell coordinates for scroll");
    }
}

/// Saves the current state of the spreadsheet to the undo stack.
///
/// This function captures the current state of the spreadsheet's cells and dependency graph
/// and adds it to the undo stack. It ensures the stack does not exceed the maximum history size.
///
/// # Arguments
/// * `sheet` - A mutable reference to the spreadsheet.
///
/// # Example
/// ```
/// let mut sheet = create_sheet(5, 5, true).unwrap();
/// sheet.cells[0][0].value = 10;
/// save_state(&mut sheet);
/// assert_eq!(sheet.undo_stack.len(), 1);
/// ```
pub fn save_state(sheet: &mut Sheet) {
    if !sheet.extension_enabled {
        return;
    }

    // Clear redo stack when a new action is performed
    sheet.redo_stack.clear();

    // Create a snapshot of the current state
    let state = SheetState {
        cells: sheet.cells.clone(),
        dependency_graph: sheet.dependency_graph.clone(),
    };

    // Add to undo stack
    sheet.undo_stack.push(state);

    // Maintain history size limit
    if sheet.undo_stack.len() > sheet.max_history_size {
        sheet.undo_stack.remove(0);
    }
}

/// Reverts the spreadsheet to the previous state in the undo stack.
///
/// This function restores the spreadsheet to the most recent state in the undo stack,
/// saving the current state to the redo stack. It only works if extensions are enabled.
///
/// # Arguments
/// * `sheet` - A mutable reference to the spreadsheet.
///
/// # Returns
/// An `Option<String>` with an error message if the operation fails, or `None` on success.
///
/// # Example
/// ```
/// let mut sheet = create_sheet(5, 5, true).unwrap();
/// sheet.cells[0][0].value = 10;
/// save_state(&mut sheet);
/// sheet.cells[0][0].value = 20;
/// undo(&mut sheet);
/// assert_eq!(sheet.cells[0][0].value, 10);
/// ```
pub fn undo(sheet: &mut Sheet) -> bool {
    if !sheet.extension_enabled || sheet.undo_stack.is_empty() {
        return false;
    }

    // Save current state to redo stack before undoing
    let current_state = SheetState {
        cells: sheet.cells.clone(),
        dependency_graph: sheet.dependency_graph.clone(),
    };
    sheet.redo_stack.push(current_state);

    // Get the previous state
    let previous_state = sheet.undo_stack.pop().unwrap();

    // Restore the sheet to the previous state
    sheet.cells = previous_state.cells;
    sheet.dependency_graph = previous_state.dependency_graph;

    true
}

/// Restores the spreadsheet to the next state in the redo stack.
///
/// This function reapplies the most recent state from the redo stack, saving the current
/// state to the undo stack. It only works if extensions are enabled.
///
/// # Arguments
/// * `sheet` - A mutable reference to the spreadsheet.
///
/// # Returns
/// An `Option<String>` with an error message if the operation fails, or `None` on success.
///
/// # Example
/// ```
/// let mut sheet = create_sheet(5, 5, true).unwrap();
/// sheet.cells[0][0].value = 10;
/// save_state(&mut sheet);
/// sheet.cells[0][0].value = 20;
/// undo(&mut sheet);
/// redo(&mut sheet);
/// assert_eq!(sheet.cells[0][0].value, 20);
/// ```
pub fn redo(sheet: &mut Sheet) -> bool {
    if !sheet.extension_enabled || sheet.redo_stack.is_empty() {
        return false;
    }

    // Save current state to undo stack before redoing
    let current_state = SheetState {
        cells: sheet.cells.clone(),
        dependency_graph: sheet.dependency_graph.clone(),
    };
    sheet.undo_stack.push(current_state);

    // Get the next state
    let next_state = sheet.redo_stack.pop().unwrap();

    // Restore the sheet to the next state
    sheet.cells = next_state.cells;
    sheet.dependency_graph = next_state.dependency_graph;

    true
}

/// Processes a user command to manipulate the spreadsheet.
///
/// # Description
/// Interprets and executes a command string, such as scrolling, setting cell values, or advanced operations like sorting or autofilling.
/// Supports commands like `w`, `s`, `a`, `d` for scrolling, `undo`, `redo`, `FORMULA`, `ROWDEL`, `COLDEL`, `COPY`, `CUT`, `PASTE`, `GRAPH`, `scroll_to`, and cell assignments (e.g., `A1=5`).
/// Returns an error message if the command is invalid or fails, or `None` on success.
///
/// # Arguments
/// * `sheet` - A mutable reference to the spreadsheet.
/// * `command` - The command string to process.
///
/// # Returns
/// An `Option<String>` containing an error message if the command fails, or `None` if successful.
///
/// # Example
/// ```
/// use spreadsheet::{create_sheet, process_command};
/// let mut sheet = create_sheet(5, 5, true).unwrap();
/// assert_eq!(process_command(&mut sheet, "A1=5"), None);
/// assert_eq!(sheet.cells[0][0].value, 5);
/// assert_eq!(process_command(&mut sheet, "s"), None); // Scroll down
/// assert_eq!(sheet.view_row, 0); // Limited by sheet size
/// assert_eq!(process_command(&mut sheet, "invalid"), Some("Invalid command format".to_string()));
/// ```
pub fn process_command(sheet: &mut Sheet, command: &str) -> Option<String> {
    if command.is_empty() {
        return None;
    }

    if command.len() == 1 {
        match command.chars().next().unwrap() {
            'w' => {
                scroll_sheet(sheet, 'w');
                return None;
            }
            'a' => {
                scroll_sheet(sheet, 'a');
                return None;
            }
            's' => {
                scroll_sheet(sheet, 's');
                return None;
            }
            'd' => {
                scroll_sheet(sheet, 'd');
                return None;
            }
            'q' => std::process::exit(0),
            _ => return Some("Invalid single-character command".to_string()),
        }
    }

    if command == "disable_output" {
        sheet.output_enabled = false;
        return None;
    }
    if command == "enable_output" {
        sheet.output_enabled = true;
        return None;
    }

    if sheet.extension_enabled {
        if command == "undo" {
            if !undo(sheet) {
                return Some("Nothing to undo".to_string());
            }
            return None;
        }
        if command == "redo" {
            if !redo(sheet) {
                return Some("Nothing to redo".to_string());
            }
            return None;
        }

        if let Some(stripped) = command.strip_prefix("FORMULA ") {
            let cell_ref = &stripped.trim();
            if let Some((row, col)) = parse_cell_reference(sheet, cell_ref) {
                let cell = &sheet.cells[row as usize][col as usize];
                if let Some(formula) = &cell.formula {
                    return Some(format!("Formula in cell {}: {}", cell_ref, formula));
                } else {
                    return Some(format!("No formula stored in cell {}", cell_ref));
                }
            } else {
                return Some(format!("Invalid cell reference: {}", cell_ref));
            }
        }

        if let Some(stripped) = command.strip_prefix("ROWDEL") {
            save_state(sheet);
            let row_str = &stripped.trim();
            if let Ok(row) = row_str.parse::<i32>() {
                if row >= 1 && row <= sheet.rows {
                    for col in 0..sheet.cols {
                        let cell = &mut sheet.cells[(row - 1) as usize][col as usize];
                        cell.value = 0;
                        cell.formula = None;
                        cell.is_formula = false;
                        cell.is_error = false;
                        cell.is_bold = false;
                        cell.is_italic = false;
                        cell.is_underline = false;
                        if let Some(cell_deps) = sheet.dependency_graph.remove(&((row - 1), col)) {
                            for dep in cell_deps.dependencies {
                                match dep {
                                    DependencyType::Single { row: r, col: c } => {
                                        remove_dependency(sheet, r, c, row - 1, col, true);
                                    }
                                    DependencyType::Range {
                                        start_row,
                                        start_col,
                                        end_row,
                                        end_col,
                                    } => {
                                        for i in start_row..=end_row {
                                            for j in start_col..=end_col {
                                                remove_dependency(sheet, i, j, row - 1, col, true);
                                            }
                                        }
                                    }
                                }
                            }
                            for dep in cell_deps.dependents {
                                match dep {
                                    DependencyType::Single { row: r, col: c } => {
                                        remove_dependency(sheet, r, c, row - 1, col, false);
                                    }
                                    DependencyType::Range { .. } => {}
                                }
                            }
                        }
                    }

                    return None;
                } else {
                    return Some(format!("Invalid row number: {}", row));
                }
            } else {
                return Some("Invalid ROWDEL format: use ROWDEL <number>".to_string());
            }
        }

        if let Some(stripped) = command.strip_prefix("COLDEL") {
            save_state(sheet);
            let col_str = &stripped.trim();
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
                        if let Some(cell_deps) = sheet.dependency_graph.remove(&(row, col)) {
                            for dep in cell_deps.dependencies {
                                match dep {
                                    DependencyType::Single { row: r, col: c } => {
                                        remove_dependency(sheet, r, c, row, col, true);
                                    }
                                    DependencyType::Range {
                                        start_row,
                                        start_col,
                                        end_row,
                                        end_col,
                                    } => {
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
                    return None;
                } else {
                    return Some(format!("Invalid column reference: {}", col_str));
                }
            } else {
                return Some("Invalid COLDEL format: use COLDEL <column>".to_string());
            }
        }

        if let Some(stripped) = command.strip_prefix("COPY ") {
            let range = &stripped;
            if let Some((start_row, start_col, end_row, end_col)) = parse_range(sheet, range) {
                if copy_range(sheet, start_row, start_col, end_row, end_col) {
                    return Some("Copied to clipboard".to_string());
                } else {
                    return Some("Invalid range for copy".to_string());
                }
            } else {
                return Some("Invalid COPY format: use COPY <range>".to_string());
            }
        }

        if let Some(stripped) = command.strip_prefix("CUT ") {
            save_state(sheet);
            let range = &stripped;
            if let Some((start_row, start_col, end_row, end_col)) = parse_range(sheet, range) {
                if cut_range(sheet, start_row, start_col, end_row, end_col) {
                    return Some("Cut to clipboard".to_string());
                } else {
                    return Some("Invalid range for cut".to_string());
                }
            } else {
                return Some("Invalid CUT format: use CUT <range>".to_string());
            }
        }

        if let Some(stripped) = command.strip_prefix("PASTE ") {
            save_state(sheet);
            let cell_ref = &stripped;
            if let Some((row, col)) = parse_cell_reference(sheet, cell_ref) {
                if paste_range(sheet, row, col) {
                    return Some("Pasted from clipboard".to_string());
                } else {
                    return Some("Nothing to paste or invalid target".to_string());
                }
            } else {
                return Some("Invalid PASTE format: use PASTE <cell>".to_string());
            }
        }

        if command.starts_with("GRAPH ") {
            let parts: Vec<&str> = command.split_whitespace().collect();
            if parts.len() == 3 {
                let graph_type = match parts[1].to_uppercase().as_str() {
                    "(BAR)" => GraphType::Bar,
                    "(SCATTER)" => GraphType::Scatter,
                    _ => return Some("Invalid graph type. Use (BAR) or (SCATTER)".to_string()),
                };
                if let Some((start_row, start_col, end_row, end_col)) = parse_range(sheet, parts[2])
                {
                    let graph_output =
                        display_graph(sheet, graph_type, start_row, start_col, end_row, end_col);
                    return Some(graph_output);
                } else {
                    return Some("Invalid range for graph".to_string());
                }
            } else {
                return Some("Usage: GRAPH <type> <range> (e.g., GRAPH (BAR) A1:A10)".to_string());
            }
        }
    }

    if let Some(stripped) = command.strip_prefix("scroll_to ") {
        let cell_ref = &stripped;
        if let Some((row, col)) = parse_cell_reference(sheet, cell_ref) {
            scroll_to_cell(sheet, row, col);
            return None;
        } else {
            return Some("Invalid cell reference for scroll".to_string());
        }
    }

    if let Some((cell_ref, formula)) = command.split_once('=') {
        let cell_ref = cell_ref.trim();
        let formula = formula.trim();
        if let Some((row, col)) = parse_cell_reference(sheet, cell_ref) {
            if sheet.extension_enabled {
                save_state(sheet);
            }
            if let Some((func_name, args)) = formula.split_once('(') {
                if let Some(range_arg) = args.strip_suffix(')') {
                    let range_arg = range_arg.trim();
                    if func_name.trim().to_uppercase() == "SORTA"
                        || func_name.trim().to_uppercase() == "SORTD"
                    {
                        if let Some((start_row, start_col, end_row, end_col)) =
                            parse_range(sheet, range_arg)
                        {
                            if start_col == end_col {
                                let mut values: Vec<(i32, i32)> = Vec::new();
                                for i in start_row..=end_row {
                                    values.push((
                                        i,
                                        sheet.cells[i as usize][start_col as usize].value,
                                    ));
                                }
                                if func_name.trim().to_uppercase() == "SORTA" {
                                    values.sort_by(|a, b| a.1.cmp(&b.1));
                                } else {
                                    values.sort_by(|a, b| b.1.cmp(&a.1));
                                }
                                let mut original_values: Vec<CellAttributes> = Vec::new();
                                for i in start_row..=end_row {
                                    let cell = &sheet.cells[i as usize][start_col as usize];
                                    original_values.push((
                                        cell.value,
                                        cell.formula.clone(),
                                        cell.is_formula,
                                        cell.is_error,
                                        cell.is_bold,
                                        cell.is_italic,
                                        cell.is_underline,
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
                                    values.push((
                                        j,
                                        sheet.cells[start_row as usize][j as usize].value,
                                    ));
                                }
                                if func_name.trim().to_uppercase() == "SORTA" {
                                    values.sort_by(|a, b| a.1.cmp(&b.1));
                                } else {
                                    values.sort_by(|a, b| b.1.cmp(&a.1));
                                }
                                let mut original_values: Vec<CellAttributes> = Vec::new();
                                for j in start_col..=end_col {
                                    let cell = &sheet.cells[start_row as usize][j as usize];
                                    original_values.push((
                                        cell.value,
                                        cell.formula.clone(),
                                        cell.is_formula,
                                        cell.is_error,
                                        cell.is_bold,
                                        cell.is_italic,
                                        cell.is_underline,
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
                                    save_state(sheet);
                                    all_values.sort();
                                } else {
                                    save_state(sheet);
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
                            return None;
                        }
                    } else if func_name.trim().to_uppercase() == "AUTOFILL" {
                        save_state(sheet);
                        if let Some((start_row, start_col, end_row, end_col)) =
                            parse_range(sheet, range_arg)
                        {
                            if start_col != end_col && start_row != end_row {
                                return None;
                            }
                            if start_col == end_col {
                                let mut values = Vec::new();
                                for i in (0.max(start_row - 5)..start_row).rev() {
                                    values.push(sheet.cells[i as usize][start_col as usize].value);
                                }
                                if values.is_empty() {
                                    return None;
                                }
                                let pattern =
                                    detect_pattern(sheet, start_row, start_col, end_row, end_col);
                                match pattern {
                                    PatternType::Constant(value) => {
                                        for i in start_row..=end_row {
                                            let cell =
                                                &mut sheet.cells[i as usize][start_col as usize];
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
                                            let cell =
                                                &mut sheet.cells[i as usize][start_col as usize];
                                            cell.value = new_value;
                                            cell.formula = None;
                                            cell.is_formula = false;
                                            cell.is_error = false;
                                        }
                                    }
                                    PatternType::Fibonacci(mut penult, mut last) => {
                                        for i in start_row..=end_row {
                                            let new_value = penult + last;
                                            let cell =
                                                &mut sheet.cells[i as usize][start_col as usize];
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
                                            let new_value = (last_value as f64 * ratio.powi(offset))
                                                .round()
                                                as i32;
                                            let cell =
                                                &mut sheet.cells[i as usize][start_col as usize];
                                            cell.value = new_value;
                                            cell.formula = None;
                                            cell.is_formula = false;
                                            cell.is_error = false;
                                        }
                                    }
                                    PatternType::Factorial(_last_value, mut next_index) => {
                                        for i in start_row..=end_row {
                                            let new_value = factorial(next_index);
                                            let cell =
                                                &mut sheet.cells[i as usize][start_col as usize];
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
                                            let cell =
                                                &mut sheet.cells[i as usize][start_col as usize];
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
                                    return None;
                                }
                                let pattern =
                                    detect_pattern(sheet, start_row, start_col, end_row, end_col);
                                match pattern {
                                    PatternType::Constant(value) => {
                                        for j in start_col..=end_col {
                                            let cell =
                                                &mut sheet.cells[start_row as usize][j as usize];
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
                                            let cell =
                                                &mut sheet.cells[start_row as usize][j as usize];
                                            cell.value = new_value;
                                            cell.formula = None;
                                            cell.is_formula = false;
                                            cell.is_error = false;
                                        }
                                    }
                                    PatternType::Fibonacci(mut penult, mut last) => {
                                        for j in start_col..=end_col {
                                            let new_value = penult + last;
                                            let cell =
                                                &mut sheet.cells[start_row as usize][j as usize];
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
                                            let new_value = (last_value as f64 * ratio.powi(offset))
                                                .round()
                                                as i32;
                                            let cell =
                                                &mut sheet.cells[start_row as usize][j as usize];
                                            cell.value = new_value;
                                            cell.formula = None;
                                            cell.is_formula = false;
                                            cell.is_error = false;
                                        }
                                    }
                                    PatternType::Factorial(_last_value, mut next_index) => {
                                        for j in start_col..=end_col {
                                            let new_value = factorial(next_index);
                                            let cell =
                                                &mut sheet.cells[start_row as usize][j as usize];
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
                                            let cell =
                                                &mut sheet.cells[start_row as usize][j as usize];
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
                            return None;
                        }
                    } else if let Some(cell_arg) = args.strip_suffix(')') {
                        save_state(sheet);
                        let cell_arg = cell_arg.trim();
                        if let Some((row, col)) = parse_cell_reference(sheet, cell_arg) {
                            match func_name.trim().to_uppercase().as_str() {
                                "BOLD" => {
                                    sheet.cells[row as usize][col as usize].is_bold = true;
                                    return None;
                                }
                                "ITALIC" => {
                                    sheet.cells[row as usize][col as usize].is_italic = true;
                                    return None;
                                }
                                "UNDERLINE" => {
                                    sheet.cells[row as usize][col as usize].is_underline = true;
                                    return None;
                                }
                                _ => {}
                            }
                        }
                    }
                }
            }
            update_cell(sheet, row, col, formula);
            None
        } else {
            Some("Invalid cell reference".to_string())
        }
    } else {
        Some("Invalid command format".to_string())
    }
}

/// Displays the current view of the spreadsheet in the terminal.
///
/// This function outputs a portion of the spreadsheet (based on the current view) to the terminal,
/// showing column headers, row numbers, and cell values. It respects the output_enabled flag and
/// highlights errors or circular dependencies.
///
/// # Arguments
/// * `sheet` - A reference to the spreadsheet.
///
/// # Example
/// ```
/// let mut sheet = create_sheet(10, 10, false).unwrap();
/// sheet.cells[0][0].value = 5;
/// display_sheet(&sheet);
/// // Outputs a 10x10 grid with "5" in A1 and zeros elsewhere
/// ```
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
                if val == 0 {
                    1
                } else {
                    val.to_string().len()
                }
            };
            max_widths[(j - sheet.view_col) as usize] =
                max_widths[(j - sheet.view_col) as usize].max(width);
        }
    }

    print!("     ");
    for j in sheet.view_col..(sheet.view_col + DISPLAY_SIZE).min(sheet.cols) {
        let mut col_header = String::new();
        encode_column(j, &mut col_header);
        print!(
            "{:width$} ",
            col_header,
            width = max_widths[(j - sheet.view_col) as usize]
        );
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

/// Generates an ASCII-based graph for a range of cells.
///
/// # Description
/// Creates a bar or scatter plot for the values in the specified cell range, outputting the result as a string.
/// The graph uses ASCII characters to represent data points or bars, with labels derived from cell references.
/// Suitable for terminal-based visualization of data.
///
/// # Arguments
/// * `sheet` - A mutable reference to the spreadsheet.
/// * `graph_type` - The type of graph (`GraphType::Bar` or `GraphType::Scatter`).
/// * `start_row` - The starting row of the range (0-based).
/// * `start_col` - The starting column of the range (0-based).
/// * `end_row` - The ending row of the range (0-based).
/// * `end_col` - The ending column of the range (0-based).
///
/// # Returns
/// A `String` containing the ASCII representation of the graph.
///
/// # Example
/// ```
/// use spreadsheet::{create_sheet, display_graph, GraphType};
/// let mut sheet = create_sheet(5, 5, false).unwrap();
/// sheet.cells[0][0].value = 5;
/// sheet.cells[1][0].value = 3;
/// let output = display_graph(&mut sheet, GraphType::Bar, 0, 0, 1, 0);
/// println!("{}", output);
/// // Outputs an ASCII bar graph for A1:A2 with values 5 and 3
/// ```
pub fn display_graph(
    sheet: &mut Sheet,
    graph_type: GraphType,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
) -> String {
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

    let mut output = String::new();

    match graph_type {
        GraphType::Bar => {
            output.push_str("\nBar Graph for range:\n".to_string().as_str());
            for value in (1..=max_val).rev() {
                output.push_str(&format!("{:2} |", value));
                for &cell_value in &values {
                    if cell_value >= value {
                        output.push_str(&format!("{:^width$}", "█", width = column_width));
                    } else {
                        output.push_str(&format!("{:^width$}", " ", width = column_width));
                    }
                }
                output.push('\n');
            }
            output.push_str("---+");
            for _ in &values {
                output.push_str(&"-".repeat(column_width));
            }
            output.push('\n');
            output.push_str("   |");
            for label in &labels {
                output.push_str(&format!("{:^width$}", label, width = column_width));
            }
            output.push('\n');
        }
        GraphType::Scatter => {
            output.push_str("\nScatter Plot for range:\n".to_string().as_str());
            for value in (1..=max_val).rev() {
                output.push_str(&format!("{:2} |", value));
                for &cell_value in &values {
                    let center = column_width / 2;
                    if cell_value == value {
                        output.push_str(&format!("{:width$}", " ", width = center));
                        output.push('*');
                        output.push_str(&format!(
                            "{:width$}",
                            " ",
                            width = column_width - center - 1
                        ));
                    } else {
                        output.push_str(&format!("{:^width$}", " ", width = column_width));
                    }
                }
                output.push('\n');
            }
            output.push_str("---+");
            for _ in &values {
                output.push_str(&"-".repeat(column_width));
            }
            output.push('\n');
            output.push_str("   |");
            for label in &labels {
                output.push_str(&format!("{:^width$}", label, width = column_width));
            }
            output.push('\n');
        }
    }
    output
}

impl Sheet {
    /// Retrieves a range of cells from the spreadsheet.
    ///
    /// # Description
    /// Returns a 2D vector of `Cell` clones for the specified range, defined by start and end row/column indices.
    /// Used for operations like copying or analyzing a range of cells.
    ///
    /// # Arguments
    /// * `start_row` - The starting row of the range (0-based).
    /// * `start_col` - The starting column of the range (0-based).
    /// * `end_row` - The ending row of the range (0-based).
    /// * `end_col` - The ending column of the range (0-based).
    ///
    /// # Returns
    /// A `Vec<Vec<Cell>>` containing the cells in the specified range.
    ///
    /// # Example
    /// ```
    /// use spreadsheet::{create_sheet, Sheet};
    /// let mut sheet = create_sheet(5, 5, false).unwrap();
    /// sheet.cells[0][0].value = 5;
    /// let range = sheet.get_cell_range(0, 0, 1, 1);
    /// assert_eq!(range[0][0].value, 5);
    /// assert_eq!(range[1][1].value, 0);
    /// ```
    pub fn get_cell_range(
        &self,
        start_row: i32,
        start_col: i32,
        end_row: i32,
        end_col: i32,
    ) -> Vec<Vec<Cell>> {
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

    /// Sets a range of cells in the spreadsheet.
    ///
    /// # Description
    /// Updates the spreadsheet with the provided 2D vector of cells, starting at the specified row and column.
    /// If a cell contains a formula, it is re-evaluated using `update_cell`. Ensures updates stay within spreadsheet bounds.
    ///
    /// # Arguments
    /// * `start_row` - The starting row for pasting the cells (0-based).
    /// * `start_col` - The starting column for pasting the cells (0-based).
    /// * `values` - A 2D vector of `Cell` values to set.
    ///
    /// # Example
    /// ```
    /// use spreadsheet::{create_sheet, Sheet, Cell};
    /// let mut sheet = create_sheet(5, 5, false).unwrap();
    /// let mut cell = Cell::new();
    /// cell.value = 10;
    /// let values = vec![vec![cell; 2]; 2];
    /// sheet.set_cell_range(0, 0, &values);
    /// assert_eq!(sheet.cells[0][0].value, 10);
    /// assert_eq!(sheet.cells[1][1].value, 10);
    /// ```
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

/// Copies a range of cells to the clipboard.
///
/// # Description
/// Copies the specified range of cells to the global clipboard, preserving cell data (values, formulas, formatting).
/// Returns `false` if the range is invalid (e.g., out of bounds or start indices exceed end indices).
///
/// # Arguments
/// * `sheet` - A mutable reference to the spreadsheet.
/// * `start_row` - The starting row of the range (0-based).
/// * `start_col` - The starting column of the range (0-based).
/// * `end_row` - The ending row of the range (0-based).
/// * `end_col` - The ending column of the range (0-based).
///
/// # Returns
/// A `bool` indicating success (`true`) or failure (`false`).
///
/// # Example
/// ```
/// use spreadsheet::{create_sheet, copy_range};
/// let mut sheet = create_sheet(5, 5, false).unwrap();
/// sheet.cells[0][0].value = 5;
/// assert!(copy_range(&mut sheet, 0, 0, 1, 1));
/// assert!(!copy_range(&mut sheet, 0, 0, 6, 6)); // Out of bounds
/// ```
pub fn copy_range(
    sheet: &mut Sheet,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
) -> bool {
    if start_row >= 0
        && start_col >= 0
        && end_row < sheet.rows
        && end_col < sheet.cols
        && start_row <= end_row
        && start_col <= end_col
    {
        let contents = sheet.get_cell_range(start_row, start_col, end_row, end_col);
        *CLIPBOARD.lock().unwrap() = Some(Clipboard {
            contents,
            is_cut: false,
            // source_range: None,
        });
        true
    } else {
        false
    }
}

/// Cuts a range of cells to the clipboard, clearing the source.
///
/// # Description
/// Copies the specified range of cells to the global clipboard and clears the source cells in the spreadsheet.
/// Returns `false` if the range is invalid (e.g., out of bounds or start indices exceed end indices).
///
/// # Arguments
/// * `sheet` - A mutable reference to the spreadsheet.
/// * `start_row` - The starting row of the range (0-based).
/// * `start_col` - The starting column of the range (0-based).
/// * `end_row` - The ending row of the range (0-based).
/// * `end_col` - The ending column of the range (0-based).
///
/// # Returns
/// A `bool` indicating success (`true`) or failure (`false`).
///
/// # Example
/// ```
/// use spreadsheet::{create_sheet, cut_range};
/// let mut sheet = create_sheet(5, 5, false).unwrap();
/// sheet.cells[0][0].value = 5;
/// assert!(cut_range(&mut sheet, 0, 0, 1, 1));
/// assert_eq!(sheet.cells[0][0].value, 0); // Source cleared
/// assert!(!cut_range(&mut sheet, 0, 0, 6, 6)); // Out of bounds
/// ```
pub fn cut_range(
    sheet: &mut Sheet,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
) -> bool {
    if start_row >= 0
        && start_col >= 0
        && end_row < sheet.rows
        && end_col < sheet.cols
        && start_row <= end_row
        && start_col <= end_col
    {
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
            //source_range: Some((start_row, start_col, end_row, end_col)),
        });
        true
    } else {
        false
    }
}

/// Pastes the clipboard contents to the specified location.
///
/// # Description
/// Pastes the contents of the global clipboard to the spreadsheet starting at the specified cell.
/// If the operation is a cut, the clipboard is cleared after pasting. Displays the updated sheet if successful.
/// Returns `false` if the clipboard is empty or the target location is invalid.
///
/// # Arguments
/// * `sheet` - A mutable reference to the spreadsheet.
/// * `start_row` - The starting row for pasting (0-based).
/// * `start_col` - The starting column for pasting (0-based).
///
/// # Returns
/// A `bool` indicating success (`true`) or failure (`false`).
///
/// # Example
/// ```
/// use spreadsheet::{create_sheet, copy_range, paste_range};
/// let mut sheet = create_sheet(5, 5, false).unwrap();
/// sheet.cells[0][0].value = 5;
/// copy_range(&mut sheet, 0, 0, 0, 0);
/// assert!(paste_range(&mut sheet, 1, 1));
/// assert_eq!(sheet.cells[1][1].value, 5);
/// assert!(!paste_range(&mut sheet, 6, 6)); // Invalid target
/// ```
pub fn paste_range(sheet: &mut Sheet, start_row: i32, start_col: i32) -> bool {
    let success = {
        let mut clipboard = CLIPBOARD.lock().unwrap();
        if let Some(clipboard_data) = &*clipboard {
            if start_row >= 0 && start_col >= 0 && start_row < sheet.rows && start_col < sheet.cols
            {
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
