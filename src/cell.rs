use std::thread::sleep;
use std::time::Duration;
use crate::types::{Sheet, DependencyType};
use crate::dependencies::{has_circular_dependency, add_dependency, remove_dependency, clear_dependencies, recalculate_dependents};
use crate::utils::{parse_cell_reference, parse_range, calculate_range_function, evaluate_arithmetic, is_valid_formula};

pub fn update_cell(sheet: &mut Sheet, cell_ref: &str, formula: &str) {
    if parse_cell_reference(sheet, cell_ref).is_none() || !is_valid_formula(sheet, formula) {
        return; // Reject invalid updates completely
    }

    if let Some((row, col)) = parse_cell_reference(sheet, cell_ref) {
        // Phase 1: Remove existing dependencies
        let mut to_remove = Vec::new();
        {
            let cell = &sheet.cells[row as usize][col as usize];
            let mut dep = cell.dependencies.as_ref();
            while let Some(d) = dep {
                match d.dependency {
                    DependencyType::Single { row: r, col: c } => to_remove.push((r, c)),
                    DependencyType::Range { start_row, start_col, end_row, end_col } => {
                        for i in start_row..=end_row {
                            for j in start_col..=end_col {
                                to_remove.push((i, j));
                            }
                        }
                    }
                }
                dep = d.next.as_ref();
            }
        }
        for (dep_row, dep_col) in to_remove {
            let mut dependents = sheet.cells[dep_row as usize][dep_col as usize].dependents.take();
            remove_dependency(&mut dependents, row, col);
            sheet.cells[dep_row as usize][dep_col as usize].dependents = dependents;
        }
        clear_dependencies(&mut sheet.cells[row as usize][col as usize].dependencies);

        // Phase 2: Check for circular dependencies
        if has_circular_dependency(sheet, cell_ref, formula) {
            let cell = &mut sheet.cells[row as usize][col as usize];
            cell.formula = Some(formula.to_string());
            cell.is_formula = true;
            recalculate_dependents(sheet, cell_ref);
            return;
        }

        // Phase 3: Collect new dependencies (no individual cell updates for ranges)
        let mut new_dependencies = None;
        let tokens: Vec<&str> = formula.split(&['+', '-', '*', '/', '(', ')', ' '][..]).collect();
        for token in tokens {
            if token.contains(':') {
                if let Some((start_row, start_col, end_row, end_col)) = parse_range(sheet, token) {
                    add_dependency(&mut new_dependencies, DependencyType::Range {
                        start_row,
                        start_col,
                        end_row,
                        end_col,
                    });
                }
            } else if token.chars().next().map_or(false, |c| c.is_alphabetic()) {
                if let Some((dep_row, dep_col)) = parse_cell_reference(sheet, token) {
                    add_dependency(&mut new_dependencies, DependencyType::Single { row: dep_row, col: dep_col });
                    let mut dependents = sheet.cells[dep_row as usize][dep_col as usize].dependents.take();
                    add_dependency(&mut dependents, DependencyType::Single { row, col });
                    sheet.cells[dep_row as usize][dep_col as usize].dependents = dependents;
                }
            }
        }

        // Phase 4: Update cell with new dependencies and formula
        {
            let cell = &mut sheet.cells[row as usize][col as usize];
            cell.formula = Some(formula.to_string());
            cell.is_formula = true;
            cell.dependencies = new_dependencies;
        }

        // Phase 5: Evaluate and update cell value
        let (value, is_error) = evaluate_expression(sheet, formula, cell_ref);
        {
            let cell = &mut sheet.cells[row as usize][col as usize];
            cell.value = value;
            cell.is_error = is_error;
        }

        // Phase 6: Recalculate dependents and reset circular flag
        recalculate_dependents(sheet, cell_ref);
        crate::dependencies::reset_circular_dependency_flag(sheet);
    }
}

pub fn evaluate_expression(sheet: &mut Sheet, expr: &str, _current_cell: &str) -> (i32, bool) {
    let mut is_error = false;

    if let Ok(value) = expr.parse::<i32>() {
        return (value, false);
    }

    if expr.chars().next().map_or(false, |c| c.is_alphabetic()) && !expr.contains(&['+', '-', '*', '/', '('][..]) {
        if let Some((row, col)) = parse_cell_reference(sheet, expr) {
            let cell = &sheet.cells[row as usize][col as usize];
            return (cell.value, cell.is_error);
        }
    }

    if let Some((function, args)) = expr.split_once('(').map(|(f, a)| (f, &a[..a.len()-1])) {
        if function == "SLEEP" {
            let (duration, error) = evaluate_expression(sheet, args, _current_cell);
            if error {
                return (0, true);
            }
            sleep(Duration::from_secs(duration as u64));
            return (duration, false);
        }
        
        if let Some((start_row, start_col, end_row, end_col)) = parse_range(sheet, args) {
            let mut max = i32::MIN;
            let mut sum = 0;
            let mut has_error = false;
            for i in start_row..=end_row {
                for j in start_col..=end_col {
                    let cell = &sheet.cells[i as usize][j as usize];
                    if cell.is_error {
                        has_error = true;
                        break;
                    }
                    max = max.max(cell.value);
                    sum += cell.value;
                }
                if has_error {
                    break;
                }
            }
            if has_error {
                return (0, true);
            }
            match function {
                "MAX" => return (max, false),
                "SUM" => return (sum, false),
                _ => return (calculate_range_function(sheet, function, args) as i32, false),
            }
        }
    }

    let mut final_expr = String::new();
    let mut pos = 0;
    while pos < expr.len() {
        let c = expr.chars().nth(pos).unwrap();
        if c.is_alphabetic() {
            let mut token_end = pos;
            while token_end < expr.len() && expr.chars().nth(token_end).map_or(false, |c| c.is_alphanumeric()) {
                token_end += 1;
            }
            let token = &expr[pos..token_end];
            if let Some((row, col)) = parse_cell_reference(sheet, token) {
                let cell = &sheet.cells[row as usize][col as usize];
                if cell.is_error {
                    return (0, true);
                }
                final_expr.push_str(&cell.value.to_string());
            } else {
                final_expr.push_str(token);
            }
            pos = token_end;
        } else if c.is_digit(10) || (c == '-' && (pos == 0 || "+-*/(".contains(expr.chars().nth(pos-1).unwrap_or(' ')))) {
            let mut token_end = pos;
            while token_end < expr.len() && expr.chars().nth(token_end).map_or(false, |c| c.is_digit(10)) {
                token_end += 1;
            }
            final_expr.push_str(&expr[pos..token_end]);
            pos = token_end;
        } else if "+-*/()".contains(c) {
            final_expr.push(' ');
            final_expr.push(c);
            final_expr.push(' ');
            pos += 1;
        } else {
            pos += 1;
        }
    }
    (evaluate_arithmetic(&final_expr, &mut is_error), is_error)
}