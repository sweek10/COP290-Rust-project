use std::thread::sleep;
use std::time::Duration;
use crate::types::{Sheet, DependencyType};
use crate::dependencies::{has_circular_dependency, remove_dependency, clear_dependencies, recalculate_dependents};
use crate::utils::{parse_cell_reference, parse_range, calculate_range_function, evaluate_arithmetic, is_valid_formula};

pub fn update_cell(sheet: &mut Sheet, cell_ref: &str, formula: &str) {
    if parse_cell_reference(sheet, cell_ref).is_none() || !is_valid_formula(sheet, formula) {
        return;
    }

    if let Some((row, col)) = parse_cell_reference(sheet, cell_ref) {
        let mut to_remove = Vec::new();
        {
            let cell = &sheet.cells[row as usize][col as usize];
            for dep in &cell.dependencies {
                match dep {
                    DependencyType::Single { row: r, col: c } => to_remove.push((*r, *c)),
                    DependencyType::Range { start_row, start_col, end_row, end_col } => {
                        for i in *start_row..=*end_row {
                            for j in *start_col..=*end_col {
                                to_remove.push((i, j));
                            }
                        }
                    }
                }
            }
        }
        
        for (dep_row, dep_col) in to_remove {
            remove_dependency(&mut sheet.cells[dep_row as usize][dep_col as usize].dependents, row, col);
        }
        
        clear_dependencies(&mut sheet.cells[row as usize][col as usize].dependencies);

        if has_circular_dependency(sheet, cell_ref, formula) {
            let cell = &mut sheet.cells[row as usize][col as usize];
            cell.formula = Some(formula.to_string());
            cell.is_formula = true;
            recalculate_dependents(sheet, cell_ref);
            return;
        }

        let mut new_dependencies = Vec::new();
        let tokens: Vec<&str> = formula.split(&['+', '-', '*', '/', '(', ')', ' '][..]).collect();
        for token in tokens {
            if token.contains(':') {
                if let Some((start_row, start_col, end_row, end_col)) = parse_range(sheet, token) {
                    new_dependencies.push(DependencyType::Range {
                        start_row,
                        start_col,
                        end_row,
                        end_col,
                    });
                }
            } else if token.chars().next().map_or(false, |c| c.is_alphabetic()) {
                if let Some((dep_row, dep_col)) = parse_cell_reference(sheet, token) {
                    new_dependencies.push(DependencyType::Single { row: dep_row, col: dep_col });
                    sheet.cells[dep_row as usize][dep_col as usize].dependents.push(
                        DependencyType::Single { row, col }
                    );
                }
            }
        }

        {
            let cell = &mut sheet.cells[row as usize][col as usize];
            cell.formula = Some(formula.to_string());
            cell.is_formula = true;
            cell.dependencies = new_dependencies;
        }

        let (value, is_error) = evaluate_expression(sheet, formula, cell_ref);
        {
            let cell = &mut sheet.cells[row as usize][col as usize];
            cell.value = value;
            cell.is_error = is_error;
        }

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
        let function = function.trim().to_uppercase();
        if function == "SLEEP" {
            let (duration, error) = evaluate_expression(sheet, args, _current_cell);
            if error {
                return (0, true);
            }
            sleep(Duration::from_secs(duration as u64));
            return (duration, false);
        }
        
        if parse_range(sheet, args).is_some() {
            match calculate_range_function(sheet, &function, args) {
                Ok(result) => {
                    if result.is_nan() || result.is_infinite() {
                        return (0, true);
                    }
                    return (result as i32, false);
                }
                Err(()) => return (0, true), 
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
