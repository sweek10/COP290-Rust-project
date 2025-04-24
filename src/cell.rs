use crate::dependencies::{has_circular_dependency, recalculate_dependents};
use crate::types::CellDependencies;
use crate::types::{DependencyType, Sheet};
use crate::utils::{
    calculate_range_function, evaluate_arithmetic, is_valid_formula, parse_cell_reference,
    parse_range,
};
use std::thread::sleep;
use std::time::Duration;

pub fn update_cell(sheet: &mut Sheet, row: i32, col: i32, formula: &str) {
    if row < 0
        || row >= sheet.rows
        || col < 0
        || col >= sheet.cols
        || !is_valid_formula(sheet, formula)
    {
        return;
    }

    if has_circular_dependency(sheet, row, col, formula) {
        let cell = &mut sheet.cells[row as usize][col as usize];
        cell.formula = Some(formula.to_string());
        cell.is_formula = true;
        cell.has_circular = true;
        recalculate_dependents(sheet, row, col);
        return;
    }

    // Parse new dependencies
    let mut new_dependencies = Vec::new();
    let tokens: Vec<&str> = formula
        .split(&['+', '-', '*', '/', '(', ')', ' '][..])
        .collect();
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
        } else if token.chars().next().is_some_and(|c| c.is_alphabetic()) {
            if let Some((dep_row, dep_col)) = parse_cell_reference(sheet, token) {
                new_dependencies.push(DependencyType::Single {
                    row: dep_row,
                    col: dep_col,
                });
            }
        }
    }

    let (value, is_error) = evaluate_expression(sheet, formula, row, col);
    {
        let cell = &mut sheet.cells[row as usize][col as usize];
        cell.formula = Some(formula.to_string());
        cell.is_formula = true;
        cell.value = value;
        cell.is_error = is_error;
    }

    // Get and remove old dependencies
    let old_cell_deps = sheet.dependency_graph.remove(&(row, col));

    // Remove this cell from dependents of old Single dependencies
    if let Some(old_cell_deps) = &old_cell_deps {
        for dep in &old_cell_deps.dependencies {
            if let DependencyType::Single { row: r, col: c } = dep {
                if let Some(dep_cell_deps) = sheet.dependency_graph.get_mut(&(*r, *c)) {
                    dep_cell_deps.dependents.retain(|d| !matches!(d, DependencyType::Single { row: r2, col: c2 } if *r2 == row && *c2 == col));
                }
            }
            // Range dependencies are handled in BFS, so no need to update dependents here
        }
    }

    // Create new cell dependencies, preserving existing dependents if any
    let new_cell_deps = CellDependencies {
        dependencies: new_dependencies.clone(),
        dependents: old_cell_deps.map_or(Vec::new(), |d| d.dependents),
    };

    // Add this cell to dependents of new Single dependencies
    for dep in &new_cell_deps.dependencies {
        if let DependencyType::Single { row: r, col: c } = dep {
            let dep_cell_deps =
                sheet
                    .dependency_graph
                    .entry((*r, *c))
                    .or_insert_with(|| CellDependencies {
                        dependencies: Vec::new(),
                        dependents: Vec::new(),
                    });
            dep_cell_deps
                .dependents
                .push(DependencyType::Single { row, col });
        }
        // Range dependencies will be detected dynamically in BFS
    }

    // Insert the updated dependencies into the graph
    sheet.dependency_graph.insert((row, col), new_cell_deps);

    // Clean up if both dependencies and dependents are empty
    if new_dependencies.is_empty() && sheet.dependency_graph[&(row, col)].dependents.is_empty() {
        sheet.dependency_graph.remove(&(row, col));
    }

    recalculate_dependents(sheet, row, col);
    crate::dependencies::reset_circular_dependency_flag(sheet);
}
pub fn evaluate_expression(sheet: &mut Sheet, expr: &str, _row: i32, _col: i32) -> (i32, bool) {
    let mut is_error = false;

    // Handle numeric literals
    if let Ok(value) = expr.parse::<i32>() {
        return (value, false);
    }

    // Handle single cell reference
    if expr.chars().next().is_some_and(|c| c.is_alphabetic())
        && !expr.contains(&['+', '-', '*', '/', '('][..])
    {
        if let Some((r, c)) = parse_cell_reference(sheet, expr) {
            let cell = &sheet.cells[r as usize][c as usize];
            return (cell.value, cell.is_error);
        }
    }

    // Handle functions like SLEEP, SUM, AVG, etc.
    if let Some((function, args)) = expr.split_once('(').map(|(f, a)| (f, &a[..a.len() - 1])) {
        let function = function.trim().to_uppercase();
        if function == "SLEEP" {
            let (duration, error) = evaluate_expression(sheet, args, _row, _col);
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

    // Handle arithmetic expressions
    let mut final_expr = String::new();
    let mut pos = 0;
    while pos < expr.len() {
        let c = expr.chars().nth(pos).unwrap();
        if c.is_alphabetic() {
            let mut token_end = pos;
            while token_end < expr.len()
                && expr
                    .chars()
                    .nth(token_end)
                    .is_some_and(|c| c.is_alphanumeric())
            {
                token_end += 1;
            }
            let token = &expr[pos..token_end];
            if let Some((r, c)) = parse_cell_reference(sheet, token) {
                let cell = &sheet.cells[r as usize][c as usize];
                if cell.is_error {
                    return (0, true);
                }
                final_expr.push_str(&cell.value.to_string());
            } else {
                return (0, true); // Invalid cell reference
            }
            pos = token_end;
        } else if c.is_ascii_digit() {
            let mut token_end = pos;
            while token_end < expr.len() && expr.chars().nth(token_end).unwrap().is_ascii_digit() {
                token_end += 1;
            }
            final_expr.push_str(&expr[pos..token_end]);
            pos = token_end;
        } else if c == '-'
            && (pos == 0 || "+-*/(".contains(expr.chars().nth(pos - 1).unwrap_or(' ')))
        {
            let mut token_end = pos + 1;
            while token_end < expr.len() && expr.chars().nth(token_end).unwrap().is_ascii_digit() {
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

    // Evaluate the constructed arithmetic expression
    let result = evaluate_arithmetic(&final_expr, &mut is_error);
    (result, is_error)
}
