use crate::types::{CellDependency, Sheet, DependencyType};
use crate::utils::{parse_cell_reference, encode_column, parse_range};
use crate::cell::evaluate_expression;

pub fn create_dependency(dependency: DependencyType) -> Box<CellDependency> {
    Box::new(CellDependency {
        dependency,
        next: None,
    })
}

pub fn add_dependency(head: &mut Option<Box<CellDependency>>, dependency: DependencyType) {
    let new_dep = create_dependency(dependency);
    if let Some(ref mut current) = head {
        let mut curr = current;
        while curr.next.is_some() {
            curr = curr.next.as_mut().unwrap();
        }
        curr.next = Some(new_dep);
    } else {
        *head = Some(new_dep);
    }
}

pub fn remove_dependency(head: &mut Option<Box<CellDependency>>, row: i32, col: i32) {
    if head.is_none() {
        return;
    }
    let mut current = head.as_mut().unwrap();
    if matches!(current.dependency, DependencyType::Single { row: r, col: c } if r == row && c == col) {
        *head = current.next.take();
        return;
    }
    while let Some(ref mut next) = current.next {
        if matches!(next.dependency, DependencyType::Single { row: r, col: c } if r == row && c == col) {
            current.next = next.next.take();
            return;
        }
        current = current.next.as_mut().unwrap();
    }
}

pub fn clear_dependencies(head: &mut Option<Box<CellDependency>>) {
    *head = None;
}

fn is_in_path(path: &[i32], row: i32, col: i32) -> bool {
    path.chunks(2).any(|chunk| chunk[0] == row && chunk[1] == col)
}

fn check_circular_recursive(
    sheet: &mut Sheet,
    curr_row: i32,
    curr_col: i32,
    path: &mut Vec<i32>,
    new_formula: &str,
) -> bool {
    if is_in_path(path, curr_row, curr_col) {
        return true;
    }
    path.push(curr_row);
    path.push(curr_col);

    if path.len() == 2 {
        let tokens: Vec<&str> = new_formula.split(&['+', '-', '*', '/', '(', ')', ' '][..]).collect();
        for token in tokens {
            if token.contains(':') {
                if let Some((start_row, start_col, end_row, end_col)) = parse_range(sheet, token) {
                    for i in start_row..=end_row {
                        for j in start_col..=end_col {
                            if check_circular_recursive(sheet, i, j, path, new_formula) {
                                return true;
                            }
                        }
                    }
                }
            } else if token.chars().next().map_or(false, |c| c.is_alphabetic()) {
                if let Some((dep_row, dep_col)) = parse_cell_reference(sheet, token) {
                    if check_circular_recursive(sheet, dep_row, dep_col, path, new_formula) {
                        return true;
                    }
                }
            }
        }
    } else {
        let cell = sheet.cells[curr_row as usize][curr_col as usize].clone();
        let mut dep = cell.dependencies.as_ref();
        while let Some(d) = dep {
            match d.dependency {
                DependencyType::Single { row, col } => {
                    if check_circular_recursive(sheet, row, col, path, new_formula) {
                        return true;
                    }
                }
                DependencyType::Range { start_row, start_col, end_row, end_col } => {
                    for i in start_row..=end_row {
                        for j in start_col..=end_col {
                            if check_circular_recursive(sheet, i, j, path, new_formula) {
                                return true;
                            }
                        }
                    }
                }
            }
            dep = d.next.as_ref();
        }
    }
    path.pop();
    path.pop();
    false
}

pub fn has_circular_dependency(sheet: &mut Sheet, cell_ref: &str, formula: &str) -> bool {
    if let Some((start_row, start_col)) = parse_cell_reference(sheet, cell_ref) {
        let mut path = Vec::new();
        let result = check_circular_recursive(sheet, start_row, start_col, &mut path, formula);
        if result {
            sheet.cells[start_row as usize][start_col as usize].has_circular = true;
            sheet.circular_dependency_detected = true;
        }
        result
    } else {
        true
    }
}

pub fn reset_circular_dependency_flag(sheet: &mut Sheet) {
    for row in &mut sheet.cells {
        for cell in row {
            if cell.has_circular {
                return;
            }
        }
    }
    sheet.circular_dependency_detected = false;
}

pub fn recalculate_dependents(sheet: &mut Sheet, cell_ref: &str) {
    if let Some((row, col)) = parse_cell_reference(sheet, cell_ref) {
        let mut updates = Vec::new();
        // Check explicit dependents
        {
            let cell = &sheet.cells[row as usize][col as usize];
            let mut dep = cell.dependents.as_ref();
            while let Some(d) = dep {
                if let DependencyType::Single { row: r, col: c } = d.dependency {
                    updates.push((r, c));
                }
                dep = d.next.as_ref();
            }
        }
        // Check all cells for range dependencies that include this cell
        for i in 0..sheet.rows {
            for j in 0..sheet.cols {
                let cell = &sheet.cells[i as usize][j as usize];
                let mut dep = cell.dependencies.as_ref();
                while let Some(d) = dep {
                    match d.dependency {
                        DependencyType::Range { start_row, start_col, end_row, end_col } => {
                            if row >= start_row && row <= end_row && col >= start_col && col <= end_col {
                                updates.push((i, j));
                            }
                        }
                        _ => {}
                    }
                    dep = d.next.as_ref();
                }
            }
        }
        // Remove duplicates to avoid redundant updates
        updates.sort_unstable();
        updates.dedup();

        // Recalculate affected cells
        for (dep_row, dep_col) in updates {
            if let Some(formula) = sheet.cells[dep_row as usize][dep_col as usize].formula.clone() {
                let (new_value, is_error) = evaluate_expression(sheet, &formula, cell_ref);
                let dep_cell = &mut sheet.cells[dep_row as usize][dep_col as usize];
                dep_cell.value = new_value;
                dep_cell.is_error = is_error;
                let mut dep_ref = String::new();
                encode_column(dep_col, &mut dep_ref);
                dep_ref.push_str(&(dep_row + 1).to_string());
                // Recursively recalculate dependents, but avoid infinite loops
                if !(dep_row == row && dep_col == col) {
                    recalculate_dependents(sheet, &dep_ref);
                }
            }
        }
    }
}