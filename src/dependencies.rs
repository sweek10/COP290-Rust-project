use crate::types::{DependencyType, Sheet};
use crate::utils::{parse_cell_reference, parse_range};
use std::collections::{HashSet, VecDeque};
use crate::cell::evaluate_expression;

pub fn remove_dependency(dependents: &mut Vec<DependencyType>, row: i32, col: i32) {
    dependents.retain(|dep| !matches!(dep, DependencyType::Single { row: r, col: c } if *r == row && *c == col));
}

pub fn clear_dependencies(dependencies: &mut Vec<DependencyType>) {
    dependencies.clear();
}

pub fn has_circular_dependency(sheet: &mut Sheet, cell_ref: &str, formula: &str) -> bool {
    if let Some((start_row, start_col)) = parse_cell_reference(sheet, cell_ref) {
        let mut new_deps = Vec::new();
        if !formula.is_empty() {
            let tokens: Vec<&str> = formula.split(&['+', '-', '*', '/', '(', ')', ' '][..]).collect();
            for token in tokens {
                if token.contains(':') {
                    if let Some((start_row, start_col, end_row, end_col)) = parse_range(sheet, token) {
                        new_deps.push(DependencyType::Range { start_row, start_col, end_row, end_col });
                    }
                } else if token.chars().next().map_or(false, |c| c.is_alphabetic()) {
                    if let Some((dep_row, dep_col)) = parse_cell_reference(sheet, token) {
                        new_deps.push(DependencyType::Single { row: dep_row, col: dep_col });
                    }
                }
            }
        }

        let cell = &mut sheet.cells[start_row as usize][start_col as usize];
        let old_deps = std::mem::replace(&mut cell.dependencies, new_deps.clone());

        let mut visited = HashSet::new();
        
        fn dfs(
            sheet: &Sheet,
            curr_row: i32,
            curr_col: i32,
            start_row: i32,
            start_col: i32,
            path: &mut HashSet<(i32, i32)>,
            visited: &mut HashSet<(i32, i32)>,
        ) -> bool {
            if path.contains(&(curr_row, curr_col)) {
                return true;
            }
            
            if visited.contains(&(curr_row, curr_col)) {
                return false;
            }
            
            path.insert((curr_row, curr_col));
            
            let dependencies = &sheet.cells[curr_row as usize][curr_col as usize].dependencies;
            
            for dep in dependencies {
                match dep {
                    DependencyType::Single { row, col } => {
                        if (*row == start_row && *col == start_col) || 
                           dfs(sheet, *row, *col, start_row, start_col, path, visited) {
                            return true;
                        }
                    }
                    DependencyType::Range { start_row: s_row, start_col: s_col, end_row: e_row, end_col: e_col } => {
                        // Check if the range includes the starting cell
                        if start_row >= *s_row && start_row <= *e_row && 
                           start_col >= *s_col && start_col <= *e_col {
                            return true;
                        }
                        // Check key cells in the range (e.g., corners) to reduce iterations
                        let corners = [
                            (*s_row, *s_col),
                            (*s_row, *e_col),
                            (*e_row, *s_col),
                            (*e_row, *e_col),
                        ];
                        for (i, j) in corners.iter() {
                            if dfs(sheet, *i, *j, start_row, start_col, path, visited) {
                                return true;
                            }
                        }
                    }
                }
            }
            
            path.remove(&(curr_row, curr_col));
            visited.insert((curr_row, curr_col));
            
            false
        }
        
        let mut path = HashSet::new();
        let mut has_cycle = false;
        
        for dep in &new_deps {
            match dep {
                DependencyType::Single { row, col } => {
                    if *row == start_row && *col == start_col {
                        has_cycle = true;
                        break;
                    }
                    
                    path.clear();
                    path.insert((start_row, start_col));
                    if dfs(sheet, *row, *col, start_row, start_col, &mut path, &mut visited) {
                        has_cycle = true;
                        break;
                    }
                }
                DependencyType::Range { start_row: s_row, start_col: s_col, end_row: e_row, end_col: e_col } => {
                    if start_row >= *s_row && start_row <= *e_row && 
                       start_col >= *s_col && start_col <= *e_col {
                        has_cycle = true;
                        break;
                    }
                    
                    let corners = [
                        (*s_row, *s_col),
                        (*s_row, *e_col),
                        (*e_row, *s_col),
                        (*e_row, *e_col),
                    ];
                    let mut found_cycle = false;
                    for (i, j) in corners.iter() {
                        path.clear();
                        path.insert((start_row, start_col));
                        if dfs(sheet, *i, *j, start_row, start_col, &mut path, &mut visited) {
                            found_cycle = true;
                            break;
                        }
                    }
                    
                    if found_cycle {
                        has_cycle = true;
                        break;
                    }
                }
            }
        }
        
        if has_cycle {
            sheet.cells[start_row as usize][start_col as usize].has_circular = true;
            sheet.circular_dependency_detected = true;
        }
        
        sheet.cells[start_row as usize][start_col as usize].dependencies = old_deps;
        
        has_cycle
    } else {
        false
    }
}

pub fn recalculate_dependents(sheet: &mut Sheet, cell_ref: &str) {
    if let Some((start_row, start_col)) = parse_cell_reference(sheet, cell_ref) {
        let mut visited = HashSet::new();
        let mut queue = VecDeque::new();
        queue.push_back((start_row, start_col));
        visited.insert((start_row, start_col));

        let mut all_dependents = Vec::new();
        while let Some((row, col)) = queue.pop_front() {
            all_dependents.push((row, col));
            
            for dep in &sheet.cells[row as usize][col as usize].dependents {
                match dep {
                    DependencyType::Single { row: r, col: c } => {
                        if visited.insert((*r, *c)) {
                            queue.push_back((*r, *c));
                        }
                    }
                    DependencyType::Range { start_row, start_col, end_row, end_col } => {
                        for i in *start_row..=*end_row {
                            for j in *start_col..=*end_col {
                                if visited.insert((i, j)) {
                                    queue.push_back((i, j));
                                }
                            }
                        }
                    }
                }
            }
            
            for i in 0..sheet.rows {
                for j in 0..sheet.cols {
                    for dep in &sheet.cells[i as usize][j as usize].dependencies {
                        if let DependencyType::Range { start_row: s_row, start_col: s_col, end_row: e_row, end_col: e_col } = dep {
                            if row >= *s_row && row <= *e_row && col >= *s_col && col <= *e_col {
                                if visited.insert((i, j)) {
                                    queue.push_back((i, j));
                                }
                            }
                        }
                    }
                }
            }
        }

        for (row, col) in all_dependents {
            if row != start_row || col != start_col {
                if let Some(formula) = sheet.cells[row as usize][col as usize].formula.clone() {
                    let (new_value, is_error) = evaluate_expression(sheet, &formula, cell_ref);
                    let cell = &mut sheet.cells[row as usize][col as usize];
                    cell.value = new_value;
                    cell.is_error = is_error;
                }
            }
        }
    }
}

pub fn reset_circular_dependency_flag(sheet: &mut Sheet) {
    // First reset all individual cell flags
    for row in &mut sheet.cells {
        for cell in row {
            cell.has_circular = false;
        }
    }
    
    // Then check if any circular dependencies still exist
    let mut has_circular = false;
    for row in &sheet.cells {
        for cell in row {
            if cell.has_circular {
                has_circular = true;
                break;
            }
        }
        if has_circular {
            break;
        }
    }
    
    sheet.circular_dependency_detected = has_circular;
}
