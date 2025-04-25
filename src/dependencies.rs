use crate::cell::evaluate_expression;
use crate::types::CellDependencies;
use crate::types::{DependencyType, Sheet};
use crate::utils::{parse_cell_reference, parse_range};
use std::collections::{HashMap, HashSet, VecDeque};

/// Removes a specific dependency or dependent relationship from the dependency graph for a given cell.
///
/// This function updates the dependency graph by removing a single cell dependency (or dependent)
/// identified by `(row, col)` from the dependencies or dependents of the cell at `(dep_row, dep_col)`.
/// If the dependency lists become empty, the cell's entry is removed from the graph.
///
/// # Arguments
/// * `sheet` - A mutable reference to the spreadsheet.
/// * `dep_row` - The row index of the cell whose dependencies are being modified.
/// * `dep_col` - The column index of the cell whose dependencies are being modified.
/// * `row` - The row index of the cell to remove from the dependencies or dependents.
/// * `col` - The column index of the cell to remove from the dependencies or dependents.
/// * `is_dependent` - If true, modifies the dependents list; if false, modifies the dependencies list.
///
/// # Example
/// ```
/// let mut sheet = create_sheet(10, 10, false).unwrap();
/// // Assume A1 depends on B2
/// sheet.dependency_graph.insert((0, 0), CellDependencies {
///     dependencies: vec![DependencyType::Single { row: 1, col: 1 }],
///     dependents: vec![],
/// });
/// remove_dependency(&mut sheet, 0, 0, 1, 1, false);
/// // A1 no longer depends on B2
/// assert!(sheet.dependency_graph.get(&(0, 0)).is_none());
/// ```
pub fn remove_dependency(
    sheet: &mut Sheet,
    dep_row: i32,
    dep_col: i32,
    row: i32,
    col: i32,
    is_dependent: bool,
) {
    if let Some(cell_deps) = sheet.dependency_graph.get_mut(&(dep_row, dep_col)) {
        let deps = if is_dependent {
            &mut cell_deps.dependents
        } else {
            &mut cell_deps.dependencies
        };
        deps.retain(|dep| !matches!(dep, DependencyType::Single { row: r, col: c } if *r == row && *c == col));
        if deps.is_empty() && cell_deps.dependencies.is_empty() && cell_deps.dependents.is_empty() {
            sheet.dependency_graph.remove(&(dep_row, dep_col));
        }
    }
}

/// Checks if a formula in a cell introduces a circular dependency.
///
/// This function evaluates the formula at the specified cell `(start_row, start_col)` to determine
/// if it creates a cycle in the dependency graph. It temporarily adds the new dependencies to the graph,
/// performs a depth-first search (DFS) to detect cycles, and restores the original state afterward.
/// If a circular dependency is found, the cell is marked as having a circular dependency.
///
/// # Arguments
/// * `sheet` - A mutable reference to the spreadsheet.
/// * `start_row` - The row index of the cell with the formula.
/// * `start_col` - The column index of the cell with the formula.
/// * `formula` - The formula to check for circular dependencies.
///
/// # Returns
/// A boolean indicating whether a circular dependency was detected.
///
/// # Example
/// ```
/// let mut sheet = create_sheet(10, 10, false).unwrap();
/// // A1 references A1 itself
/// let has_circular = has_circular_dependency(&mut sheet, 0, 0, "=A1");
/// assert!(has_circular);
/// assert!(sheet.cells[0][0].has_circular);
/// ```
pub fn has_circular_dependency(
    sheet: &mut Sheet,
    start_row: i32,
    start_col: i32,
    formula: &str,
) -> bool {
    if start_row < 0 || start_row >= sheet.rows || start_col < 0 || start_col >= sheet.cols {
        return false;
    }

    let mut new_deps = Vec::new();
    if !formula.is_empty() {
        let tokens: Vec<&str> = formula
            .split(&['+', '-', '*', '/', '(', ')', ' '][..])
            .collect();
        for token in tokens {
            if token.contains(':') {
                if let Some((start_row, start_col, end_row, end_col)) = parse_range(sheet, token) {
                    new_deps.push(DependencyType::Range {
                        start_row,
                        start_col,
                        end_row,
                        end_col,
                    });
                }
            } else if token.chars().next().is_some_and(|c| c.is_alphabetic()) {
                if let Some((dep_row, dep_col)) = parse_cell_reference(sheet, token) {
                    new_deps.push(DependencyType::Single {
                        row: dep_row,
                        col: dep_col,
                    });
                }
            }
        }
    }

    // Temporarily add new dependencies
    let old_deps = sheet.dependency_graph.remove(&(start_row, start_col));
    sheet.dependency_graph.insert(
        (start_row, start_col),
        CellDependencies {
            dependencies: new_deps.clone(),
            dependents: old_deps
                .as_ref()
                .map_or(Vec::new(), |d| d.dependents.clone()),
        },
    );

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

        if let Some(cell_deps) = sheet.dependency_graph.get(&(curr_row, curr_col)) {
            for dep in &cell_deps.dependencies {
                match dep {
                    DependencyType::Single { row, col } => {
                        if (*row == start_row && *col == start_col)
                            || dfs(sheet, *row, *col, start_row, start_col, path, visited)
                        {
                            return true;
                        }
                    }
                    DependencyType::Range {
                        start_row: s_row,
                        start_col: s_col,
                        end_row: e_row,
                        end_col: e_col,
                    } => {
                        if start_row >= *s_row
                            && start_row <= *e_row
                            && start_col >= *s_col
                            && start_col <= *e_col
                        {
                            return true;
                        }
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
                if dfs(
                    sheet,
                    *row,
                    *col,
                    start_row,
                    start_col,
                    &mut path,
                    &mut visited,
                ) {
                    has_cycle = true;
                    break;
                }
            }
            DependencyType::Range {
                start_row: s_row,
                start_col: s_col,
                end_row: e_row,
                end_col: e_col,
            } => {
                if start_row >= *s_row
                    && start_row <= *e_row
                    && start_col >= *s_col
                    && start_col <= *e_col
                {
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

    // Restore old dependencies
    if let Some(old_deps) = old_deps {
        sheet.dependency_graph.insert(
            (start_row, start_col),
            CellDependencies {
                dependencies: old_deps.dependencies,
                dependents: old_deps.dependents,
            },
        );
    } else {
        sheet.dependency_graph.remove(&(start_row, start_col));
    }

    has_cycle
}

/// Recalculates the values of all cells that depend on the specified cell.
///
/// This function uses a breadth-first search (BFS) to identify all cells that depend on the cell
/// at `(start_row, start_col)`, either directly or indirectly. It then performs a topological sort
/// to determine the order of recalculation and updates the values of dependent cells based on their
/// formulas. The starting cell itself is not recalculated.
///
/// # Arguments
/// * `sheet` - A mutable reference to the spreadsheet.
/// * `start_row` - The row index of the cell whose dependents need recalculation.
/// * `start_col` - The column index of the cell whose dependents need recalculation.
///
/// # Example
/// ```
/// let mut sheet = create_sheet(10, 10, false).unwrap();
/// // A1 = 5, B1 = A1 + 1
/// update_cell(&mut sheet, 0, 0, "5");
/// update_cell(&mut sheet, 0, 1, "=A1+1");
/// // Change A1 to 10
/// update_cell(&mut sheet, 0, 0, "10");
/// recalculate_dependents(&mut sheet, 0, 0);
/// // B1 should now be 11
/// assert_eq!(sheet.cells[0][1].value, 11);
/// ```
pub fn recalculate_dependents(sheet: &mut Sheet, start_row: i32, start_col: i32) {
    if start_row < 0 || start_row >= sheet.rows || start_col < 0 || start_col >= sheet.cols {
        return;
    }

    // Collect dependents using BFS
    let mut dependents = Vec::new();
    let mut visited = HashSet::new();
    let mut queue = VecDeque::new();
    queue.push_back((start_row, start_col));
    visited.insert((start_row, start_col));

    while let Some((row, col)) = queue.pop_front() {
        dependents.push((row, col));

        if let Some(cell_deps) = sheet.dependency_graph.get(&(row, col)) {
            for dep in &cell_deps.dependents {
                match dep {
                    DependencyType::Single { row: r, col: c } => {
                        if visited.insert((*r, *c)) {
                            queue.push_back((*r, *c));
                        }
                    }
                    DependencyType::Range { .. } => {}
                }
            }
        }

        for ((r, c), cell_deps) in &sheet.dependency_graph {
            for dep in &cell_deps.dependencies {
                if let DependencyType::Range {
                    start_row,
                    start_col,
                    end_row,
                    end_col,
                } = dep
                {
                    if row >= *start_row
                        && row <= *end_row
                        && col >= *start_col
                        && col <= *end_col
                        && visited.insert((*r, *c))
                    {
                        queue.push_back((*r, *c));
                    }
                }
            }
        }
    }

    // Topological sort
    let mut graph = HashMap::new();
    let mut in_degree = HashMap::new();

    for &(row, col) in &dependents {
        let node = (row, col);
        in_degree.entry(node).or_insert(0);

        if let Some(cell_deps) = sheet.dependency_graph.get(&(row, col)) {
            for dep in &cell_deps.dependencies {
                match dep {
                    DependencyType::Single { row: r, col: c } => {
                        if dependents.contains(&(*r, *c)) {
                            graph.entry((*r, *c)).or_insert_with(Vec::new).push(node);
                            *in_degree.entry(node).or_insert(0) += 1;
                        }
                    }
                    DependencyType::Range {
                        start_row: s_row,
                        start_col: s_col,
                        end_row: e_row,
                        end_col: e_col,
                    } => {
                        for i in *s_row..=*e_row {
                            for j in *s_col..=*e_col {
                                if dependents.contains(&(i, j)) {
                                    graph.entry((i, j)).or_insert_with(Vec::new).push(node);
                                    *in_degree.entry(node).or_insert(0) += 1;
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    let mut topo_order = Vec::new();
    let mut queue = VecDeque::new();

    for &node in &dependents {
        if in_degree.get(&node).unwrap_or(&0) == &0 {
            queue.push_back(node);
        }
    }

    while let Some(node) = queue.pop_front() {
        topo_order.push(node);
        if let Some(neighbors) = graph.get(&node) {
            for &neighbor in neighbors {
                let degree = in_degree.get_mut(&neighbor).unwrap();
                *degree -= 1;
                if *degree == 0 {
                    queue.push_back(neighbor);
                }
            }
        }
    }

    // Recalculate cells
    for &(row, col) in &topo_order {
        if row != start_row || col != start_col {
            if let Some(formula) = sheet.cells[row as usize][col as usize].formula.clone() {
                let (new_value, is_error) = evaluate_expression(sheet, &formula, row, col);
                let cell = &mut sheet.cells[row as usize][col as usize];
                cell.value = new_value;
                cell.is_error = is_error;
            }
        }
    }
}

/// Resets the circular dependency flags for all cells in the spreadsheet.
///
/// This function clears the `has_circular` flag for each cell and the global
/// `circular_dependency_detected` flag in the sheet, effectively resetting the circular
/// dependency state after a recalculation or update.
///
/// # Arguments
/// * `sheet` - A mutable reference to the spreadsheet.
///
/// # Example
/// ```
/// let mut sheet = create_sheet(10, 10, false).unwrap();
/// sheet.cells[0][0].has_circular = true;
/// sheet.circular_dependency_detected = true;
/// reset_circular_dependency_flag(&mut sheet);
/// assert!(!sheet.cells[0][0].has_circular);
/// assert!(!sheet.circular_dependency_detected);
/// ```
pub fn reset_circular_dependency_flag(sheet: &mut Sheet) {
    for row in &mut sheet.cells {
        for cell in row {
            cell.has_circular = false;
        }
    }

    sheet.circular_dependency_detected = false;
}
