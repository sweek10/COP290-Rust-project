use std::sync::Mutex;

#[derive(Clone, Debug)]
pub enum DependencyType {
    Single { row: i32, col: i32 },
    Range { start_row: i32, start_col: i32, end_row: i32, end_col: i32 },
}

#[derive(Clone, Debug)]
pub struct CellDependency {
    pub dependency: DependencyType,
    pub next: Option<Box<CellDependency>>,
}
#[derive(Debug)]
pub enum PatternType {
    Constant(i32),           // All values are the same
    Arithmetic(i32, i32),    // (initial_value, difference)
    Fibonacci(i32, i32),     // (penultimate, last) for Fibonacci sequence
    Unknown,                 // No recognized pattern
}

#[derive(Clone, Debug)]
pub struct Cell {
    pub value: i32,
    pub formula: Option<String>,
    pub is_formula: bool,
    pub is_error: bool,
    pub dependencies: Option<Box<CellDependency>>,
    
    pub dependents: Option<Box<CellDependency>>,
    pub has_circular: bool,
    pub is_bold: bool,
    pub is_italic: bool,
    pub is_underline: bool,
}
impl Cell {
    // Add a default constructor or update existing creation logic
    pub fn new() -> Self {
        Cell {
            value: 0,
            formula: None,
            is_formula: false,
            is_error: false,
            dependencies: None,
            dependents: None,
            has_circular: false,
            is_bold: false,
            is_italic: false,
            is_underline: false,
        }
    }
}

pub struct Sheet {
    pub cells: Vec<Vec<Cell>>,
    pub rows: i32,
    pub cols: i32,
    pub view_row: i32,
    pub view_col: i32,
    pub output_enabled: bool,
    pub circular_dependency_detected: bool,
    pub extension_enabled: bool,
}

lazy_static::lazy_static! {
    pub static ref SHEET: Mutex<Option<Sheet>> = Mutex::new(None);
}