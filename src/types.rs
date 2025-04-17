use std::sync::Mutex;

#[derive(Clone, Debug)]
pub enum DependencyType {
    Single { row: i32, col: i32 },
    Range { start_row: i32, start_col: i32, end_row: i32, end_col: i32 },
}

#[derive(Debug)]
pub enum PatternType {
    Constant(i32),           // All values are the same
    Arithmetic(i32, i32),    // (initial_value, difference)
    Fibonacci(i32, i32),     // (penultimate, last) for Fibonacci sequence
    Geometric(i32, f64),
    Factorial(i32, i32),     // (last_value, next_index) e.g., (6, 4) for next = 24
    Triangular(i32, i32),
    Unknown,                 // No recognized pattern
}

#[derive(Clone, Debug)]
pub struct Cell {
    pub value: i32,
    pub formula: Option<String>,
    pub is_formula: bool,
    pub is_error: bool,
    pub dependencies: Vec<DependencyType>,  
    pub dependents: Vec<DependencyType>,    
    pub has_circular: bool,
    pub is_bold: bool,
    pub is_italic: bool,
    pub is_underline: bool,
}

impl Cell {
    pub fn new() -> Self {
        Cell {
            value: 0,
            formula: None,
            is_formula: false,
            is_error: false,
            dependencies: Vec::new(),      // Initialize empty vector
            dependents: Vec::new(),        // Initialize empty vector
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
    pub command_history: Vec<String>,
    pub command_position: usize,
    pub max_history_size: usize,
}

#[derive(Debug, Clone, Copy)]
pub enum GraphType {
    Bar,
    Scatter,
   
}

#[derive(Clone, Debug)]
pub struct Clipboard {
    pub contents: Vec<Vec<Cell>>,
    pub is_cut: bool,
    pub source_range: Option<(i32, i32, i32, i32)>,
}


lazy_static::lazy_static! {
    pub static ref SHEET: Mutex<Option<Sheet>> = Mutex::new(None);
    pub static ref CLIPBOARD: Mutex<Option<Clipboard>> = Mutex::new(None);
}
