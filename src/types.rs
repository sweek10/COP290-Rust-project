use std::collections::HashMap;
use std::sync::Mutex;

#[derive(Clone, Debug)]
pub enum DependencyType {
    Single {
        row: i32,
        col: i32,
    },
    Range {
        start_row: i32,
        start_col: i32,
        end_row: i32,
        end_col: i32,
    },
}

#[derive(Clone, Debug)]
pub struct CellDependencies {
    pub dependencies: Vec<DependencyType>, // Cells this cell depends on
    pub dependents: Vec<DependencyType>,   // Cells that depend on this cell
}

#[derive(Debug)]
pub enum PatternType {
    Constant(i32),
    Arithmetic(i32, i32),
    Fibonacci(i32, i32),
    Geometric(i32, f64),
    Factorial(i32, i32),
    Triangular(i32, i32),
    Unknown,
}

#[derive(Clone, Debug)]
pub struct Cell {
    pub value: i32,
    pub formula: Option<String>,
    pub is_formula: bool,
    pub is_error: bool,
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
            has_circular: false,
            is_bold: false,
            is_italic: false,
            is_underline: false,
        }
    }
}

#[derive(Clone)]
pub struct SheetState {
    pub cells: Vec<Vec<Cell>>,
    pub dependency_graph: HashMap<(i32, i32), CellDependencies>,
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
    pub max_history_size: usize,
    pub dependency_graph: HashMap<(i32, i32), CellDependencies>,
    pub undo_stack: Vec<SheetState>,
    pub redo_stack: Vec<SheetState>,
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
}

lazy_static::lazy_static! {
    pub static ref SHEET: Mutex<Option<Sheet>> = Mutex::new(None);
    pub static ref CLIPBOARD: Mutex<Option<Clipboard>> = Mutex::new(None);
}
