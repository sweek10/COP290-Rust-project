// utils.rs
use std::str::FromStr;
use crate::types::SHEET;
use crate::types::Sheet;

pub fn parse_cell_reference(sheet: &mut Sheet,ref_str: &str) -> Option<(i32, i32)> {
    // Trim the reference string to handle any whitespace
    // println!("Parsing cell reference: '{}'", ref_str);
    let ref_str = ref_str.trim();
    
    // Find the position where digits start
    let num_start = ref_str.chars().position(|c| c.is_ascii_digit())?;
    // println!("num_start: {}", num_start);
    
    // Split into column letters and row number
    let (col_str, row_str) = ref_str.split_at(num_start);
    // println!("col_str: '{}', row_str: '{}'", col_str, row_str);
    
    // Ensure the column string is not empty
    if col_str.is_empty() {
        return None;
    }
    
    // Convert column letters to column index
    let col = decode_column(col_str);
    // println!("col: {}", col);
    
    // Parse row number, subtracting 1 to convert to 0-indexed
    let row = i32::from_str(row_str).ok()? - 1;
    // println!("row: {}", row);
    
    // let sheet = SHEET.lock().unwrap();
    // let sheet = sheet.as_ref().unwrap();
    
    // Check if the cell is within bounds
    if row >= 0 && row < sheet.rows && col >= 0 && col < sheet.cols {
        Some((row, col))
    } else {
        None
    }
}
pub fn parse_range(sheet:&mut Sheet,range: &str) -> Option<(i32, i32, i32, i32)> {
    let (start, end) = range.split_once(':')?;
    let (start_row, start_col) = parse_cell_reference(sheet,start)?;
    let (end_row, end_col) = parse_cell_reference(sheet,end)?;
    if start_row <= end_row && start_col <= end_col {
        Some((start_row, start_col, end_row, end_col))
    } else {
        None
    }
}

pub fn calculate_range_function(sheet:&mut Sheet,function: &str, range: &str) -> f64 {
    let (start_row, start_col, end_row, end_col) = parse_range(sheet,range).unwrap_or((0, 0, 0, 0));
    // let sheet = SHEET.lock().unwrap();
    // let sheet = sheet.as_ref().unwrap();
    
    let mut count = 0;
    let mut sum = 0.0;
    let mut min = i32::MAX as f64;
    let mut max = i32::MIN as f64;
    
    for i in start_row..=end_row {
        for j in start_col..=end_col {
            let value = sheet.cells[i as usize][j as usize].value as f64;
            sum += value;
            min = min.min(value);
            max = max.max(value);
            count += 1;
        }
    }
    
    if count == 0 { return 0.0; }
    let mean = sum / count as f64;
    
    match function {
        "STDEV" => {
            let variance: f64 = (start_row..=end_row)
                .flat_map(|i| (start_col..=end_col).map(move |j| i as usize * j as usize))
                .map(|idx| {
                    let diff = sheet.cells[idx / (end_col - start_col + 1) as usize][idx % (end_col - start_col + 1) as usize].value as f64 - mean;
                    diff * diff
                })
                .sum();
            (variance / count as f64).sqrt()
        }
        "MIN" => min,
        "MAX" => max,
        "SUM" => sum,
        "AVG" => mean,
        _ => 0.0,
    }
}

pub fn evaluate_arithmetic(expr: &str, is_error: &mut bool) -> i32 {
    let tokens: Vec<&str> = expr.split_whitespace().collect();
    if tokens.len() == 1 {
        return tokens[0].parse().unwrap_or(0);
    }

    let mut result = tokens[0].parse::<i32>().unwrap_or(0);
    let mut i = 1;
    while i < tokens.len() - 1 {
        let op = tokens[i];
        let b = tokens[i + 1].parse::<i32>().unwrap_or(0);
        match op {
            "+" => result += b,
            "-" => result -= b,
            "*" => result *= b,
            "/" => {
                if b == 0 {
                    *is_error = true;
                    return 0;
                }
                result /= b;
            }
            _ => {}
        }
        i += 2;
    }
    result
}

pub fn decode_column(col_str: &str) -> i32 {
    let mut result = 0;
    for c in col_str.chars() {
        result = result * 26 + (c.to_ascii_uppercase() as i32 - 'A' as i32 + 1);
    }
    result - 1
}

pub fn encode_column(col: i32, col_str: &mut String) {
    let mut col = col + 1;
    while col > 0 {
        col -= 1;
        col_str.push((b'A' + (col % 26) as u8) as char);
        col /= 26;
    }
    let chars: Vec<char> = col_str.chars().rev().collect();
    *col_str = chars.into_iter().collect();
}

pub fn is_valid_formula(sheet: &mut Sheet, formula: &str) -> bool {
    // Trim whitespace from the formula
    let formula = formula.trim();

    // Check if it's a function call (e.g., SUM(A1:B2) or SLEEP(2))
    if let Some((func_name, args)) = formula.split_once('(') {
        if let Some(args) = args.strip_suffix(')') {
            let func_name = func_name.trim().to_uppercase();
            
            // Validate all supported functions
            match func_name.as_str() {
                "SUM" | "AVG" | "MAX" | "MIN" | "STDEV" => {
                    // These functions require a range argument
                    return parse_range(sheet, args.trim()).is_some();
                },
                "SLEEP" => {
                    // SLEEP accepts either a number or cell reference
                    let arg = args.trim();
                    return arg.parse::<i32>().is_ok() || 
                           parse_cell_reference(sheet, arg).is_some();
                },
                _ => {
                    // Reject unsupported functions
                    return false;
                }
            }
        }
    }

    // Check for simple arithmetic expressions (e.g., A1+B2)
    if formula.contains('+') || formula.contains('-') || 
       formula.contains('*') || formula.contains('/') {
        let parts: Vec<&str> = formula.split(|c| "+-*/".contains(c))
            .map(|s| s.trim())
            .filter(|s| !s.is_empty())
            .collect();

        return parts.iter().all(|part| {
            part.parse::<i32>().is_ok() ||  // Number
            parse_cell_reference(sheet, part).is_some()  // Cell reference
        });
    }

    // Check for plain cell references or numbers
    parse_cell_reference(sheet, formula).is_some() ||
    formula.parse::<i32>().is_ok()
}
pub fn is_valid_command(sheet: &mut Sheet, command: &str) -> bool {
    if command.len() == 1 && "wasdq".contains(command) {
        return true;
    }
    if command == "disable_output" || command == "enable_output" {
        return true;
    }
    if command.starts_with("scroll_to ") {
        return parse_cell_reference(sheet, &command[10..]).is_some();
    }
    command.split_once('=').map_or(false, |(ref_str, formula)| {
        parse_cell_reference(sheet, ref_str).is_some() && 
        !formula.is_empty() && 
        is_valid_formula(sheet, formula) // Validate the formula too!
    })
}