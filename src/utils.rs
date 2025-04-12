// utils.rs
use std::str::FromStr;
use crate::types::Sheet;
use crate::types::PatternType;

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

pub fn factorial(n: i32) -> i32 {
    if n <= 1 { 1 } else { n * factorial(n - 1) }
}

pub fn triangular(n: i32) -> i32 {
    n * (n + 1) / 2
}

pub fn is_factorial_sequence(values: &[i32]) -> Option<(i32, i32)> {
    let forward_values: Vec<i32> = values.iter().rev().copied().collect();
    if forward_values.is_empty() {
        return None;
    }

    // Find the starting n for the first value
    let first_value = forward_values[0];
    let mut start_n = 0; // Start from 0! = 1
    while factorial(start_n) < first_value {
        start_n += 1;
    }
    if factorial(start_n) != first_value {
        return None; // First value isn’t a factorial
    }

    // Check if the sequence follows factorials from start_n
    for (i, &val) in forward_values.iter().enumerate() {
        let n = start_n + i as i32;
        if val != factorial(n) {
            return None;
        }
    }

    let last_value = forward_values.last().unwrap();
    let next_index = start_n + forward_values.len() as i32;
    Some((*last_value, next_index))
}

pub fn is_triangular_sequence(values: &[i32]) -> Option<(i32, i32)> {
    let forward_values: Vec<i32> = values.iter().rev().copied().collect();
    if forward_values.is_empty() {
        return None;
    }

    // Find the starting n for the first value
    let first_value = forward_values[0];
    let mut start_n = 1;
    while triangular(start_n) < first_value {
        start_n += 1;
    }
    if triangular(start_n) != first_value {
        return None; // First value isn’t a triangular number
    }

    // Check if the sequence follows triangular numbers from start_n
    for (i, &val) in forward_values.iter().enumerate() {
        let n = start_n + i as i32;
        if val != triangular(n) {
            return None;
        }
    }

    let last_value = forward_values.last().unwrap();
    let next_index = start_n + forward_values.len() as i32;
    Some((*last_value, next_index))
}
pub fn detect_pattern(sheet: &Sheet, start_row: i32, start_col: i32, end_row: i32, end_col: i32) -> PatternType {
    let mut values = Vec::new();

    if start_row == end_row { // Horizontal autofill
        for j in (0.max(start_col - 5)..start_col).rev() {
            values.push(sheet.cells[start_row as usize][j as usize].value);
        }
    } else if start_col == end_col { // Vertical autofill
        for i in (0.max(start_row - 5)..start_row).rev() {
            values.push(sheet.cells[i as usize][start_col as usize].value);
        }
    } else {
        return PatternType::Unknown;
    }

    if values.is_empty() {
        return PatternType::Unknown;
    }

    println!("Collected values: {:?}", values); // Debug

    // Constant pattern
    if values.len() >= 2 && values.iter().all(|&v| v == values[0]) {
        return PatternType::Constant(values[0]);
    }

    // Arithmetic pattern
    if values.len() >= 2 {
        let diffs: Vec<i32> = values.windows(2).map(|w| w[1] - w[0]).collect();
        if diffs.len() >= 1 && diffs.iter().all(|&d| d == diffs[0]) {
            return PatternType::Arithmetic(values[0], diffs[0]);
        }
    }

    // Triangular pattern
    if values.len() >= 2 {
        if let Some((last_value, next_index)) = is_triangular_sequence(&values) {
            return PatternType::Triangular(last_value, next_index);
        }
    }

    // Factorial pattern
    if values.len() >= 2 {
        if let Some((last_value, next_index)) = is_factorial_sequence(&values) {
            return PatternType::Factorial(last_value, next_index);
        }
    }

    // Fibonacci pattern
    if values.len() >= 3 {
        let forward_values: Vec<i32> = values.clone().into_iter().rev().collect();
        let is_fibonacci = forward_values.windows(3).all(|w| w[2] == w[0] + w[1]);
        if is_fibonacci {
            return PatternType::Fibonacci(forward_values[forward_values.len() - 2], forward_values[forward_values.len() - 1]);
        }
    }

    // Geometric pattern
    if values.len() >= 2 {
        let forward_values: Vec<i32> = values.clone().into_iter().rev().collect();
        let ratios: Vec<f64> = forward_values.windows(2)
            .map(|w| if w[0] == 0 { f64::INFINITY } else { w[1] as f64 / w[0] as f64 })
            .collect();
        if ratios.len() >= 1 && ratios.iter().all(|&r| (r - ratios[0]).abs() < 1e-10) && ratios[0].is_finite() {
            return PatternType::Geometric(forward_values[0], ratios[0]);
        }
    }

    PatternType::Unknown
}
    
    // Default to constant pattern with the last value if no other pattern is detected
pub fn is_valid_formula(sheet: &mut Sheet, formula: &str) -> bool {
    let formula = formula.trim();
    if sheet.extension_enabled{
    if let Some((func_name, args)) = formula.split_once('(') {
        if let Some(args) = args.strip_suffix(')') {
            let func_name = func_name.trim().to_uppercase();
            match func_name.as_str() {
                "SUM" | "AVG" | "MAX" | "MIN" | "STDEV" => {
                    return parse_range(sheet, args.trim()).is_some();
                }
                "SLEEP" => {
                    return args.parse::<i32>().is_ok() || parse_cell_reference(sheet, args.trim()).is_some();
                }
                "BOLD" | "ITALIC" | "UNDERLINE" => {
                    return parse_cell_reference(sheet, args.trim()).is_some();
                }
                "AUTOFILL" | "SORTA" | "SORTD" => {
                    return parse_range(sheet, args.trim()).is_some();}
                _ => return false,
            }
        }
    }
}
else { if let Some((func_name, args)) = formula.split_once('(') {
    if let Some(args) = args.strip_suffix(')') {
        let func_name = func_name.trim().to_uppercase();
        match func_name.as_str() {
            "SUM" | "AVG" | "MAX" | "MIN" | "STDEV" => {
                return parse_range(sheet, args.trim()).is_some();
            }
            "SLEEP" => {
                return args.parse::<i32>().is_ok() || parse_cell_reference(sheet, args.trim()).is_some();
            }
            // "BOLD" | "ITALIC" | "UNDERLINE" => {
            //     return parse_cell_reference(sheet, args.trim()).is_some();
            // }
            // "AUTOFILL" => {
            //     return parse_range(sheet, args.trim()).is_some();}
            _ => return false,
        }
    }
}
}

    if formula.contains('+') || formula.contains('-') || 
       formula.contains('*') || formula.contains('/') {
        let parts: Vec<&str> = formula.split(|c| "+-*/".contains(c))
            .map(|s| s.trim())
            .filter(|s| !s.is_empty())
            .collect();
        return parts.iter().all(|part| {
            part.parse::<i32>().is_ok() || parse_cell_reference(sheet, part).is_some()
        });
    }

    parse_cell_reference(sheet, formula).is_some() || formula.parse::<i32>().is_ok()
}

pub fn is_valid_command(sheet: &mut Sheet, command: &str) -> bool {
    if command.len() == 1 && "wasdq".contains(command) {
        return true;
    }
    if command == "disable_output" || command == "enable_output" {
        return true;
    }
    if sheet.extension_enabled && (command == "undo" || command == "redo" || 
    command.starts_with("FORMULA ") || command.starts_with("ROWDEL ") || 
    command.starts_with("COLDEL ")) {
        if command.starts_with("FORMULA ") {
            return parse_cell_reference(sheet, &command[8..].trim()).is_some();
        }
        if command.starts_with("ROWDEL ") {
            let row_str = &command[7..].trim();
            return row_str.parse::<i32>().map_or(false, |r| r >= 1 && r <= sheet.rows);
        }
        if command.starts_with("COLDEL ") {
            let col_str = &command[7..].trim();
            return col_str.chars().all(|c| c.is_ascii_alphabetic());
        }
        return true;
    }
    if command.starts_with("scroll_to ") {
        return parse_cell_reference(sheet, &command[10..]).is_some();
    }
    if command.starts_with("GRAPH ") {
        let parts: Vec<&str> = command.split_whitespace().collect();
        return parts.len() == 3 && 
               ["(BAR)", "(SCATTER)"].contains(&parts[1].to_uppercase().as_str()) && 
               parse_range(sheet, parts[2]).is_some();
    }

    if command.starts_with("COPY") {
        let range = if command.starts_with("COPY ") {
            &command[5..]
        } else {
            &command[4..]
        };
        return parse_range(sheet, range).is_some();
    }
    

    if command.starts_with("CUT") {
        let range = if command.starts_with("CUT ") {
            &command[4..]
        } else {
            &command[3..]
        };
        return parse_range(sheet, range).is_some();
    }
    

    if command.starts_with("PASTE") {
        let cell_ref = if command.starts_with("PASTE ") {
            &command[6..]
        } else {
            &command[5..]
        };
        return parse_cell_reference(sheet, cell_ref).is_some();
    }
    
    command.split_once('=').map_or(false, |(ref_str, formula)| {
        parse_cell_reference(sheet, ref_str.trim()).is_some() && 
        !formula.is_empty() && 
        is_valid_formula(sheet, formula)
    })
}
