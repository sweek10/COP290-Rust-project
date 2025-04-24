use crate::types::{PatternType, Sheet};
use std::str::FromStr;

pub fn parse_cell_reference(sheet: &mut Sheet, ref_str: &str) -> Option<(i32, i32)> {
    let ref_str = ref_str.trim();
    let num_start = ref_str.chars().position(|c| c.is_ascii_digit())?;
    let (col_str, row_str) = ref_str.split_at(num_start);

    if col_str.is_empty() {
        return None;
    }

    let col = decode_column(col_str);
    let row = i32::from_str(row_str).ok()? - 1;

    if row >= 0 && row < sheet.rows && col >= 0 && col < sheet.cols {
        Some((row, col))
    } else {
        None
    }
}

pub fn parse_range(sheet: &mut Sheet, range: &str) -> Option<(i32, i32, i32, i32)> {
    let (start, end) = range.split_once(':')?;
    let (start_row, start_col) = parse_cell_reference(sheet, start)?;
    let (end_row, end_col) = parse_cell_reference(sheet, end)?;
    if start_row <= end_row && start_col <= end_col {
        Some((start_row, start_col, end_row, end_col))
    } else {
        None
    }
}

pub fn calculate_range_function(sheet: &mut Sheet, function: &str, range: &str) -> Result<f64, ()> {
    let (start_row, start_col, end_row, end_col) = match parse_range(sheet, range) {
        Some(range) => range,
        None => return Err(()),
    };

    let function = function.to_uppercase();
    let mut count: usize = 0;
    let mut sum: f64 = 0.0;
    let mut min: f64 = f64::MAX;
    let mut max: f64 = f64::MIN;
    // For STDEV: Welford's online algorithm variables
    let mut mean: f64 = 0.0;
    let mut m2: f64 = 0.0;

    for i in start_row..=end_row {
        for j in start_col..=end_col {
            let cell = &sheet.cells[i as usize][j as usize];
            if cell.is_error {
                return Err(());
            }
            let value = cell.value as f64;
            count += 1;

            // Update aggregates
            sum += value;
            min = min.min(value);
            max = max.max(value);

            // Welford's algorithm for variance
            if function == "STDEV" {
                let delta = value - mean;
                mean += delta / count as f64;
                let delta2 = value - mean;
                m2 += delta * delta2;
            }
        }
    }

    if count == 0 {
        return Err(());
    }

    match function.as_str() {
        "SUM" => Ok(sum),
        "AVG" => Ok(sum / count as f64),
        "MIN" => Ok(min),
        "MAX" => Ok(max),
        "STDEV" => {
            if count <= 1 {
                Ok(0.0) // Consistent with original behavior
            } else {
                let variance = m2 / count as f64;
                Ok(variance.sqrt().round())
            }
        }
        _ => Err(()),
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
    if n <= 1 {
        1
    } else {
        n * factorial(n - 1)
    }
}

pub fn triangular(n: i32) -> i32 {
    n * (n + 1) / 2
}

pub fn is_factorial_sequence(values: &[i32]) -> Option<(i32, i32)> {
    let forward_values: Vec<i32> = values.iter().rev().copied().collect();
    if forward_values.is_empty() {
        return None;
    }

    let first_value = forward_values[0];
    let mut possible_start_n = vec![];

    // Find all possible start_n where factorial(start_n) == first_value
    let mut start_n = 0;
    while factorial(start_n) <= first_value {
        if factorial(start_n) == first_value {
            possible_start_n.push(start_n);
        }
        start_n += 1;
        if start_n > 20 { // Prevent excessive computation
            break;
        }
    }

    // Try each possible start_n to find a matching sequence
    for start_n in possible_start_n {
        let mut is_valid = true;
        for (i, &val) in forward_values.iter().enumerate() {
            let n = start_n + i as i32;
            if val != factorial(n) {
                is_valid = false;
                break;
            }
        }
        if is_valid {
            let last_value = *forward_values.last().unwrap();
            let next_index = start_n + forward_values.len() as i32;
            return Some((last_value, next_index));
        }
    }

    None
}

pub fn is_triangular_sequence(values: &[i32]) -> Option<(i32, i32)> {
    let forward_values: Vec<i32> = values.iter().rev().copied().collect();
    if forward_values.is_empty() {
        return None;
    }

    let first_value = forward_values[0];
    let mut start_n = 1;
    while triangular(start_n) < first_value {
        start_n += 1;
    }
    if triangular(start_n) != first_value {
        return None;
    }

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

pub fn detect_pattern(
    sheet: &Sheet,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
) -> PatternType {
    let mut values = Vec::new();

    if start_row == end_row {
        for j in (0.max(start_col - 5)..start_col).rev() {
            values.push(sheet.cells[start_row as usize][j as usize].value);
        }
    } else if start_col == end_col {
        for i in (0.max(start_row - 5)..start_row).rev() {
            values.push(sheet.cells[i as usize][start_col as usize].value);
        }
    } else {
        return PatternType::Unknown;
    }

    if values.is_empty() {
        return PatternType::Unknown;
    }

    if values.len() >= 2 && values.iter().all(|&v| v == values[0]) {
        return PatternType::Constant(values[0]);
    }

    if values.len() >= 2 {
        let diffs: Vec<i32> = values.windows(2).map(|w| w[1] - w[0]).collect();
        if diffs.iter().all(|&d| d == diffs[0]) {
            return PatternType::Arithmetic(values[0], diffs[0]);
        }
    }

    if values.len() >= 2 {
        if let Some((last_value, next_index)) = is_triangular_sequence(&values) {
            return PatternType::Triangular(last_value, next_index);
        }
    }

    if values.len() >= 2 {
        if let Some((last_value, next_index)) = is_factorial_sequence(&values) {
            return PatternType::Factorial(last_value, next_index);
        }
    }

    if values.len() >= 3 {
        let forward_values: Vec<i32> = values.clone().into_iter().rev().collect();
        let is_fibonacci = forward_values.windows(3).all(|w| w[2] == w[0] + w[1]);
        if is_fibonacci {
            return PatternType::Fibonacci(
                forward_values[forward_values.len() - 2],
                forward_values[forward_values.len() - 1],
            );
        }
    }

    if values.len() >= 2 {
        let forward_values: Vec<i32> = values.clone().into_iter().rev().collect();
        let ratios: Vec<f64> = forward_values
            .windows(2)
            .map(|w| {
                if w[0] == 0 {
                    f64::INFINITY
                } else {
                    w[1] as f64 / w[0] as f64
                }
            })
            .collect();
        if ratios.iter().all(|&r| (r - ratios[0]).abs() < 1e-10) && ratios[0].is_finite() {
            return PatternType::Geometric(forward_values[0], ratios[0]);
        }
    }

    PatternType::Unknown
}

pub fn is_valid_formula(sheet: &mut Sheet, formula: &str) -> bool {
    let formula = formula.trim();
    if sheet.extension_enabled {
        if let Some((func_name, args)) = formula.split_once('(') {
            if let Some(args) = args.strip_suffix(')') {
                let func_name = func_name.trim().to_uppercase();
                match func_name.as_str() {
                    "SUM" | "AVG" | "MAX" | "MIN" | "STDEV" | "SORTA" | "SORTD" | "AUTOFILL" => {
                        return parse_range(sheet, args.trim()).is_some();
                    }
                    "SLEEP" => {
                        return args.parse::<i32>().is_ok()
                            || parse_cell_reference(sheet, args.trim()).is_some();
                    }
                    "BOLD" | "ITALIC" | "UNDERLINE" => {
                        return parse_cell_reference(sheet, args.trim()).is_some();
                    }
                    _ => return false,
                }
            }
        }
    } else if let Some((func_name, args)) = formula.split_once('(') {
        if let Some(args) = args.strip_suffix(')') {
            let func_name = func_name.trim().to_uppercase();
            match func_name.as_str() {
                "SUM" | "AVG" | "MAX" | "MIN" | "STDEV" => {
                    return parse_range(sheet, args.trim()).is_some();
                }
                "SLEEP" => {
                    return args.parse::<i32>().is_ok()
                        || parse_cell_reference(sheet, args.trim()).is_some();
                }
                _ => return false,
            }
        }
    }

    if formula.contains('+')
        || formula.contains('-')
        || formula.contains('*')
        || formula.contains('/')
    {
        let parts: Vec<&str> = formula
            .split(|c| "+-*/".contains(c))
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
    if sheet.extension_enabled && (command == "undo" || command == "redo") {
        return true;
    }
    if sheet.extension_enabled {
        if let Some(stripped) = command.strip_prefix("FORMULA ") {
            return parse_cell_reference(sheet, stripped.trim()).is_some();
        }
    }
    if sheet.extension_enabled {
        if let Some(stripped) = command.strip_prefix("ROWDEL ") {
            return stripped
                .trim()
                .parse::<i32>()
                .is_ok_and(|r| r >= 1 && r <= sheet.rows);
        }
    }
    if sheet.extension_enabled {
        if let Some(stripped) = command.strip_prefix("COLDEL ") {
            return stripped.trim().chars().all(|c| c.is_ascii_alphabetic());
        }
    }
    if let Some(stripped) = command.strip_prefix("scroll_to ") {
        return parse_cell_reference(sheet, stripped).is_some();
    }
    if sheet.extension_enabled && command.starts_with("GRAPH ") {
        let parts: Vec<&str> = command.split_whitespace().collect();
        return parts.len() == 3
            && ["(BAR)", "(SCATTER)"].contains(&parts[1].to_uppercase().as_str())
            && parse_range(sheet, parts[2]).is_some();
    }
    if sheet.extension_enabled {
        let range = if let Some(stripped) = command.strip_prefix("COPY ") {
            stripped
        } else {
            &command[4..]
        };
        return parse_range(sheet, range).is_some();
    }
    if sheet.extension_enabled {
        let range = if let Some(stripped) = command.strip_prefix("CUT ") {
            stripped
        } else {
            &command[3..]
        };
        return parse_range(sheet, range).is_some();
    }
    if sheet.extension_enabled {
        let cell_ref = if let Some(stripped) = command.strip_prefix("PASTE ") {
            stripped
        } else {
            &command[5..]
        };
        return parse_cell_reference(sheet, cell_ref).is_some();
    }
    command.split_once('=').is_some_and(|(ref_str, formula)| {
        parse_cell_reference(sheet, ref_str.trim()).is_some()
            && is_valid_formula(sheet, formula.trim())
    })
}
