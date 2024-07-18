use std::env;
use std::process;

#[derive(Debug)]
enum ConversionError {
    InvalidColumnLetter,
    InvalidColumnNumber,
}

fn excel_column_to_number(column: &str) -> Result<u32, ConversionError> {
    let mut number: u32 = 0;
    for c in column.chars() {
        if !c.is_ascii_alphabetic() {
            return Err(ConversionError::InvalidColumnLetter);
        }
        number = number * 26 + (c.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
    }
    Ok(number)
}

fn number_to_excel_column(mut number: u32) -> Result<String, ConversionError> {
    if number == 0 {
        return Err(ConversionError::InvalidColumnNumber);
    }
    let mut column = String::new();
    while number > 0 {
        number -= 1;
        column.insert(0, ((number % 26) as u8 + b'A') as char);
        number /= 26;
    }
    Ok(column)
}

fn main() {
    let args: Vec<String> = env::args().collect();

    if args.len() != 2 {
        eprintln!("usage: xcol [col_letter] or xcol [col_number]");
        process::exit(1);
    }

    let input = &args[1];

    if let Ok(number) = input.parse::<u32>() {
        // Number to Column Letter
        match number_to_excel_column(number) {
            Ok(column) => println!("{}", column),
            Err(_) => {
                eprintln!("Invalid column number");
                process::exit(1);
            }
        }
    } else {
        // Column Letter to Number
        match excel_column_to_number(input) {
            Ok(number) => println!("{}", number),
            Err(_) => {
                eprintln!("Invalid column letter");
                process::exit(1);
            }
        }
    }
}
