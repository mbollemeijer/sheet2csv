use std::{env, fs::{self, File}, path::PathBuf};

use calamine::{DataType, Reader, Xlsx, open_workbook};
use serde::{Deserialize, Serialize};

use std::io::{BufWriter, Write};

#[derive(Serialize, Deserialize, Debug)]
#[serde( rename_all = "camelCase")]
struct SheetSettings {
    sheet_name: String,
    output_file_name: String,
    start_row_index: i8,
}

fn main() {

    // For debug purposes
    let args: Vec<String> = env::args().collect();
    println!("all arguments in a list are {:?}", &args);
    
    // Grabbing the arguments
    let source_file = get_program_argument(&args, "--source");
    let output_path = get_program_argument(&args, "--out");
    let config_path = get_program_argument(&args, "--config");
    let settings = get_settings(&config_path);

    let _result = convert_workbook_to_csv(&source_file, &output_path, &settings);
}

fn convert_workbook_to_csv(source_file: &str, output_path: &str, settings: &Vec<SheetSettings>) -> std::io::Result<()> {

    for setting in settings {

        let mut workbook: Xlsx<_> = open_workbook(source_file).expect("Cannot open file");
        // Read whole worksheet data and provide some statistics
        if let Some(Ok(sheet)) = workbook.worksheet_range(&setting.sheet_name) {
            let sce = PathBuf::from(output_path.to_owned() + "/" + &setting.output_file_name);
            let dest = sce.with_extension("csv");
            let mut dest = BufWriter::new(File::create(dest).unwrap());

            
            for (index, row) in sheet.rows().enumerate() {
                if index >= setting.start_row_index as usize {
                    for (_i, c) in row.iter().enumerate() {

                        match *c {
                            DataType::Empty => write!(dest, ";"),
                            DataType::String(ref s) => write!(dest, "\"{}\";", s),
                            DataType::Int(ref i) => write!(dest, "{};", i),
                            DataType::Float(ref f) => write!(dest, "{};", f),
                            DataType::Bool(ref b) => write!(dest, "{};", b),
                            DataType::DateTime(ref dt) => write!(dest, "{};",dt),
                            DataType::Error(ref e) => write!(dest, "{};", e),
                        }?;
                    }
                    write!(dest, "\n")?;
                }
            }
        }
    }
    Ok(())
}

fn get_settings(path: &str) -> Vec<SheetSettings> {

    let settings = fs::read_to_string(path).unwrap();
    let vec_of_sheet_settings: Vec<SheetSettings> = serde_json::from_str(&settings).unwrap();
    return vec_of_sheet_settings;
}

fn get_program_argument(args: &Vec<String>, program_argument: &str) -> String {

    let value = args.iter()
        .filter(|e| e.contains("--"))
        .find(|e| e.contains(program_argument));

    match value {
        None => format!("argument {} is mandatory but was not found", program_argument),
        Some(value) => {
            let val = value.split("=").nth(1);
            match val {
                None => format!("found argument {} but no value was given", program_argument),
                Some(val) => val.to_string()
            }
        }
    }
}

