use std::{env, fs::{self, File}, path::PathBuf};

use calamine::{DataType, Reader, Xlsx, open_workbook};
use serde::{Deserialize, Serialize};

use std::io::{BufWriter, Write};

#[derive(Serialize, Deserialize, Debug)]
#[serde( rename_all = "camelCase")]
struct SheetSettings {
    sheet_name: String,
    output_file_name: String,
    start_row_index: Option<i32>,
    end_row_index: Option<i32>,
    separator: Option<String>
}

impl SheetSettings {

    /// Returns the value configured for this sheet
    /// or the default index row which is zero (0)
    fn start_index_or_default(&self) -> i32 {
        match self.start_row_index {
            None => 0,
            Some(value) => value
        }
    }

    /// Returns the value configured for this sheet
    /// or the defaul value which is -1
    fn end_index_or_default(&self) -> i32 {
        match self.end_row_index {
            None => -1,
            Some(value) => value
        }
    }

    fn separator_or_default(&self) -> String {
        match &self.separator {
            None => String::from(";"),
            Some(value) => value.to_string()
        }
    }
}

fn main() {

    // For debug purposes
    let args: Vec<String> = env::args().collect();
    
    // Grabbing the arguments
    let source_file = get_program_argument(&args, "--source");
    let output_path = get_program_argument(&args, "--out");
    let config_path = get_program_argument(&args, "--config");
    let settings = get_settings(&config_path);

    let _result = convert_workbook_to_csv(&source_file, &output_path, &settings);
}

fn convert_workbook_to_csv(source_file: &str, output_path: &str, settings: &Vec<SheetSettings>) -> std::io::Result<()> {

    // Read whole worksheet data and provide some statistics
    let mut workbook: Xlsx<_> = open_workbook(source_file).expect("Cannot open file");
    for setting in settings {
    
        if let Some(Ok(sheet)) = workbook.worksheet_range(&setting.sheet_name) {
            let sce = PathBuf::from(output_path.to_owned() + "/" + &setting.output_file_name);
            let dest = sce.with_extension("csv");
            let mut dest = BufWriter::new(File::create(dest).unwrap());

            for (index, row) in sheet.rows().enumerate() {
                if index >= setting.start_index_or_default() as usize 
                    && index <= setting.end_index_or_default() as usize {
                    for (_i, c) in row.iter().enumerate() {
                        match *c {
                            DataType::Empty => write!(dest, "{}", setting.separator_or_default()),
                            DataType::String(ref s) => write!(dest, "\"{}\"{}", s, setting.separator_or_default()),
                            DataType::Int(ref i) => write!(dest, "{}{}", i, setting.separator_or_default()),
                            DataType::Float(ref f) => write!(dest, "{}{}", f, setting.separator_or_default()),
                            DataType::Bool(ref b) => write!(dest, "{}{}", b, setting.separator_or_default()),
                            DataType::DateTime(ref dt) => write!(dest, "{}{}",dt, setting.separator_or_default()),
                            DataType::Error(ref e) => write!(dest, "{}{}", e, setting.separator_or_default()),
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

