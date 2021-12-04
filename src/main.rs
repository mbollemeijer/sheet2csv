use std::{env, fs::{self, File}, path::PathBuf, time::Instant};

use calamine::{DataType, Reader, Xlsx, open_workbook};
use serde::{Deserialize, Serialize};

use std::io::{BufWriter, Write};

#[derive(Serialize, Deserialize, Debug)]
#[serde( rename_all = "camelCase")]
struct Settings {
    source_file: String,
    output_path: String,
    sheet_settings: Vec<SheetSettings>,
    substitutions: Option<Vec<Substitution>>,
    wrap_strings: Option<bool>,
    wrap_string_char: Option<String>
}

#[derive(Serialize, Deserialize, Debug)]
#[serde( rename_all = "camelCase")]
struct SheetSettings {
    sheet_name: String,
    output_file_name: String,
    start_row_index: Option<i32>,
    end_row_index: Option<i32>,
    separator: Option<String>,
    substitutions: Option<Vec<Substitution>>,
    wrap_strings: Option<bool>,
    wrap_string_char: Option<String>
}

#[derive(Serialize, Deserialize, Debug)]
#[serde( rename_all = "camelCase")]
struct Substitution {
    match_on: String,
    replace_with: String
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
    let settings = get_settings(config_path, source_file, output_path);

    let result = convert_workbook_to_csv(settings);

    println!("{:?}", result);
}

fn convert_workbook_to_csv(settings: Settings) -> std::io::Result<()> {

    let start_on_workbook_time = Instant::now();
    // Read whole worksheet data and provide some statistics
    println!("opening workbook: {}", &settings.source_file);
    let mut workbook: Xlsx<_> = open_workbook(&settings.source_file).expect(format!("Could not open file @ {} ", &settings.source_file).as_str());

    for sheet_settings in settings.sheet_settings {
        let sheet_time = Instant::now();
        if let Some(Ok(sheet)) = workbook.worksheet_range(&sheet_settings.sheet_name) {
            let sce = PathBuf::from(settings.output_path.to_owned() + "/" + &sheet_settings.output_file_name);
            let dest = sce.with_extension("csv");
            let mut dest = BufWriter::new(File::create(dest).unwrap());
            

            let wrap = match sheet_settings.wrap_strings {
                Some(value) => value,
                None => match settings.wrap_strings {
                    None => true,
                    Some(value) => value
                }
            };

            let wrap_char = match sheet_settings.wrap_string_char.as_ref() {
                None =>  match settings.wrap_string_char.as_ref() {
                    None => String::from("\""),
                    Some(value) => value.to_string()
                }
                Some(v) => v.to_string()
            };

            for (index, row) in sheet.rows().enumerate() {
                if index >= sheet_settings.start_index_or_default() as usize 
                    && index <= sheet_settings.end_index_or_default() as usize {
                    for (_i, c) in row.iter().enumerate() {
                        match *c {
                            DataType::Empty => write!(dest, "{}", sheet_settings.separator_or_default()),
                            DataType::String(ref string_value) 
                                => write!(dest, "{}{s}", 
                                          process_cell(&settings.substitutions, &wrap, &wrap_char, string_value), s = sheet_settings.separator_or_default()), 
                            DataType::Int(ref i) => write!(dest, "{}{}", i, sheet_settings.separator_or_default()),
                            DataType::Float(ref f) => write!(dest, "{}{}", f, sheet_settings.separator_or_default()),
                            DataType::Bool(ref b) => write!(dest, "{}{}", b, sheet_settings.separator_or_default()),
                            DataType::DateTime(ref dt) => write!(dest, "{}{}",dt, sheet_settings.separator_or_default()),
                            DataType::Error(ref e) => write!(dest, "{}{}", e, sheet_settings.separator_or_default()),
                        }?;
                    }
                    write!(dest, "\r\n")?;
                } 
            }

            println!("{} {}s", sheet_settings.sheet_name, sheet_time.elapsed().as_secs());
        }
    }
    println!("Total: {}s", start_on_workbook_time.elapsed().as_secs());
    Ok(())
}

fn process_cell(subs: &Option<Vec<Substitution>>, wrap: &bool, wrap_char: &String, cell_value: &str) -> String {

    let subbed_value = sub_values(&subs, cell_value);
    match wrap {
        true => format!("{wrap_char}{value}{wrap_char}", value = subbed_value, wrap_char = wrap_char),
        false => subbed_value
    }
}

fn sub_values(subs: &Option<Vec<Substitution>>, cell_value: &str) -> String {

    let mut subbed_value = String::from(cell_value);
    if subs.is_some() {
        for sub in subs.as_ref().unwrap() {
            if subbed_value.contains(&sub.match_on) {
                subbed_value = subbed_value.replace(&sub.match_on, &sub.replace_with);
            }
        }
    }

   subbed_value 
}


fn get_settings(config_path: Option<String>, source_file: Option<String>, output_path: Option<String>) -> Settings {

    let settings_file = fs::read_to_string(config_path.unwrap()).expect("Could not read file ");
    let mut settings: Settings = serde_json::from_str(&settings_file).expect("Could not parse settings file");
    if source_file.is_some() {
        settings.source_file = source_file.unwrap();
    }

    if output_path.is_some() {
        settings.output_path = output_path.unwrap();
    }

    return settings;
}

fn get_program_argument(args: &Vec<String>, program_argument: &str) -> Option<String> {

    let value = args.iter()
        .filter(|e| e.contains("--"))
        .find(|e| e.contains(program_argument));

    match value {
        None => None,
        Some(value) => {
            let val = value.split("=").nth(1);
            match val {
                None => None,
                Some(val) => Some(val.to_string())
            }
        }
    }
}

