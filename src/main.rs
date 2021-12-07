use std::{env, fs::{self, File}, path::PathBuf, io::Result};

use calamine::{DataType, Reader, Xlsx, open_workbook, Range};
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

    let args: Vec<String> = env::args().collect();

    // Grabbing the arguments
    let source_file = get_program_argument(&args, "--source");
    let output_path = get_program_argument(&args, "--out");
    let config_path = get_program_argument(&args, "--config");
    let settings = get_settings(config_path, source_file, output_path);

    // let result = convert_workbook_to_csv(settings);
    process_workbook(&settings);
}

fn process_workbook(main_setting: &Settings) {
    let mut workbook: Xlsx<_> = open_workbook(&main_setting.source_file).expect(format!("Could not open file @ {} ", &main_setting.source_file).as_str());
    
    for sheet_settings in &main_setting.sheet_settings {
        if let Some(Ok(sheet)) = workbook.worksheet_range(&sheet_settings.sheet_name) {
            let _res = process_sheet(sheet, main_setting, sheet_settings);
        }
    }
}

fn process_sheet(sheet: Range<DataType>, main_setting: &Settings, sheet_settings: &SheetSettings) -> Result<()> {

    // These settings are combined and can change from one sheet to another
    let subs = get_subs(main_setting, sheet_settings);
    let wrap_settings = get_wrap_settings(main_setting, sheet_settings);

    let sce = PathBuf::from(main_setting.output_path.to_owned() + "/" + &sheet_settings.output_file_name);
    let dest = sce.with_extension("csv");
    let mut dest = BufWriter::new(File::create(dest).unwrap());
    
    for (index, row) in sheet.rows().enumerate() {

        if index >= sheet_settings.start_index_or_default() as usize && index <= sheet_settings.end_index_or_default() as usize{
            process_row(row, &subs, &wrap_settings, &mut dest)?;
            write!(dest, "\r\n")?;
        }
    }

    Ok(())
}

fn process_row(row: &[DataType], subs: &Option<Vec<&Substitution>>, wrap_settings: &(bool, String, String), dest: &mut BufWriter<File>) -> Result<()>{

    for (_index, cell) in row.iter().enumerate() {
        // Process cell values
        let cell_value = process_cell(cell, subs);
        // Check if we need to wrap the value
        if wrap_settings.0 {
            write!(dest, "{wrap}{val}{wrap}{sep}", wrap = wrap_settings.1, val = cell_value, sep = wrap_settings.2)?;
        }

        write!(dest, "{val}{sep}", val=cell_value, sep= wrap_settings.2)?;
    }

    Ok(())
}

fn process_cell(cell: &DataType, subs: &Option<Vec<&Substitution>>) -> String { 

    match *cell {
        DataType::Empty => String::new(),
        DataType::String(ref string_value) => format!("{}", sub_values(subs, &string_value)),
        DataType::Int(ref i) => format!("{}", i),
        DataType::Float(ref f) => format!("{}", f),
        DataType::Bool(ref b) => format!("{}", b),
        DataType::DateTime(ref dt) => format!("{}",dt),
        DataType::Error(ref e) => format!("{}", e)
    }
}

fn sub_values(subs: &Option<Vec<&Substitution>>, cell_value: &str) -> String {

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

fn get_wrap_settings(main_setting: &Settings, sheet_settings: &SheetSettings) -> (bool, String, String) {

    let wrap = match sheet_settings.wrap_strings {
        Some(value) => value,
        None => match main_setting.wrap_strings {
            None => true,
            Some(value) => value
        }
    };

    let wrap_char = match sheet_settings.wrap_string_char.as_ref() {
        None =>  match main_setting.wrap_string_char.as_ref() {
            None => String::from("\""),
            Some(value) => value.to_string()
        }
        Some(v) => v.to_string()
    };

    (wrap, wrap_char, sheet_settings.separator_or_default())
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

fn get_subs<'a>(main_setting: &'a Settings, sheet_settings: &'a SheetSettings) -> Option<Vec<&'a Substitution>> {

   let main_subs = main_setting.substitutions.as_ref();

    let main_subs_cloned = main_subs.clone();
    let mut subs: Vec<&Substitution> = Vec::new();
    for sub in main_subs_cloned.unwrap() {
        subs.push(sub.clone());
    }

    if sheet_settings.substitutions.is_some() {
        for sheet_sub in sheet_settings.substitutions.as_ref().unwrap() {
            subs.push(&sheet_sub);
        }
    }

    Some(subs)
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
