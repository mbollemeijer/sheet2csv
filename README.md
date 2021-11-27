# sheet2csv

sheet2csv is a tool for those who need to convert Excel sheets to csv's on a regular basis.

### Why Rust? Aren't other languages better suited?   
Maybe, but I wanted to try Rust and thought this would be a fun use case to try and implement.

---

### Requirements
 - Rust (https://www.rust-lang.org/tools/install)
 - cargo, that hips with Rust.
 - Git

### Dependencies
- calamine = "0.18.0"
- serde = { version = "1.0", features = ["derive"] }
- serde_json = "1.0"

### Installation

1. Clone this repo
2. cd into the cloned repo.
3. run `cargo build` 

### Features / TO-DOS
- [x] Converts an Excel sheet to a csv file based on configuration
- [x] Configurable through a config file  
   - [x] Grab sheet by name
   - [x] Start from a row index 
   - [x] Set output file name
   - [x] Specify row end index
   - [x] Separator configurable 
   - [ ] Filters on column index and its contents
        - [ ] Operator: Equals
        - [ ] Operator: Not equals
        - [ ] Operator: Contains
        - [ ] Operator: Not contains
- [ ] Run without config
- [ ] CI/CD
- [ ] Make it available in AUR
- [ ] Cross platform ?
- [ ] Optimize code

**NOTE:** Strings are (for) now, always wrapped in double quotes -> ""   
**NOTE:** Default Separator is currently semicolon -> ;

### How to use sheet2csv

Currently, only source code compilation is supported, so you will need ```cargo```.


### Program arguments
|arguments       |description|
|----------|-----------|
|--source  | path to excel file
|--out     | Where the converted  sheets will be written
|--config  | path to config file.

**NOTE:** The above arguments are mandatory


One unique thing (I think) is that you can customize the conversion based on config, let me show you.
Example config for the sheet that resides in `<project-root>/examples/test.xlsx`  

### Settings

| property          | example value| description |
|--------------     |--------------|------------|
| sheetName         | Sheet1       | Name of the sheet (tab) in the excel file|
| outputFileName    | converted-files.csv      | Name of the file that the sheet will be converted into|
| startRowIndex     | 10           | Numeric value, 0 when omitted |
| endRowIndex       | 200          | Numeric value, -1 when omitted|
| separator         | ```,```      | Character used for separating values  |

Example `config.json` 
```json 
[
    {
        "sheetName": "Sheet1",
        "outputFileName": "test-sheet1.csv",
        "startRowIndex": 5,
        "endRowIndex": 100,
        "separator": ","
    },
    {
        "sheetName": "Sheet2",
        "outputFileName": "test-sheet2.csv"
    }
]
```

### Run the program
1. `cd` into folder where you cloned this repo
2. run `cargo run -- --source=<path to xlsx> --out=<path to output dir> --config=<path to config file>`
---
