# sheet2csv

sheet2csv is a handy tool for those who need to convert large Excel sheets to csv's on a regular basis.

### Why ?
For one of the projects I work on on a daily basis, we occasionally need to convert certain .xlsx files to separate .csv files.   
From one .xlsx with multiple sheets into several single .csv files And because this process was done manually, copy N amount of rows and pasting them into a new workbook and saving them as a .csv This is of course an error prone process and I wanted to automate it.

### Why Rust arent other languages better suited?   
Maybe, but I wanted to try Rust and thought this would be a fun usecase to try and implement.

---

### Requirements
 - Rust (https://www.rust-lang.org/tools/install)
 - cargo, that ships with Rus.
 - Git

### Dependencies
- calamine = "0.18.0"
- serde = { version = "1.0", features = ["derive"] }
- serde_json = "1.0"

### Installation

1. Clone this repo
2. cd into the cloned repo.
3. run `cargo build` 

### Features / TODOS
- [x] Converts an Excel sheet to a csv file based on configuration
- [x] Configurable through a config file  
   - [x] Grab sheet by name
   - [x] Start from a row index 
   - [x] Set output file name
   - [x] Specify row end index
   - [ ] Filters on column index and its contents
        - [ ] Operator: Equals
        - [ ] Operator: Not eqauls
        - [ ] Operator: Contains
        - [ ] Operator: Not contains
   - [ ] Seperator configurable 
- [ ] Run without config
- [ ] CI/CD
- [ ] Make it available in AUR
- [ ] Cross platform ?

**NOTE: Strings are (for) now, always wrappend in double quotes -> "**   
**NOTE: Default Seperator is currently semicolon -> ;**

### How to use sheet2csv

Currently only source code compilation is supported.

1. `cd` into folder where you cloned this repo
2. run `cargo run -- --source=<path to xlsx> --out=<path to output dir> --config=<path to config file>`
3. Profit

One unqiue thing (I think) is that you can customize the conversion based on config, let me show you.
Example config for the sheet that resides in `<project-root>/examples/test.xlsx`  

### Settings

| property          | example value| description |
|--------------     |--------------|------------|
| sheetName         | Sheet1       | Name of the sheet (tab) in the excel file|
| outputFileName    | converted-files.csv      | Name of the file that the sheet will be converted into|
| startRowIndex     | 10           | Numeric value, 0 when omitted |
| endRowIndex       | 200          | Numeric value, -1 when omitted|

`config.json`
```json 
[
    {
        "sheetName": "Sheet1",
        "outputFileName": "test-sheet1.csv",
        "startRowIndex": 5,
        "endRowIndex": 100
    },
    {
        "sheetName": "Sheet2",
        "outputFileName": "test-sheet2.csv"
    }
]
```
---
