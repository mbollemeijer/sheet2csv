# sheet2csv

sheet2csv converts Excel/Workbook files into CSV files through a JSON config file.
This makes repetitive conversions less error prone and more efficient.

### Why Rust? Aren't other languages better suited?   
Maybe, but I wanted to try Rust and thought this would be a fun use case to try and implement.

---

### Requirements
 - Rust (https://www.rust-lang.org/tools/install)
 - cargo, that ships with Rust.
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
   - [x] Substitute strings
   - [x] Wrap String values
   - [x] Specify wrap character(s)
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
|arguments |description| Mandatory |
|----------|-----------| ----------|
|--source  | path to excel file | No, see main config
|--out     | Where the converted  sheets will be written| No, see main config
|--config  | path to config file.| Yes, for now |

----
### Settings

#### Main Config
| property | example | default | description|
| ---------|---------|---------|----------- |
| `sourceFile`| `/home/username/path-to-file`| No defaults| Path to workbook - can also be specified by giving `--source=<path-to-file>` as command line argument, overriding the property set in the config|
| `ouputPath`| `/home/username/output`| No defaults| Path to output folder - can also be specified by giving `--out=<path-to-output` as command line argument, overriding the property set in the config|
| `wrapStrings`| `true` or `false`| `true` | When true, cells that identify as string would be wrapped in the csv 
| `wrapStringChar` | any char(s): `123`| `"` | When specified all string values are wrapped by the specified char `123example123`|
| `sheetSettings`| Array of sheet settings| No defaults | Mandatory, see SheetSettings for more details|
|`substitutions` | Array of subs | No defaults| See subs for more details|

```json 
{
    "sourceFile": "<path>",
    "outputPath": "<path>",
    "wrapStrings": true,
    "wrapStringChar": "\"",
    "sheetSettings": [], # Sheet settings ommited 
    "substitutions ": [] # Subs ommited
}
```

#### SheetSettings

| property | example | default | description |
| -------- | ------- | ------- | ----------- |
| `sheetName`| `Sheet1`  | No default| The sheet/tab you want to convert|
| `outputFileName`| `sheet1.csv`| No default| The name you wnat your output file to have|
| `startRowIndex`| `5`| `0`| The index you want to start the processing of your sheet| 
| `endRowIndex` | `15`| `-1` | The index you want the processing of your sheet to stop| 
| `wrapStrings`| `false`| `true` | This setting is inherited by the main config, but can be overriden|
| `wrapStringChar`| any char(s) `<`| `"`| Also inherited by the main config if not specified on sheet settings|
| `substitutions`| Array of subs | No default| Subs specified by the main config are also inherited on each sheet config

Below an partial example of a sheet configuration
```json 
{
    "sheetName": "Sheet1",
    "outputFileName": "fancy-name-of-file1.csv",
    "startRowIndex": 5,
    "endRowIndex": 10,
    "wrapStrings": true,
    "wrapStringChar": "<>",
    "separator": ",",
    "substitutions": [] # Subs omitted
}
```
#### Substitutions
| Property | Description | 
| -------- | ----------- |
| `matchOn`| The character(s) that should be substituted|
| `replaceWith`| The character(s) that will replace the matched character(s)|

Example
```json
{
    "matchOn": "\"",
    "replaceWith": "\"\""
}
```
### All together
```json 
{
	"sourceFile": "/home/username/path/to/some/excel.xlsx",
	"outputPath": "/home/username/path/to/some/output/folder",
	"wrapStrings": true,
	"wrapStringChar": "\"",
	"substitutions": [{
		"matchOn": "\"",
		"replaceWith": "\"\""
	}],
	"sheetSettings": [{
		"sheetName": "Sheet1",
		"outputFileName": "fancy-name-of-file1.csv",
		"startRowIndex": 5,
		"wrapStrings": true,
		"wrapStringChar": "<>"
	}, {
		"sheetName": "Sheet2",
		"outputFileName": "second-fancy-name-of-file.csv",
		"startRowIndex": 4,
		"wrapStrings": false
	}]
}
```

### Run the program
1. `cd` into folder where you cloned this repo
2. run `cargo run -- --config=<path to config file>`
---
