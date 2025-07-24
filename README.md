# Direct.CSV.Library
Libary that provides new methods to work with CSV files

## Main features:
### Import CSV To DataTable

**Description**:
- Parses provided CSV file line by line to separate it by given delimiter. If delimiter was used in an individual value but was correctly escaped with "" then the value will be included in the result.

**Expected parameters:**
- `File Path` - Full file path to CSV file
- `Delimiter` - Delimiter used to separate values in CSV file
- `Has Header` - Whether or not CSV file contains Header row and if it should be added as Column Row to result DataTable

**Returns**:
- `DataTable` with values from CSV file

**Where to find**:
- `Library Objects` -> `CSV Actions`

### Convert CSV to Xlsx

**Description**:
- Converts given CSV file into Excel xlsx file

**Expected parameters:**
- `File Path` - Full file path to CSV file
- `Delimiter` - Delimiter used to separate values in CSV file
- `Excel File Path` - Full file path where created Xlsx should be saved

**Returns**:
- `True` - when the conversion has finished successfully
- `False` - when the conversion has failed

**Where to find**:
- `Library Objects` -> `CSV Actions`

### Import CSV to List of Rows

**Description**:
- Parses provided CSV file line by line to separate it by given delimiter. If delimiter was used in an individual value but was correctly escaped with "" then the value will be included in the result.

**Expected parameters:**
- `File Path` - Full file path to CSV file
- `Delimiter` - Delimiter used to separate values in CSV file
- `Has Header` - Whether or not CSV file contains Header row and if it should be added as first item into result List of Rows

**Returns**:
- `List of Rows` with values from CSV file

**Where to find**:
- `Library Objects` -> `CSV Actions`

### Export List Of Rows to CSV

**Description**  
- Exports the provided **List Of Rows** to a CSV file at the specified path.
  - Overwrites the file if it exists.
  - Any occurrence of the chosen delimiter inside a cell is replaced with the provided *delimiterReplacer*.  
  - Quotes inside cells are doubled, and fields containing quotes or CR/LF are wrapped in quotes.

**Parameters**
- **rows** – The List Of Rows to export (`DirectCollection<DirectRow>`).  
- **filePath** – Full path of the output CSV file.  
- **delimiter** – Delimiter between values (default `","`).  
- **delimiterReplacer** – String used to replace the delimiter when it appears inside a cell (default `"."`).  
- **header** – Optional single header line to write as the first row (empty or null to skip).

**Returns**  
- **bool** – `true` on success; `false` on failure (errors are logged).

**Where to find**:
- `Library Objects` -> `CSV Actions`

### Install:

- Copy Direct.CSV.Library.dll to NICE Designer and NICE Client installation directory
- Reference Dll in project: http://apa-onlinehelp.s3-website-us-west-2.amazonaws.com/72/content/library%20objects%20sdk/installing%20the%20library%20objects%20project.htm

### Verified Compatibility:

- NICE APA 7.3
- NICE APA 7.4
- NICE APA 7.5
- NICE Automation Studio

Disclaimer: this is a product of PAteam meant for the NICE community and is not created or supported by NICE