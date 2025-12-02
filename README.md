# Excel Find and Log

A Python script that searches through Excel files for matching values across two columns and logs all matches to timestamped result files. Perfect for data validation, quality control, and automated Excel data processing.

## 📖 What Does This Script Do?

The script searches through Excel workbooks in a specified folder and looks for rows where:
1. A specific value exists in one column (the "search column")
2. AND another specific value exists in a different column (the "check column") in the same row

When both conditions are met, it logs the match with detailed information.

### Example Use Case

Imagine you have product spreadsheets where:
- Column M contains product codes (like "NXT0015", "NXT2277")
- Column BK contains colors (like "Black", "Red", "Blue")

You want to find all rows where product "NXT0015" has color "Black". This script will:
- Open each Excel file in your folder
- Search the specified sheet
- Find every row matching your criteria
- Log the results with filename, sheet name, and matched values

## ✨ Features

- **Multiple Search Rules** - Define unlimited search rules in a single configuration
- **Named Rules** - Give each rule a descriptive name for easy tracking in logs
- **Case-Insensitive Matching** - Automatically handles case differences (e.g., "BLACK" = "black")
- **Whitespace Handling** - Ignores leading/trailing spaces in values
- **Batch Processing** - Processes entire folders of Excel files automatically
- **Dual Logging** - Separate logs for results and errors
- **Robust Error Handling** - Continues processing even if individual files fail
- **Progress Reporting** - Real-time console feedback showing what's being processed
- **Enable/Disable Rules** - Turn rules on/off without deleting them
- **Cross-Platform** - Works on Windows, macOS, and Linux

## 📋 Requirements

### Software Requirements
- **Python 3.x** (Python 3.7 or higher recommended)
- **Microsoft Excel** must be installed on your system (Windows or macOS)
- **xlwings library** for Excel integration

### Installation

1. **Install Python** (if not already installed):
   - Windows: Download from [python.org](https://www.python.org/downloads/)
   - macOS: `brew install python3`
   - Linux: `sudo apt-get install python3`

2. **Install xlwings**:
   ```bash
   pip install xlwings
   ```

3. **Download this script** to a folder on your computer

## ⚙️ Configuration

### config.json File

All settings are configured in `config.json`. Here's what each section does:

```json
{
  "max_rows_to_process": 300,
  "folder_paths": {
    "windows": "C:\\Users\\YourName\\Documents\\ExcelFiles",
    "mac": "/Users/yourname/Documents/ExcelFiles",
    "linux": "/path/to/excel/files"
  },
  "search_rules": [
    {
      "name": "NXT0015 Black Check",
      "sheet_name": "Std Cricket Products Upload",
      "search_column": "M",
      "search_value": "NXT0015",
      "check_column": "BK",
      "check_value": "Black",
      "enabled": true
    }
  ]
}
```

### Configuration Parameters

#### max_rows_to_process
- **Type**: Number
- **Description**: Maximum number of rows to process per sheet (prevents slow processing on huge sheets)
- **Example**: `300` means it will check the first 300 rows of each sheet
- **Recommended**: `300-1000` for most use cases

#### folder_paths
- **Description**: The folder containing your Excel files
- **Platform-specific**: Set the path for your operating system
- **Windows format**: Use double backslashes `\\` or forward slashes `/`
- **Mac/Linux format**: Standard Unix paths

#### search_rules
Each rule defines one search criteria. You can have as many rules as needed.

**Rule Parameters**:
- `name` (optional): A descriptive name for tracking in logs. If omitted, auto-generates "Rule 1", "Rule 2", etc.
- `sheet_name` (required): The exact name of the Excel sheet to search
- `search_column` (required): Column letter to search (e.g., "A", "M", "AA", "BK")
- `search_value` (required): The value to find in the search column
- `check_column` (required): Column letter to check when search_value is found
- `check_value` (required): The value that must exist in check_column for a match
- `enabled` (optional): Set to `false` to temporarily disable a rule. Default is `true`.

### Column Naming

Use Excel column letters:
- Single letters: `A`, `B`, `C`, ... `Z`
- Double letters: `AA`, `AB`, `AC`, ... `AZ`, `BA`, `BB`, ... `ZZ`
- Triple letters: `AAA`, `AAB`, etc.

**Example**: Column BK is the 63rd column in Excel.

## 🚀 How to Use

### Step 1: Set Up Your Configuration

1. Open `config.json` in a text editor
2. Set the correct folder path for your operating system
3. Define your search rules
4. Save the file

### Step 2: Prepare Your Excel Files

1. Place all Excel files you want to search in the folder specified in `config.json`
2. Make sure the files are not open in Excel
3. Ensure the sheet names in your rules match the actual sheet names in your files

### Step 3: Run the Script

**Windows**:
```bash
python main.py
```

**Mac/Linux**:
```bash
python3 main.py
```

### Step 4: Review the Results

The script creates two timestamped log files:

1. **search_results_YYYYMMDD_HHMMSS.txt** - All matches found
2. **errors_YYYYMMDD_HHMMSS.txt** - Any errors encountered

## 📊 Understanding the Output

### Console Output

While running, you'll see:
```
Running on: Windows 10
Results log: search_results_20231215_143022.txt
Error log: errors_20231215_143022.txt

Looking for Excel files in: C:\Users\YourName\Documents\ExcelFiles
Found 5 Excel files

Active search rules: 2 total
  Sheet 'Std Cricket Products Upload': 2 rule(s)
    Column pair M->BK: 2 rule(s)
      - NXT0015 Black Check: 'NXT0015' -> 'Black'
      - NXT2277 Black Check: 'NXT2277' -> 'Black'

File 1/5: ProductFile1.xlsx
  Processing ProductFile1.xlsx
  Found 6 sheets
    Checking sheet: Std Cricket Products Upload
      Found 2 rule(s) in 1 column pair(s)
      Processing 300 rows with 1 column pair(s)
      Found 3 match(es)
    Checking sheet: Other Sheet
      Skipping (no rules for this sheet)

File 2/5: ProductFile2.xlsx
...

============================================================
SEARCH COMPLETE!
Total files processed: 5
Total matches found: 8

Results saved to: search_results_20231215_143022.txt
Errors logged to: errors_20231215_143022.txt
============================================================
```

### Results Log Format

Each match is logged with full details:
```
2023-12-15 14:30:22 - [NXT0015 Black Check] Match - File: ProductFile1.xlsx, Sheet: Std Cricket Products Upload, Value 1: NXT0015, Value 2: Black
2023-12-15 14:30:22 - [NXT2277 Black Check] Match - File: ProductFile1.xlsx, Sheet: Std Cricket Products Upload, Value 1: NXT2277, Value 2: Black
```

**Log Fields**:
- **Timestamp**: When the match was found
- **Rule Name**: Which rule found the match
- **File**: Excel filename containing the match
- **Sheet**: Sheet name where match was found
- **Value 1**: The actual value found in the search column
- **Value 2**: The actual value found in the check column

### Error Log Format

Any errors are logged with context:
```
2023-12-15 14:30:25 - Failed to open workbook CorruptFile.xlsx: [Errno 13] Permission denied
2023-12-15 14:30:26 - Failed to disable Excel calculations for LockedFile.xlsx: (-2147352567, 'Exception occurred.')
```

## 🔍 Matching Behavior

### Exact Match with Normalization

The script performs **exact matching** with some helpful normalization:

**✅ Will Match**:
- `"NXT0015"` matches `"nxt0015"` (case-insensitive)
- `"Black"` matches `" Black "` (whitespace trimmed)
- `"Black"` matches `"BLACK"` (case-insensitive)

**❌ Will NOT Match**:
- `"NXT0015"` vs `"NXT00152"` (extra characters)
- `"Black"` vs `"Black Color"` (extra words)
- `"NXT0015"` vs `"Contains NXT0015"` (partial/substring not matched)

### Empty Cell Handling

- Empty cells are skipped (no match)
- Cells with only whitespace are treated as empty

## 🛠️ Advanced Usage

### Multiple Rules for Same Sheet

You can define multiple rules that search the same sheet:

```json
"search_rules": [
  {
    "name": "Check 1",
    "sheet_name": "Products",
    "search_column": "A",
    "search_value": "Item1",
    "check_column": "B",
    "check_value": "Active",
    "enabled": true
  },
  {
    "name": "Check 2",
    "sheet_name": "Products",
    "search_column": "A",
    "search_value": "Item2",
    "check_column": "B",
    "check_value": "Active",
    "enabled": true
  }
]
```

The script automatically groups rules that use the same column pairs for efficient processing.

### Multiple Sheets

You can search different sheets in the same files:

```json
"search_rules": [
  {
    "sheet_name": "Sheet1",
    "search_column": "A",
    "search_value": "X",
    "check_column": "B",
    "check_value": "Y"
  },
  {
    "sheet_name": "Sheet2",
    "search_column": "C",
    "search_value": "Z",
    "check_column": "D",
    "check_value": "W"
  }
]
```

### Temporarily Disable Rules

Set `"enabled": false` to disable a rule without deleting it:

```json
{
  "name": "Disabled Check",
  "enabled": false,
  "sheet_name": "Products",
  ...
}
```

## ⚠️ Common Issues and Solutions

### Issue: "xlwings is not installed"
**Solution**: Run `pip install xlwings`

### Issue: "Configuration file not found"
**Solution**: Make sure `config.json` is in the same folder as `main.py`

### Issue: "Invalid folder path"
**Solution**: Check that the folder path in `config.json` exists and is correct for your OS

### Issue: "Failed to start Excel application"
**Solution**:
- Make sure Microsoft Excel is installed
- Close any open Excel instances
- On Mac, ensure Excel has accessibility permissions

### Issue: "Failed to close workbook"
**Solution**: This is usually harmless. The script logs it but continues processing. If it happens frequently:
- Close all Excel windows before running
- Check if files are being used by another program

### Issue: "No matches found" but you expect matches
**Solution**:
- Verify sheet names match exactly (case-sensitive)
- Check that column letters are correct
- Ensure `max_rows_to_process` is high enough
- Remember: matching is exact (see Matching Behavior section)

## 📝 Tips and Best Practices

1. **Test First**: Start with a small folder of 2-3 files to verify your rules work correctly
2. **Use Descriptive Names**: Name your rules clearly so logs are easy to understand
3. **Check Sheet Names**: Sheet names are case-sensitive and must match exactly
4. **Backup Important Files**: The script opens files as read-only, but backup important data
5. **Close Excel**: Close Excel before running to avoid conflicts
6. **Review Error Log**: Always check the error log after processing
7. **Adjust max_rows**: If your data is beyond row 300, increase `max_rows_to_process`

## 🔒 Safety Features

- **Read-Only Mode**: Files are opened as read-only; your data is never modified
- **Error Isolation**: If one file fails, others continue processing
- **Detailed Error Logging**: All errors are captured for review
- **Graceful Shutdown**: Properly closes Excel even if errors occur

## 📄 File Types Supported

- `.xlsx` (Excel 2007+)
- `.xlsm` (Excel with macros)
- `.xls` (Excel 97-2003)

Temporary files starting with `~$` are automatically skipped.

## 🤝 Support

If you encounter issues:
1. Check the error log file for detailed error messages
2. Verify your configuration matches the examples
3. Ensure all requirements are installed
4. Try with a single test file first

## 📜 License

This script is provided as-is for data processing tasks.
