"""
Excel Column Search Script
Searches for a target value in one column, then checks another column for a second target value
Logs matches with sheet name and cell values
"""

import os
import sys
import platform
from datetime import datetime
import logging
import json

try:
    import xlwings as xw
except ImportError:
    print("ERROR: xlwings is not installed!")
    print("Please install it with: pip install xlwings")
    print("You also need Microsoft Excel installed on your system")
    sys.exit(1)

# Global configuration
CONFIG = {}

def load_configuration(config_path=None):
    """Load configuration from JSON file"""
    global CONFIG

    if config_path is None:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        config_path = os.path.join(script_dir, "config.json")

    if not os.path.exists(config_path):
        print(f"ERROR: Configuration file not found: {config_path}")
        sys.exit(1)

    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            CONFIG = json.load(f)
        print(f"Configuration loaded from {config_path}")
        return True
    except json.JSONDecodeError as e:
        print(f"ERROR: Invalid JSON in configuration file: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"ERROR: Failed to load configuration: {e}")
        sys.exit(1)

def setup_logging():
    """Set up logging to track matches"""
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    log_filename = f"search_results_{timestamp}.txt"

    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(message)s',
        handlers=[
            logging.FileHandler(log_filename, encoding='utf-8'),
            logging.StreamHandler()
        ]
    )

    return log_filename

def column_letter_to_index(col_letter):
    """Convert column letter(s) to zero-based index"""
    col_letter = col_letter.upper()
    result = 0
    for char in col_letter:
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result - 1

def search_sheet(sheet, search_col, search_value, check_col, check_value, max_rows):
    """
    Search a sheet for matching values in two columns

    Returns list of matches with format:
    [(sheet_name, first_column_value, second_column_value), ...]
    """
    matches = []
    sheet_name = sheet.name

    try:
        # Get used range
        used_range = sheet.used_range
        if used_range is None:
            return matches

        rows, cols = used_range.shape
        rows_to_process = min(rows, max_rows)

        # Convert column letters to indices
        search_col_idx = column_letter_to_index(search_col)
        check_col_idx = column_letter_to_index(check_col)

        # Search through rows
        for row_idx in range(rows_to_process):
            try:
                # Get first column value
                first_cell = used_range[row_idx, search_col_idx]
                first_value = first_cell.value

                if not first_value:
                    continue

                # Normalize and compare
                first_value_str = str(first_value).strip()
                if first_value_str.lower() == search_value.lower():
                    # Match found in first column, check second column
                    second_cell = used_range[row_idx, check_col_idx]
                    second_value = second_cell.value

                    if second_value:
                        second_value_str = str(second_value).strip()
                        if second_value_str.lower() == check_value.lower():
                            # Both columns match - log it
                            matches.append((sheet_name, first_value_str, second_value_str))

            except Exception:
                continue

        return matches

    except Exception as e:
        print(f"      Error searching sheet: {e}")
        return matches

def process_excel_file(filepath, search_rules):
    """
    Process Excel file with search rules

    search_rules format:
    {
        "sheet_name": {
            "search_column": "A",
            "search_value": "Product123",
            "check_column": "B",
            "check_value": "Active"
        }
    }
    """
    app = None
    wb = None
    all_matches = []

    try:
        print(f"  Processing {os.path.basename(filepath)}")

        # Start Excel
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False

        # Open workbook
        wb = app.books.open(filepath, update_links=False, notify=False, read_only=True)

        print(f"  Found {len(wb.sheets)} sheets")

        max_rows = CONFIG.get('max_rows_to_process', 1000)

        for sheet in wb.sheets:
            sheet_name = sheet.name
            print(f"    Checking sheet: {sheet_name}")

            # Check if we have rules for this sheet
            if sheet_name in search_rules:
                rule = search_rules[sheet_name]
                matches = search_sheet(
                    sheet,
                    rule['search_column'],
                    rule['search_value'],
                    rule['check_column'],
                    rule['check_value'],
                    max_rows
                )

                if matches:
                    print(f"      Found {len(matches)} match(es)")
                    for match in matches:
                        sheet_n, first_val, second_val = match
                        logging.info(f"Match - Sheet: {sheet_n}, Column {rule['search_column']}: {first_val}, Column {rule['check_column']}: {second_val}")
                    all_matches.extend(matches)
                else:
                    print(f"      No matches found")
            else:
                print(f"      Skipping (no rules for this sheet)")

        return all_matches

    except Exception as e:
        print(f"  ERROR: {str(e)}")
        logging.error(f"Failed to process {os.path.basename(filepath)}: {e}")
        return all_matches

    finally:
        if wb:
            wb.close()
        if app:
            app.quit()

def main():
    """Main execution function"""

    # Load configuration
    if not load_configuration():
        return

    print(f"Running on: {platform.system()} {platform.release()}")

    # Setup logging
    log_file = setup_logging()
    logging.info(f"Starting Excel column search operation")
    print(f"Results log: {log_file}")

    # Get directory
    folder_paths = CONFIG.get('folder_paths', {})
    system = platform.system()

    if system == "Windows":
        directory = folder_paths.get('windows', '')
    elif system == "Darwin":
        directory = folder_paths.get('mac', '')
    else:
        directory = folder_paths.get('linux', '')

    if not directory or not os.path.exists(directory):
        print(f"ERROR: Invalid folder path for {system}")
        return

    print(f"\nLooking for Excel files in: {directory}")

    # Get Excel files
    try:
        excel_files = [f for f in os.listdir(directory)
                      if f.endswith(('.xlsx', '.xlsm', '.xls')) and not f.startswith('~$')]
        print(f"Found {len(excel_files)} Excel files")
    except Exception as e:
        print(f"Error reading directory: {e}")
        return

    if not excel_files:
        print("No Excel files found")
        return

    # Build search rules
    search_rules_config = CONFIG.get('search_rules', [])
    search_rules = {}

    for rule in search_rules_config:
        if not rule.get('enabled', True):
            continue

        sheet_name = rule.get('sheet_name')
        if sheet_name:
            search_rules[sheet_name] = {
                'search_column': rule.get('search_column'),
                'search_value': rule.get('search_value'),
                'check_column': rule.get('check_column'),
                'check_value': rule.get('check_value')
            }

    if not search_rules:
        print("ERROR: No enabled search rules found in configuration")
        return

    print(f"\nActive search rules:")
    for sheet_name, rule in search_rules.items():
        print(f"  Sheet '{sheet_name}': Search col {rule['search_column']} for '{rule['search_value']}', check col {rule['check_column']} for '{rule['check_value']}'")

    # Process files
    total_matches = 0
    os.chdir(directory)

    for i, filename in enumerate(excel_files, 1):
        filepath = os.path.join(directory, filename)
        print(f"\nFile {i}/{len(excel_files)}: {filename}")

        matches = process_excel_file(filepath, search_rules)
        total_matches += len(matches)

    # Summary
    print("\n" + "="*60)
    print("SEARCH COMPLETE!")
    print(f"Total files processed: {len(excel_files)}")
    print(f"Total matches found: {total_matches}")
    print(f"\nResults saved to: {log_file}")
    print("="*60)

if __name__ == "__main__":
    main()
