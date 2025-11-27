"""
Excel Column Search Script - OPTIMIZED VERSION
Searches for target values in columns and logs matches

OPTIMIZATIONS:
- Batch column reading: Reads entire columns at once instead of cell-by-cell (10-50x faster)
- Multiple rules per sheet: Supports unlimited rules for the same sheet
- Grouped processing: Rules with same column pairs processed together
- Excel performance: Disables calculations and events during processing
- Smart lookup: Uses dictionaries for O(1) value matching

FEATURES:
- Multiple search rules per sheet
- Named rules for better tracking
- Progress reporting
- Case-insensitive matching
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

def search_sheet_optimized(sheet, rules_by_column_pair, max_rows):
    """
    OPTIMIZED: Search a sheet for matching values using batch column reads

    Processes multiple rules grouped by column pairs in a single pass

    rules_by_column_pair format:
    {
        ('A', 'B'): [
            {'rule_name': 'Rule 1', 'search_value': 'val1', 'check_value': 'val2'},
            {'rule_name': 'Rule 2', 'search_value': 'val3', 'check_value': 'val4'}
        ]
    }

    Returns list of matches with format:
    [(rule_name, sheet_name, first_column_value, second_column_value), ...]
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

        print(f"      Processing {rows_to_process} rows with {len(rules_by_column_pair)} column pair(s)")

        # Process each column pair group
        for (search_col, check_col), rule_group in rules_by_column_pair.items():
            try:
                # Convert column letters to indices
                search_col_idx = column_letter_to_index(search_col)
                check_col_idx = column_letter_to_index(check_col)

                # OPTIMIZATION: Read entire columns at once
                search_range = sheet.range((1, search_col_idx + 1), (rows_to_process, search_col_idx + 1))
                check_range = sheet.range((1, check_col_idx + 1), (rows_to_process, check_col_idx + 1))

                # Get all values as lists (much faster than cell-by-cell)
                search_values = search_range.value
                check_values = check_range.value

                # Ensure lists (single cell returns single value, not list)
                if not isinstance(search_values, list):
                    search_values = [search_values]
                if not isinstance(check_values, list):
                    check_values = [check_values]

                # Build lookup dictionary for all rules in this column pair
                # Format: {normalized_search_value: [(rule_name, check_value), ...]}
                search_lookup = {}
                for rule in rule_group:
                    search_val_normalized = str(rule['search_value']).strip().lower()
                    if search_val_normalized not in search_lookup:
                        search_lookup[search_val_normalized] = []
                    search_lookup[search_val_normalized].append({
                        'rule_name': rule['rule_name'],
                        'check_value': str(rule['check_value']).strip().lower()
                    })

                # SINGLE PASS through rows for this column pair
                for row_idx in range(len(search_values)):
                    search_val = search_values[row_idx]
                    check_val = check_values[row_idx]

                    if not search_val or not check_val:
                        continue

                    # Normalize values
                    search_val_str = str(search_val).strip()
                    search_val_normalized = search_val_str.lower()
                    check_val_str = str(check_val).strip()
                    check_val_normalized = check_val_str.lower()

                    # Check if search value matches any rule
                    if search_val_normalized in search_lookup:
                        # Check all rules for this search value
                        for rule_info in search_lookup[search_val_normalized]:
                            if check_val_normalized == rule_info['check_value']:
                                # Match found!
                                matches.append((
                                    rule_info['rule_name'],
                                    sheet_name,
                                    search_val_str,
                                    check_val_str
                                ))

            except Exception as e:
                print(f"      Warning: Error processing column pair {search_col}->{check_col}: {e}")
                continue

        return matches

    except Exception as e:
        print(f"      Error searching sheet: {e}")
        return matches

def process_excel_file(filepath, sheet_rules):
    """
    OPTIMIZED: Process Excel file with search rules

    sheet_rules format:
    {
        "sheet_name": {
            ('A', 'B'): [
                {'rule_name': 'Rule 1', 'search_value': 'val1', 'check_value': 'val2'},
                {'rule_name': 'Rule 2', 'search_value': 'val3', 'check_value': 'val4'}
            ]
        }
    }
    """
    app = None
    wb = None
    all_matches = []

    try:
        print(f"  Processing {os.path.basename(filepath)}")

        # Start Excel with performance optimizations
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False

        # Disable calculation and events for speed
        calc_prev = None
        events_prev = None
        try:
            calc_prev = app.api.Calculation
            app.api.Calculation = -4135  # xlCalculationManual
            events_prev = app.api.EnableEvents
            app.api.EnableEvents = False
        except Exception:
            pass

        # Open workbook
        wb = app.books.open(filepath, update_links=False, notify=False, read_only=True)

        print(f"  Found {len(wb.sheets)} sheets")

        max_rows = CONFIG.get('max_rows_to_process', 1000)

        for sheet in wb.sheets:
            sheet_name = sheet.name
            print(f"    Checking sheet: {sheet_name}")

            # Check if we have rules for this sheet
            if sheet_name in sheet_rules:
                rules_by_column_pair = sheet_rules[sheet_name]
                total_rules = sum(len(rules) for rules in rules_by_column_pair.values())
                print(f"      Found {total_rules} rule(s) in {len(rules_by_column_pair)} column pair(s)")

                matches = search_sheet_optimized(sheet, rules_by_column_pair, max_rows)

                if matches:
                    print(f"      Found {len(matches)} match(es)")
                    for match in matches:
                        rule_name, sheet_n, first_val, second_val = match
                        logging.info(f"[{rule_name}] Match - Sheet: {sheet_n}, Value 1: {first_val}, Value 2: {second_val}")
                    all_matches.extend(matches)
                else:
                    print(f"      No matches found")
            else:
                print(f"      Skipping (no rules for this sheet)")

        # Restore Excel settings
        try:
            if calc_prev is not None:
                app.api.Calculation = calc_prev
            if events_prev is not None:
                app.api.EnableEvents = events_prev
        except Exception:
            pass

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

    # Build search rules - NEW: Support multiple rules per sheet
    # Data structure: { sheet_name: { (search_col, check_col): [rules] } }
    search_rules_config = CONFIG.get('search_rules', [])
    sheet_rules = {}

    rule_count = 0
    for idx, rule in enumerate(search_rules_config):
        if not rule.get('enabled', True):
            continue

        sheet_name = rule.get('sheet_name')
        if not sheet_name:
            continue

        search_col = rule.get('search_column')
        check_col = rule.get('check_column')
        search_val = rule.get('search_value')
        check_val = rule.get('check_value')

        # Auto-generate rule name if not provided
        rule_name = rule.get('name', f"Rule {idx + 1}")

        # Initialize sheet if not exists
        if sheet_name not in sheet_rules:
            sheet_rules[sheet_name] = {}

        # Group by column pair
        col_pair = (search_col, check_col)
        if col_pair not in sheet_rules[sheet_name]:
            sheet_rules[sheet_name][col_pair] = []

        # Add rule to group
        sheet_rules[sheet_name][col_pair].append({
            'rule_name': rule_name,
            'search_value': search_val,
            'check_value': check_val
        })
        rule_count += 1

    if not sheet_rules:
        print("ERROR: No enabled search rules found in configuration")
        return

    print(f"\nActive search rules: {rule_count} total")
    for sheet_name, rules_by_col_pair in sheet_rules.items():
        total_rules_for_sheet = sum(len(rules) for rules in rules_by_col_pair.values())
        print(f"  Sheet '{sheet_name}': {total_rules_for_sheet} rule(s)")
        for (search_col, check_col), rules in rules_by_col_pair.items():
            print(f"    Column pair {search_col}->{check_col}: {len(rules)} rule(s)")
            for rule in rules:
                print(f"      - {rule['rule_name']}: '{rule['search_value']}' -> '{rule['check_value']}'")

    # Process files
    total_matches = 0
    os.chdir(directory)

    for i, filename in enumerate(excel_files, 1):
        filepath = os.path.join(directory, filename)
        print(f"\nFile {i}/{len(excel_files)}: {filename}")

        matches = process_excel_file(filepath, sheet_rules)
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
