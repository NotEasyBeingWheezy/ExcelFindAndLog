"""
Excel Single Column Search Script - OPTIMIZED VERSION
Searches for target values in a column and logs specified columns from matching rows

OPTIMIZATIONS:
- Batch column reading: Reads entire columns at once instead of cell-by-cell (10-50x faster)
- Multiple rules per sheet: Supports unlimited rules for the same sheet
- Excel performance: Disables calculations and events during processing
- Smart lookup: Uses dictionaries for O(1) value matching
- Reused Excel instance: One Excel app for all files (20-40% faster)

FEATURES:
- Search for specific values in a column
- Log multiple columns (A, B, C, etc.) when value is found
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
        config_path = os.path.join(script_dir, "config_single.json")

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
    """Set up logging to track matches and errors"""
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    log_filename = f"search_results_{timestamp}.txt"
    error_log_filename = f"errors_{timestamp}.txt"

    # Create main logger for results
    logger = logging.getLogger('results')
    logger.setLevel(logging.INFO)

    # Create error logger
    error_logger = logging.getLogger('errors')
    error_logger.setLevel(logging.ERROR)

    # Format for logs
    formatter = logging.Formatter('%(asctime)s - %(message)s')

    # Results log handlers
    results_file_handler = logging.FileHandler(log_filename, encoding='utf-8')
    results_file_handler.setFormatter(formatter)
    results_console_handler = logging.StreamHandler()
    results_console_handler.setFormatter(formatter)

    logger.addHandler(results_file_handler)
    logger.addHandler(results_console_handler)

    # Error log handlers
    error_file_handler = logging.FileHandler(error_log_filename, encoding='utf-8')
    error_file_handler.setFormatter(formatter)
    error_console_handler = logging.StreamHandler()
    error_console_handler.setFormatter(formatter)

    error_logger.addHandler(error_file_handler)
    error_logger.addHandler(error_console_handler)

    return log_filename, error_log_filename

def column_letter_to_index(col_letter):
    """Convert column letter(s) to zero-based index"""
    col_letter = col_letter.upper()
    result = 0
    for char in col_letter:
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result - 1

def search_sheet_optimized(sheet, rules_by_search_column, max_rows):
    """
    OPTIMIZED: Search a sheet for matching values and log specified columns

    Processes multiple rules grouped by search column in a single pass

    rules_by_search_column format:
    {
        'M': [
            {'rule_name': 'Rule 1', 'search_value': '1KBP0059', 'log_columns': ['A', 'B', 'C']},
            {'rule_name': 'Rule 2', 'search_value': 'NXT2277', 'log_columns': ['A', 'B', 'C']}
        ]
    }

    Returns list of matches with format:
    [(rule_name, sheet_name, search_value, {col: value, ...}), ...]
    """
    matches = []
    sheet_name = sheet.name
    error_logger = logging.getLogger('errors')  # Cache logger reference

    try:
        # Get used range
        used_range = sheet.used_range
        if used_range is None:
            return matches

        rows, cols = used_range.shape
        rows_to_process = min(rows, max_rows)

        print(f"      Processing {rows_to_process} rows with {len(rules_by_search_column)} search column(s)")

        # Process each search column group
        for search_col, rule_group in rules_by_search_column.items():
            try:
                # Convert column letter to index
                search_col_idx = column_letter_to_index(search_col)

                # Collect all unique log columns needed for this search column
                all_log_columns = set()
                for rule in rule_group:
                    all_log_columns.update(rule['log_columns'])

                # OPTIMIZATION: Read search column at once
                search_range = sheet.range((1, search_col_idx + 1), (rows_to_process, search_col_idx + 1))
                search_values = search_range.value

                # Ensure list (single cell returns single value, not list)
                if not isinstance(search_values, list):
                    search_values = [search_values]

                # Read all log columns at once
                log_column_data = {}
                for log_col in all_log_columns:
                    log_col_idx = column_letter_to_index(log_col)
                    log_range = sheet.range((1, log_col_idx + 1), (rows_to_process, log_col_idx + 1))
                    log_values = log_range.value
                    if not isinstance(log_values, list):
                        log_values = [log_values]
                    log_column_data[log_col] = log_values

                # Build lookup dictionary for all rules in this search column
                # Format: {normalized_search_value: [rule_info, ...]}
                search_lookup = {}
                for rule in rule_group:
                    search_val_normalized = str(rule['search_value']).strip().lower()
                    if search_val_normalized not in search_lookup:
                        search_lookup[search_val_normalized] = []
                    search_lookup[search_val_normalized].append({
                        'rule_name': rule['rule_name'],
                        'log_columns': rule['log_columns']
                    })

                # SINGLE PASS through rows for this search column
                for row_idx in range(len(search_values)):
                    search_val = search_values[row_idx]

                    if not search_val:
                        continue

                    # Normalize search value
                    search_val_str = str(search_val).strip()
                    search_val_normalized = search_val_str.lower()

                    # Check if search value matches any rule
                    if search_val_normalized in search_lookup:
                        # Check all rules for this search value
                        for rule_info in search_lookup[search_val_normalized]:
                            # Get values from log columns
                            logged_values = {}
                            for log_col in rule_info['log_columns']:
                                col_val = log_column_data[log_col][row_idx]
                                logged_values[log_col] = str(col_val).strip() if col_val else ""

                            # Match found!
                            matches.append((
                                rule_info['rule_name'],
                                sheet_name,
                                search_val_str,
                                logged_values
                            ))

            except Exception as e:
                error_msg = f"Error processing search column {search_col} in sheet {sheet_name}: {e}"
                print(f"      Warning: {error_msg}")
                error_logger.error(error_msg)
                continue

        return matches

    except Exception as e:
        error_msg = f"Error searching sheet {sheet_name}: {e}"
        print(f"      {error_msg}")
        error_logger.error(error_msg)
        return matches

def process_excel_file(filepath, sheet_rules, app=None):
    """
    OPTIMIZED: Process Excel file with search rules

    Args:
        filepath: Path to the Excel file
        sheet_rules: Dictionary of sheet rules
        app: Optional xlwings App instance (for reusing across files)

    sheet_rules format:
    {
        "sheet_name": {
            'M': [
                {'rule_name': 'Rule 1', 'search_value': 'val1', 'log_columns': ['A', 'B', 'C']},
                {'rule_name': 'Rule 2', 'search_value': 'val2', 'log_columns': ['A', 'B', 'C']}
            ]
        }
    }
    """
    wb = None
    all_matches = []
    app_created = False  # Track if we created the app instance

    # Cache logger references
    error_logger = logging.getLogger('errors')
    results_logger = logging.getLogger('results')

    try:
        print(f"  Processing {os.path.basename(filepath)}")

        # Start Excel with performance optimizations (if not provided)
        if app is None:
            try:
                app = xw.App(visible=False, add_book=False)
                app.display_alerts = False
                app.screen_updating = False
                app_created = True
            except Exception as e:
                error_msg = f"Failed to start Excel application for {os.path.basename(filepath)}: {e}"
                print(f"  ERROR: {error_msg}")
                error_logger.error(error_msg)
                return all_matches

        # Disable calculation and events for speed
        calc_prev = None
        events_prev = None
        try:
            calc_prev = app.api.Calculation
            app.api.Calculation = -4135  # xlCalculationManual
            events_prev = app.api.EnableEvents
            app.api.EnableEvents = False
        except Exception as e:
            error_logger.error(f"Failed to disable Excel calculations for {os.path.basename(filepath)}: {e}")

        # Open workbook
        try:
            wb = app.books.open(filepath, update_links=False, notify=False, read_only=True)
        except Exception as e:
            error_msg = f"Failed to open workbook {os.path.basename(filepath)}: {e}"
            print(f"  ERROR: {error_msg}")
            error_logger.error(error_msg)
            if app and app_created:
                try:
                    app.quit()
                except:
                    pass
            return all_matches

        print(f"  Found {len(wb.sheets)} sheets")

        max_rows = CONFIG.get('max_rows_to_process', 300)

        for sheet in wb.sheets:
            sheet_name = sheet.name

            # Skip sheets without rules early (performance optimization)
            if sheet_name not in sheet_rules:
                continue

            print(f"    Checking sheet: {sheet_name}")
            rules_by_search_column = sheet_rules[sheet_name]
            total_rules = sum(len(rules) for rules in rules_by_search_column.values())
            print(f"      Found {total_rules} rule(s) in {len(rules_by_search_column)} search column(s)")

            matches = search_sheet_optimized(sheet, rules_by_search_column, max_rows)

            if matches:
                print(f"      Found {len(matches)} match(es)")
                filename = os.path.basename(filepath)
                for match in matches:
                    rule_name, sheet_n, search_val, logged_values = match
                    # Format logged columns for display
                    columns_str = ", ".join([f"{col}: {val}" for col, val in logged_values.items()])
                    results_logger.info(f"[{rule_name}] Match - File: {filename}, Sheet: {sheet_n}, Search Value: {search_val}, {columns_str}")
                all_matches.extend(matches)
            else:
                print(f"      No matches found")

        # Restore Excel settings
        try:
            if calc_prev is not None:
                app.api.Calculation = calc_prev
            if events_prev is not None:
                app.api.EnableEvents = events_prev
        except Exception as e:
            error_logger.error(f"Failed to restore Excel settings for {os.path.basename(filepath)}: {e}")

        return all_matches

    except Exception as e:
        error_msg = f"Failed to process {os.path.basename(filepath)}: {e}"
        print(f"  ERROR: {error_msg}")
        error_logger.error(error_msg)
        return all_matches

    finally:
        # Close workbook with error handling
        if wb:
            try:
                wb.close()
            except Exception as e:
                error_msg = f"Failed to close workbook {os.path.basename(filepath)}: {e}"
                print(f"  Warning: {error_msg}")
                error_logger.error(error_msg)
                # Try to force close if regular close fails
                try:
                    wb.api.Close(False)  # False = don't save changes
                except:
                    pass

        # Quit Excel application only if we created it (not if it was passed in)
        if app and app_created:
            try:
                app.quit()
            except Exception as e:
                error_msg = f"Failed to quit Excel application after processing {os.path.basename(filepath)}: {e}"
                print(f"  Warning: {error_msg}")
                error_logger.error(error_msg)
                # Try to force kill the Excel instance
                try:
                    app.api.Quit()
                except:
                    pass

def main():
    """Main execution function"""

    # Load configuration
    if not load_configuration():
        return

    print(f"Running on: {platform.system()} {platform.release()}")

    # Setup logging
    log_file, error_log_file = setup_logging()
    results_logger = logging.getLogger('results')
    results_logger.info(f"Starting Excel column search operation")
    print(f"Results log: {log_file}")
    print(f"Error log: {error_log_file}")

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

    # Build search rules - Support multiple rules per sheet
    # Data structure: { sheet_name: { search_column: [rules] } }
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
        search_val = rule.get('search_value')
        log_columns = rule.get('log_columns', [])

        # Auto-generate rule name if not provided
        rule_name = rule.get('name', f"Rule {idx + 1}")

        # Initialize sheet if not exists
        if sheet_name not in sheet_rules:
            sheet_rules[sheet_name] = {}

        # Group by search column
        if search_col not in sheet_rules[sheet_name]:
            sheet_rules[sheet_name][search_col] = []

        # Add rule to group
        sheet_rules[sheet_name][search_col].append({
            'rule_name': rule_name,
            'search_value': search_val,
            'log_columns': log_columns
        })
        rule_count += 1

    if not sheet_rules:
        print("ERROR: No enabled search rules found in configuration")
        return

    print(f"\nActive search rules: {rule_count} total")
    for sheet_name, rules_by_col in sheet_rules.items():
        total_rules_for_sheet = sum(len(rules) for rules in rules_by_col.values())
        print(f"  Sheet '{sheet_name}': {total_rules_for_sheet} rule(s)")
        for search_col, rules in rules_by_col.items():
            print(f"    Search column {search_col}: {len(rules)} rule(s)")
            for rule in rules:
                log_cols_str = ", ".join(rule['log_columns'])
                print(f"      - {rule['rule_name']}: '{rule['search_value']}' → log [{log_cols_str}]")

    # Process files with reused Excel instance for performance
    total_matches = 0
    os.chdir(directory)

    # Create Excel instance once and reuse for all files (major performance boost)
    app = None
    error_logger = logging.getLogger('errors')

    try:
        print("\nStarting Excel application...")
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False

        # Disable calculation and events for speed
        try:
            app.api.Calculation = -4135  # xlCalculationManual
            app.api.EnableEvents = False
        except Exception as e:
            error_logger.error(f"Failed to disable Excel calculations: {e}")

        print("Excel ready. Processing files...\n")

        for i, filename in enumerate(excel_files, 1):
            filepath = os.path.join(directory, filename)
            print(f"\nFile {i}/{len(excel_files)}: {filename}")

            matches = process_excel_file(filepath, sheet_rules, app)
            total_matches += len(matches)

    except Exception as e:
        error_msg = f"Failed to initialize Excel application: {e}"
        print(f"ERROR: {error_msg}")
        error_logger.error(error_msg)
        print("Falling back to processing files without shared Excel instance...")

        # Fallback: process files without shared instance
        for i, filename in enumerate(excel_files, 1):
            filepath = os.path.join(directory, filename)
            print(f"\nFile {i}/{len(excel_files)}: {filename}")
            matches = process_excel_file(filepath, sheet_rules)
            total_matches += len(matches)

    finally:
        # Close shared Excel instance
        if app:
            print("\nClosing Excel application...")
            try:
                app.quit()
            except Exception as e:
                error_logger.error(f"Failed to quit Excel application: {e}")
                try:
                    app.api.Quit()
                except:
                    pass

    # Summary
    print("\n" + "="*60)
    print("SEARCH COMPLETE!")
    print(f"Total files processed: {len(excel_files)}")
    print(f"Total matches found: {total_matches}")
    print(f"\nResults saved to: {log_file}")
    print(f"Errors logged to: {error_log_file}")
    print("="*60)

if __name__ == "__main__":
    main()
