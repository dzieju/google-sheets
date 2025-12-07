"""
quadra_service.py
Service module for Quadra feature: checking DBF order numbers against Google Sheets.

Functions:
- read_dbf_column(dbf_path, column_identifier) - Read values from a DBF column
- search_dbf_values_in_sheets(drive_service, sheets_service, dbf_values, spreadsheet_id, 
                                search_mode, sheet_names, column_names) - Search DBF values in Google Sheets
- format_quadra_results(results) - Format results for display and export
"""

import logging
import re
from typing import List, Dict, Any, Optional, Union, Tuple
from dbfread import DBF
from sheets_search import (
    normalize_number_string,
    normalize_header_name,
    find_all_column_indices_by_name,
    col_index_to_a1,
)

logger = logging.getLogger(__name__)

# DBF field name mappings for automatic detection (case-insensitive)
DBF_STAWKA_FIELD_NAMES = ['STAWKA', 'STAW', 'RATE', 'PRICE']
DBF_CZESCI_FIELD_NAMES = ['CZESCI', 'PARTS']


def column_letter_to_index(column: str) -> int:
    """
    Convert column letter (A, B, C, ..., AA, AB, ...) to 0-based index.
    
    Args:
        column: Column letter (case-insensitive)
    
    Returns:
        0-based column index
    
    Examples:
        'A' -> 0, 'B' -> 1, 'Z' -> 25, 'AA' -> 26
    """
    column = column.upper().strip()
    result = 0
    for char in column:
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result - 1


def parse_column_identifier(column_id: Union[str, int]) -> int:
    """
    Parse column identifier which can be:
    - A letter (A, B, C, ..., AA, AB, ...) - most common for user input
    - A 1-based number as string ('1', '2', '3', ...) - treat as 1-based user input
    - An integer - treat as 1-based if it's from user input context
    
    Args:
        column_id: Column identifier
    
    Returns:
        0-based column index
    """
    if isinstance(column_id, int):
        # If it's an integer, assume 1-based (user input) if > 0
        return column_id - 1 if column_id > 0 else 0
    
    # For strings, first try to parse as letter (most common case)
    column_str = str(column_id).strip()
    
    # Check if it's purely alphabetic (column letter like 'A', 'B', 'AA')
    if column_str.isalpha():
        return column_letter_to_index(column_str)
    
    # Try to parse as integer (1-based number like '1', '2', '3')
    try:
        num = int(column_str)
        return num - 1 if num > 0 else 0
    except (ValueError, TypeError):
        pass
    
    # If all else fails, try letter conversion
    try:
        return column_letter_to_index(column_str)
    except Exception:
        # Default to column 0
        return 0


def detect_dbf_field_name(field_names: List[str], possible_names: List[str]) -> Optional[str]:
    """
    Detect field name from a list of possible alternatives (case-insensitive).
    
    Args:
        field_names: List of field names in DBF table
        possible_names: List of possible field names to match
    
    Returns:
        First matching field name (original case) or None if not found
    
    Examples:
        >>> detect_dbf_field_name(['NUMER', 'STAWKA', 'DATA'], ['STAWKA', 'STAW', 'RATE'])
        'STAWKA'
        >>> detect_dbf_field_name(['NUMER', 'RATE', 'DATA'], ['STAWKA', 'STAW', 'RATE'])
        'RATE'
    """
    if not field_names or not possible_names:
        return None
    
    # Normalize field names for comparison (uppercase, trim)
    normalized_fields = {name.upper().strip(): name for name in field_names}
    
    # Check each possible name
    for possible in possible_names:
        normalized_possible = possible.upper().strip()
        if normalized_possible in normalized_fields:
            return normalized_fields[normalized_possible]
    
    return None


def map_dbf_record_to_result(record: Dict[str, Any], field_names: List[str]) -> Dict[str, Any]:
    """
    Map a DBF record to a result dictionary with detected fields.
    
    Detects and extracts:
    - 'stawka' from field names: STAWKA, STAW, RATE, PRICE
    - 'czesci' from field names: CZESCI, PARTS
    
    Args:
        record: DBF record as dictionary
        field_names: List of all field names in the DBF table
    
    Returns:
        Dictionary with 'stawka' and 'czesci' keys (empty strings if not found)
    
    Examples:
        >>> record = {'NUMER': '12345', 'STAWKA': '150.00', 'CZESCI': 'ABC'}
        >>> field_names = ['NUMER', 'STAWKA', 'CZESCI']
        >>> map_dbf_record_to_result(record, field_names)
        {'stawka': '150.00', 'czesci': 'ABC'}
    """
    # Detect field names using module-level constants
    stawka_field = detect_dbf_field_name(field_names, DBF_STAWKA_FIELD_NAMES)
    czesci_field = detect_dbf_field_name(field_names, DBF_CZESCI_FIELD_NAMES)
    
    # Extract values
    stawka_value = ''
    if stawka_field and stawka_field in record:
        val = record[stawka_field]
        if val is not None:
            stawka_value = str(val).strip()
    
    czesci_value = ''
    if czesci_field and czesci_field in record:
        val = record[czesci_field]
        if val is not None:
            czesci_value = str(val).strip()
    
    return {
        'stawka': stawka_value,
        'czesci': czesci_value
    }


def read_dbf_column(dbf_path: str, column_identifier: Union[str, int] = 'B') -> List[Any]:
    """
    Read values from a specific column in a DBF file.
    
    Args:
        dbf_path: Path to the DBF file
        column_identifier: Column to read - can be:
            - Letter: 'A', 'B', 'C', etc.
            - 1-based index: 1, 2, 3, etc.
            - Default: 'B' (second column)
    
    Returns:
        List of values from the specified column (excluding None/empty values)
    
    Raises:
        FileNotFoundError: If DBF file doesn't exist
        ValueError: If column identifier is invalid or column doesn't exist
    """
    from dbfread.exceptions import DBFNotFound
    
    try:
        table = DBF(dbf_path, encoding='cp1250')  # Polish encoding, adjust if needed
    except DBFNotFound as e:
        raise FileNotFoundError(f"DBF file not found: {dbf_path}")
    except Exception as e:
        raise ValueError(f"Error opening DBF file: {e}")
    
    # Get column index
    col_index = parse_column_identifier(column_identifier)
    
    # Get field names
    field_names = table.field_names
    
    if col_index < 0 or col_index >= len(field_names):
        raise ValueError(
            f"Column index {col_index} (from '{column_identifier}') is out of range. "
            f"DBF has {len(field_names)} columns: {', '.join(field_names)}"
        )
    
    field_name = field_names[col_index]
    logger.info(f"Reading column '{field_name}' (index {col_index}) from DBF file")
    
    # Extract values
    values = []
    for record in table:
        value = record.get(field_name)
        # Skip None and empty values
        if value is not None and str(value).strip():
            values.append(value)
    
    logger.info(f"Read {len(values)} non-empty values from DBF column '{field_name}'")
    return values


def read_dbf_records_with_extra_fields(dbf_path: str, column_identifier: Union[str, int] = 'B') -> List[Dict[str, Any]]:
    """
    Read DBF records with main column value and additional fields (stawka, czesci).
    
    Args:
        dbf_path: Path to the DBF file
        column_identifier: Column to read main values from (default 'B')
    
    Returns:
        List of dictionaries with 'value', 'stawka', and 'czesci' keys
        
    Example:
        >>> records = read_dbf_records_with_extra_fields('orders.dbf', 'B')
        >>> records[0]
        {'value': '12345', 'stawka': '150.00', 'czesci': 'ABC'}
    
    Raises:
        FileNotFoundError: If DBF file doesn't exist
        ValueError: If column identifier is invalid or column doesn't exist
    """
    from dbfread.exceptions import DBFNotFound
    
    try:
        table = DBF(dbf_path, encoding='cp1250')  # Polish encoding
    except DBFNotFound as e:
        raise FileNotFoundError(f"DBF file not found: {dbf_path}")
    except Exception as e:
        raise ValueError(f"Error opening DBF file: {e}")
    
    # Get column index for main value
    col_index = parse_column_identifier(column_identifier)
    
    # Get field names
    field_names = table.field_names
    
    if col_index < 0 or col_index >= len(field_names):
        raise ValueError(
            f"Column index {col_index} (from '{column_identifier}') is out of range. "
            f"DBF has {len(field_names)} columns: {', '.join(field_names)}"
        )
    
    main_field_name = field_names[col_index]
    logger.info(f"Reading records from DBF, main column: '{main_field_name}' (index {col_index})")
    
    # Read all records
    records = []
    for record in table:
        # Get main value
        main_value = record.get(main_field_name)
        if main_value is None or not str(main_value).strip():
            continue  # Skip empty records
        
        # Map additional fields
        extra_fields = map_dbf_record_to_result(record, field_names)
        
        records.append({
            'value': main_value,
            'stawka': extra_fields['stawka'],
            'czesci': extra_fields['czesci']
        })
    
    logger.info(f"Read {len(records)} records from DBF with extra fields")
    return records


def normalize_value_for_comparison(value: Any, mode: str = 'exact') -> str:
    """
    Normalize a value for comparison.
    
    Args:
        value: Value to normalize
        mode: Comparison mode - 'exact' or 'substring'
    
    Returns:
        Normalized string value
    """
    if value is None:
        return ""
    
    # Convert to string
    value_str = str(value)
    
    # Try to normalize as number
    normalized_num = normalize_number_string(value)
    if normalized_num:
        # If it looks like a number, use the normalized version
        value_str = normalized_num
    else:
        # Otherwise just trim and lowercase
        value_str = value_str.strip().lower()
    
    return value_str


def values_match(dbf_value: Any, sheet_value: Any, mode: str = 'exact') -> bool:
    """
    Check if two values match according to the specified mode.
    
    Args:
        dbf_value: Value from DBF file
        sheet_value: Value from Google Sheet
        mode: Comparison mode - 'exact' (trim, case-insensitive, or numeric) or 'substring'
    
    Returns:
        True if values match, False otherwise
    """
    if dbf_value is None and sheet_value is None:
        return True
    if dbf_value is None or sheet_value is None:
        return False
    
    # Normalize both values
    dbf_normalized = normalize_value_for_comparison(dbf_value, mode)
    sheet_normalized = normalize_value_for_comparison(sheet_value, mode)
    
    if not dbf_normalized or not sheet_normalized:
        return False
    
    if mode == 'substring':
        # Check if either value contains the other
        return dbf_normalized in sheet_normalized or sheet_normalized in dbf_normalized
    else:
        # Exact match (after normalization)
        return dbf_normalized == sheet_normalized


def search_value_in_sheet_data(
    target_value: Any,
    sheet_values: List[List[Any]],
    sheet_name: str,
    mode: str = 'exact',
    column_names: Optional[List[str]] = None,
    header_row_index: int = 0
) -> Optional[Dict[str, Any]]:
    """
    Search for a value in sheet data (2D array).
    
    Args:
        target_value: Value to search for
        sheet_values: 2D array of sheet values
        sheet_name: Name of the sheet
        mode: Comparison mode - 'exact' or 'substring'
        column_names: Optional list of column names to restrict search
        header_row_index: Index of header row (default 0)
    
    Returns:
        Dictionary with match details if found, None otherwise
        {
            'sheetName': str,
            'columnIndex': int,
            'columnName': str,
            'rowIndex': int,
            'value': Any
        }
    """
    if not sheet_values or len(sheet_values) == 0:
        return None
    
    # Get headers (first row by default)
    headers = sheet_values[header_row_index] if len(sheet_values) > header_row_index else []
    
    # Determine which columns to search
    columns_to_search = []
    if column_names:
        # Search only in specified columns
        for col_name in column_names:
            indices = find_all_column_indices_by_name(headers, col_name)
            columns_to_search.extend(indices)
    else:
        # Search all columns
        columns_to_search = list(range(len(headers)))
    
    # Search in data rows (after header)
    for row_idx in range(header_row_index + 1, len(sheet_values)):
        row = sheet_values[row_idx]
        for col_idx in columns_to_search:
            if col_idx < len(row):
                cell_value = row[col_idx]
                if values_match(target_value, cell_value, mode):
                    # Found a match
                    col_name = headers[col_idx] if col_idx < len(headers) else f"Column {col_index_to_a1(col_idx)}"
                    return {
                        'sheetName': sheet_name,
                        'columnIndex': col_idx,
                        'columnName': col_name,
                        'rowIndex': row_idx,
                        'value': cell_value
                    }
    
    return None


def search_dbf_values_in_sheets(
    drive_service,
    sheets_service,
    dbf_values: Union[List[Any], List[Dict[str, Any]]],
    spreadsheet_id: str,
    mode: str = 'exact',
    sheet_names: Optional[List[str]] = None,
    column_names: Optional[List[str]] = None,
    header_row_index: int = 0
) -> List[Dict[str, Any]]:
    """
    Search DBF values in Google Sheets and return results.
    
    Args:
        drive_service: Google Drive service instance
        sheets_service: Google Sheets service instance
        dbf_values: List of values or dicts with {'value', 'stawka', 'czesci'} from DBF to search for
        spreadsheet_id: ID of the spreadsheet to search in
        mode: Comparison mode - 'exact' or 'substring'
        sheet_names: Optional list of sheet names to search (None = all sheets)
        column_names: Optional list of column names to search (None = all columns)
        header_row_index: Index of header row (default 0)
    
    Returns:
        List of results, one per DBF value:
        {
            'dbfValue': Any,
            'found': bool,
            'sheetName': str or None,
            'columnIndex': int or None,
            'columnName': str or None,
            'rowIndex': int or None,
            'matchedValue': Any or None,
            'notes': str,
            'stawka': str (from DBF record),
            'czesci': str (from DBF record)
        }
    """
    # Get spreadsheet metadata
    try:
        metadata = sheets_service.spreadsheets().get(
            spreadsheetId=spreadsheet_id,
            fields='sheets.properties'
        ).execute()
    except Exception as e:
        logger.error(f"Error fetching spreadsheet metadata: {e}")
        raise
    
    # Get list of sheets to search
    all_sheets = metadata.get('sheets', [])
    if sheet_names:
        # Filter to only specified sheets
        sheets_to_search = [
            sh for sh in all_sheets 
            if sh['properties']['title'] in sheet_names
        ]
    else:
        # Search all sheets
        sheets_to_search = all_sheets
    
    logger.info(f"Searching {len(dbf_values)} DBF values in {len(sheets_to_search)} sheets")
    
    # Load all sheet data upfront
    sheet_data = {}
    for sheet in sheets_to_search:
        sheet_name = sheet['properties']['title']
        try:
            result = sheets_service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=sheet_name,
                majorDimension='ROWS'
            ).execute()
            sheet_data[sheet_name] = result.get('values', [])
        except Exception as e:
            logger.warning(f"Error loading sheet '{sheet_name}': {e}")
            sheet_data[sheet_name] = []
    
    # Search each DBF value
    results = []
    for dbf_item in dbf_values:
        # Handle both simple values and record dicts
        if isinstance(dbf_item, dict):
            dbf_value = dbf_item.get('value')
            stawka = dbf_item.get('stawka', '')
            czesci = dbf_item.get('czesci', '')
        else:
            dbf_value = dbf_item
            stawka = ''
            czesci = ''
        
        found = False
        match_info = None
        
        # Search in each sheet until found
        for sheet_name in sheet_data:
            match_info = search_value_in_sheet_data(
                target_value=dbf_value,
                sheet_values=sheet_data[sheet_name],
                sheet_name=sheet_name,
                mode=mode,
                column_names=column_names,
                header_row_index=header_row_index
            )
            
            if match_info:
                found = True
                break
        
        # Build result
        if found and match_info:
            result = {
                'dbfValue': dbf_value,
                'found': True,
                'sheetName': match_info['sheetName'],
                'columnIndex': match_info['columnIndex'],
                'columnName': match_info['columnName'],
                'rowIndex': match_info['rowIndex'],
                'matchedValue': match_info['value'],
                'notes': f"Found in {match_info['sheetName']} at {col_index_to_a1(match_info['columnIndex'])}{match_info['rowIndex'] + 1}",
                'stawka': stawka,
                'czesci': czesci
            }
        else:
            result = {
                'dbfValue': dbf_value,
                'found': False,
                'sheetName': None,
                'columnIndex': None,
                'columnName': None,
                'rowIndex': None,
                'matchedValue': None,
                'notes': 'Missing',
                'stawka': stawka,
                'czesci': czesci
            }
        
        results.append(result)
    
    logger.info(f"Search completed: {sum(1 for r in results if r['found'])} found, {sum(1 for r in results if not r['found'])} missing")
    return results


def format_quadra_result_for_table(result: Dict[str, Any]) -> List[str]:
    """
    Format a Quadra result for display in GUI table.
    
    Args:
        result: Result dictionary from search_dbf_values_in_sheets
    
    Returns:
        List of strings for table row: [Numer z DBF, Stawka, Status, Arkusz, Kolumna, Wiersz, Czesci_extra, Uwagi]
    """
    status = "Found" if result['found'] else "Missing"
    sheet_name = result.get('sheetName', '') or ''
    column_name = result.get('columnName', '') or ''
    row_index = result.get('rowIndex')
    row_display = str(row_index + 1) if row_index is not None else ''
    notes = result.get('notes', '') or ''
    stawka = result.get('stawka', '') or ''
    czesci = result.get('czesci', '') or ''
    
    return [
        str(result['dbfValue']),
        stawka,
        status,
        sheet_name,
        column_name,
        row_display,
        czesci,
        notes
    ]


def export_quadra_results_to_json(results: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Format Quadra results for JSON export.
    
    Args:
        results: List of result dictionaries from search_dbf_values_in_sheets
    
    Returns:
        List of dictionaries ready for JSON serialization
    """
    export_list = []
    for result in results:
        export_obj = {
            'dbfValue': str(result['dbfValue']),
            'stawka': result.get('stawka', ''),
            'status': 'Found' if result['found'] else 'Missing',
            'sheetName': result.get('sheetName', ''),
            'columnName': result.get('columnName', ''),
            'columnIndex': result.get('columnIndex'),
            'rowIndex': result.get('rowIndex'),
            'matchedValue': str(result.get('matchedValue', '')) if result.get('matchedValue') is not None else '',
            'czesci': result.get('czesci', ''),
            'notes': result.get('notes', '')
        }
        export_list.append(export_obj)
    
    return export_list


def export_quadra_results_to_csv(results: List[Dict[str, Any]]) -> str:
    """
    Format Quadra results for CSV export.
    
    Args:
        results: List of result dictionaries from search_dbf_values_in_sheets
    
    Returns:
        CSV string with header and data rows
    """
    import csv
    import io
    
    output = io.StringIO()
    writer = csv.writer(output)
    
    # Write header
    writer.writerow(['DBF_Value', 'Stawka', 'Status', 'SheetName', 'ColumnName', 'ColumnIndex', 'RowIndex', 'MatchedValue', 'Czesci', 'Notes'])
    
    # Write data
    for result in results:
        writer.writerow([
            str(result['dbfValue']),
            result.get('stawka', ''),
            'Found' if result['found'] else 'Missing',
            result.get('sheetName', ''),
            result.get('columnName', ''),
            result.get('columnIndex', ''),
            result.get('rowIndex', ''),
            str(result.get('matchedValue', '')) if result.get('matchedValue') is not None else '',
            result.get('czesci', ''),
            result.get('notes', '')
        ])
    
    return output.getvalue()


def write_quadra_results_to_sheet(
    sheets_service,
    spreadsheet_id: str,
    sheet_name: str,
    results: List[Dict[str, Any]],
    start_row: int = 1
) -> None:
    """
    Write Quadra results to a Google Sheet with columns I (Stawka) and J (Czesci_extra).
    
    Creates columns I and J if they don't exist, sets headers, and writes data.
    
    Args:
        sheets_service: Google Sheets service instance
        spreadsheet_id: ID of the spreadsheet
        sheet_name: Name of the sheet tab
        results: List of result dictionaries from search_dbf_values_in_sheets
        start_row: Row number to start writing data (1-based, default=1 for header row)
    
    Example:
        >>> write_quadra_results_to_sheet(
        ...     sheets_service, 'spreadsheet_id', 'Sheet1', results, start_row=1
        ... )
    
    Note:
        - Column I (index 8): Stawka
        - Column J (index 9): Czesci_extra
        - Preserves existing data in other columns
    """
    if not results:
        logger.warning("No results to write to sheet")
        return
    
    # Get sheet ID
    try:
        metadata = sheets_service.spreadsheets().get(
            spreadsheetId=spreadsheet_id,
            fields='sheets(properties(sheetId,title))'
        ).execute()
        
        sheet_id = None
        for sheet in metadata.get('sheets', []):
            if sheet['properties']['title'] == sheet_name:
                sheet_id = sheet['properties']['sheetId']
                break
        
        if sheet_id is None:
            raise ValueError(f"Sheet '{sheet_name}' not found in spreadsheet")
            
    except Exception as e:
        logger.error(f"Error getting sheet metadata: {e}")
        raise
    
    # Prepare data for columns I and J
    # Column I (index 8): Stawka
    # Column J (index 9): Czesci_extra
    
    # Build values array: [[header_I, header_J], [row1_I, row1_J], [row2_I, row2_J], ...]
    values = [
        ['Stawka', 'Czesci_extra']  # Header row
    ]
    
    for result in results:
        stawka = result.get('stawka', '') or ''
        czesci = result.get('czesci', '') or ''
        values.append([stawka, czesci])
    
    # Write to columns I and J starting at start_row
    # A1 notation: I{start_row}:J{start_row + len(results)}
    range_notation = f"{sheet_name}!I{start_row}:J{start_row + len(results)}"
    
    try:
        body = {
            'values': values
        }
        
        result = sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=range_notation,
            valueInputOption='RAW',
            body=body
        ).execute()
        
        logger.info(
            f"Wrote {result.get('updatedCells', 0)} cells to {range_notation} "
            f"in sheet '{sheet_name}'"
        )
        
    except Exception as e:
        logger.error(f"Error writing results to sheet: {e}")
        raise
