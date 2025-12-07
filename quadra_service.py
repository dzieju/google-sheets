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
    dbf_values: List[Any],
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
        dbf_values: List of values from DBF to search for
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
            'notes': str
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
    for dbf_value in dbf_values:
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
                'notes': f"Found in {match_info['sheetName']} at {col_index_to_a1(match_info['columnIndex'])}{match_info['rowIndex'] + 1}"
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
                'notes': 'Missing'
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
        List of strings for table row: [DBF_value, Status, SheetName, Column, Row, Notes]
    """
    status = "Found" if result['found'] else "Missing"
    sheet_name = result.get('sheetName', '') or ''
    column_name = result.get('columnName', '') or ''
    row_index = result.get('rowIndex')
    row_display = str(row_index + 1) if row_index is not None else ''
    notes = result.get('notes', '') or ''
    
    return [
        str(result['dbfValue']),
        status,
        sheet_name,
        column_name,
        row_display,
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
            'status': 'Found' if result['found'] else 'Missing',
            'sheetName': result.get('sheetName', ''),
            'columnName': result.get('columnName', ''),
            'columnIndex': result.get('columnIndex'),
            'rowIndex': result.get('rowIndex'),
            'matchedValue': str(result.get('matchedValue', '')) if result.get('matchedValue') is not None else '',
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
    writer.writerow(['DBF_Value', 'Status', 'SheetName', 'ColumnName', 'ColumnIndex', 'RowIndex', 'MatchedValue', 'Notes'])
    
    # Write data
    for result in results:
        writer.writerow([
            str(result['dbfValue']),
            'Found' if result['found'] else 'Missing',
            result.get('sheetName', ''),
            result.get('columnName', ''),
            result.get('columnIndex', ''),
            result.get('rowIndex', ''),
            str(result.get('matchedValue', '')) if result.get('matchedValue') is not None else '',
            result.get('notes', '')
        ])
    
    return output.getvalue()
