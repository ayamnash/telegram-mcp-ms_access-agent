import os
import pyodbc
from fastmcp import FastMCP
from win32com.client import Dispatch
import win32com.client
import uuid
import random
import tempfile
import re
import gc
import pythoncom
import time
import logging
from typing import Callable, Tuple, Optional, List, Dict, Any

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

mcp = FastMCP("Flexible Access DB MCP")

# --- Configuration ---
class Config:
    """Configuration settings for the MCP server"""
    LOCK_TIMEOUT = 10  # seconds to wait for lock release
    CLEANUP_DELAY = 0.5  # seconds to wait after cleanup
    MAX_RETRIES = 3  # maximum retry attempts for transient errors
    RETRY_DELAY = 1.0  # seconds between retries
    POLL_INTERVAL = 0.5  # seconds between lock file checks

# --- State Tracking ---
_template_generated = False
_last_template_type = None
_batch_mode_db = None
_batch_mode_access = None

# --- Helper Functions ---

def _ensure_access_closed():
    """Force close all Access instances and clean up COM objects"""
    try:
        access = win32com.client.GetActiveObject("Access.Application")
        try:
            access.Quit(1)  # acQuitSaveAll
            logger.debug("Successfully closed active Access instance")
        except win32com.client.pywintypes.com_error as e:
            logger.warning(f"COM error while closing Access: {e}")
        except Exception as e:
            logger.warning(f"Unexpected error while closing Access: {e}")
        finally:
            del access
    except win32com.client.pywintypes.com_error:
        # No active Access instance - this is expected
        logger.debug("No active Access instance to close")
    except Exception as e:
        logger.warning(f"Unexpected error in _ensure_access_closed: {e}")
    
    # Force COM cleanup
    try:
        pythoncom.CoUninitialize()
    except Exception as e:
        logger.debug(f"CoUninitialize error (may be expected): {e}")
    
    try:
        pythoncom.CoInitialize()
    except Exception as e:
        logger.debug(f"CoInitialize error (may be expected): {e}")
    
    gc.collect()
    time.sleep(Config.CLEANUP_DELAY)

def _with_access_database(db_name: str, operation_func: Callable) -> Any:
    """Context manager pattern for Access operations with automatic cleanup
    
    Args:
        db_name: Database name or path
        operation_func: Function that takes access object and returns result
        
    Returns:
        Result from operation_func
        
    Raises:
        Exception: If operation fails after retries
    """
    path = get_db_path(db_name)
    access = None
    
    try:
        # Check if in batch mode
        global _batch_mode_access, _batch_mode_db
        if _batch_mode_access and _batch_mode_db == db_name:
            logger.debug(f"Using existing batch connection for {db_name}")
            return operation_func(_batch_mode_access)
        
        # Normal mode - open, execute, close
        logger.info(f"Opening database: {path}")
        access = win32com.client.Dispatch("Access.Application")
        access.Visible = False
        access.OpenCurrentDatabase(path)
        
        result = operation_func(access)
        
        # Save and close
        try:
            access.DoCmd.Save()
            logger.debug("Database saved successfully")
        except Exception as e:
            logger.debug(f"Save not needed or failed (may be expected): {e}")
        
        access.CloseCurrentDatabase()
        access.Quit(1)
        logger.info(f"Database closed successfully: {path}")
        
        return result
        
    except win32com.client.pywintypes.com_error as e:
        logger.error(f"COM error in database operation: {e}")
        raise Exception(f"COM error: {str(e)}")
    except Exception as e:
        logger.error(f"Error in database operation: {e}")
        raise
    finally:
        if access and not _batch_mode_access:
            try:
                access.Quit(1)
            except Exception as e:
                logger.debug(f"Error during final quit (may be expected): {e}")
            del access
            _ensure_access_closed()

def is_database_locked(db_path: str) -> bool:
    """Check if database has an active lock file
    
    Args:
        db_path: Full path to database file
        
    Returns:
        True if lock file exists, False otherwise
    """
    lock_file = db_path.replace('.accdb', '.laccdb')
    locked = os.path.exists(lock_file)
    if locked:
        logger.warning(f"Database is locked: {lock_file}")
    return locked

def wait_for_lock_release(db_path: str, timeout: Optional[int] = None) -> Tuple[bool, str]:
    """Wait for lock file to be released
    
    Args:
        db_path: Full path to database file
        timeout: Maximum seconds to wait (default: Config.LOCK_TIMEOUT)
        
    Returns:
        Tuple of (success: bool, message: str)
    """
    if timeout is None:
        timeout = Config.LOCK_TIMEOUT
        
    lock_file = db_path.replace('.accdb', '.laccdb')
    
    if not os.path.exists(lock_file):
        return True, "Database is not locked"
    
    logger.info(f"Waiting for lock release: {lock_file} (timeout: {timeout}s)")
    start_time = time.time()
    
    while os.path.exists(lock_file):
        elapsed = time.time() - start_time
        if elapsed > timeout:
            msg = f"Timeout: Database still locked after {timeout} seconds. Please close MS Access manually."
            logger.error(msg)
            return False, msg
        time.sleep(Config.POLL_INTERVAL)
    
    logger.info(f"Lock released after {time.time() - start_time:.1f} seconds")
    return True, "Lock released"

def _validate_module_name(module_name: str) -> Tuple[bool, str]:
    """Validate VBA module name
    
    Args:
        module_name: Name to validate
        
    Returns:
        Tuple of (is_valid: bool, error_message: str)
    """
    if not module_name or not module_name.strip():
        return False, "Module name cannot be empty"
    
    if not re.match(r'^[a-zA-Z_][a-zA-Z0-9_]*$', module_name):
        return False, "Module name must be a valid VBA identifier (letters, numbers, underscore; cannot start with number)"
    
    if len(module_name) > 64:
        return False, "Module name too long (max 64 characters)"
    
    # VBA reserved words
    reserved = ['Sub', 'Function', 'End', 'If', 'Then', 'Else', 'For', 'Next', 'Do', 'Loop', 
                'While', 'Select', 'Case', 'Dim', 'As', 'Integer', 'String', 'Boolean']
    if module_name in reserved:
        return False, f"Module name '{module_name}' is a VBA reserved word"
    
    return True, ""

def _validate_database_name(db_name: str) -> Tuple[bool, str]:
    """Validate database name/path
    
    Args:
        db_name: Database name or path to validate
        
    Returns:
        Tuple of (is_valid: bool, error_message: str)
    """
    if not db_name or not db_name.strip():
        return False, "Database name cannot be empty"
    
    # Check for path traversal attempts
    if '..' in db_name:
        return False, "Database name cannot contain '..' (path traversal)"
    
    return True, ""

# IMPROVED get_db_path function with better path detection
def get_db_path(db_name: str) -> str:
    """Gets the full path for the database. Handles both absolute and relative paths.
    Now includes better path detection and validation."""
    
    # If the path is already absolute (e.g., "F:\...") use it directly.
    if os.path.isabs(db_name):
        if not db_name.lower().endswith(".accdb"):
            db_name += ".accdb"
        return db_name
    
    # For relative paths, try multiple locations in order of preference:
    if not db_name.lower().endswith(".accdb"):
        db_name += ".accdb"
    
    # 1. Current working directory (most common for development)
    current_dir_path = os.path.join(os.getcwd(), db_name)
    if os.path.exists(current_dir_path):
        return current_dir_path
    
    # 2. User's home directory (original behavior)
    home_dir_path = os.path.join(os.path.expanduser("~"), db_name)
    if os.path.exists(home_dir_path):
        return home_dir_path
    
    # 3. If neither exists, default to current directory (for new database creation)
    return current_dir_path

def get_driver() -> str:
    """Finds a suitable Microsoft Access ODBC driver."""
    drivers = pyodbc.drivers()
    for d in [
        "Microsoft Access Driver (*.mdb, *.accdb)",
        "Microsoft Access Driver (*.accdb)",
        "Microsoft Access Driver (*.mdb)"
    ]:
        if d in drivers:
            return d
    raise Exception("Access ODBC driver not found")



def _run_query_internal(db_name: str, sql: str) -> str:
    """Internal helper to run any SQL query."""
    path = get_db_path(db_name)
    driver = get_driver()
    conn_str = f"DRIVER={{{driver}}};DBQ={path};"

    try:
        with pyodbc.connect(conn_str) as conn:
            cursor = conn.cursor()
            cursor.execute(sql)

            if sql.strip().lower().startswith("select"):
                columns = [col[0] for col in cursor.description]
                rows = cursor.fetchall()
                if rows:
                    result = f"Query Results ({len(rows)} rows):\n"
                    result += " | ".join(f"{col:<15}" for col in columns) + "\n"
                    result += "-" * (len(columns) * 17) + "\n"
                    for row in rows:
                        result += " | ".join(f"{str(val):<15}" for val in row) + "\n"
                    return result
                else:
                    return "No results found"
            else:
                conn.commit()
                return "Query executed successfully"
    except Exception as e:
        return f"Error: {str(e)}"

def _get_table_schema(db_name: str, table_name: str) -> list[str]:
    """Internal helper to get column names for a table or query."""
    path = get_db_path(db_name)
    driver = get_driver()
    conn_str = f"DRIVER={{{driver}}};DBQ={path};"
    try:
        with pyodbc.connect(conn_str) as conn:
            cursor = conn.cursor()
            # Try to get schema by running a SELECT query, which works for both tables and queries
            cursor.execute(f"SELECT * FROM [{table_name}] WHERE 1=0")
            columns = [col[0] for col in cursor.description]
            if not columns:
                raise ValueError(f"Table or query '{table_name}' not found or has no columns.")
            return columns
    except Exception as e:
        raise ValueError(f"Could not retrieve schema for table or query '{table_name}'. Error: {e}")
def sanitize_vba_code(code: str) -> str:
    """Clean VBA code by removing duplicate declarations that Access adds automatically
    
    Args:
        code: Raw VBA code string
        
    Returns:
        Cleaned VBA code
    """
    if not code:
        return code
    
    lines = code.split('\n')
    cleaned_lines = []
    
    # Track if we've seen these declarations (Access adds them automatically)
    seen_option_compare = False
    seen_option_explicit = False
    
    for line in lines:
        stripped = line.strip()
        
        # Skip duplicate "Option Compare Database" (Access adds this automatically)
        if stripped.lower() == "option compare database":
            if not seen_option_compare:
                seen_option_compare = True
                # Skip it - Access will add it automatically
                continue
            else:
                # Duplicate found, skip it
                logger.info("Removed duplicate 'Option Compare Database'")
                continue
        
        # Keep "Option Explicit" if present (it's useful)
        if stripped.lower() == "option explicit":
            if not seen_option_explicit:
                seen_option_explicit = True
                cleaned_lines.append(line)
            else:
                # Duplicate found, skip it
                logger.info("Removed duplicate 'Option Explicit'")
            continue
        
        # Keep all other lines
        cleaned_lines.append(line)
    
    return '\n'.join(cleaned_lines)

def sanitize_access_schema(schema: str) -> str:
    replacements = {
        r"\bAUTOINCREMENT\b": "COUNTER",
        r"\bINTEGER\b": "LONG",
        r"\bINT\b": "LONG",
        r"\bBIGINT\b": "LONG",
        r"\bBOOLEAN\b": "YESNO",
        r"\bBIT\b": "YESNO",
        r"\bLONGTEXT\b": "MEMO",
        r"\bTEXT\(MAX\)": "MEMO",
        r"\bDECIMAL\([^)]+\)": "CURRENCY",
        r"\bNUMERIC\([^)]+\)": "CURRENCY",
    }
    for pattern, repl in replacements.items():
        schema = re.sub(pattern, repl, schema, flags=re.IGNORECASE)
    
    # Remove DEFAULT clauses that Access doesn't handle well in CREATE TABLE
    schema = re.sub(r"DEFAULT\s+NOW\(\)", "", schema, flags=re.IGNORECASE)
    schema = re.sub(r"DEFAULT\s+CURRENT_TIMESTAMP", "", schema, flags=re.IGNORECASE)
    schema = re.sub(r"DEFAULT\s+TRUE", "", schema, flags=re.IGNORECASE)
    schema = re.sub(r"DEFAULT\s+-1", "", schema, flags=re.IGNORECASE)
    schema = re.sub(r"DEFAULT\s+0", "", schema, flags=re.IGNORECASE)
    schema = re.sub(r"DEFAULT\s+'[^']*'", "", schema, flags=re.IGNORECASE)
    
    # Wrap reserved words in brackets
    reserved_words = ["Status", "Notes", "Description", "Name", "Date", "User"]
    for word in reserved_words:
        schema = re.sub(rf"\b{word}\b(?!\])", f"[{word}]", schema, flags=re.IGNORECASE)
    
    # Clean up extra spaces and fix malformed parentheses
    schema = re.sub(r"\s{2,}", " ", schema)
    schema = re.sub(r",\s*\)", ")", schema)
    schema = re.sub(r"\(\s*,", "(", schema)
    
    return schema.strip()

def check_vba_compilation_errors(access_app) -> Tuple[bool, str]:
    """Check if there are VBA compilation errors in the current database
    
    Args:
        access_app: Active Access.Application COM object
        
    Returns:
        Tuple of (has_errors: bool, error_message: str)
    """
    try:
        # Try to access the VBA project
        project = access_app.VBE.VBProjects(1)
        
        # Try to compile the project
        # Note: This doesn't actually compile, but accessing modules can reveal errors
        for i in range(1, project.VBComponents.Count + 1):
            try:
                component = project.VBComponents(i)
                # Try to access the code module
                if component.CodeModule.CountOfLines > 0:
                    # Just accessing it can trigger compilation
                    _ = component.CodeModule.Lines(1, 1)
            except Exception as comp_ex:
                error_msg = str(comp_ex)
                if "compile" in error_msg.lower() or "syntax" in error_msg.lower():
                    logger.warning(f"VBA compilation error detected in {component.Name}: {error_msg}")
                    return True, f"VBA compilation error in {component.Name}: {error_msg}"
        
        return False, "No VBA compilation errors detected"
        
    except Exception as e:
        # If we can't check, assume no errors (or VBA is protected)
        logger.info(f"Could not check VBA compilation (may be protected): {e}")
        return False, "VBA check skipped (protected or no VBA)"

@mcp.tool
def insert_data(db_name: str, table: str, rows: list[dict]) -> str:
    """Insert rows into a table. Example: [{'ID': 1, 'Name': 'Ali'}]"""
    path = get_db_path(db_name)
    driver = get_driver()
    conn_str = f"DRIVER={{{driver}}};DBQ={path};"
    with pyodbc.connect(conn_str) as conn:
        cursor = conn.cursor()
        for row in rows:
            columns = ', '.join(f"[{c}]" for c in row.keys())
            placeholders = ', '.join('?' for _ in row)
            values = list(row.values())
            sql = f"INSERT INTO {table} ({columns}) VALUES ({placeholders})"
            cursor.execute(sql, values)
        conn.commit()
        return f"Inserted {len(rows)} rows into '{table}'"

@mcp.tool
def run_query(db_name: str, sql: str) -> str:
    """Run a SELECT or action query (INSERT, UPDATE, DELETE)."""
    return _run_query_internal(db_name, sql)


@mcp.tool
def list_tables(db_name: str) -> str:
    """List all tables in the database"""
    path = get_db_path(db_name)
    driver = get_driver()
    conn_str = f"DRIVER={{{driver}}};DBQ={path};"
    try:
        with pyodbc.connect(conn_str) as conn:
            cursor = conn.cursor()
            tables = cursor.tables(tableType='TABLE')
            table_names = [row.table_name for row in tables if not row.table_name.startswith('MSys')]
            if table_names:
                return "Tables:\n" + "\n".join(f"- {name}" for name in table_names)
            else:
                return "No tables found"
    except Exception as e:
        return f"Error: {str(e)}"
def fix_access_sql_syntax(sql: str) -> str:
    """
    Automatically fix common Access SQL syntax issues:
    1. Convert double quotes to single quotes for string literals
    2. Keep double quotes only for special cases like Format functions
    3. Fix multiple JOIN syntax by adding proper parentheses
    """
    # Pattern to match string literals that should use single quotes
    # This matches double quotes that are NOT part of function calls like Format("yyyy-mm-dd")
    
    # First, protect Format function quotes and similar cases
    protected_patterns = []
    
    # Find and temporarily replace Format function quotes
    format_pattern = r'(Format\s*\([^,]+,\s*)"([^"]+)"'
    def protect_format(match):
        placeholder = f"__PROTECTED_QUOTE_{len(protected_patterns)}__"
        protected_patterns.append(f'"{match.group(2)}"')
        return f'{match.group(1)}{placeholder}'
    
    sql = re.sub(format_pattern, protect_format, sql, flags=re.IGNORECASE)
    
    # Now convert remaining double quotes to single quotes for string literals
    # This pattern matches double quotes around values (not in function contexts)
    sql = re.sub(r'=\s*"([^"]*)"', r"= '\1'", sql)  # = "value" -> = 'value'
    sql = re.sub(r'<>\s*"([^"]*)"', r"<> '\1'", sql)  # <> "value" -> <> 'value'
    sql = re.sub(r'IN\s*\(\s*"([^"]*)"', r"IN ('\1'", sql, flags=re.IGNORECASE)  # IN ("value" -> IN ('value'
    sql = re.sub(r'LIKE\s*"([^"]*)"', r"LIKE '\1'", sql, flags=re.IGNORECASE)  # LIKE "value" -> LIKE 'value'
    
    # Fix multiple JOIN syntax for Access
    # Access requires parentheses around multiple JOINs
    # Pattern: FROM table1 INNER JOIN table2 ON ... INNER JOIN table3 ON ...
    # Should become: FROM (table1 INNER JOIN table2 ON ...) INNER JOIN table3 ON ...
    
    # Find FROM clause with multiple INNER JOINs
    from_pattern = r'FROM\s+([^()]+?)\s+INNER\s+JOIN\s+([^()]+?)\s+ON\s+([^()]+?)\s+INNER\s+JOIN'
    if re.search(from_pattern, sql, re.IGNORECASE):
        # Replace the pattern to add parentheses around the first JOIN
        sql = re.sub(
            from_pattern,
            r'FROM (\1 INNER JOIN \2 ON \3) INNER JOIN',
            sql,
            flags=re.IGNORECASE
        )
    
    # Handle LEFT JOIN cases too
    from_pattern_left = r'FROM\s+([^()]+?)\s+LEFT\s+JOIN\s+([^()]+?)\s+ON\s+([^()]+?)\s+(?:INNER|LEFT)\s+JOIN'
    if re.search(from_pattern_left, sql, re.IGNORECASE):
        sql = re.sub(
            from_pattern_left,
            r'FROM (\1 LEFT JOIN \2 ON \3) INNER JOIN' if 'INNER JOIN' in sql.upper() else r'FROM (\1 LEFT JOIN \2 ON \3) LEFT JOIN',
            sql,
            flags=re.IGNORECASE
        )
    
    # Restore protected quotes
    for i, protected in enumerate(protected_patterns):
        sql = sql.replace(f"__PROTECTED_QUOTE_{i}__", protected)
    
    return sql


def _generate_report_template_internal(db_name: str, record_source: str, report_type: str = "tabular") -> str:
    """Internal helper function to generate report template without MCP tool wrapper."""
    try:
        # Validate record source and get fields
        fields = _get_table_schema(db_name, record_source)
        
        report_guid = str(uuid.uuid4()).replace('-', '')
        
        # Generate controls based on report type
        if report_type.lower() == "columnar":
            # Columnar layout - fields stacked vertically
            controls_text = ""
            namemap_entries = []
            y_pos = 500
            
            for i, field in enumerate(fields):
                controls_text += f"""
                Begin Label
                    OverlapFlags =85
                    Left =500
                    Top ={y_pos}
                    Width =2000
                    Height =315
                    Name ="{field}_Label"
                    Caption ="{field}:"
                    GUID = Begin
                        0x{uuid.uuid4().hex}
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =2700
                    Top ={y_pos}
                    Width =4000
                    Height =315
                    Name ="{field}"
                    ControlSource ="{field}"
                    GUID = Begin
                        0x{uuid.uuid4().hex}
                    End
                End"""
                
                # Add to NameMap
                rand_hex = ''.join(random.choices('0123456789abcdef', k=32))
                field_hex = f"{field}_Label".encode('utf-16le').hex()
                namemap_entries.append(f"0x{rand_hex}{len(f'{field}_Label'):02x}000000{field_hex}")
                
                rand_hex = ''.join(random.choices('0123456789abcdef', k=32))
                field_hex = field.encode('utf-16le').hex()
                namemap_entries.append(f"0x{rand_hex}{len(field):02x}000000{field_hex}")
                
                y_pos += 400
                
        else:  # tabular layout (default)
            # Header controls
            header_controls = ""
            detail_controls = ""
            namemap_entries = []
            x_pos = 500
            
            for i, field in enumerate(fields):
                # Header label
                header_controls += f"""
                Begin Label
                    OverlapFlags =85
                    Left ={x_pos}
                    Top =200
                    Width =1500
                    Height =315
                    Name ="{field}_Header"
                    Caption ="{field}"
                    GUID = Begin
                        0x{uuid.uuid4().hex}
                    End
                End"""
                
                # Detail textbox
                detail_controls += f"""
                Begin TextBox
                    OverlapFlags =85
                    Left ={x_pos}
                    Top =200
                    Width =1500
                    Height =315
                    Name ="{field}"
                    ControlSource ="{field}"
                    GUID = Begin
                        0x{uuid.uuid4().hex}
                    End
                End"""
                
                # Add to NameMap
                rand_hex = ''.join(random.choices('0123456789abcdef', k=32))
                field_hex = f"{field}_Header".encode('utf-16le').hex()
                namemap_entries.append(f"0x{rand_hex}{len(f'{field}_Header'):02x}000000{field_hex}")
                
                rand_hex = ''.join(random.choices('0123456789abcdef', k=32))
                field_hex = field.encode('utf-16le').hex()
                namemap_entries.append(f"0x{rand_hex}{len(field):02x}000000{field_hex}")
                
                x_pos += 1600
            
            controls_text = f"""
        Begin Section
            Height =600
            Name ="ReportHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =500
                    Top =200
                    Width =6000
                    Height =400
                    Name ="Title"
                    Caption ="__REPORT_NAME_PLACEHOLDER__"
                    FontSize =14
                    FontWeight =700
                    GUID = Begin
                        0x{uuid.uuid4().hex}
                    End
                End
            End
        End
        Begin Section
            Height =600
            Name ="PageHeader"
            Begin
                {header_controls}
            End
        End
        Begin Section
            Height =400
            Name ="Detail"
            Begin
                {detail_controls}
            End
        End"""
        
        # Add Title to NameMap
        rand_hex = ''.join(random.choices('0123456789abcdef', k=32))
        field_hex = "Title".encode('utf-16le').hex()
        namemap_entries.append(f"0x{rand_hex}05000000{field_hex}")
        
        # NameMap
        namemap_text = ",\n        ".join(namemap_entries) + ",\n        0x000000000000000000000000000000000c000000050000000000000000000000000000000000"
        
        if report_type.lower() == "columnar":
            template = f"""Version =21
VersionRequired =20
PublishOption =1
Checksum ={random.randint(-2000000000, 2000000000)}
Begin Report
    Width =7400
    PictureAlignment =2
    GUID = Begin
        0x{report_guid}
    End
    NameMap = Begin
        {namemap_text}
    End
    RecordSource ="{record_source}"
    Caption ="__REPORT_NAME_PLACEHOLDER__"
    Begin
        Begin Section
            Height ={y_pos + 200}
            Name ="Detail"
            Begin
                {controls_text}
            End
        End
    End
End"""
        else:  # tabular
            template = f"""Version =21
VersionRequired =20
PublishOption =1
Checksum ={random.randint(-2000000000, 2000000000)}
Begin Report
    Width =7400
    PictureAlignment =2
    GUID = Begin
        0x{report_guid}
    End
    NameMap = Begin
        {namemap_text}
    End
    RecordSource ="{record_source}"
    Caption ="__REPORT_NAME_PLACEHOLDER__"
    Begin
        {controls_text}
    End
End"""
        
        return template
        
    except Exception as e:
        raise Exception(f"Error generating report template: {e}")

def _create_report_from_template_internal(db_name: str, report_name: str, report_text: str) -> str:
    """Internal helper function to create report from template using new pattern."""
    # Replace placeholder if it exists
    if "__REPORT_NAME_PLACEHOLDER__" in report_text:
        report_text = report_text.replace("__REPORT_NAME_PLACEHOLDER__", report_name)
    
    def operation(access):
        """Inner function to create the report"""
        logger.info(f"Creating report: {report_name}")
        
        # Write to temp file
        temp_file_path = None
        try:
            with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix=".txt", encoding='utf-8') as tf:
                tf.write(report_text)
                temp_file_path = tf.name

            AC_REPORT = 3
            
            # Delete existing report if it exists
            try:
                access.DoCmd.DeleteObject(AC_REPORT, report_name)
                logger.debug(f"Deleted existing report: {report_name}")
            except Exception:
                logger.debug(f"Report {report_name} doesn't exist (creating new)")

            # Load report from text file
            access.LoadFromText(AC_REPORT, report_name, temp_file_path)
            logger.info(f"Report '{report_name}' created successfully")

            return f"Report '{report_name}' created successfully in database '{db_name}'."
            
        finally:
            if temp_file_path and os.path.exists(temp_file_path):
                os.remove(temp_file_path)
                logger.debug("Temp file cleaned up")
    
    try:
        path = get_db_path(db_name)
        
        # Check for lock
        if is_database_locked(path):
            success, message = wait_for_lock_release(path)
            if not success:
                raise Exception(message)
        
        result = _with_access_database(db_name, operation)
        return result
        
    except Exception as e:
        logger.error(f"Error creating report '{report_name}': {e}")
        raise Exception(f"Error creating report from template: {e}")




            
if __name__ == "__main__":
    mcp.run()
