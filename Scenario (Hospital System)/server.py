import os
import json
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

from operations import execute_registered_operation

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

mcp = FastMCP("Flexible Access DB MCP")


@mcp.tool()
def execute_operation(operation: str, payload: dict, domain: str) -> str:
    """Execute a registered business operation using a closed operation catalog.

    This is the recommended production entrypoint for AI-driven database work.
    The AI should choose only:
    - domain
    - operation
    - payload

    The AI must not generate SQL or perform multi-step database writes itself.
    """
    result = execute_registered_operation(operation=operation, payload=payload, domain=domain)
    return json.dumps(result, ensure_ascii=False)

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
@mcp.tool()
def save_and_close_access_database(db_name: str, force_close: bool = False) -> dict:
    """
    Save all changes and close the MS Access database.
    If Access is not running, returns a safe success message.
    
    Args:
        db_name: Database name or path
        force_close: If True, force close even if there are VBA compilation errors
    
    Returns:
        dict with success status and message
    """
    try:
        access_app = win32com.client.GetActiveObject("Access.Application")
        current_db = access_app.CurrentDb()

        if current_db is None:
            return {"success": False, "message": "No database is currently open in Access."}

        current_path = current_db.Name

        if db_name.lower() not in current_path.lower():
            return {
                "success": False,
                "message": f"The open database '{current_path}' does not match '{db_name}'."
            }

        # Try to save first
        save_attempted = False
        save_error = None
        try:
            access_app.DoCmd.Save()
            save_attempted = True
            logger.info("Database saved successfully")
        except Exception as save_ex:
            save_error = str(save_ex)
            logger.warning(f"Could not save database (may have VBA errors): {save_error}")
            
            # If force_close is True, we'll continue to close anyway
            if not force_close:
                return {
                    "success": False,
                    "message": f"Cannot save database (VBA compilation errors?): {save_error}. Use force_close=True to close without saving.",
                    "vba_error": True
                }

        # Try to close gracefully with save
        close_method = None
        try:
            if force_close or save_error:
                # Force close without saving if there were save errors
                logger.info("Attempting force close (acQuitSaveNone)")
                access_app.Quit(2)  # acQuitSaveNone = 2 (don't save)
                close_method = "force_close_no_save"
            else:
                # Normal close with save
                logger.info("Attempting normal close (acQuitSaveAll)")
                access_app.Quit(1)  # acQuitSaveAll = 1
                close_method = "normal_close_with_save"
        except Exception as quit_ex:
            logger.warning(f"Quit command failed: {quit_ex}, trying alternative method")
            try:
                # Alternative: Close current database then quit
                access_app.CloseCurrentDatabase()
                access_app.Quit()
                close_method = "alternative_close"
            except Exception as alt_ex:
                logger.error(f"Alternative close also failed: {alt_ex}")
                return {
                    "success": False,
                    "message": f"Could not close Access: {alt_ex}. Please close manually.",
                    "close_error": True
                }

        # Wait a moment for Access to close
        time.sleep(0.5)
        
        lock_file = current_path.replace('.accdb', '.laccdb')
        lock_released = not os.path.exists(lock_file)

        return {
            "success": True,
            "message": f"'{current_path}' closed successfully using {close_method}.",
            "lock_file_released": lock_released,
            "save_attempted": save_attempted,
            "save_error": save_error,
            "force_close_used": force_close,
            "warning": "Database closed without saving due to VBA errors" if save_error else None
        }

    except win32com.client.pywintypes.com_error:
        return {"success": True, "message": "MS Access was not running. Nothing to close."}
    except Exception as e:
        logger.error(f"Unexpected error in save_and_close: {e}")
        return {"success": False, "message": f"Unexpected error: {str(e)}"}

@mcp.tool()
def force_close_access(db_name: str = None) -> dict:
    """
    Force close MS Access without saving, useful when there are VBA compilation errors.
    This is a convenience wrapper around save_and_close_access_database with force_close=True.
    
    Args:
        db_name: Optional database name for verification (if None, closes any open database)
    
    Returns:
        dict with success status and message
    """
    try:
        access_app = win32com.client.GetActiveObject("Access.Application")
        
        if db_name:
            current_db = access_app.CurrentDb()
            if current_db:
                current_path = current_db.Name
                if db_name.lower() not in current_path.lower():
                    return {
                        "success": False,
                        "message": f"The open database '{current_path}' does not match '{db_name}'."
                    }
        
        logger.info("Force closing Access without saving")
        
        try:
            # Force quit without saving
            access_app.Quit(2)  # acQuitSaveNone = 2
            time.sleep(0.5)
            return {
                "success": True,
                "message": "Access force closed successfully (no save).",
                "warning": "Database was NOT saved before closing"
            }
        except Exception as quit_ex:
            logger.warning(f"Quit(2) failed: {quit_ex}, trying alternative")
            try:
                access_app.CloseCurrentDatabase()
                access_app.Quit()
                time.sleep(0.5)
                return {
                    "success": True,
                    "message": "Access closed using alternative method (no save).",
                    "warning": "Database was NOT saved before closing"
                }
            except Exception as alt_ex:
                return {
                    "success": False,
                    "message": f"Could not force close: {alt_ex}. Please close manually."
                }
                
    except win32com.client.pywintypes.com_error:
        return {"success": True, "message": "MS Access was not running. Nothing to close."}
    except Exception as e:
        logger.error(f"Unexpected error in force_close: {e}")
        return {"success": False, "message": f"Unexpected error: {str(e)}"}


    

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
def find_database(db_name: str) -> str:
    """Debug tool to find where a database file actually exists"""
    possible_paths = []
    
    # Add the resolved path from get_db_path
    resolved_path = get_db_path(db_name)
    possible_paths.append(("get_db_path() result", resolved_path, os.path.exists(resolved_path)))
    
    # Add current directory
    if not db_name.lower().endswith('.accdb'):
        db_name_with_ext = db_name + '.accdb'
    else:
        db_name_with_ext = db_name
    
    current_dir = os.path.join(os.getcwd(), db_name_with_ext)
    possible_paths.append(("Current directory", current_dir, os.path.exists(current_dir)))
    
    # Add home directory
    home_dir = os.path.join(os.path.expanduser("~"), db_name_with_ext)
    possible_paths.append(("Home directory", home_dir, os.path.exists(home_dir)))
    
    # If db_name looks like an absolute path, check it
    if os.path.isabs(db_name):
        possible_paths.append(("Absolute path (as-is)", db_name, os.path.exists(db_name)))
        if not db_name.lower().endswith('.accdb'):
            abs_with_ext = db_name + '.accdb'
            possible_paths.append(("Absolute path + .accdb", abs_with_ext, os.path.exists(abs_with_ext)))
    
    result = f"Database search results for '{db_name}':\n"
    result += f"Current working directory: {os.getcwd()}\n\n"
    
    found_any = False
    for description, path, exists in possible_paths:
        status = "✓ EXISTS" if exists else "✗ Not found"
        result += f"{description}: {status}\n  {path}\n\n"
        if exists:
            found_any = True
    
    if found_any:
        result += "✓ Database found in at least one location."
    else:
        result += "✗ Database not found in any checked location."
    
    return result

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

@mcp.tool
def begin_batch_operation(db_name: str) -> str:
    """Start a batch operation - keeps database open for multiple commands.
    
    Use this when you need to perform multiple operations (create tables, forms, VBA modules)
    in sequence. This is much faster than individual operations.
    
    IMPORTANT: You MUST call commit_batch_operation() when done!
    """
    global _batch_mode_db, _batch_mode_access
    
    if _batch_mode_access:
        return f"Error: Batch operation already in progress for '{_batch_mode_db}'"
    
    try:
        path = get_db_path(db_name)
        
        # Check for lock
        if is_database_locked(path):
            success, message = wait_for_lock_release(path, timeout=10)
            if not success:
                return f"Error: {message}"
        
        _batch_mode_access = win32com.client.Dispatch("Access.Application")
        _batch_mode_access.Visible = False
        _batch_mode_access.OpenCurrentDatabase(path)
        _batch_mode_db = db_name
        
        return f"✓ Batch operation started for '{db_name}'. Database will stay open until you call commit_batch_operation()."
    
    except Exception as e:
        _batch_mode_access = None
        _batch_mode_db = None
        return f"Error starting batch operation: {str(e)}"

@mcp.tool
def commit_batch_operation() -> str:
    """End batch operation, save all changes, and close database.
    
    Call this after you've completed all operations in a batch.
    """
    global _batch_mode_db, _batch_mode_access
    
    if not _batch_mode_access:
        return "Error: No batch operation in progress"
    
    db_name = _batch_mode_db
    
    try:
        # Save all changes
        _batch_mode_access.DoCmd.Save()
        
        # Close database
        _batch_mode_access.CloseCurrentDatabase()
        _batch_mode_access.Quit(1)
        
        # Clear state
        _batch_mode_db = None
        _batch_mode_access = None
        
        # Force cleanup
        _ensure_access_closed()
        
        return f"✓ Batch operation committed successfully for '{db_name}'. Database closed and saved."
    
    except Exception as e:
        # Try to cleanup even on error
        try:
            if _batch_mode_access:
                _batch_mode_access.Quit(1)
        except:
            pass
        
        _batch_mode_db = None
        _batch_mode_access = None
        _ensure_access_closed()
        
        return f"Error committing batch operation: {str(e)}"

@mcp.tool
def rollback_batch_operation() -> str:
    """Cancel batch operation without saving changes and close database.
    
    Use this if something went wrong and you want to discard all changes.
    """
    global _batch_mode_db, _batch_mode_access
    
    if not _batch_mode_access:
        return "Error: No batch operation in progress"
    
    db_name = _batch_mode_db
    
    try:
        # Close without saving
        _batch_mode_access.CloseCurrentDatabase()
        _batch_mode_access.Quit(0)  # acQuitSaveNone
        
        _batch_mode_db = None
        _batch_mode_access = None
        
        _ensure_access_closed()
        
        return f"✓ Batch operation rolled back for '{db_name}'. Changes discarded."
    
    except Exception as e:
        _batch_mode_db = None
        _batch_mode_access = None
        _ensure_access_closed()
        
        return f"Error rolling back batch operation: {str(e)}"

            
if __name__ == "__main__":
    mcp.run()

