# db_context.py
# Edit this file whenever your database schema changes.

DB = r"k:\bot_call_mcpmsaccess\invoice.accdb"

# Exact table/column definitions
# Set "auto": True for AUTOINCREMENT columns — these are NEVER inserted
TABLES = {
    "acctable": {
        "columns": {
            "AccID":    {"type": "int",    "auto": True},
            "AccName":  {"type": "text",   "auto": False, "required": True},
            "AccType":  {"type": "text",   "auto": False, "required": True,
                         "allowed": ["customer", "supplier", "other"]},
            "Phone":    {"type": "text",   "auto": False},
            "Address":  {"type": "text",   "auto": False},
        }
    },
    "items": {
        "columns": {
            "ItemID":    {"type": "int",    "auto": True},
            "ItemCode":  {"type": "text",   "auto": False, "required": True},
            "ItemName":  {"type": "text",   "auto": False, "required": True},
            "Unit":      {"type": "text",   "auto": False, "required": True},
            "SalePrice": {"type": "number", "auto": False},
        }
    },
    "invoices": {
        "columns": {
            "InvoiceID":   {"type": "int",    "auto": True},
            "InvoiceDate": {"type": "date",   "auto": False, "required": True},
            "AccID":       {"type": "int",    "auto": False, "required": True},
            "InvType":     {"type": "text",   "auto": False, "required": True,
                            "allowed": ["sale", "purchase"]},
            "PayType":     {"type": "text",   "auto": False, "required": True,
                            "allowed": ["cash", "credit"]},
            "TotalAmount": {"type": "number", "auto": False, "required": True},
            "Notes":       {"type": "text",   "auto": False},
        }
    },
    "itemstrans": {
        "columns": {
            "TransID":   {"type": "int",    "auto": True},
            "InvoiceID": {"type": "int",    "auto": False, "required": True},
            "ItemID":    {"type": "int",    "auto": False, "required": True},
            "Qty":       {"type": "number", "auto": False, "required": True},
            "UnitPrice": {"type": "number", "auto": False, "required": True},
            "TaxPct":    {"type": "number", "auto": False},
            "LineTotal": {"type": "number", "auto": True},  # calculated by Access
        }
    },
}

def get_insert_columns(table: str) -> list[str]:
    """Returns only the columns that can be inserted (non-auto)."""
    return [
        col for col, info in TABLES[table]["columns"].items()
        if not info.get("auto")
    ]

def schema_summary() -> str:
    """Human-readable schema for AI prompts."""
    lines = []
    for table, info in TABLES.items():
        cols = []
        for col, meta in info["columns"].items():
            note = "AUTO-skip" if meta.get("auto") else ("required" if meta.get("required") else "optional")
            allowed = f" [{'/'.join(meta['allowed'])}]" if "allowed" in meta else ""
            cols.append(f"    {col} ({meta['type']}, {note}){allowed}")
        lines.append(f"  {table}:\n" + "\n".join(cols))
    return "\n".join(lines)