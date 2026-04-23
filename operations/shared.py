import datetime as dt
import decimal
import os
from typing import Any

import pyodbc


class OperationValidationError(Exception):
    def __init__(
        self,
        message: str,
        *,
        missing_fields: list[str] | None = None,
        details: dict[str, Any] | None = None,
        status: str = "needs_input",
    ) -> None:
        super().__init__(message)
        self.message = message
        self.missing_fields = missing_fields or []
        self.details = details or {}
        self.status = status

    def to_result(self, operation: str, domain: str) -> dict[str, Any]:
        return {
            "status": self.status,
            "domain": domain,
            "operation": operation,
            "message": self.message,
            "missing_fields": self.missing_fields,
            "details": self.details,
        }


def get_db_path(db_name: str) -> str:
    if os.path.isabs(db_name):
        if not db_name.lower().endswith(".accdb"):
            db_name += ".accdb"
        return db_name

    if not db_name.lower().endswith(".accdb"):
        db_name += ".accdb"

    current_dir_path = os.path.join(os.getcwd(), db_name)
    if os.path.exists(current_dir_path):
        return current_dir_path

    home_dir_path = os.path.join(os.path.expanduser("~"), db_name)
    if os.path.exists(home_dir_path):
        return home_dir_path

    return current_dir_path


def get_driver() -> str:
    drivers = pyodbc.drivers()
    for driver_name in [
        "Microsoft Access Driver (*.mdb, *.accdb)",
        "Microsoft Access Driver (*.accdb)",
        "Microsoft Access Driver (*.mdb)",
    ]:
        if driver_name in drivers:
            return driver_name
    raise RuntimeError("Access ODBC driver not found")


def connect_access(db_name: str) -> pyodbc.Connection:
    db_path = get_db_path(db_name)
    conn_str = f"DRIVER={{{get_driver()}}};DBQ={db_path};"
    return pyodbc.connect(conn_str, autocommit=False)


def serialize_value(value: Any) -> Any:
    if isinstance(value, decimal.Decimal):
        return float(value)
    if isinstance(value, (dt.datetime, dt.date, dt.time)):
        return value.isoformat()
    return value


def rows_to_dicts(cursor: pyodbc.Cursor, rows: list[Any]) -> list[dict[str, Any]]:
    columns = [column[0] for column in cursor.description]
    serialized_rows: list[dict[str, Any]] = []
    for row in rows:
        serialized_rows.append(
            {column: serialize_value(value) for column, value in zip(columns, row)}
        )
    return serialized_rows


def first_row_to_dict(cursor: pyodbc.Cursor, row: Any) -> dict[str, Any]:
    return rows_to_dicts(cursor, [row])[0]


def normalize_text(value: str) -> str:
    return value.strip()


def require_fields(payload: dict[str, Any], fields: list[str]) -> None:
    missing = []
    for field in fields:
        value = payload.get(field)
        if value is None:
            missing.append(field)
            continue
        if isinstance(value, str) and not value.strip():
            missing.append(field)
            continue
        if isinstance(value, list) and not value:
            missing.append(field)
    if missing:
        raise OperationValidationError(
            f"Missing required fields: {', '.join(missing)}",
            missing_fields=missing,
        )


def parse_iso_date(value: str, field_name: str) -> str:
    try:
        return dt.date.fromisoformat(value.strip()).isoformat()
    except Exception as exc:  # pragma: no cover - defensive validation
        raise OperationValidationError(
            f"Field '{field_name}' must use YYYY-MM-DD format.",
            missing_fields=[field_name],
        ) from exc


def as_positive_number(value: Any, field_name: str, *, allow_zero: bool = False) -> float:
    try:
        number = float(value)
    except (TypeError, ValueError) as exc:
        raise OperationValidationError(f"Field '{field_name}' must be a number.") from exc

    if allow_zero:
        if number < 0:
            raise OperationValidationError(f"Field '{field_name}' must be >= 0.")
    elif number <= 0:
        raise OperationValidationError(f"Field '{field_name}' must be > 0.")

    return number
