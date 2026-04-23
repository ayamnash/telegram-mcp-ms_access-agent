import os
from typing import Any

from .shared import (
    OperationValidationError,
    as_positive_number,
    connect_access,
    first_row_to_dict,
    normalize_text,
    parse_iso_date,
    require_fields,
    rows_to_dicts,
)


DB_NAME = os.environ.get("DEFAULT_DB", "invoice.accdb")
DOMAIN = "accounting"


def _success(message: str, *, data: Any = None) -> dict[str, Any]:
    result: dict[str, Any] = {"status": "success", "message": message}
    if data is not None:
        result["data"] = data
    return result


def _get_customer_by_name(cursor, customer_name: str) -> dict[str, Any] | None:
    cursor.execute(
        """
        SELECT AccID, AccName, AccType, Phone
        FROM acctable
        WHERE UCASE(AccName) = UCASE(?)
        """,
        customer_name,
    )
    row = cursor.fetchone()
    if not row:
        return None
    return first_row_to_dict(cursor, row)


def _get_item_by_name(cursor, item_name: str) -> dict[str, Any] | None:
    cursor.execute(
        """
        SELECT ItemID, ItemCode, ItemName, Unit, SalePrice
        FROM items
        WHERE UCASE(ItemName) = UCASE(?)
        """,
        item_name,
    )
    row = cursor.fetchone()
    if not row:
        return None
    return first_row_to_dict(cursor, row)


def list_items(payload: dict[str, Any]) -> dict[str, Any]:
    with connect_access(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute(
            """
            SELECT ItemID, ItemCode, ItemName, Unit, SalePrice
            FROM items
            ORDER BY ItemName
            """
        )
        rows = rows_to_dicts(cursor, cursor.fetchall())
    return _success(f"Found {len(rows)} item(s).", data=rows)


def list_customers(payload: dict[str, Any]) -> dict[str, Any]:
    with connect_access(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute(
            """
            SELECT AccID, AccName, AccType, Phone
            FROM acctable
            ORDER BY AccName
            """
        )
        rows = rows_to_dicts(cursor, cursor.fetchall())
    return _success(f"Found {len(rows)} customer/account record(s).", data=rows)


def list_recent_invoices(payload: dict[str, Any]) -> dict[str, Any]:
    limit = int(payload.get("limit", 10) or 10)
    if limit < 1 or limit > 100:
        raise OperationValidationError("Field 'limit' must be between 1 and 100.")

    with connect_access(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute(
            f"""
            SELECT TOP {limit}
                i.InvoiceID,
                i.InvoiceDate,
                a.AccName,
                i.InvType,
                i.PayType,
                i.TotalAmount,
                i.Notes
            FROM invoices AS i
            INNER JOIN acctable AS a ON i.AccID = a.AccID
            ORDER BY i.InvoiceID DESC
            """
        )
        rows = rows_to_dicts(cursor, cursor.fetchall())
    return _success(f"Found {len(rows)} recent invoice(s).", data=rows)


def find_item(payload: dict[str, Any]) -> dict[str, Any]:
    require_fields(payload, ["item_name"])
    item_name = normalize_text(payload["item_name"])
    match_mode = (payload.get("match_mode") or "contains").strip().lower()

    with connect_access(DB_NAME) as conn:
        cursor = conn.cursor()
        if match_mode == "exact":
            cursor.execute(
                """
                SELECT ItemID, ItemCode, ItemName, Unit, SalePrice
                FROM items
                WHERE UCASE(ItemName) = UCASE(?)
                ORDER BY ItemName
                """,
                item_name,
            )
        else:
            cursor.execute(
                """
                SELECT ItemID, ItemCode, ItemName, Unit, SalePrice
                FROM items
                WHERE UCASE(ItemName) LIKE UCASE(?)
                ORDER BY ItemName
                """,
                f"%{item_name}%",
            )
        rows = rows_to_dicts(cursor, cursor.fetchall())
    return _success(f"Found {len(rows)} item(s) for '{item_name}'.", data=rows)


def find_customer(payload: dict[str, Any]) -> dict[str, Any]:
    require_fields(payload, ["customer_name"])
    customer_name = normalize_text(payload["customer_name"])
    match_mode = (payload.get("match_mode") or "contains").strip().lower()

    with connect_access(DB_NAME) as conn:
        cursor = conn.cursor()
        if match_mode == "exact":
            cursor.execute(
                """
                SELECT AccID, AccName, AccType, Phone
                FROM acctable
                WHERE UCASE(AccName) = UCASE(?)
                ORDER BY AccName
                """,
                customer_name,
            )
        else:
            cursor.execute(
                """
                SELECT AccID, AccName, AccType, Phone
                FROM acctable
                WHERE UCASE(AccName) LIKE UCASE(?)
                ORDER BY AccName
                """,
                f"%{customer_name}%",
            )
        rows = rows_to_dicts(cursor, cursor.fetchall())
    return _success(f"Found {len(rows)} customer/account record(s).", data=rows)


def find_invoices_by_customer(payload: dict[str, Any]) -> dict[str, Any]:
    require_fields(payload, ["customer_name"])
    customer_name = normalize_text(payload["customer_name"])

    with connect_access(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute(
            """
            SELECT
                i.InvoiceID,
                i.InvoiceDate,
                a.AccName,
                i.InvType,
                i.PayType,
                i.TotalAmount,
                i.Notes
            FROM invoices AS i
            INNER JOIN acctable AS a ON i.AccID = a.AccID
            WHERE UCASE(a.AccName) = UCASE(?)
            ORDER BY i.InvoiceID DESC
            """,
            customer_name,
        )
        rows = rows_to_dicts(cursor, cursor.fetchall())
    return _success(f"Found {len(rows)} invoice(s) for '{customer_name}'.", data=rows)


def get_invoice_details(payload: dict[str, Any]) -> dict[str, Any]:
    require_fields(payload, ["invoice_id"])
    invoice_id = int(payload["invoice_id"])

    with connect_access(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute(
            """
            SELECT
                i.InvoiceID,
                i.InvoiceDate,
                a.AccName,
                i.InvType,
                i.PayType,
                i.TotalAmount,
                i.Notes
            FROM invoices AS i
            INNER JOIN acctable AS a ON i.AccID = a.AccID
            WHERE i.InvoiceID = ?
            """,
            invoice_id,
        )
        header_row = cursor.fetchone()
        if not header_row:
            return _success(f"Invoice #{invoice_id} was not found.", data={"header": None, "lines": []})

        header = first_row_to_dict(cursor, header_row)

        cursor.execute(
            """
            SELECT
                t.TransID,
                t.Qty,
                t.UnitPrice,
                t.TaxPct,
                t.LineTotal,
                it.ItemID,
                it.ItemName,
                it.Unit
            FROM itemstrans AS t
            INNER JOIN items AS it ON t.ItemID = it.ItemID
            WHERE t.InvoiceID = ?
            ORDER BY t.TransID
            """,
            invoice_id,
        )
        lines = rows_to_dicts(cursor, cursor.fetchall())

    return _success(f"Loaded invoice #{invoice_id}.", data={"header": header, "lines": lines})


def list_invoices_by_date_range(payload: dict[str, Any]) -> dict[str, Any]:
    require_fields(payload, ["start_date", "end_date"])
    start_date = parse_iso_date(payload["start_date"], "start_date")
    end_date = parse_iso_date(payload["end_date"], "end_date")

    with connect_access(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute(
            """
            SELECT
                i.InvoiceID,
                i.InvoiceDate,
                a.AccName,
                i.InvType,
                i.PayType,
                i.TotalAmount,
                i.Notes
            FROM invoices AS i
            INNER JOIN acctable AS a ON i.AccID = a.AccID
            WHERE i.InvoiceDate BETWEEN ? AND ?
            ORDER BY i.InvoiceDate DESC, i.InvoiceID DESC
            """,
            start_date,
            end_date,
        )
        rows = rows_to_dicts(cursor, cursor.fetchall())
    return _success(f"Found {len(rows)} invoice(s) between {start_date} and {end_date}.", data=rows)


def total_sales_per_customer(payload: dict[str, Any]) -> dict[str, Any]:
    with connect_access(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute(
            """
            SELECT
                a.AccName,
                COUNT(i.InvoiceID) AS InvoiceCount,
                SUM(i.TotalAmount) AS TotalSales
            FROM invoices AS i
            INNER JOIN acctable AS a ON i.AccID = a.AccID
            WHERE i.InvType = 'sale'
            GROUP BY a.AccName
            ORDER BY SUM(i.TotalAmount) DESC
            """
        )
        rows = rows_to_dicts(cursor, cursor.fetchall())
    return _success(f"Calculated sales totals for {len(rows)} customer(s).", data=rows)


def create_customer(payload: dict[str, Any]) -> dict[str, Any]:
    require_fields(payload, ["customer_name"])
    customer_name = normalize_text(payload["customer_name"])
    acc_type = (payload.get("acc_type") or "customer").strip().lower()
    phone = (payload.get("phone") or "").strip() or None

    if acc_type not in {"customer", "supplier", "other"}:
        raise OperationValidationError("Field 'acc_type' must be customer, supplier, or other.")

    with connect_access(DB_NAME) as conn:
        cursor = conn.cursor()
        existing_customer = _get_customer_by_name(cursor, customer_name)
        if existing_customer:
            conn.rollback()
            return _success(f"Customer/account '{customer_name}' already exists.", data=existing_customer)

        cursor.execute(
            """
            INSERT INTO acctable (AccName, AccType, Phone)
            VALUES (?, ?, ?)
            """,
            customer_name,
            acc_type,
            phone,
        )
        cursor.execute("SELECT @@IDENTITY")
        acc_id = int(cursor.fetchone()[0])
        conn.commit()

    return _success(
        f"Customer/account '{customer_name}' created successfully.",
        data={"AccID": acc_id, "AccName": customer_name, "AccType": acc_type, "Phone": phone},
    )


def create_item(payload: dict[str, Any]) -> dict[str, Any]:
    require_fields(payload, ["item_name", "item_code", "unit"])
    item_name = normalize_text(payload["item_name"])
    item_code = normalize_text(payload["item_code"])
    unit = normalize_text(payload["unit"])
    sale_price = as_positive_number(payload.get("sale_price", 0), "sale_price", allow_zero=True)

    with connect_access(DB_NAME) as conn:
        cursor = conn.cursor()
        existing_item = _get_item_by_name(cursor, item_name)
        if existing_item:
            conn.rollback()
            return _success(f"Item '{item_name}' already exists.", data=existing_item)

        cursor.execute(
            """
            INSERT INTO items (ItemCode, ItemName, Unit, SalePrice)
            VALUES (?, ?, ?, ?)
            """,
            item_code,
            item_name,
            unit,
            sale_price,
        )
        cursor.execute("SELECT @@IDENTITY")
        item_id = int(cursor.fetchone()[0])
        conn.commit()

    return _success(
        f"Item '{item_name}' created successfully.",
        data={
            "ItemID": item_id,
            "ItemCode": item_code,
            "ItemName": item_name,
            "Unit": unit,
            "SalePrice": sale_price,
        },
    )


def create_invoice(payload: dict[str, Any]) -> dict[str, Any]:
    require_fields(payload, ["invoice_date", "customer_name", "inv_type", "pay_type", "items"])

    invoice_date = parse_iso_date(payload["invoice_date"], "invoice_date")
    customer_name = normalize_text(payload["customer_name"])
    inv_type = normalize_text(payload["inv_type"]).lower()
    pay_type = normalize_text(payload["pay_type"]).lower()
    notes = (payload.get("notes") or "").strip() or None
    items_payload = payload["items"]

    if inv_type not in {"sale", "purchase"}:
        raise OperationValidationError("Field 'inv_type' must be sale or purchase.")
    if pay_type not in {"cash", "credit"}:
        raise OperationValidationError("Field 'pay_type' must be cash or credit.")
    if not isinstance(items_payload, list) or not items_payload:
        raise OperationValidationError(
            "Field 'items' must be a non-empty list.",
            missing_fields=["items"],
        )

    normalized_items: list[dict[str, Any]] = []
    missing_fields: list[str] = []
    for index, item in enumerate(items_payload):
        if not isinstance(item, dict):
            raise OperationValidationError(f"Each item must be an object. Invalid entry at index {index}.")

        item_missing_fields: list[str] = []
        item_name = (item.get("item_name") or "").strip()
        if not item_name:
            item_missing_fields.append(f"items[{index}].item_name")

        quantity = item.get("quantity")
        unit_price = item.get("unit_price")
        if quantity in (None, ""):
            item_missing_fields.append(f"items[{index}].quantity")
        if unit_price in (None, ""):
            item_missing_fields.append(f"items[{index}].unit_price")

        if item_missing_fields:
            missing_fields.extend(item_missing_fields)
            continue

        tax_pct = item.get("tax_pct", 0)
        normalized_items.append(
            {
                "item_name": normalize_text(item_name),
                "quantity": as_positive_number(quantity, f"items[{index}].quantity"),
                "unit_price": as_positive_number(
                    unit_price,
                    f"items[{index}].unit_price",
                    allow_zero=True,
                ),
                "tax_pct": as_positive_number(
                    tax_pct,
                    f"items[{index}].tax_pct",
                    allow_zero=True,
                ),
            }
        )

    if missing_fields:
        raise OperationValidationError(
            f"Missing required invoice item fields: {', '.join(missing_fields)}",
            missing_fields=missing_fields,
        )

    with connect_access(DB_NAME) as conn:
        cursor = conn.cursor()

        customer = _get_customer_by_name(cursor, customer_name)
        if not customer:
            raise OperationValidationError(
                f"Customer/account '{customer_name}' was not found.",
                details={"unknown_customer": customer_name, "suggested_operation": "accounting.create_customer"},
            )

        resolved_items: list[dict[str, Any]] = []
        unknown_items: list[str] = []
        for item in normalized_items:
            existing_item = _get_item_by_name(cursor, item["item_name"])
            if not existing_item:
                unknown_items.append(item["item_name"])
                continue
            resolved_items.append({**item, "item_id": existing_item["ItemID"], "unit": existing_item["Unit"]})

        if unknown_items:
            raise OperationValidationError(
                "Some items were not found.",
                details={"unknown_items": unknown_items, "suggested_operation": "accounting.create_item"},
            )

        total_amount = 0.0
        for item in resolved_items:
            total_amount += item["quantity"] * item["unit_price"] * (1 + (item["tax_pct"] / 100))

        try:
            cursor.execute(
                """
                INSERT INTO invoices (InvoiceDate, AccID, InvType, PayType, TotalAmount, Notes)
                VALUES (?, ?, ?, ?, ?, ?)
                """,
                invoice_date,
                customer["AccID"],
                inv_type,
                pay_type,
                round(total_amount, 2),
                notes,
            )
            cursor.execute("SELECT @@IDENTITY")
            invoice_id = int(cursor.fetchone()[0])

            for item in resolved_items:
                cursor.execute(
                    """
                    INSERT INTO itemstrans (InvoiceID, ItemID, Qty, UnitPrice, TaxPct)
                    VALUES (?, ?, ?, ?, ?)
                    """,
                    invoice_id,
                    item["item_id"],
                    item["quantity"],
                    item["unit_price"],
                    item["tax_pct"],
                )

            conn.commit()
        except Exception:
            conn.rollback()
            raise

    return _success(
        f"Invoice #{invoice_id} created successfully.",
        data={
            "invoice_id": invoice_id,
            "invoice_date": invoice_date,
            "customer_name": customer["AccName"],
            "inv_type": inv_type,
            "pay_type": pay_type,
            "total_amount": round(total_amount, 2),
            "items": resolved_items,
        },
    )


def register_operations(registry) -> None:
    registry.register("accounting.list_items", DOMAIN, list_items)
    registry.register("accounting.list_customers", DOMAIN, list_customers)
    registry.register("accounting.list_recent_invoices", DOMAIN, list_recent_invoices)
    registry.register("accounting.find_item", DOMAIN, find_item)
    registry.register("accounting.find_customer", DOMAIN, find_customer)
    registry.register("accounting.find_invoices_by_customer", DOMAIN, find_invoices_by_customer)
    registry.register("accounting.get_invoice_details", DOMAIN, get_invoice_details)
    registry.register("accounting.list_invoices_by_date_range", DOMAIN, list_invoices_by_date_range)
    registry.register("accounting.total_sales_per_customer", DOMAIN, total_sales_per_customer)
    registry.register("accounting.create_customer", DOMAIN, create_customer)
    registry.register("accounting.create_item", DOMAIN, create_item)
    registry.register("accounting.create_invoice", DOMAIN, create_invoice)
