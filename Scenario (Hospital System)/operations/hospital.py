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


DB_NAME = os.environ.get("DEFAULT_DB", "Hospital.accdb")
DOMAIN = "hospital"


def _success(message: str, *, data: Any = None) -> dict[str, Any]:
    result: dict[str, Any] = {"status": "success", "message": message}
    if data is not None:
        result["data"] = data
    return result


def _get_patient_by_id(cursor, patient_id: int) -> dict[str, Any] | None:
    cursor.execute(
        """
        SELECT PatientID, PatientName
        FROM Patients
        WHERE PatientID = ?
        """,
        patient_id,
    )
    row = cursor.fetchone()
    if not row:
        return None
    return first_row_to_dict(cursor, row)


def _get_patients_by_name(cursor, patient_name: str, *, exact: bool) -> list[dict[str, Any]]:
    if exact:
        cursor.execute(
            """
            SELECT PatientID, PatientName
            FROM Patients
            WHERE UCASE(PatientName) = UCASE(?)
            ORDER BY PatientID
            """,
            patient_name,
        )
    else:
        cursor.execute(
            """
            SELECT PatientID, PatientName
            FROM Patients
            WHERE UCASE(PatientName) LIKE UCASE(?)
            ORDER BY PatientName, PatientID
            """,
            f"%{patient_name}%",
        )
    return rows_to_dicts(cursor, cursor.fetchall())


def _resolve_patient(cursor, payload: dict[str, Any]) -> dict[str, Any]:
    patient_id = payload.get("patient_id")
    patient_name = (payload.get("patient_name") or "").strip()

    if patient_id not in (None, ""):
        try:
            patient_id_int = int(patient_id)
        except (TypeError, ValueError) as exc:
            raise OperationValidationError("Field 'patient_id' must be an integer.") from exc

        patient = _get_patient_by_id(cursor, patient_id_int)
        if not patient:
            raise OperationValidationError(
                f"Patient with id {patient_id_int} was not found.",
                details={"unknown_patient_id": patient_id_int},
            )
        return patient

    if not patient_name:
        raise OperationValidationError(
            "You must provide either patient_id or patient_name.",
            missing_fields=["patient_id_or_patient_name"],
        )

    matches = _get_patients_by_name(cursor, patient_name, exact=True)
    if not matches:
        raise OperationValidationError(
            f"Patient '{patient_name}' was not found.",
            details={
                "unknown_patient_name": patient_name,
                "suggested_operation": "hospital.create_patient",
            },
        )
    if len(matches) > 1:
        raise OperationValidationError(
            f"More than one patient matched '{patient_name}'. Please provide patient_id.",
            details={"matching_patients": matches},
        )
    return matches[0]


def list_patients(payload: dict[str, Any]) -> dict[str, Any]:
    with connect_access(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute(
            """
            SELECT PatientID, PatientName
            FROM Patients
            ORDER BY PatientName, PatientID
            """
        )
        rows = rows_to_dicts(cursor, cursor.fetchall())
    return _success(f"Found {len(rows)} patient(s).", data=rows)


def find_patient(payload: dict[str, Any]) -> dict[str, Any]:
    require_fields(payload, ["patient_name"])
    patient_name = normalize_text(payload["patient_name"])
    match_mode = (payload.get("match_mode") or "contains").strip().lower()

    with connect_access(DB_NAME) as conn:
        cursor = conn.cursor()
        rows = _get_patients_by_name(cursor, patient_name, exact=(match_mode == "exact"))
    return _success(f"Found {len(rows)} patient(s) for '{patient_name}'.", data=rows)


def create_patient(payload: dict[str, Any]) -> dict[str, Any]:
    require_fields(payload, ["patient_name"])
    patient_name = normalize_text(payload["patient_name"])

    with connect_access(DB_NAME) as conn:
        cursor = conn.cursor()
        existing = _get_patients_by_name(cursor, patient_name, exact=True)
        if existing:
            conn.rollback()
            return _success(f"Patient '{patient_name}' already exists.", data=existing[0])

        cursor.execute(
            """
            INSERT INTO Patients (PatientName)
            VALUES (?)
            """,
            patient_name,
        )
        cursor.execute("SELECT @@IDENTITY")
        patient_id = int(cursor.fetchone()[0])
        conn.commit()

    return _success(
        f"Patient '{patient_name}' created successfully.",
        data={"PatientID": patient_id, "PatientName": patient_name},
    )


def list_inventory(payload: dict[str, Any]) -> dict[str, Any]:
    with connect_access(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute(
            """
            SELECT MedicineID, CurrentQty
            FROM Inventory
            ORDER BY MedicineID
            """
        )
        rows = rows_to_dicts(cursor, cursor.fetchall())
    return _success(f"Found {len(rows)} inventory item(s).", data=rows)


def list_recent_visits(payload: dict[str, Any]) -> dict[str, Any]:
    limit = int(payload.get("limit", 10) or 10)
    if limit < 1 or limit > 100:
        raise OperationValidationError("Field 'limit' must be between 1 and 100.")

    with connect_access(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute(
            f"""
            SELECT TOP {limit}
                v.VisitID,
                v.PatientID,
                p.PatientName,
                v.DoctorID,
                v.VisitDate
            FROM Visits AS v
            INNER JOIN Patients AS p ON v.PatientID = p.PatientID
            ORDER BY v.VisitID DESC
            """
        )
        rows = rows_to_dicts(cursor, cursor.fetchall())
    return _success(f"Found {len(rows)} recent visit(s).", data=rows)


def get_visit_details(payload: dict[str, Any]) -> dict[str, Any]:
    require_fields(payload, ["visit_id"])
    try:
        visit_id = int(payload["visit_id"])
    except (TypeError, ValueError) as exc:
        raise OperationValidationError("Field 'visit_id' must be an integer.") from exc

    with connect_access(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute(
            """
            SELECT
                v.VisitID,
                v.PatientID,
                p.PatientName,
                v.DoctorID,
                v.VisitDate
            FROM Visits AS v
            INNER JOIN Patients AS p ON v.PatientID = p.PatientID
            WHERE v.VisitID = ?
            """,
            visit_id,
        )
        visit_row = cursor.fetchone()
        if not visit_row:
            return _success(f"Visit #{visit_id} was not found.", data={"visit": None, "services": [], "prescriptions": []})

        visit = first_row_to_dict(cursor, visit_row)

        cursor.execute(
            """
            SELECT ServiceID, VisitID, ServiceName, ServiceCost
            FROM VisitServices
            WHERE VisitID = ?
            ORDER BY ServiceID
            """,
            visit_id,
        )
        services = rows_to_dicts(cursor, cursor.fetchall())

        cursor.execute(
            """
            SELECT PrescriptionID, VisitID, MedicineID, Qty
            FROM Prescriptions
            WHERE VisitID = ?
            ORDER BY PrescriptionID
            """,
            visit_id,
        )
        prescriptions = rows_to_dicts(cursor, cursor.fetchall())

        cursor.execute(
            """
            SELECT EntryID, PatientID, EntryDate, Debit, Credit, ReferenceID
            FROM PatientLedger
            WHERE ReferenceID = ?
            ORDER BY EntryID
            """,
            visit_id,
        )
        ledger_entries = rows_to_dicts(cursor, cursor.fetchall())

        cursor.execute(
            """
            SELECT LogID, DoctorID, LogDate, Activity, ReferenceID
            FROM DoctorLog
            WHERE ReferenceID = ?
            ORDER BY LogID
            """,
            visit_id,
        )
        doctor_logs = rows_to_dicts(cursor, cursor.fetchall())

    return _success(
        f"Loaded visit #{visit_id}.",
        data={
            "visit": visit,
            "services": services,
            "prescriptions": prescriptions,
            "ledger_entries": ledger_entries,
            "doctor_logs": doctor_logs,
        },
    )


def create_visit(payload: dict[str, Any]) -> dict[str, Any]:
    require_fields(payload, ["doctor_id", "visit_date", "services"])

    try:
        doctor_id = int(payload["doctor_id"])
    except (TypeError, ValueError) as exc:
        raise OperationValidationError("Field 'doctor_id' must be an integer.") from exc

    visit_date = parse_iso_date(payload["visit_date"], "visit_date")
    services_payload = payload["services"]
    prescriptions_payload = payload.get("prescriptions", []) or []

    if not isinstance(services_payload, list) or not services_payload:
        raise OperationValidationError(
            "Field 'services' must be a non-empty list.",
            missing_fields=["services"],
        )
    if not isinstance(prescriptions_payload, list):
        raise OperationValidationError("Field 'prescriptions' must be a list if provided.")

    normalized_services: list[dict[str, Any]] = []
    service_missing_fields: list[str] = []
    for index, service in enumerate(services_payload):
        if not isinstance(service, dict):
            raise OperationValidationError(f"Each service must be an object. Invalid entry at index {index}.")

        item_missing_fields: list[str] = []
        service_name = (service.get("service_name") or "").strip()
        service_cost = service.get("service_cost")

        if not service_name:
            item_missing_fields.append(f"services[{index}].service_name")
        if service_cost in (None, ""):
            item_missing_fields.append(f"services[{index}].service_cost")

        if item_missing_fields:
            service_missing_fields.extend(item_missing_fields)
            continue

        normalized_services.append(
            {
                "service_name": normalize_text(service_name),
                "service_cost": as_positive_number(
                    service_cost,
                    f"services[{index}].service_cost",
                    allow_zero=True,
                ),
            }
        )

    if service_missing_fields:
        raise OperationValidationError(
            f"Missing required service fields: {', '.join(service_missing_fields)}",
            missing_fields=service_missing_fields,
        )

    normalized_prescriptions: list[dict[str, Any]] = []
    prescription_missing_fields: list[str] = []
    for index, prescription in enumerate(prescriptions_payload):
        if not isinstance(prescription, dict):
            raise OperationValidationError(f"Each prescription must be an object. Invalid entry at index {index}.")

        item_missing_fields: list[str] = []
        medicine_id = prescription.get("medicine_id")
        qty = prescription.get("qty")

        if medicine_id in (None, ""):
            item_missing_fields.append(f"prescriptions[{index}].medicine_id")
        if qty in (None, ""):
            item_missing_fields.append(f"prescriptions[{index}].qty")

        if item_missing_fields:
            prescription_missing_fields.extend(item_missing_fields)
            continue

        try:
            medicine_id_int = int(medicine_id)
        except (TypeError, ValueError) as exc:
            raise OperationValidationError(
                f"Field 'prescriptions[{index}].medicine_id' must be an integer."
            ) from exc

        normalized_prescriptions.append(
            {
                "medicine_id": medicine_id_int,
                "qty": as_positive_number(qty, f"prescriptions[{index}].qty"),
            }
        )

    if prescription_missing_fields:
        raise OperationValidationError(
            f"Missing required prescription fields: {', '.join(prescription_missing_fields)}",
            missing_fields=prescription_missing_fields,
        )

    with connect_access(DB_NAME) as conn:
        cursor = conn.cursor()
        patient = _resolve_patient(cursor, payload)

        inventory_snapshot: list[dict[str, Any]] = []
        for item in normalized_prescriptions:
            cursor.execute(
                """
                SELECT MedicineID, CurrentQty
                FROM Inventory
                WHERE MedicineID = ?
                """,
                item["medicine_id"],
            )
            stock_row = cursor.fetchone()
            if not stock_row:
                raise OperationValidationError(
                    f"Medicine {item['medicine_id']} was not found in inventory.",
                    details={"unknown_medicine_id": item["medicine_id"]},
                )
            stock = first_row_to_dict(cursor, stock_row)
            current_qty = float(stock["CurrentQty"])
            if current_qty < item["qty"]:
                raise OperationValidationError(
                    f"Medicine {item['medicine_id']} does not have enough stock.",
                    details={
                        "medicine_id": item["medicine_id"],
                        "current_qty": current_qty,
                        "requested_qty": item["qty"],
                    },
                )
            inventory_snapshot.append(stock)

        total_cost = sum(service["service_cost"] for service in normalized_services)

        try:
            cursor.execute(
                """
                INSERT INTO Visits (PatientID, DoctorID, VisitDate)
                VALUES (?, ?, ?)
                """,
                patient["PatientID"],
                doctor_id,
                visit_date,
            )
            cursor.execute("SELECT @@IDENTITY")
            visit_id = int(cursor.fetchone()[0])

            for service in normalized_services:
                cursor.execute(
                    """
                    INSERT INTO VisitServices (VisitID, ServiceName, ServiceCost)
                    VALUES (?, ?, ?)
                    """,
                    visit_id,
                    service["service_name"],
                    service["service_cost"],
                )

            for prescription in normalized_prescriptions:
                cursor.execute(
                    """
                    INSERT INTO Prescriptions (VisitID, MedicineID, Qty)
                    VALUES (?, ?, ?)
                    """,
                    visit_id,
                    prescription["medicine_id"],
                    prescription["qty"],
                )
                cursor.execute(
                    """
                    UPDATE Inventory
                    SET CurrentQty = CurrentQty - ?
                    WHERE MedicineID = ?
                    """,
                    prescription["qty"],
                    prescription["medicine_id"],
                )

            cursor.execute(
                """
                INSERT INTO PatientLedger (PatientID, EntryDate, Debit, Credit, ReferenceID)
                VALUES (?, ?, ?, ?, ?)
                """,
                patient["PatientID"],
                visit_date,
                total_cost,
                0,
                visit_id,
            )

            cursor.execute(
                """
                INSERT INTO DoctorLog (DoctorID, LogDate, Activity, ReferenceID)
                VALUES (?, ?, ?, ?)
                """,
                doctor_id,
                visit_date,
                "Patient Consultation",
                visit_id,
            )

            conn.commit()
        except Exception:
            conn.rollback()
            raise

    return _success(
        f"Visit #{visit_id} created successfully.",
        data={
            "visit_id": visit_id,
            "patient_id": patient["PatientID"],
            "patient_name": patient["PatientName"],
            "doctor_id": doctor_id,
            "visit_date": visit_date,
            "services": normalized_services,
            "prescriptions": normalized_prescriptions,
            "total_cost": total_cost,
        },
    )


def register_operations(registry) -> None:
    registry.register("hospital.list_patients", DOMAIN, list_patients)
    registry.register("hospital.find_patient", DOMAIN, find_patient)
    registry.register("hospital.create_patient", DOMAIN, create_patient)
    registry.register("hospital.list_inventory", DOMAIN, list_inventory)
    registry.register("hospital.list_recent_visits", DOMAIN, list_recent_visits)
    registry.register("hospital.get_visit_details", DOMAIN, get_visit_details)
    registry.register("hospital.create_visit", DOMAIN, create_visit)
