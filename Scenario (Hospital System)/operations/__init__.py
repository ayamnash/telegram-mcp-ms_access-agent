from .hospital import register_operations as register_hospital_operations
from .registry import OperationRegistry


registry = OperationRegistry()
register_hospital_operations(registry)


def execute_registered_operation(operation: str, payload: dict, domain: str) -> dict:
    return registry.execute(operation, payload or {}, domain)
