from .accounting import register_operations as register_accounting_operations
from .registry import OperationRegistry


registry = OperationRegistry()
register_accounting_operations(registry)


def execute_registered_operation(operation: str, payload: dict, domain: str) -> dict:
    return registry.execute(operation, payload or {}, domain)
