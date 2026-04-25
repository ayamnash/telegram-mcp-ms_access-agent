from collections.abc import Callable
from typing import Any

from .shared import OperationValidationError


OperationHandler = Callable[[dict[str, Any]], dict[str, Any]]


class OperationRegistry:
    def __init__(self) -> None:
        self._operations: dict[str, dict[str, Any]] = {}

    def register(self, operation: str, domain: str, handler: OperationHandler) -> None:
        self._operations[operation] = {"domain": domain, "handler": handler}

    def execute(self, operation: str, payload: dict[str, Any], domain: str) -> dict[str, Any]:
        metadata = self._operations.get(operation)
        if metadata is None:
            return {
                "status": "error",
                "domain": domain,
                "operation": operation,
                "message": f"Unknown operation '{operation}'.",
            }

        if metadata["domain"] != domain:
            return {
                "status": "error",
                "domain": domain,
                "operation": operation,
                "message": f"Operation '{operation}' does not belong to domain '{domain}'.",
            }

        try:
            result = metadata["handler"](payload)
        except OperationValidationError as exc:
            return exc.to_result(operation, domain)
        except Exception as exc:  # pragma: no cover - runtime safety
            return {
                "status": "error",
                "domain": domain,
                "operation": operation,
                "message": str(exc),
            }

        result.setdefault("status", "success")
        result.setdefault("domain", domain)
        result.setdefault("operation", operation)
        return result
