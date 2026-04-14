from __future__ import annotations

from dataclasses import dataclass

from .modbus_service import ModbusService
from .models import ConnectionSettings, RegisterFormat, RegisterValueType
from .value_encoder import ValueEncoder


@dataclass(slots=True)
class FloatReadResult:
    address: int
    registers: list[int]
    value: float


class ModbusClientService(ModbusService):
    pass


def value_to_registers(value: int | float, value_type: RegisterValueType, register_format: RegisterFormat) -> list[int]:
    return ValueEncoder.encode_value(value, value_type, register_format)
