from __future__ import annotations

import math
import struct

from .models import RegisterFormat, RegisterValueType


class ValueEncodingError(ValueError):
    pass


class ValueEncoder:
    @staticmethod
    def encode_value(
        value: int | float,
        value_type: RegisterValueType,
        register_format: RegisterFormat,
    ) -> list[int]:
        if value_type == RegisterValueType.INT16:
            raw = ValueEncoder._pack_int(int(value), bits=16, signed=True)
        elif value_type == RegisterValueType.UINT16:
            raw = ValueEncoder._pack_int(int(value), bits=16, signed=False)
        elif value_type == RegisterValueType.INT32:
            raw = ValueEncoder._pack_int(int(value), bits=32, signed=True)
        elif value_type == RegisterValueType.UINT32:
            raw = ValueEncoder._pack_int(int(value), bits=32, signed=False)
        elif value_type == RegisterValueType.FLOAT32:
            raw = ValueEncoder._pack_float(float(value), bits=32)
        elif value_type == RegisterValueType.INT64:
            raw = ValueEncoder._pack_int(int(value), bits=64, signed=True)
        elif value_type == RegisterValueType.UINT64:
            raw = ValueEncoder._pack_int(int(value), bits=64, signed=False)
        elif value_type == RegisterValueType.FLOAT64:
            raw = ValueEncoder._pack_float(float(value), bits=64)
        else:
            raise ValueEncodingError(f"Nicht unterstuetzter Datentyp: {value_type}")
        return ValueEncoder._bytes_to_registers(raw, register_format)

    @staticmethod
    def coerce_text_value(raw_text: str, value_type: RegisterValueType) -> int | float:
        text = raw_text.strip().replace(",", ".")
        if not text:
            raise ValueEncodingError("Leerer Wert kann nicht kodiert werden")
        if value_type.requires_integer:
            number = float(text)
            if not number.is_integer():
                raise ValueEncodingError(f"{value_type.label} erwartet eine Ganzzahl, erhalten: {raw_text}")
            return int(number)
        return float(text)

    @staticmethod
    def _bytes_to_registers(raw: bytes, register_format: RegisterFormat) -> list[int]:
        if len(raw) % 2 != 0:
            raise ValueEncodingError("Bytefolge muss aus 16-Bit-Worten bestehen")
        words = [bytearray(raw[index:index + 2]) for index in range(0, len(raw), 2)]
        if register_format in {RegisterFormat.LITTLE, RegisterFormat.LITTLE_BYTE_SWAP}:
            for word in words:
                word.reverse()
        if register_format in {RegisterFormat.BIG_BYTE_SWAP, RegisterFormat.LITTLE_BYTE_SWAP}:
            words.reverse()
        return [int.from_bytes(bytes(word), byteorder="big", signed=False) for word in words]

    @staticmethod
    def _pack_int(value: int, bits: int, signed: bool) -> bytes:
        minimum = -(2 ** (bits - 1)) if signed else 0
        maximum = (2 ** (bits - 1) - 1) if signed else (2 ** bits - 1)
        if not minimum <= value <= maximum:
            raise ValueEncodingError(f"Wert {value} liegt ausserhalb des Bereichs fuer {bits}-Bit {'signed' if signed else 'unsigned'}")
        return value.to_bytes(bits // 8, byteorder="big", signed=signed)

    @staticmethod
    def _pack_float(value: float, bits: int) -> bytes:
        if not math.isfinite(value):
            raise ValueEncodingError(f"Float-Wert ist nicht endlich: {value}")
        if bits == 32:
            return struct.pack(">f", value)
        if bits == 64:
            return struct.pack(">d", value)
        raise ValueEncodingError(f"Nicht unterstuetzte Float-Breite: {bits}")
