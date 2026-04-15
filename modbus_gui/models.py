from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum


class RegisterFormat(str, Enum):
    BIG = "big"
    LITTLE = "little"
    BIG_BYTE_SWAP = "big_byte_swap"
    LITTLE_BYTE_SWAP = "little_byte_swap"

    @property
    def label(self) -> str:
        return {
            RegisterFormat.BIG: "Big Endian",
            RegisterFormat.LITTLE: "Little Endian",
            RegisterFormat.BIG_BYTE_SWAP: "Big Endian Byte Swap",
            RegisterFormat.LITTLE_BYTE_SWAP: "Little Endian Byte Swap",
        }[self]

    @classmethod
    def from_legacy(cls, byte_order: str | None, word_order: str | None) -> "RegisterFormat":
        mapping = {
            ("big", "normal"): cls.BIG,
            ("big", "big"): cls.BIG,
            ("little", "normal"): cls.LITTLE,
            ("little", "big"): cls.LITTLE,
            ("big", "swap"): cls.BIG_BYTE_SWAP,
            ("big", "little"): cls.BIG_BYTE_SWAP,
            ("little", "swap"): cls.LITTLE_BYTE_SWAP,
            ("little", "little"): cls.LITTLE_BYTE_SWAP,
        }
        return mapping.get(((byte_order or "big").lower(), (word_order or "normal").lower()), cls.BIG)


class RegisterValueType(str, Enum):
    INT16 = "int16"
    UINT16 = "uint16"
    INT32 = "int32"
    UINT32 = "uint32"
    FLOAT32 = "float32"
    INT64 = "int64"
    UINT64 = "uint64"
    FLOAT64 = "float64"

    @property
    def label(self) -> str:
        return {
            RegisterValueType.INT16: "Int16",
            RegisterValueType.UINT16: "UInt16",
            RegisterValueType.INT32: "Int32",
            RegisterValueType.UINT32: "UInt32",
            RegisterValueType.FLOAT32: "Float32",
            RegisterValueType.INT64: "Int64",
            RegisterValueType.UINT64: "UInt64",
            RegisterValueType.FLOAT64: "Float64",
        }[self]

    @property
    def register_count(self) -> int:
        return {
            RegisterValueType.INT16: 1,
            RegisterValueType.UINT16: 1,
            RegisterValueType.INT32: 2,
            RegisterValueType.UINT32: 2,
            RegisterValueType.FLOAT32: 2,
            RegisterValueType.INT64: 4,
            RegisterValueType.UINT64: 4,
            RegisterValueType.FLOAT64: 4,
        }[self]

    @property
    def requires_integer(self) -> bool:
        return self in {
            RegisterValueType.INT16,
            RegisterValueType.UINT16,
            RegisterValueType.INT32,
            RegisterValueType.UINT32,
            RegisterValueType.INT64,
            RegisterValueType.UINT64,
        }


@dataclass(slots=True)
class ConnectionSettings:
    host: str = "127.0.0.1"
    port: int = 502
    slave_id: int = 1
    register_format: RegisterFormat = RegisterFormat.BIG
    keepalive_interval_seconds: int = 20


@dataclass(slots=True)
class RemoteConnectionSettings:
    enabled: bool = False
    auto_control: bool = True
    host: str = "127.0.0.1"
    port: int = 5025
    filename: str = ""


@dataclass(slots=True)
class ChannelConfig:
    name: str
    label: str = ""
    start_register: int = 0
    value_type: RegisterValueType = RegisterValueType.FLOAT32


@dataclass(slots=True)
class ChannelWrite:
    channel: ChannelConfig
    original_text: str
    value: int | float


@dataclass(slots=True)
class StageExecution:
    row_index: int
    stage_number: int
    duration_seconds: int
    writes: list[ChannelWrite] = field(default_factory=list)
    skipped_channels: list[str] = field(default_factory=list)


STATUS_DISCONNECTED = "disconnected"
STATUS_CONNECTED = "connected"
STATUS_RUNNING = "running"
STATUS_PAUSED = "paused"
STATUS_ERROR = "error"
STATUS_FINISHED = "finished"
