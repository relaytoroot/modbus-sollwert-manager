from __future__ import annotations

import logging
from dataclasses import dataclass

from pymodbus.client import ModbusTcpClient

from .models import ChannelConfig, ConnectionSettings, RegisterFormat, RegisterValueType
from .value_encoder import ValueEncoder


logger = logging.getLogger(__name__)


@dataclass(slots=True)
class WriteResult:
    channel_name: str
    slave_id: int
    start_register: int
    value_type: RegisterValueType
    original_value: int | float
    register_format: RegisterFormat
    registers: list[int]
    function_code: str
    response_text: str


class ModbusService:
    def __init__(self) -> None:
        self._client: ModbusTcpClient | None = None
        self._settings: ConnectionSettings | None = None

    @property
    def is_connected(self) -> bool:
        return self._client is not None and self._settings is not None

    @property
    def settings(self) -> ConnectionSettings | None:
        return self._settings

    def connect(self, settings: ConnectionSettings) -> None:
        self.disconnect()
        client = ModbusTcpClient(settings.host, port=settings.port, timeout=3.0)
        if not client.connect():
            raise RuntimeError(f"Verbindung zu {settings.host}:{settings.port} fehlgeschlagen")
        self._client = client
        self._settings = settings
        logger.info("Verbunden host=%s port=%s slave_id=%s", settings.host, settings.port, settings.slave_id)

    def disconnect(self) -> None:
        if self._client is not None:
            self._client.close()
        self._client = None
        self._settings = None
        logger.info("Verbindung getrennt")

    def update_runtime_settings(self, settings: ConnectionSettings) -> None:
        if self.is_connected:
            self._settings = settings

    def write_channel(self, channel: ChannelConfig, value: int | float) -> WriteResult:
        client = self._require_client()
        settings = self._require_settings()
        start_register = self.normalize_address(channel.start_register)
        registers = ValueEncoder.encode_value(value, channel.value_type, settings.register_format)
        function_code = "FC06" if len(registers) == 1 else "FC16"
        logger.debug(
            "write_request host=%s port=%s slave_id=%s channel=%s start_register=%s type=%s format=%s value=%s registers=%s function=%s",
            settings.host,
            settings.port,
            settings.slave_id,
            channel.name,
            start_register,
            channel.value_type.value,
            settings.register_format.value,
            value,
            registers,
            function_code,
        )
        try:
            if len(registers) == 1:
                response = client.write_register(start_register, registers[0], device_id=settings.slave_id)
            else:
                response = client.write_registers(start_register, registers, device_id=settings.slave_id)
        except Exception as exc:
            logger.exception(
                "write_exception slave_id=%s channel=%s start_register=%s function=%s registers=%s",
                settings.slave_id,
                channel.name,
                start_register,
                function_code,
                registers,
            )
            raise RuntimeError(
                f"Schreibfehler {channel.name} @ {start_register} ({function_code}): {exc}"
            ) from exc
        logger.debug(
            "write_response slave_id=%s channel=%s start_register=%s function=%s response=%s",
            settings.slave_id,
            channel.name,
            start_register,
            function_code,
            response,
        )
        if response.isError():
            raise RuntimeError(
                f"Modbus-Fehler {channel.name} @ {start_register} ({function_code}): {response}"
            )
        return WriteResult(
            channel_name=channel.name,
            slave_id=settings.slave_id,
            start_register=start_register,
            value_type=channel.value_type,
            original_value=value,
            register_format=settings.register_format,
            registers=registers,
            function_code=function_code,
            response_text=str(response),
        )

    def keep_alive(self, address: int = 0, count: int = 1) -> str:
        client = self._require_client()
        settings = self._require_settings()
        start_register = self.normalize_address(address)
        logger.debug(
            "keepalive_request host=%s port=%s slave_id=%s start_register=%s count=%s function=FC03",
            settings.host,
            settings.port,
            settings.slave_id,
            start_register,
            count,
        )
        try:
            response = client.read_holding_registers(start_register, count=count, device_id=settings.slave_id)
        except Exception as exc:
            logger.exception(
                "keepalive_exception slave_id=%s start_register=%s count=%s",
                settings.slave_id,
                start_register,
                count,
            )
            raise RuntimeError(f"Keepalive fehlgeschlagen @ {start_register}: {exc}") from exc
        logger.debug(
            "keepalive_response slave_id=%s start_register=%s count=%s response=%s",
            settings.slave_id,
            start_register,
            count,
            response,
        )
        if response.isError():
            raise RuntimeError(f"Keepalive Modbus-Fehler @ {start_register}: {response}")
        registers = getattr(response, "registers", [])
        return f"Register {start_register}, Anzahl {count}, Rueckgabe {list(registers)}"

    @staticmethod
    def normalize_address(address: int) -> int:
        if address < 0:
            raise RuntimeError("Registeradresse darf nicht negativ sein")
        return address

    def _require_client(self) -> ModbusTcpClient:
        if self._client is None:
            raise RuntimeError("Keine aktive Modbus-Verbindung")
        return self._client

    def _require_settings(self) -> ConnectionSettings:
        if self._settings is None:
            raise RuntimeError("Keine aktiven Verbindungseinstellungen")
        return self._settings
