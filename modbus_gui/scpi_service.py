from __future__ import annotations

import logging
import socket
from dataclasses import dataclass

from .models import ScpiSettings


logger = logging.getLogger(__name__)


@dataclass(slots=True)
class ScpiResult:
    command: str
    response_text: str = ""


class ScpiService:
    MAX_RESPONSE_BYTES = 65_536

    def __init__(self) -> None:
        self._socket: socket.socket | None = None
        self._settings: ScpiSettings | None = None

    @property
    def is_connected(self) -> bool:
        return self._socket is not None and self._settings is not None

    @property
    def settings(self) -> ScpiSettings | None:
        return self._settings

    def connect(self, settings: ScpiSettings) -> None:
        self.disconnect()
        try:
            client = socket.create_connection((settings.host, settings.port), timeout=settings.timeout_seconds)
            client.settimeout(settings.timeout_seconds)
        except OSError as exc:
            raise RuntimeError(f"SCPI-Verbindung zu {settings.host}:{settings.port} fehlgeschlagen: {exc}") from exc
        self._socket = client
        self._settings = settings
        logger.info(
            "scpi_connected host=%s port=%s timeout=%.2f",
            settings.host,
            settings.port,
            settings.timeout_seconds,
        )

    def disconnect(self) -> None:
        if self._socket is not None:
            try:
                self._socket.close()
            except OSError:
                logger.debug("scpi_socket_close_failed", exc_info=True)
        self._socket = None
        self._settings = None
        logger.info("scpi_disconnected")

    def write(self, command: str) -> ScpiResult:
        normalized_command = self._normalize_command(command)
        payload = self._build_payload(normalized_command)
        client = self._require_socket()
        try:
            client.sendall(payload)
        except OSError as exc:
            raise RuntimeError(f"SCPI-Kommando konnte nicht gesendet werden: {exc}") from exc
        logger.debug("scpi_write command=%s", normalized_command)
        return ScpiResult(command=normalized_command)

    def query(self, command: str) -> ScpiResult:
        result = self.write(command)
        response_text = self._read_response()
        logger.debug("scpi_query command=%s response=%s", result.command, response_text)
        return ScpiResult(command=result.command, response_text=response_text)

    def _build_payload(self, command: str) -> bytes:
        settings = self._require_settings()
        return f"{command}{settings.write_termination}".encode(settings.encoding)

    def _read_response(self) -> str:
        client = self._require_socket()
        settings = self._require_settings()
        terminator = settings.read_termination.encode(settings.encoding) if settings.read_termination else b""
        buffer = bytearray()
        while True:
            try:
                chunk = client.recv(4096)
            except socket.timeout as exc:
                if buffer:
                    break
                raise RuntimeError("SCPI-Antwort blieb aus") from exc
            except OSError as exc:
                raise RuntimeError(f"SCPI-Antwort konnte nicht gelesen werden: {exc}") from exc
            if not chunk:
                break
            buffer.extend(chunk)
            if len(buffer) > self.MAX_RESPONSE_BYTES:
                raise RuntimeError("SCPI-Antwort ist zu gross")
            if terminator and terminator in buffer:
                break
        if not buffer:
            raise RuntimeError("SCPI-Antwort blieb leer")
        payload = bytes(buffer)
        if terminator and terminator in payload:
            payload = payload.split(terminator, 1)[0]
        else:
            payload = payload.rstrip(b"\r\n")
        return payload.decode(settings.encoding, errors="replace")

    def _require_socket(self) -> socket.socket:
        if self._socket is None:
            raise RuntimeError("Keine aktive SCPI-Verbindung")
        return self._socket

    def _require_settings(self) -> ScpiSettings:
        if self._settings is None:
            raise RuntimeError("Keine aktiven SCPI-Einstellungen")
        return self._settings

    @staticmethod
    def _normalize_command(command: str) -> str:
        normalized = command.strip()
        if not normalized:
            raise RuntimeError("SCPI-Kommando darf nicht leer sein")
        return normalized
