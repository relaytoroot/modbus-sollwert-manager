from __future__ import annotations

import socket

from .models import RemoteConnectionSettings


class ScpiService:
    def __init__(self, timeout_seconds: float = 5.0) -> None:
        self._timeout_seconds = timeout_seconds

    def set_filename(self, settings: RemoteConnectionSettings) -> None:
        filename = settings.filename.strip()
        if not filename:
            raise RuntimeError("Bitte einen Dateinamen fuer die Remote-Messung eintragen.")
        sanitized_filename = filename.replace('"', "'")
        self.send_command(settings, f':STORe:FILE:NAME "{sanitized_filename}"')

    def start_measurement(self, settings: RemoteConnectionSettings) -> None:
        self.send_command(settings, ":STORe:STARt")

    def stop_measurement(self, settings: RemoteConnectionSettings) -> None:
        self.send_command(settings, ":STORe:STOP")

    def send_command(self, settings: RemoteConnectionSettings, command: str) -> None:
        if not settings.host.strip():
            raise RuntimeError("Bitte eine IP-Adresse fuer die Remote-Verbindung eintragen.")
        if settings.port <= 0:
            raise RuntimeError("Bitte einen gueltigen Remote-Port eintragen.")
        payload = f"{command}\n".encode("ascii", errors="strict")
        with socket.create_connection((settings.host.strip(), settings.port), timeout=self._timeout_seconds) as client:
            client.sendall(payload)
