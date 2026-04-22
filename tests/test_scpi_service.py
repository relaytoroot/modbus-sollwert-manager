from __future__ import annotations

import socket
import unittest
from unittest.mock import patch

from modbus_gui.models import ScpiSettings
from modbus_gui.scpi_service import ScpiService


class _FakeSocket:
    def __init__(self, responses: list[bytes] | None = None, timeout_on_recv: bool = False) -> None:
        self.responses = responses or []
        self.timeout_on_recv = timeout_on_recv
        self.sent_payloads: list[bytes] = []
        self.timeout_value: float | None = None
        self.closed = False

    def settimeout(self, value: float) -> None:
        self.timeout_value = value

    def sendall(self, payload: bytes) -> None:
        self.sent_payloads.append(payload)

    def recv(self, _size: int) -> bytes:
        if self.timeout_on_recv:
            raise socket.timeout()
        if self.responses:
            return self.responses.pop(0)
        return b""

    def close(self) -> None:
        self.closed = True


class ScpiServiceTests(unittest.TestCase):
    def test_query_reads_until_linefeed_and_strips_termination(self) -> None:
        fake_socket = _FakeSocket([b"FGH,PSU,", b"1234,1.0\n"])
        service = ScpiService()

        with patch("modbus_gui.scpi_service.socket.create_connection", return_value=fake_socket):
            service.connect(ScpiSettings(host="192.168.0.20", port=5025, timeout_seconds=1.5))
            result = service.query("*IDN?")

        self.assertEqual(fake_socket.timeout_value, 1.5)
        self.assertEqual(fake_socket.sent_payloads, [b"*IDN?\n"])
        self.assertEqual(result.command, "*IDN?")
        self.assertEqual(result.response_text, "FGH,PSU,1234,1.0")

    def test_query_raises_when_no_response_arrives(self) -> None:
        fake_socket = _FakeSocket(timeout_on_recv=True)
        service = ScpiService()

        with patch("modbus_gui.scpi_service.socket.create_connection", return_value=fake_socket):
            service.connect(ScpiSettings())
            with self.assertRaisesRegex(RuntimeError, "SCPI-Antwort blieb aus"):
                service.query("MEAS:VOLT?")


if __name__ == "__main__":
    unittest.main()
