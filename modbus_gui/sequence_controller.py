from __future__ import annotations

from PyQt5.QtCore import QObject, QTimer, pyqtSignal

from .modbus_service import ModbusService, WriteResult
from .models import (
    StageExecution,
    STATUS_CONNECTED,
    STATUS_DISCONNECTED,
    STATUS_ERROR,
    STATUS_FINISHED,
    STATUS_PAUSED,
    STATUS_RUNNING,
)


class SequenceController(QObject):
    log_message = pyqtSignal(str)
    state_changed = pyqtSignal(str, str)
    stage_changed = pyqtSignal(int, int)
    timing_changed = pyqtSignal(int, int, int)
    finished = pyqtSignal()
    error_occurred = pyqtSignal(str)
    stage_write_completed = pyqtSignal(int, str)

    def __init__(self, modbus_service: ModbusService, parent: QObject | None = None) -> None:
        super().__init__(parent)
        self._modbus_service = modbus_service
        self._plan: list[StageExecution] = []
        self._current_index = -1
        self._current_stage_remaining = 0
        self._total_remaining = 0
        self._total_planned = 0
        self._is_paused = False
        self._timer = QTimer(self)
        self._timer.setInterval(1000)
        self._timer.timeout.connect(self._on_tick)

    @property
    def is_running(self) -> bool:
        return self._timer.isActive()

    @property
    def is_paused(self) -> bool:
        return self._is_paused and bool(self._plan)

    @property
    def has_active_plan(self) -> bool:
        return bool(self._plan)

    @property
    def current_stage_index(self) -> int:
        return self._current_index

    @property
    def total_planned(self) -> int:
        return self._total_planned

    @property
    def total_remaining(self) -> int:
        return self._total_remaining

    @property
    def current_stage_remaining(self) -> int:
        return self._current_stage_remaining

    @property
    def plan(self) -> list[StageExecution]:
        return list(self._plan)

    def start(self, plan: list[StageExecution]) -> None:
        if not self._modbus_service.is_connected:
            raise RuntimeError("Keine aktive Modbus-Verbindung")
        if not plan:
            self.log_message.emit("Teststart abgebrochen: keine aktive Stufe mit gueltiger Zeit gefunden.")
            raise RuntimeError("Es ist keine aktive Stufe mit gueltiger Zeit vorhanden")
        self.stop(reset_state=False)
        self._plan = plan
        self._current_index = 0
        self._current_stage_remaining = plan[0].duration_seconds
        self._total_planned = sum(stage.duration_seconds for stage in plan)
        self._total_remaining = self._total_planned
        self._is_paused = False
        self.log_message.emit(f"Testlauf gestartet: {len(plan)} aktive Stufen, Gesamtzeit {self._total_planned} s.")
        self.state_changed.emit(STATUS_RUNNING, "Test läuft")
        self._write_current_stage()
        self._emit_timing()
        self._timer.start()

    def pause(self) -> None:
        if not self._plan or not self._timer.isActive():
            return
        self._timer.stop()
        self._is_paused = True
        self.log_message.emit("Testlauf pausiert.")
        self.state_changed.emit(STATUS_PAUSED, "Pausiert")

    def resume(self) -> None:
        if not self._plan or not self._is_paused:
            return
        if not self._modbus_service.is_connected:
            raise RuntimeError("Keine aktive Modbus-Verbindung")
        self._is_paused = False
        self._timer.start()
        self.log_message.emit("Testlauf fortgesetzt.")
        self.state_changed.emit(STATUS_RUNNING, "Test läuft")
        self._emit_timing()

    def skip_current_stage(self) -> None:
        if not self._plan or not (0 <= self._current_index < len(self._plan)):
            return
        skipped_stage = self._plan[self._current_index]
        skipped_remaining = max(0, self._current_stage_remaining)
        if skipped_remaining > 0:
            self._total_remaining = max(0, self._total_remaining - skipped_remaining)
        was_paused = self._is_paused
        self._timer.stop()
        self.log_message.emit(
            f"Stufe {skipped_stage.stage_number} vorzeitig beendet. Springe zur naechsten Stufe."
        )
        self._advance_stage()
        if not self._plan or self._current_index < 0:
            return
        if was_paused:
            self._is_paused = True
            self.state_changed.emit(STATUS_PAUSED, "Pausiert")
        else:
            self._is_paused = False
            self._timer.start()
            self.state_changed.emit(STATUS_RUNNING, "Test läuft")

    def stop(self, reset_state: bool = True) -> None:
        self._timer.stop()
        self._plan = []
        self._current_index = -1
        self._current_stage_remaining = 0
        self._total_remaining = 0
        self._total_planned = 0
        self._is_paused = False
        if reset_state:
            status = STATUS_CONNECTED if self._modbus_service.is_connected else STATUS_DISCONNECTED
            text = "Bereit" if self._modbus_service.is_connected else "Getrennt"
            self.state_changed.emit(status, text)
            self.stage_changed.emit(-1, -1)
            self.timing_changed.emit(0, 0, 0)

    def _on_tick(self) -> None:
        try:
            self._handle_tick()
        except Exception as exc:
            self._handle_error(str(exc))

    def _handle_tick(self) -> None:
        if not self._plan or self._current_index < 0:
            self.stop()
            return
        if self._current_stage_remaining > 0:
            self._current_stage_remaining -= 1
        if self._total_remaining > 0:
            self._total_remaining -= 1
        if self._current_stage_remaining <= 0:
            self._advance_stage()
            return
        self._emit_timing()

    def _advance_stage(self) -> None:
        self._current_index += 1
        if self._current_index >= len(self._plan):
            self._timer.stop()
            self._is_paused = False
            self.log_message.emit("Testlauf abgeschlossen.")
            self.state_changed.emit(STATUS_FINISHED, "Fertig")
            self.stage_changed.emit(-1, -1)
            self.timing_changed.emit(self._total_planned, 0, 0)
            self.finished.emit()
            self._plan = []
            self._current_index = -1
            self._current_stage_remaining = 0
            return
        self._current_stage_remaining = self._plan[self._current_index].duration_seconds
        self._write_current_stage()
        self._emit_timing()

    def _write_current_stage(self) -> None:
        if not (0 <= self._current_index < len(self._plan)):
            raise RuntimeError("Keine aktive Stufe vorhanden")
        stage = self._plan[self._current_index]
        self.log_message.emit(f"Stufe {stage.stage_number} gestartet: Laufzeit {stage.duration_seconds} s.")
        self.stage_changed.emit(self._current_index, stage.row_index)
        for channel_name in stage.skipped_channels:
            self.log_message.emit(f"Stufe {stage.stage_number}, {channel_name}: kein Sollwert eingetragen, Kanal uebersprungen.")
        stage_errors: list[str] = []
        for write in stage.writes:
            try:
                result = self._modbus_service.write_channel(write.channel, write.value)
                self._log_write_result(stage.stage_number, result)
                self.stage_write_completed.emit(stage.stage_number, f"{write.channel.name} geschrieben")
            except Exception as exc:
                message = f"Stufe {stage.stage_number}, Kanal {write.channel.name}: Schreiben fehlgeschlagen. {exc}"
                self.log_message.emit(message)
                stage_errors.append(message)
        if stage_errors:
            raise RuntimeError("; ".join(stage_errors))

    def _handle_error(self, message: str) -> None:
        self._timer.stop()
        self._is_paused = False
        self.log_message.emit(f"Testlauf mit Fehler beendet: {message}")
        self.state_changed.emit(STATUS_ERROR, "Fehler")
        self.error_occurred.emit(message)
        self._plan = []
        self._current_index = -1
        self._current_stage_remaining = 0
        self._total_remaining = 0

    def _log_write_result(self, stage_number: int, result: WriteResult) -> None:
        self.log_message.emit(
            f"Stufe {stage_number}, {result.channel_name}: Wert {result.original_value} "
            f"als {result.value_type.label} auf Register {result.start_register} geschrieben "
            f"(Geraete-ID {result.slave_id}, Format {result.register_format.label}, "
            f"Registerwerte {result.registers}, Funktion {result.function_code})."
        )

    def _emit_timing(self) -> None:
        self.timing_changed.emit(self._total_planned, self._total_remaining, self._current_stage_remaining)
