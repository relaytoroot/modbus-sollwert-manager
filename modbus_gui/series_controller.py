from __future__ import annotations

from dataclasses import dataclass
from typing import Callable

from PyQt5.QtCore import QObject, pyqtSignal

from .models import AutomationJob


@dataclass(slots=True)
class _AutomationRun:
    job: AutomationJob
    repeat_index: int
    repeat_total: int


class SeriesController(QObject):
    log_message = pyqtSignal(str)
    job_started = pyqtSignal(int, str, int, int, int, int)
    job_finished = pyqtSignal(int, str, int, int, int, int)
    progress_changed = pyqtSignal(int, int)
    finished = pyqtSignal()
    stopped = pyqtSignal()
    error_occurred = pyqtSignal(str)

    def __init__(self, start_job_callback: Callable[[AutomationJob], None], parent: QObject | None = None) -> None:
        super().__init__(parent)
        self._start_job_callback = start_job_callback
        self._runs: list[_AutomationRun] = []
        self._current_run_index = -1

    @property
    def is_running(self) -> bool:
        return bool(self._runs)

    @property
    def total_runs(self) -> int:
        return len(self._runs)

    @property
    def current_run_number(self) -> int:
        return self._current_run_index + 1 if self._current_run_index >= 0 else 0

    def start(self, jobs: list[AutomationJob]) -> None:
        if self.is_running:
            raise RuntimeError("Es laeuft bereits eine Automatisierungsserie")
        runs: list[_AutomationRun] = []
        for job in jobs:
            for repeat_index in range(1, max(1, job.repeat_count) + 1):
                runs.append(
                    _AutomationRun(
                        job=job,
                        repeat_index=repeat_index,
                        repeat_total=max(1, job.repeat_count),
                    )
                )
        if not runs:
            raise RuntimeError("Es ist kein aktiver Serienjob vorhanden")
        self._runs = runs
        self._current_run_index = -1
        self.log_message.emit(
            f"Serienlauf gestartet: {len(jobs)} Job(s), {len(runs)} geplante Durchgaenge."
        )
        self.progress_changed.emit(0, len(self._runs))
        self._start_next_run()

    def plan_finished(self) -> None:
        if not self.is_running or not (0 <= self._current_run_index < len(self._runs)):
            return
        run = self._runs[self._current_run_index]
        completed_runs = self._current_run_index + 1
        self.job_finished.emit(
            run.job.row_index,
            run.job.name,
            run.repeat_index,
            run.repeat_total,
            completed_runs,
            len(self._runs),
        )
        self.progress_changed.emit(completed_runs, len(self._runs))
        self._start_next_run()

    def plan_failed(self, message: str) -> None:
        if not self.is_running or not (0 <= self._current_run_index < len(self._runs)):
            return
        run = self._runs[self._current_run_index]
        full_message = (
            f"Serienjob {run.job.name} (Wiederholung {run.repeat_index}/{run.repeat_total}) fehlgeschlagen: {message}"
        )
        self.log_message.emit(full_message)
        self._clear()
        self.error_occurred.emit(full_message)

    def stop(self) -> None:
        if not self.is_running:
            return
        self.log_message.emit("Serienlauf manuell gestoppt.")
        self._clear()
        self.stopped.emit()

    def _start_next_run(self) -> None:
        self._current_run_index += 1
        if self._current_run_index >= len(self._runs):
            self.log_message.emit("Serienlauf abgeschlossen.")
            self._clear()
            self.finished.emit()
            return
        run = self._runs[self._current_run_index]
        current_run_number = self._current_run_index + 1
        self.job_started.emit(
            run.job.row_index,
            run.job.name,
            run.repeat_index,
            run.repeat_total,
            current_run_number,
            len(self._runs),
        )
        self.log_message.emit(
            f"Serienjob {current_run_number}/{len(self._runs)} gestartet: "
            f"{run.job.name} (Wiederholung {run.repeat_index}/{run.repeat_total})."
        )
        try:
            self._start_job_callback(run.job)
        except Exception as exc:
            self.plan_failed(str(exc))

    def _clear(self) -> None:
        self._runs = []
        self._current_run_index = -1
