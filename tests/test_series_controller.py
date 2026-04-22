from __future__ import annotations

import unittest
from pathlib import Path

from PyQt5.QtWidgets import QApplication

from modbus_gui.models import AutomationJob
from modbus_gui.series_controller import SeriesController


class SeriesControllerTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])

    def test_series_controller_runs_jobs_in_repeat_order(self) -> None:
        started_jobs: list[str] = []
        controller = SeriesController(lambda job: started_jobs.append(job.name))
        jobs = [
            AutomationJob(row_index=0, name="Plan A", source_path=Path("C:/Tests/A.xlsx"), repeat_count=1),
            AutomationJob(row_index=1, name="Plan B", source_path=Path("C:/Tests/B.xlsx"), repeat_count=2),
        ]

        controller.start(jobs)
        controller.plan_finished()
        controller.plan_finished()
        controller.plan_finished()

        self.assertEqual(started_jobs, ["Plan A", "Plan B", "Plan B"])
        self.assertFalse(controller.is_running)

    def test_series_controller_emits_error_when_job_cannot_start(self) -> None:
        errors: list[str] = []
        def failing_start(_job: AutomationJob) -> None:
            raise RuntimeError("Datei ungueltig")

        controller = SeriesController(failing_start)
        controller.error_occurred.connect(errors.append)

        controller.start([AutomationJob(row_index=0, name="Plan A", source_path=Path("C:/Tests/A.xlsx"))])

        self.assertFalse(controller.is_running)
        self.assertEqual(len(errors), 1)
        self.assertIn("Datei ungueltig", errors[0])


if __name__ == "__main__":
    unittest.main()
