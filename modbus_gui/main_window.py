from __future__ import annotations

import sys
from pathlib import Path

from openpyxl import Workbook, load_workbook
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QCloseEvent, QDoubleValidator, QColor, QIcon, QKeySequence, QPixmap
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFormLayout,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QSpinBox,
    QSplitter,
    QStyledItemDelegate,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from .modbus_service import ModbusService
from .models import (
    ChannelConfig,
    ChannelWrite,
    ConnectionSettings,
    RemoteConnectionSettings,
    RegisterFormat,
    RegisterValueType,
    StageExecution,
    STATUS_CONNECTED,
    STATUS_DISCONNECTED,
    STATUS_ERROR,
    STATUS_FINISHED,
    STATUS_PAUSED,
    STATUS_RUNNING,
)
from .app_info import (
    APP_AUTHOR,
    APP_COMPANY,
    APP_NAME,
    APP_VERSION,
    CHECKBOX_CHECK_DARK_THEME_FILE,
    CHECKBOX_CHECK_LIGHT_THEME_FILE,
    COMBO_ARROW_DARK_THEME_FILE,
    COMBO_ARROW_LIGHT_THEME_FILE,
    HEADER_LOGO_FILE,
    ICON_FILE,
    SPIN_ARROW_DOWN_DARK_THEME_FILE,
    SPIN_ARROW_DOWN_LIGHT_THEME_FILE,
    SPIN_ARROW_UP_DARK_THEME_FILE,
    SPIN_ARROW_UP_LIGHT_THEME_FILE,
    resource_path,
)
from .sequence_controller import SequenceController
from .scpi_service import ScpiService
from .value_encoder import ValueEncoder, ValueEncodingError


class NumericItemDelegate(QStyledItemDelegate):
    def createEditor(self, parent, option, index):  # type: ignore[override]
        editor = QLineEdit(parent)
        validator = QDoubleValidator(editor)
        validator.setNotation(QDoubleValidator.StandardNotation)
        validator.setDecimals(8)
        editor.setValidator(validator)
        editor.setAlignment(Qt.AlignCenter)
        return editor


class PlanTableWidget(QTableWidget):
    def keyPressEvent(self, event):  # type: ignore[override]
        if event.matches(QKeySequence.Copy):
            self._copy_selection()
            return
        if event.matches(QKeySequence.Paste):
            self._paste_selection()
            return
        if event.key() in {Qt.Key_Delete, Qt.Key_Backspace}:
            self._clear_selection()
            return
        super().keyPressEvent(event)

    def _copy_selection(self) -> None:
        ranges = self.selectedRanges()
        if not ranges:
            return
        selected = ranges[0]
        rows: list[str] = []
        for row in range(selected.topRow(), selected.bottomRow() + 1):
            values: list[str] = []
            for column in range(selected.leftColumn(), selected.rightColumn() + 1):
                item = self.item(row, column)
                values.append("" if item is None else item.text())
            rows.append("\t".join(values))
        QApplication.clipboard().setText("\n".join(rows))

    def _paste_selection(self) -> None:
        text = QApplication.clipboard().text()
        if not text:
            return
        start_row = max(0, self.currentRow())
        start_column = max(2, self.currentColumn())
        for row_offset, line in enumerate(text.splitlines()):
            for column_offset, value in enumerate(line.split("\t")):
                row = start_row + row_offset
                column = start_column + column_offset
                if row >= self.rowCount() or column >= self.columnCount() or column < 2:
                    continue
                item = self.item(row, column)
                if item is not None and (item.flags() & Qt.ItemIsEditable):
                    item.setText(value)

    def _clear_selection(self) -> None:
        for item in self.selectedItems():
            if item.column() >= 2 and (item.flags() & Qt.ItemIsEditable):
                item.setText("")


class ModbusMainWindow(QMainWindow):
    THEMES = {
        "light": {
            "window_bg": "#f4f7fa",
            "text_primary": "#1d2733",
            "text_secondary": "#607080",
            "group_bg": "#ffffff",
            "border": "#d4dde6",
            "button_bg": "#f6f8fb",
            "button_hover": "#eef3f7",
            "button_pressed": "#e1e9f1",
            "button_disabled_bg": "#eef2f5",
            "button_disabled_text": "#96a1ac",
            "input_bg": "#ffffff",
            "input_border": "#cad5df",
            "focus_ring": "#8fb2cf",
            "control_surface": "#edf2f7",
            "control_surface_hover": "#e2eaf2",
            "table_bg": "#ffffff",
            "table_alt_bg": "#f7f9fb",
            "table_grid": "#e4e9ef",
            "table_header_bg": "#eef3f7",
            "table_selection_bg": "#dfe9f4",
            "table_selection_text": "#13202c",
            "default_row_bg": "#ffffff",
            "inactive_row_bg": "#f3f5f7",
            "active_row_bg": "#e8f0f7",
            "active_row_text": "#13202c",
            "inactive_row_text": "#667585",
            "author_text": "#6b7885",
            "status_styles": {
                STATUS_DISCONNECTED: "background:#e3e8ed;color:#2f4354;",
                STATUS_CONNECTED: "background:#d9ebe1;color:#24543a;",
                STATUS_RUNNING: "background:#f3e5c8;color:#73511d;",
                STATUS_PAUSED: "background:#e7edf4;color:#37526b;",
                STATUS_ERROR: "background:#f3d9d7;color:#7d2e28;",
                STATUS_FINISHED: "background:#d8e7f4;color:#234d70;",
            },
        },
        "dark": {
            "window_bg": "#1b2128",
            "text_primary": "#f1f5f9",
            "text_secondary": "#b5c0cb",
            "group_bg": "#252d36",
            "border": "#516171",
            "button_bg": "#415161",
            "button_hover": "#4e6276",
            "button_pressed": "#5c748d",
            "button_disabled_bg": "#2b333d",
            "button_disabled_text": "#8b97a4",
            "input_bg": "#1f2730",
            "input_border": "#70859a",
            "focus_ring": "#97bddc",
            "control_surface": "#4b5d71",
            "control_surface_hover": "#5b7087",
            "table_bg": "#222932",
            "table_alt_bg": "#1d242c",
            "table_grid": "#3f4b57",
            "table_header_bg": "#33404d",
            "table_selection_bg": "#496a89",
            "table_selection_text": "#f5f8fb",
            "default_row_bg": "#252d36",
            "inactive_row_bg": "#202730",
            "active_row_bg": "#38516a",
            "active_row_text": "#f6f9fc",
            "inactive_row_text": "#9eacba",
            "author_text": "#9aa9b8",
            "status_styles": {
                STATUS_DISCONNECTED: "background:#465565;color:#edf3f8;",
                STATUS_CONNECTED: "background:#2f6448;color:#eef8f2;",
                STATUS_RUNNING: "background:#7b632a;color:#fff5dd;",
                STATUS_PAUSED: "background:#476279;color:#eef6fc;",
                STATUS_ERROR: "background:#7a3f3b;color:#ffefed;",
                STATUS_FINISHED: "background:#35607d;color:#edf7ff;",
            },
        },
    }
    DEFAULT_WINDOW_TITLE = APP_NAME
    COLUMN_ACTIVE = 0
    COLUMN_STAGE = 1
    COLUMN_VALUE_START = 2
    COLUMN_DURATION = 6
    CHANNEL_COUNT = 4

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(self.DEFAULT_WINDOW_TITLE)
        self.resize(1600, 980)
        self.logo_path = resource_path(ICON_FILE)
        self.header_logo_path = resource_path(HEADER_LOGO_FILE)

        self.modbus_service = ModbusService()
        self.scpi_service = ScpiService()
        self.sequence_controller = SequenceController(self.modbus_service, self)
        self.keepalive_timer = QTimer(self)
        self.keepalive_timer.timeout.connect(self._on_keepalive_tick)
        self.current_file_path: Path | None = None
        self._has_unsaved_changes = False
        self._remote_measurement_running = False
        self._remote_state_text = "Aus"
        self.row_checkboxes: list[QCheckBox] = []
        self.register_inputs: list[QSpinBox] = []
        self.channel_label_inputs: list[QLineEdit] = []
        self.type_inputs: list[QComboBox] = []
        self.start_value_inputs: list[QLineEdit] = []
        self.send_value_buttons: list[QPushButton] = []
        self.current_highlight_row = -1
        self._table_ui_update_in_progress = False
        self.current_theme = "light"
        self.current_status_key = STATUS_DISCONNECTED
        self.current_status_text = "Getrennt"
        self.default_row_color = QColor("#ffffff")
        self.inactive_row_color = QColor("#f1f3f5")
        self.active_stage_color = QColor("#d9ecff")
        self.active_row_text_color = QColor("#13202c")
        self.inactive_row_text_color = QColor("#667585")

        self.host_input = QLineEdit("127.0.0.1")
        self.port_input = QSpinBox()
        self.port_input.setRange(1, 65535)
        self.port_input.setValue(502)
        self.slave_id_input = QSpinBox()
        self.slave_id_input.setRange(1, 247)
        self.slave_id_input.setValue(1)
        self.keepalive_input = QSpinBox()
        self.keepalive_input.setRange(1, 3600)
        self.keepalive_input.setSuffix(" s")
        self.keepalive_input.setValue(20)

        self.register_format_combo = QComboBox()
        for register_format in RegisterFormat:
            self.register_format_combo.addItem(register_format.label, register_format)
        self.register_format_combo.setMaximumWidth(180)
        self.theme_combo = QComboBox()
        self.theme_combo.addItem("Hell", "light")
        self.theme_combo.addItem("Dunkel", "dark")
        self.theme_combo.setMaximumWidth(120)
        self.pt1_p_input = QSpinBox()
        self.pt1_p_input.setRange(0, 1_000_000)
        self.pt1_p_input.setSuffix(" ms")
        self.pt1_p_input.setMaximumWidth(92)
        self.pt1_p_register_input = QSpinBox()
        self.pt1_p_register_input.setRange(0, 65535)
        self.pt1_p_register_input.setMaximumWidth(84)
        self.pt1_q_input = QSpinBox()
        self.pt1_q_input.setRange(0, 1_000_000)
        self.pt1_q_input.setSuffix(" ms")
        self.pt1_q_input.setMaximumWidth(92)
        self.pt1_q_register_input = QSpinBox()
        self.pt1_q_register_input.setRange(0, 65535)
        self.pt1_q_register_input.setMaximumWidth(84)

        for index in range(4):
            register_input = QSpinBox()
            register_input.setRange(0, 65535)
            register_input.setValue((0, 1, 3, 5)[index])
            channel_label_input = QLineEdit()
            channel_label_input.setPlaceholderText("z. B. Wirkleistung")
            channel_label_input.setMaximumWidth(156)
            type_input = QComboBox()
            for value_type in RegisterValueType:
                type_input.addItem(value_type.label, value_type)
            type_input.setCurrentIndex(type_input.findData(RegisterValueType.FLOAT32))
            start_value_input = QLineEdit()
            start_value_input.setPlaceholderText("Startwert")
            start_value_input.setMaximumWidth(88)
            send_value_button = QPushButton("Send")
            send_value_button.setMaximumWidth(64)
            register_input.setMaximumWidth(84)
            type_input.setMaximumWidth(126)
            self.register_inputs.append(register_input)
            self.channel_label_inputs.append(channel_label_input)
            self.type_inputs.append(type_input)
            self.start_value_inputs.append(start_value_input)
            self.send_value_buttons.append(send_value_button)

        self.connect_button = QPushButton("Verbinden")
        self.disconnect_button = QPushButton("Trennen")
        self.start_button = QPushButton("Start")
        self.pause_button = QPushButton("Pause")
        self.next_stage_button = QPushButton("Naechste Stufe")
        self.stop_button = QPushButton("Stopp")
        self.save_button = QPushButton("Speichern unter")
        self.load_button = QPushButton("Laden")
        self.copy_stage_time_button = QPushButton("Zeit aus Stufe 1")
        self.copy_stage_time_button.setToolTip(
            "Uebernimmt die Zeit aus Stufe 1 in alle anderen aktiven Stufen."
        )
        self.copy_stage_time_button.setMaximumWidth(150)
        self.disconnect_button.setEnabled(False)
        self.pause_button.setEnabled(False)
        self.next_stage_button.setEnabled(False)
        self.stop_button.setEnabled(False)
        self.copy_stage_time_button.setEnabled(False)
        self.remote_enabled_checkbox = QCheckBox("Remote aktiv")
        self.remote_auto_checkbox = QCheckBox("Mit Test koppeln")
        self.remote_auto_checkbox.setChecked(True)
        self.remote_host_input = QLineEdit("127.0.0.1")
        self.remote_port_input = QSpinBox()
        self.remote_port_input.setRange(1, 65535)
        self.remote_port_input.setValue(5025)
        self.remote_port_input.setMaximumWidth(90)
        self.remote_filename_input = QLineEdit()
        self.remote_filename_input.setPlaceholderText("Hersteller_Produkt_Norm_YYYYMMDD_HHMMSS_001.dmd")
        self.remote_filename_input.setToolTip(
            "Beispiel: Bachmann_SPPC_EN_20260325_112003_001.dmd\n"
            "Die laufende Nummer am Ende sollte fuer jede Messdatei eindeutig sein."
        )
        self.remote_set_filename_button = QPushButton("Dateiname setzen")
        self.remote_start_button = QPushButton("Messung starten")
        self.remote_stop_button = QPushButton("Messung stoppen")
        self.remote_status_value = QLabel("Aus")
        self.remote_status_value.setAlignment(Qt.AlignCenter)
        self.remote_status_value.setMinimumWidth(96)
        self.remote_start_button.setMaximumWidth(150)
        self.remote_stop_button.setMaximumWidth(150)
        self.remote_set_filename_button.setMaximumWidth(150)

        self.logo_label = QLabel()
        self._configure_logo()
        self.author_label = QLabel(
            f"{APP_COMPANY} | {APP_NAME} | Version {APP_VERSION} | {APP_AUTHOR}"
        )
        self.author_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)

        self.status_badge = QLabel("Getrennt")
        self.status_badge.setAlignment(Qt.AlignCenter)
        self.status_badge.setMinimumWidth(120)
        self.current_stage_value = QLabel("-")
        self.stage_remaining_value = QLabel("00:00:00")
        self.total_remaining_value = QLabel("00:00:00")
        self.total_time_value = QLabel("00:00:00")
        self.last_write_value = QLabel("Noch keine Schreibaktion")
        self.address_info = QLabel("Registeradressen sind nullbasiert.")

        self.log_panel = QPlainTextEdit()
        self.log_panel.setReadOnly(True)
        self.log_panel.setMaximumBlockCount(1000)
        self.log_panel.setMinimumHeight(120)
        self.log_panel.setMaximumHeight(260)

        self.test_table = PlanTableWidget(40, 7)
        editable_columns = [
            *range(self.COLUMN_VALUE_START, self.COLUMN_VALUE_START + self.CHANNEL_COUNT),
            self.COLUMN_DURATION,
        ]
        for column in editable_columns:
            self.test_table.setItemDelegateForColumn(column, NumericItemDelegate(self.test_table))

        self._configure_table()
        self._build_ui()
        self._apply_styles()
        self._connect_signals()
        self._set_status(STATUS_DISCONNECTED, "Getrennt")
        self._apply_row_visuals()
        self._refresh_summary()
        self._update_manual_write_buttons()
        self._append_log("Anwendung gestartet. Bitte Verbindung pruefen und Testplan laden oder erstellen.")

    def _build_ui(self) -> None:
        central = QWidget(self)
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(18, 18, 18, 18)
        root.setSpacing(14)

        top_cards = QHBoxLayout()
        top_cards.addWidget(self._build_connection_group(), 3)
        top_cards.addWidget(self._build_format_group(), 2)
        top_cards.addWidget(self._build_control_group(), 3)
        root.addLayout(top_cards)
        root.addWidget(self._build_channel_group())
        root.addWidget(self._build_remote_group())

        self.main_splitter = QSplitter(Qt.Vertical)
        self.main_splitter.addWidget(self._build_table_group())
        self.main_splitter.addWidget(self._build_log_group())
        self.main_splitter.setChildrenCollapsible(False)
        self.main_splitter.setStretchFactor(0, 5)
        self.main_splitter.setStretchFactor(1, 1)
        self.main_splitter.setSizes([700, 180])
        root.addWidget(self.main_splitter, 1)
        root.addWidget(self._build_footer())

    def _build_connection_group(self) -> QGroupBox:
        group = QGroupBox("Verbindung")
        layout = QGridLayout(group)

        logo_col = QVBoxLayout()
        logo_col.addWidget(self.logo_label, 0, Qt.AlignLeft | Qt.AlignTop)
        logo_col.addStretch(1)
        layout.addLayout(logo_col, 0, 0, 2, 1)

        form = QFormLayout()
        form.addRow("IP-Adresse", self.host_input)
        form.addRow("Port", self.port_input)
        form.addRow("Slave-ID", self.slave_id_input)
        form.addRow("Keepalive", self.keepalive_input)
        layout.addLayout(form, 0, 1, 2, 1)

        button_col = QVBoxLayout()
        button_col.addWidget(self.connect_button)
        button_col.addWidget(self.disconnect_button)
        button_col.addStretch(1)
        layout.addLayout(button_col, 0, 2)

        status_col = QVBoxLayout()
        status_col.addWidget(QLabel("Status"))
        status_col.addWidget(self.status_badge)
        status_col.addWidget(self.address_info)
        status_col.addStretch(1)
        layout.addLayout(status_col, 0, 3)
        layout.setColumnStretch(1, 1)
        return group

    def _build_format_group(self) -> QGroupBox:
        group = QGroupBox("Datenformat")
        form = QFormLayout(group)
        form.setLabelAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        form.setFormAlignment(Qt.AlignTop)
        form.setHorizontalSpacing(10)
        group.setMaximumWidth(250)
        form.addRow("Registerformat", self.register_format_combo)
        form.addRow("Darstellung", self.theme_combo)
        return group

    def _build_control_group(self) -> QGroupBox:
        group = QGroupBox("Ablauf")
        layout = QGridLayout(group)
        layout.addWidget(self.start_button, 0, 0)
        layout.addWidget(self.pause_button, 0, 1)
        layout.addWidget(self.next_stage_button, 0, 2)
        layout.addWidget(self.stop_button, 0, 3)
        layout.addWidget(self.save_button, 1, 0)
        layout.addWidget(self.load_button, 1, 1)
        layout.addWidget(QLabel("Aktuelle Stufe"), 0, 4)
        layout.addWidget(self.current_stage_value, 0, 5)
        layout.addWidget(QLabel("Restzeit Stufe"), 1, 4)
        layout.addWidget(self.stage_remaining_value, 1, 5)
        layout.addWidget(QLabel("Restzeit gesamt"), 0, 6)
        layout.addWidget(self.total_remaining_value, 0, 7)
        layout.addWidget(QLabel("Gesamtzeit"), 1, 6)
        layout.addWidget(self.total_time_value, 1, 7)
        layout.addWidget(QLabel("Letzte Rueckmeldung"), 2, 0, 1, 4)
        layout.addWidget(self.last_write_value, 2, 4, 1, 4)
        return group

    def _build_channel_group(self) -> QGroupBox:
        group = QGroupBox("Kanalzuordnung")
        group.setMaximumHeight(190)
        outer_layout = QHBoxLayout(group)
        outer_layout.setContentsMargins(12, 12, 12, 12)
        outer_layout.setSpacing(12)

        mapping_widget = QWidget()
        mapping_widget.setMaximumWidth(600)
        layout = QGridLayout(mapping_widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setHorizontalSpacing(6)
        layout.setVerticalSpacing(4)
        layout.addWidget(QLabel("Kanal"), 0, 0)
        layout.addWidget(QLabel("Bezeichnung"), 0, 1)
        layout.addWidget(QLabel("Startregister"), 0, 2)
        layout.addWidget(QLabel("Datentyp"), 0, 3)
        layout.addWidget(QLabel("Startwert"), 0, 4)
        layout.addWidget(QLabel("Aktion"), 0, 5)
        for index, (channel_label_input, register_input, type_input, start_value_input, send_value_button) in enumerate(
            zip(self.channel_label_inputs, self.register_inputs, self.type_inputs, self.start_value_inputs, self.send_value_buttons),
            start=1,
        ):
            layout.addWidget(QLabel(f"S{index}"), index, 0)
            layout.addWidget(channel_label_input, index, 1)
            layout.addWidget(register_input, index, 2)
            layout.addWidget(type_input, index, 3)
            layout.addWidget(start_value_input, index, 4)
            layout.addWidget(send_value_button, index, 5)

        pt1_widget = QWidget()
        pt1_widget.setMaximumWidth(165)
        pt1_layout = QFormLayout(pt1_widget)
        pt1_layout.setContentsMargins(0, 0, 0, 0)
        pt1_layout.setLabelAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        pt1_layout.setHorizontalSpacing(6)
        pt1_layout.setVerticalSpacing(6)
        pt1_layout.addRow("PT1 P [ms]", self.pt1_p_input)
        pt1_layout.addRow("PT1 P Reg.", self.pt1_p_register_input)
        pt1_layout.addRow("PT1 Q [ms]", self.pt1_q_input)
        pt1_layout.addRow("PT1 Q Reg.", self.pt1_q_register_input)

        outer_layout.addWidget(mapping_widget, 0, Qt.AlignLeft | Qt.AlignTop)
        outer_layout.addWidget(pt1_widget, 0, Qt.AlignTop)
        outer_layout.addStretch(1)
        return group

    def _build_table_group(self) -> QGroupBox:
        group = QGroupBox("Testplan")
        layout = QVBoxLayout(group)
        helper_row = QHBoxLayout()
        helper_row.setContentsMargins(0, 0, 0, 0)
        helper_row.addStretch(1)
        helper_row.addWidget(self.copy_stage_time_button, 0, Qt.AlignRight)
        layout.addLayout(helper_row)
        layout.addWidget(self.test_table)
        return group

    def _build_remote_group(self) -> QGroupBox:
        group = QGroupBox("Remote Messung (SCPI)")
        group.setMaximumHeight(120)
        layout = QGridLayout(group)
        layout.setHorizontalSpacing(10)
        layout.setVerticalSpacing(8)
        layout.addWidget(self.remote_enabled_checkbox, 0, 0)
        layout.addWidget(QLabel("IP-Adresse"), 0, 1)
        layout.addWidget(self.remote_host_input, 0, 2)
        layout.addWidget(QLabel("Port"), 0, 3)
        layout.addWidget(self.remote_port_input, 0, 4)
        layout.addWidget(self.remote_auto_checkbox, 0, 5)
        layout.addWidget(QLabel("Status"), 0, 6)
        layout.addWidget(self.remote_status_value, 0, 7)
        layout.addWidget(QLabel("Messdatei"), 1, 0)
        layout.addWidget(self.remote_filename_input, 1, 1, 1, 4)
        layout.addWidget(self.remote_set_filename_button, 1, 5)
        layout.addWidget(self.remote_start_button, 1, 6)
        layout.addWidget(self.remote_stop_button, 1, 7)
        layout.setColumnStretch(2, 1)
        return group

    def _build_log_group(self) -> QGroupBox:
        group = QGroupBox("Ereignisprotokoll")
        layout = QVBoxLayout(group)
        layout.addWidget(self.log_panel)
        return group

    def _build_footer(self) -> QWidget:
        footer = QWidget()
        layout = QHBoxLayout(footer)
        layout.setContentsMargins(2, 0, 2, 0)
        layout.addStretch(1)
        layout.addWidget(self.author_label, 0, Qt.AlignRight | Qt.AlignVCenter)
        return footer

    def _configure_logo(self) -> None:
        if self.logo_path.exists():
            pixmap = QPixmap(str(self.logo_path)).scaled(
                168,
                64,
                Qt.KeepAspectRatio,
                Qt.SmoothTransformation,
            )
            self.logo_label.setPixmap(pixmap)
            self.logo_label.setMinimumSize(168, 64)
            self.logo_label.setMaximumSize(168, 64)
            self.logo_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            window_icon_path = self.header_logo_path if self.header_logo_path.exists() else self.logo_path
            self.setWindowIcon(QIcon(str(window_icon_path)))
        else:
            self.logo_label.setText("")

    def _apply_styles(self) -> None:
        self.status_badge.setObjectName("statusBadge")
        self.remote_status_value.setObjectName("remoteStatusBadge")
        self.author_label.setObjectName("authorLabel")
        self.copy_stage_time_button.setObjectName("secondaryButton")
        self.next_stage_button.setObjectName("secondaryButton")
        self.remote_set_filename_button.setObjectName("secondaryButton")
        self._apply_theme(self.current_theme)

    def _apply_theme(self, theme_name: str) -> None:
        theme = self.THEMES.get(theme_name, self.THEMES["light"])
        self.current_theme = theme_name if theme_name in self.THEMES else "light"
        self.default_row_color = QColor(theme["default_row_bg"])
        self.inactive_row_color = QColor(theme["inactive_row_bg"])
        self.active_stage_color = QColor(theme["active_row_bg"])
        self.active_row_text_color = QColor(theme["active_row_text"])
        self.inactive_row_text_color = QColor(theme["inactive_row_text"])
        combo_arrow = resource_path(
            COMBO_ARROW_DARK_THEME_FILE if self.current_theme == "dark" else COMBO_ARROW_LIGHT_THEME_FILE
        ).as_posix()
        spin_up_arrow = resource_path(
            SPIN_ARROW_UP_DARK_THEME_FILE if self.current_theme == "dark" else SPIN_ARROW_UP_LIGHT_THEME_FILE
        ).as_posix()
        spin_down_arrow = resource_path(
            SPIN_ARROW_DOWN_DARK_THEME_FILE if self.current_theme == "dark" else SPIN_ARROW_DOWN_LIGHT_THEME_FILE
        ).as_posix()
        checkbox_check = resource_path(
            CHECKBOX_CHECK_DARK_THEME_FILE if self.current_theme == "dark" else CHECKBOX_CHECK_LIGHT_THEME_FILE
        ).as_posix()

        self.setStyleSheet(
            f"""
            QWidget {{
                font-size: 13px;
                color: {theme["text_primary"]};
                background: {theme["window_bg"]};
            }}
            QMainWindow {{
                background: {theme["window_bg"]};
            }}
            QGroupBox {{
                border: 1px solid {theme["border"]};
                border-radius: 10px;
                margin-top: 10px;
                padding-top: 12px;
                background: {theme["group_bg"]};
                font-weight: 600;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 6px;
            }}
            QLabel {{
                color: {theme["text_primary"]};
            }}
            QPushButton {{
                min-height: 34px;
                border-radius: 8px;
                border: 1px solid {theme["input_border"]};
                padding: 0 14px;
                background: {theme["button_bg"]};
                color: {theme["text_primary"]};
            }}
            QPushButton:hover {{
                background: {theme["button_hover"]};
            }}
            QPushButton:pressed {{
                background: {theme["button_pressed"]};
            }}
            QPushButton:disabled {{
                color: {theme["button_disabled_text"]};
                background: {theme["button_disabled_bg"]};
            }}
            QLineEdit, QComboBox, QSpinBox, QPlainTextEdit {{
                border: 1px solid {theme["input_border"]};
                border-radius: 8px;
                padding: 6px 8px;
                background: {theme["input_bg"]};
                color: {theme["text_primary"]};
                selection-background-color: {theme["table_selection_bg"]};
                selection-color: {theme["table_selection_text"]};
            }}
            QLineEdit:focus, QComboBox:focus, QSpinBox:focus, QPlainTextEdit:focus, QPushButton:focus {{
                border: 1px solid {theme["focus_ring"]};
            }}
            QComboBox {{
                padding-right: 30px;
            }}
            QComboBox::drop-down {{
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 26px;
                border-left: 1px solid {theme["input_border"]};
                border-top-right-radius: 8px;
                border-bottom-right-radius: 8px;
                background: {theme["control_surface"]};
            }}
            QComboBox::drop-down:hover {{
                background: {theme["control_surface_hover"]};
            }}
            QComboBox::down-arrow {{
                image: url("{combo_arrow}");
                width: 14px;
                height: 14px;
            }}
            QComboBox QAbstractItemView {{
                border: 1px solid {theme["input_border"]};
                background: {theme["input_bg"]};
                color: {theme["text_primary"]};
                selection-background-color: {theme["table_selection_bg"]};
                selection-color: {theme["table_selection_text"]};
            }}
            QSpinBox {{
                padding-right: 34px;
            }}
            QSpinBox::up-button, QSpinBox::down-button {{
                width: 20px;
                background: {theme["control_surface"]};
                border-left: 1px solid {theme["input_border"]};
            }}
            QSpinBox::up-button {{
                subcontrol-origin: border;
                subcontrol-position: top right;
                border-top-right-radius: 8px;
                border-bottom: 1px solid {theme["input_border"]};
            }}
            QSpinBox::down-button {{
                subcontrol-origin: border;
                subcontrol-position: bottom right;
                border-bottom-right-radius: 8px;
            }}
            QSpinBox::up-button:hover, QSpinBox::down-button:hover {{
                background: {theme["control_surface_hover"]};
            }}
            QSpinBox::up-arrow, QSpinBox::down-arrow {{
                width: 12px;
                height: 12px;
            }}
            QSpinBox::up-arrow {{
                image: url("{spin_up_arrow}");
            }}
            QSpinBox::down-arrow {{
                image: url("{spin_down_arrow}");
            }}
            QCheckBox {{
                spacing: 8px;
                color: {theme["text_primary"]};
                background: transparent;
            }}
            QCheckBox::indicator {{
                width: 16px;
                height: 16px;
                border-radius: 4px;
                border: 1px solid {theme["input_border"]};
                background: {theme["input_bg"]};
            }}
            QCheckBox::indicator:hover {{
                background: {theme["control_surface"]};
                border: 1px solid {theme["focus_ring"]};
            }}
            QCheckBox::indicator:checked {{
                background: {theme["button_pressed"]};
                border: 1px solid {theme["focus_ring"]};
                image: url("{checkbox_check}");
            }}
            QSplitter::handle {{
                background: {theme["border"]};
            }}
            QTableWidget {{
                border: 1px solid {theme["border"]};
                background: {theme["table_bg"]};
                color: {theme["text_primary"]};
                gridline-color: {theme["table_grid"]};
                alternate-background-color: {theme["table_alt_bg"]};
                selection-background-color: {theme["table_selection_bg"]};
                selection-color: {theme["table_selection_text"]};
            }}
            QTableWidget::item:selected {{
                background: {theme["table_selection_bg"]};
                color: {theme["table_selection_text"]};
            }}
            QTableCornerButton::section {{
                background: {theme["table_header_bg"]};
                border: none;
                border-bottom: 1px solid {theme["border"]};
                border-right: 1px solid {theme["border"]};
            }}
            QHeaderView::section {{
                background: {theme["table_header_bg"]};
                color: {theme["text_primary"]};
                border: none;
                border-bottom: 1px solid {theme["border"]};
                padding: 8px;
                font-weight: 600;
            }}
            QLabel#statusBadge {{
                border-radius: 12px;
                padding: 8px 12px;
                font-weight: 700;
            }}
            QLabel#remoteStatusBadge {{
                border-radius: 10px;
                padding: 6px 10px;
                font-weight: 600;
                border: 1px solid {theme["input_border"]};
                background: {theme["control_surface"]};
                color: {theme["text_primary"]};
            }}
            QPushButton#secondaryButton {{
                padding: 0 10px;
            }}
            QLabel#authorLabel {{
                color: {theme["author_text"]};
                font-size: 9px;
                font-weight: 400;
                padding: 1px 4px;
            }}
            """
        )
        self._set_status(self.current_status_key, self.current_status_text)
        self._update_remote_controls()
        self._apply_row_visuals()

    def _connect_signals(self) -> None:
        self.connect_button.clicked.connect(self._on_connect_clicked)
        self.disconnect_button.clicked.connect(self._on_disconnect_clicked)
        self.start_button.clicked.connect(self._on_start_clicked)
        self.pause_button.clicked.connect(self._on_pause_clicked)
        self.next_stage_button.clicked.connect(self._on_next_stage_clicked)
        self.stop_button.clicked.connect(self._on_stop_clicked)
        self.save_button.clicked.connect(self._on_save_clicked)
        self.load_button.clicked.connect(self._on_load_clicked)
        self.copy_stage_time_button.clicked.connect(self._copy_first_stage_time_to_active_rows)
        self.remote_enabled_checkbox.toggled.connect(self._on_remote_settings_changed)
        self.remote_auto_checkbox.toggled.connect(self._on_remote_settings_changed)
        self.remote_set_filename_button.clicked.connect(self._on_remote_set_filename_clicked)
        self.remote_start_button.clicked.connect(self._on_remote_start_clicked)
        self.remote_stop_button.clicked.connect(self._on_remote_stop_clicked)
        self.theme_combo.currentIndexChanged.connect(self._on_theme_changed)
        self.keepalive_input.valueChanged.connect(self._update_keepalive_timer_interval)
        self.test_table.itemChanged.connect(self._on_table_item_changed)
        self.sequence_controller.log_message.connect(self._append_log)
        self.sequence_controller.state_changed.connect(self._set_status)
        self.sequence_controller.stage_changed.connect(self._on_stage_changed)
        self.sequence_controller.timing_changed.connect(self._on_timing_changed)
        self.sequence_controller.finished.connect(self._on_sequence_finished)
        self.sequence_controller.error_occurred.connect(self._on_sequence_error)
        self.sequence_controller.stage_write_completed.connect(self._on_stage_write_completed)
        for checkbox in self.row_checkboxes:
            checkbox.toggled.connect(self._apply_row_visuals)
            checkbox.toggled.connect(self._refresh_summary)
            checkbox.toggled.connect(self._update_time_copy_button)
            checkbox.toggled.connect(self._mark_dirty)
        for input_field in self.channel_label_inputs:
            input_field.textChanged.connect(self._on_channel_label_changed)
        for input_field in self.start_value_inputs:
            input_field.textChanged.connect(self._mark_dirty)
        for send_value_button in self.send_value_buttons:
            send_value_button.clicked.connect(self._on_send_start_value_clicked)
        self.pt1_p_input.valueChanged.connect(self._mark_dirty)
        self.pt1_q_input.valueChanged.connect(self._mark_dirty)
        self.pt1_p_register_input.valueChanged.connect(self._mark_dirty)
        self.pt1_q_register_input.valueChanged.connect(self._mark_dirty)
        self.remote_host_input.textChanged.connect(self._on_remote_settings_changed)
        self.remote_port_input.valueChanged.connect(self._on_remote_settings_changed)
        self.remote_filename_input.textChanged.connect(self._on_remote_settings_changed)

    def _configure_table(self) -> None:
        self._update_table_headers()
        self.test_table.setAlternatingRowColors(True)
        self.test_table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.test_table.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.test_table.setSortingEnabled(False)
        self.test_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.test_table.verticalHeader().setDefaultSectionSize(30)
        self.test_table.blockSignals(True)
        for row in range(40):
            checkbox = QCheckBox()
            checkbox.setChecked(False)
            self.row_checkboxes.append(checkbox)
            self.test_table.setCellWidget(row, self.COLUMN_ACTIVE, self._wrap_widget(checkbox))

            stage_item = QTableWidgetItem(str(row + 1))
            stage_item.setFlags(stage_item.flags() & ~Qt.ItemIsEditable)
            stage_item.setTextAlignment(Qt.AlignCenter)
            self.test_table.setItem(row, self.COLUMN_STAGE, stage_item)

            for column in range(self.COLUMN_VALUE_START, self.COLUMN_DURATION + 1):
                item = QTableWidgetItem("")
                item.setTextAlignment(Qt.AlignCenter)
                self.test_table.setItem(row, column, item)
        self.test_table.blockSignals(False)

    def _wrap_widget(self, widget: QWidget) -> QWidget:
        container = QWidget()
        layout = QHBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setAlignment(Qt.AlignCenter)
        layout.addWidget(widget)
        return container

    def _current_settings(self) -> ConnectionSettings:
        return ConnectionSettings(
            host=self.host_input.text().strip() or "127.0.0.1",
            port=self.port_input.value(),
            slave_id=self.slave_id_input.value(),
            register_format=self.register_format_combo.currentData(),
            keepalive_interval_seconds=self.keepalive_input.value(),
        )

    def _current_remote_settings(self) -> RemoteConnectionSettings:
        return RemoteConnectionSettings(
            enabled=self.remote_enabled_checkbox.isChecked(),
            auto_control=self.remote_auto_checkbox.isChecked(),
            host=self.remote_host_input.text().strip() or "127.0.0.1",
            port=self.remote_port_input.value(),
            filename=self.remote_filename_input.text().strip(),
        )

    def _channel_configs(self) -> list[ChannelConfig]:
        return [
            ChannelConfig(
                name=f"S{index}",
                label=channel_label_input.text().strip(),
                start_register=register_input.value(),
                value_type=type_input.currentData(),
            )
            for index, (channel_label_input, register_input, type_input) in enumerate(
                zip(self.channel_label_inputs, self.register_inputs, self.type_inputs),
                start=1,
            )
        ]

    def _keepalive_address(self) -> int:
        return min(spin_box.value() for spin_box in self.register_inputs) if self.register_inputs else 0

    def _start_keepalive_timer(self) -> None:
        self.keepalive_timer.start(max(1, self.keepalive_input.value()) * 1000)
        self._append_log(
            f"Keepalive aktiv: alle {self.keepalive_input.value()} s, Register {self._keepalive_address()}."
        )

    def _stop_keepalive_timer(self) -> None:
        self.keepalive_timer.stop()

    def _update_keepalive_timer_interval(self) -> None:
        if self.keepalive_timer.isActive():
            self.keepalive_timer.start(max(1, self.keepalive_input.value()) * 1000)
            self._append_log(f"Keepalive aktualisiert: Intervall {self.keepalive_input.value()} s.")

    def _on_keepalive_tick(self) -> None:
        if not self.modbus_service.is_connected:
            self.keepalive_timer.stop()
            return
        try:
            response_text = self.modbus_service.keep_alive(address=self._keepalive_address(), count=1)
        except Exception as exc:
            self._append_log(f"Keepalive fehlgeschlagen: {exc}")
            self.sequence_controller.stop()
            self._set_running_widgets(False)
            self.modbus_service.disconnect()
            self._set_connection_widgets(False)
            self._set_status(STATUS_ERROR, "Fehler")
            self._show_error("Verbindung abgebrochen", str(exc))
            return
        self._append_log(f"Keepalive erfolgreich: {response_text}")

    def _build_stage_plan(self) -> list[StageExecution]:
        plan: list[StageExecution] = []
        channel_configs = self._channel_configs()
        for row in range(self.test_table.rowCount()):
            if not self.row_checkboxes[row].isChecked():
                continue
            duration = self._parse_duration(row)
            if duration <= 0:
                continue
            writes: list[ChannelWrite] = []
            skipped_channels: list[str] = []
            for channel_index, channel in enumerate(channel_configs):
                item = self.test_table.item(row, self.COLUMN_VALUE_START + channel_index)
                text = "" if item is None else item.text().strip()
                if not text:
                    skipped_channels.append(channel.name)
                    continue
                try:
                    value = ValueEncoder.coerce_text_value(text, channel.value_type)
                except ValueEncodingError as exc:
                    raise RuntimeError(f"Stufe {row + 1} {channel.name}: {exc}") from exc
                writes.append(
                    ChannelWrite(
                        channel=ChannelConfig(
                            name=channel.name,
                            start_register=channel.start_register,
                            value_type=channel.value_type,
                        ),
                        original_text=text,
                        value=value,
                    )
                )
            if not writes:
                continue
            plan.append(StageExecution(row_index=row, stage_number=row + 1, duration_seconds=duration, writes=writes, skipped_channels=skipped_channels))
        return plan

    def _parse_duration(self, row: int) -> int:
        item = self.test_table.item(row, self.COLUMN_DURATION)
        text = "" if item is None else item.text().strip().replace(",", ".")
        if not text:
            return 0
        try:
            value = float(text)
        except ValueError as exc:
            raise RuntimeError(f"Stufe {row + 1}: Zeit ist keine Zahl") from exc
        if value < 0:
            raise RuntimeError(f"Stufe {row + 1}: Zeit darf nicht negativ sein")
        return int(round(value))

    def _on_connect_clicked(self) -> None:
        settings = self._current_settings()
        while True:
            try:
                self.modbus_service.connect(settings)
            except Exception as exc:
                self.modbus_service.disconnect()
                self._set_connection_widgets(False)
                self._set_status(STATUS_ERROR, "Fehler")
                self._append_log(
                    f"Verbindung fehlgeschlagen: {settings.host}:{settings.port}, Geraete-ID {settings.slave_id}. {exc}"
                )
                if self._show_retry_dialog("Verbindung fehlgeschlagen", str(exc)):
                    continue
                return
            break
        self._set_status(STATUS_CONNECTED, "Verbunden")
        self._append_log(
            "Verbindung hergestellt: "
            f"{settings.host}:{settings.port}, Geraete-ID {settings.slave_id}, "
            f"Format {settings.register_format.label}, Keepalive {settings.keepalive_interval_seconds} s."
        )
        self._set_connection_widgets(True)
        self._start_keepalive_timer()

    def _on_disconnect_clicked(self) -> None:
        self.sequence_controller.stop()
        self._stop_remote_measurement("Remote-Messung bei Trennen gestoppt.", show_dialog_on_error=False)
        self._stop_keepalive_timer()
        self.modbus_service.disconnect()
        self._set_connection_widgets(False)
        self._set_status(STATUS_DISCONNECTED, "Getrennt")
        self._append_log("Verbindung getrennt")
        self.last_write_value.setText("Noch keine Schreibaktion")

    def _on_start_clicked(self) -> None:
        if not self.modbus_service.is_connected:
            self._show_error("Keine Verbindung", "Bitte zuerst verbinden.")
            return
        try:
            plan = self._build_stage_plan()
        except RuntimeError as exc:
            self._show_error("Ungültiger Testplan", str(exc))
            return
        if not plan:
            self._append_log("Teststart abgebrochen: keine aktive Stufe mit Zeit und Sollwerten vorhanden.")
            self._show_error("Kein Testplan", "Mindestens eine aktive Stufe mit Zeit > 0 und mindestens einem Sollwert ist erforderlich.")
            return
        self.modbus_service.update_runtime_settings(self._current_settings())
        self._update_keepalive_timer_interval()
        remote_started = False
        remote_settings = self._current_remote_settings()
        if remote_settings.enabled and remote_settings.auto_control:
            try:
                self._start_remote_measurement(remote_settings, source_label="Teststart")
                remote_started = True
            except Exception as exc:
                self._show_error("Remote-Messung fehlgeschlagen", str(exc))
                return
        try:
            self.sequence_controller.start(plan)
        except Exception as exc:
            if remote_started:
                self._stop_remote_measurement("Remote-Messung wegen Teststart-Fehler gestoppt.", show_dialog_on_error=False)
            self._set_status(STATUS_ERROR, "Fehler")
            self._show_error("Teststart fehlgeschlagen", str(exc))
            return
        self._set_running_widgets(True)
        self._update_pause_button()

    def _on_pause_clicked(self) -> None:
        if not self.sequence_controller.has_active_plan:
            return
        try:
            if self.sequence_controller.is_paused:
                self.sequence_controller.resume()
                self.last_write_value.setText("Testlauf fortgesetzt")
            else:
                self.sequence_controller.pause()
                self.last_write_value.setText("Testlauf pausiert")
        except Exception as exc:
            self._set_status(STATUS_ERROR, "Fehler")
            self._show_error("Pause/Fortsetzen fehlgeschlagen", str(exc))
            return
        self._set_running_widgets(True)
        self._update_pause_button()

    def _on_stop_clicked(self) -> None:
        self.sequence_controller.stop()
        self._stop_remote_measurement("Remote-Messung manuell gestoppt.", show_dialog_on_error=False)
        self._set_running_widgets(False)
        if self.modbus_service.is_connected:
            self._set_status(STATUS_CONNECTED, "Bereit")
        else:
            self._set_status(STATUS_DISCONNECTED, "Getrennt")
        self.last_write_value.setText("Test gestoppt")
        self._append_log("Testlauf manuell gestoppt.")
        self._highlight_current_row(-1)
        self._refresh_summary()

    def _on_save_clicked(self) -> None:
        self._save_as()

    def _on_load_clicked(self) -> None:
        if not self._confirm_discard_unsaved_changes("Vor dem Laden speichern?", "Es gibt ungespeicherte Aenderungen."):
            return
        start_dir = self.current_file_path.parent if self.current_file_path else Path.home()
        file_path, _ = QFileDialog.getOpenFileName(self, "Testplan laden", str(start_dir), "Excel-Dateien (*.xlsx)")
        if not file_path:
            return
        try:
            self._load_from_excel(Path(file_path))
        except Exception as exc:
            self._show_error("Laden fehlgeschlagen", str(exc))
            return
        self.current_file_path = Path(file_path)
        self._set_dirty(False)
        self._update_window_title()
        self._append_log(f"Testplan geladen: {file_path}")

    def _save_as(self) -> bool:
        start_dir = self.current_file_path.parent if self.current_file_path else Path.home()
        default_name = self.current_file_path.name if self.current_file_path else "modbus_testplan.xlsx"
        file_path, _ = QFileDialog.getSaveFileName(self, "Testplan speichern", str(start_dir / default_name), "Excel-Dateien (*.xlsx)")
        if not file_path:
            return False
        path = Path(file_path)
        if path.suffix.lower() != ".xlsx":
            path = path.with_suffix(".xlsx")
        return self._save_plan_to_path(path)

    def _save_plan_to_path(self, path: Path) -> bool:
        try:
            self._save_to_excel(path)
        except Exception as exc:
            self._show_error("Speichern fehlgeschlagen", str(exc))
            return False
        self.current_file_path = path
        self._set_dirty(False)
        self._update_window_title()
        self._append_log(f"Testplan gespeichert: {path}")
        return True

    def _save_to_excel(self, path: Path) -> None:
        workbook = Workbook()
        sheet_connection = workbook.active
        sheet_connection.title = "Verbindung"
        sheet_connection.append(["Feld", "Wert"])
        sheet_connection.append(["host", self.host_input.text().strip() or "127.0.0.1"])
        sheet_connection.append(["port", self.port_input.value()])
        sheet_connection.append(["slave_id", self.slave_id_input.value()])
        sheet_connection.append(["register_format", self.register_format_combo.currentData().value])
        sheet_connection.append(["keepalive_interval_seconds", self.keepalive_input.value()])
        sheet_connection.append(["pt1_p_ms", self.pt1_p_input.value()])
        sheet_connection.append(["pt1_p_start_register", self.pt1_p_register_input.value()])
        sheet_connection.append(["pt1_q_ms", self.pt1_q_input.value()])
        sheet_connection.append(["pt1_q_start_register", self.pt1_q_register_input.value()])
        sheet_connection.append(["remote_enabled", self.remote_enabled_checkbox.isChecked()])
        sheet_connection.append(["remote_auto_control", self.remote_auto_checkbox.isChecked()])
        sheet_connection.append(["remote_host", self.remote_host_input.text().strip() or "127.0.0.1"])
        sheet_connection.append(["remote_port", self.remote_port_input.value()])
        sheet_connection.append(["remote_filename", self.remote_filename_input.text().strip()])

        sheet_channels = workbook.create_sheet("Kanaele")
        sheet_channels.append(["name", "label", "start_register", "value_type", "start_value"])
        for index, (channel_label_input, register_input, type_input, start_value_input) in enumerate(
            zip(self.channel_label_inputs, self.register_inputs, self.type_inputs, self.start_value_inputs),
            start=1,
        ):
            sheet_channels.append(
                [
                    f"S{index}",
                    channel_label_input.text().strip(),
                    register_input.value(),
                    type_input.currentData().value,
                    start_value_input.text().strip(),
                ]
            )

        sheet_plan = workbook.create_sheet("Testplan")
        sheet_plan.append(["active", "stage", "value_1", "value_2", "value_3", "value_4", "duration"])
        for row in range(self.test_table.rowCount()):
            sheet_plan.append(
                [
                    self.row_checkboxes[row].isChecked(),
                    row + 1,
                    *[
                        self._item_text(row, column)
                        for column in range(self.COLUMN_VALUE_START, self.COLUMN_VALUE_START + self.CHANNEL_COUNT)
                    ],
                    self._item_text(row, self.COLUMN_DURATION),
                ]
            )
        workbook.save(path)

    def _load_from_excel(self, path: Path) -> None:
        workbook = load_workbook(path)
        if "Verbindung" not in workbook.sheetnames or "Kanaele" not in workbook.sheetnames or "Testplan" not in workbook.sheetnames:
            raise RuntimeError("Die Excel-Datei muss die Blaetter 'Verbindung', 'Kanaele' und 'Testplan' enthalten.")

        connection_sheet = workbook["Verbindung"]
        connection = {
            str(row[0].value): row[1].value
            for row in connection_sheet.iter_rows(min_row=2, max_col=2)
            if row[0].value is not None
        }
        self.host_input.setText(connection.get("host", "127.0.0.1"))
        self.port_input.setValue(int(connection.get("port", 502)))
        self.slave_id_input.setValue(int(connection.get("slave_id", 1)))
        self.keepalive_input.setValue(int(connection.get("keepalive_interval_seconds", 20)))
        self.pt1_p_input.setValue(int(connection.get("pt1_p_ms", 0)))
        self.pt1_p_register_input.setValue(int(connection.get("pt1_p_start_register", 0)))
        self.pt1_q_input.setValue(int(connection.get("pt1_q_ms", 0)))
        self.pt1_q_register_input.setValue(int(connection.get("pt1_q_start_register", 0)))
        self.remote_enabled_checkbox.setChecked(self._coerce_bool(connection.get("remote_enabled", False)))
        self.remote_auto_checkbox.setChecked(
            self._coerce_bool(connection.get("remote_auto_control", connection.get("remote_auto_recording", True)), default=True)
        )
        self.remote_host_input.setText(str(connection.get("remote_host", "127.0.0.1")))
        self.remote_port_input.setValue(int(connection.get("remote_port", 5025)))
        self.remote_filename_input.setText("" if connection.get("remote_filename") is None else str(connection.get("remote_filename", "")))

        if "register_format" in connection:
            register_format_value = connection.get("register_format", RegisterFormat.BIG.value)
        else:
            register_format_value = RegisterFormat.from_legacy(connection.get("byte_order"), connection.get("word_order")).value
        self._set_combo_value(self.register_format_combo, register_format_value)

        channels_sheet = workbook["Kanaele"]
        channels = []
        for row in channels_sheet.iter_rows(min_row=2, max_col=5, values_only=True):
            if row[0] is None:
                continue
            if len(row) >= 5:
                channels.append(
                    {
                        "name": row[0],
                        "label": row[1],
                        "start_register": row[2],
                        "value_type": row[3],
                        "start_value": row[4],
                    }
                )
            elif len(row) >= 4:
                channels.append(
                    {
                        "name": row[0],
                        "label": row[1],
                        "start_register": row[2],
                        "value_type": row[3],
                        "start_value": "",
                    }
                )
            else:
                channels.append(
                    {
                        "name": row[0],
                        "label": "",
                        "start_register": row[1],
                        "value_type": row[2],
                        "start_value": "",
                    }
                )
        for index, channel in enumerate(channels):
            if index >= len(self.register_inputs):
                break
            self.channel_label_inputs[index].setText("" if channel.get("label") is None else str(channel.get("label", "")))
            self.register_inputs[index].setValue(int(channel.get("start_register", 0)))
            self._set_combo_value(self.type_inputs[index], channel.get("value_type", RegisterValueType.FLOAT32.value))
            self.start_value_inputs[index].setText("" if channel.get("start_value") is None else str(channel.get("start_value", "")))
        self._update_table_headers()

        plan_sheet = workbook["Testplan"]
        rows = []
        for row in plan_sheet.iter_rows(min_row=2, max_col=8, values_only=True):
            if row[1] is None:
                continue
            rows.append(
                {
                    "active": bool(row[0]),
                    "stage": row[1],
                    "values": ["" if value is None else str(value) for value in row[2:6]],
                    "duration": "" if row[6] is None else str(row[6]),
                }
            )
        self.test_table.blockSignals(True)
        for row in range(self.test_table.rowCount()):
            row_data = rows[row] if row < len(rows) else {}
            self.row_checkboxes[row].setChecked(bool(row_data.get("active", False)))
            values = row_data.get("values", ["", "", "", ""])
            for column, value in enumerate(values, start=self.COLUMN_VALUE_START):
                self._set_item_text(row, column, str(value))
            self._set_item_text(row, self.COLUMN_DURATION, str(row_data.get("duration", "")))
        self.test_table.blockSignals(False)
        self._update_time_copy_button()
        self._update_remote_controls()
        self._apply_row_visuals()
        self._refresh_summary()

    def _set_combo_value(self, combo: QComboBox, raw_value: str) -> None:
        normalized_raw = getattr(raw_value, "value", raw_value)
        for index in range(combo.count()):
            data = combo.itemData(index)
            normalized_data = getattr(data, "value", data)
            if data == raw_value or normalized_data == normalized_raw or str(normalized_data) == str(normalized_raw):
                combo.setCurrentIndex(index)
                return

    def _coerce_bool(self, raw_value: object, default: bool = False) -> bool:
        if raw_value is None:
            return default
        if isinstance(raw_value, str):
            normalized = raw_value.strip().lower()
            if normalized in {"1", "true", "yes", "ja", "on"}:
                return True
            if normalized in {"0", "false", "no", "nein", "off"}:
                return False
        return bool(raw_value)

    def _item_text(self, row: int, column: int) -> str:
        item = self.test_table.item(row, column)
        return "" if item is None else item.text().strip()

    def _set_item_text(self, row: int, column: int, value: str) -> None:
        item = self.test_table.item(row, column)
        if item is None:
            item = QTableWidgetItem()
            self.test_table.setItem(row, column, item)
        item.setText(value if value != "None" else "")
        item.setTextAlignment(Qt.AlignCenter)

    def _copy_first_stage_time_to_active_rows(self) -> None:
        source_text = self._item_text(0, self.COLUMN_DURATION)
        if not source_text:
            self._update_time_copy_button()
            return
        if not any(self.row_checkboxes[row].isChecked() for row in range(1, self.test_table.rowCount())):
            self._update_time_copy_button()
            return
        previous = self._table_ui_update_in_progress
        self._table_ui_update_in_progress = True
        self.test_table.blockSignals(True)
        try:
            for row in range(1, self.test_table.rowCount()):
                if self.row_checkboxes[row].isChecked():
                    self._set_item_text(row, self.COLUMN_DURATION, source_text)
        finally:
            self.test_table.blockSignals(False)
            self._table_ui_update_in_progress = previous
        self._mark_dirty()
        if not self.sequence_controller.has_active_plan:
            self._refresh_summary()
        self._apply_row_visuals()
        self.last_write_value.setText("Zeit aus Stufe 1 auf aktive Stufen uebernommen")
        self._update_time_copy_button()

    def _on_send_start_value_clicked(self) -> None:
        sender = self.sender()
        try:
            channel_index = self.send_value_buttons.index(sender)
        except ValueError:
            return
        if not self.modbus_service.is_connected:
            self._show_error("Keine Verbindung", "Bitte zuerst verbinden.")
            return
        raw_text = self.start_value_inputs[channel_index].text().strip()
        if not raw_text:
            self._show_error("Startwert fehlt", f"Bitte einen Startwert fuer S{channel_index + 1} eingeben.")
            return
        channel = self._channel_configs()[channel_index]
        try:
            value = ValueEncoder.coerce_text_value(raw_text, channel.value_type)
            result = self.modbus_service.write_channel(channel, value)
        except Exception as exc:
            self._show_error("Startwert senden fehlgeschlagen", str(exc))
            return
        self.last_write_value.setText(f"{channel.name}: Startwert gesendet")
        self._append_log(
            f"Startwert fuer {channel.name} gesendet: Wert {result.original_value} auf Register {result.start_register}."
        )

    def _on_next_stage_clicked(self) -> None:
        if not self.sequence_controller.has_active_plan:
            return
        self.sequence_controller.skip_current_stage()
        self.last_write_value.setText("Zur naechsten Stufe gesprungen")
        self._update_pause_button()

    def _on_table_item_changed(self, item: QTableWidgetItem) -> None:
        if self._table_ui_update_in_progress:
            return
        self._mark_dirty()
        self._table_ui_update_in_progress = True
        try:
            if item.column() >= 2:
                item.setTextAlignment(Qt.AlignCenter)
            if item.row() == 0 and item.column() == self.COLUMN_DURATION:
                self._update_time_copy_button()
            if not self.sequence_controller.has_active_plan:
                self._refresh_summary()
            self._apply_row_visuals()
        finally:
            self._table_ui_update_in_progress = False

    def _refresh_summary(self) -> None:
        if self.sequence_controller.has_active_plan:
            return
        try:
            plan = self._build_stage_plan()
        except RuntimeError:
            plan = []
        total = sum(stage.duration_seconds for stage in plan)
        self.total_time_value.setText(self._format_seconds(total))
        self.total_remaining_value.setText(self._format_seconds(total))
        self.stage_remaining_value.setText("00:00:00")
        self.current_stage_value.setText("-")

    def _apply_row_visuals(self) -> None:
        previous = self._table_ui_update_in_progress
        self._table_ui_update_in_progress = True
        try:
            for row in range(self.test_table.rowCount()):
                active = self.row_checkboxes[row].isChecked()
                for column in range(1, self.test_table.columnCount()):
                    item = self.test_table.item(row, column)
                    if item is None:
                        continue
                    flags = item.flags() | Qt.ItemIsSelectable | Qt.ItemIsEnabled
                    if column >= 2:
                        if active and not self.sequence_controller.has_active_plan:
                            item.setFlags(flags | Qt.ItemIsEditable)
                        else:
                            item.setFlags(flags & ~Qt.ItemIsEditable)
                    color = self.active_stage_color if row == self.current_highlight_row else (self.default_row_color if active else self.inactive_row_color)
                    item.setBackground(color)
                    item.setForeground(self.active_row_text_color if active else self.inactive_row_text_color)
        finally:
            self._table_ui_update_in_progress = previous

    def _set_connection_widgets(self, connected: bool) -> None:
        self.connect_button.setEnabled(not connected)
        self.disconnect_button.setEnabled(connected)
        self.host_input.setEnabled(not connected)
        self.port_input.setEnabled(not connected)
        self.slave_id_input.setEnabled(not connected)
        self.register_format_combo.setEnabled(not connected)
        self.keepalive_input.setEnabled(not connected)
        self._update_manual_write_buttons()
        self._update_remote_controls()

    def _set_running_widgets(self, running: bool) -> None:
        self.start_button.setEnabled(not running)
        self.pause_button.setEnabled(running)
        self.next_stage_button.setEnabled(running)
        self.stop_button.setEnabled(running)
        self.save_button.setEnabled(not running)
        self.load_button.setEnabled(not running)
        self.copy_stage_time_button.setEnabled(False)
        self.test_table.setEnabled(True)
        for spin_box in self.register_inputs:
            spin_box.setEnabled(not running)
        for input_field in self.channel_label_inputs:
            input_field.setEnabled(not running)
        for combo in self.type_inputs:
            combo.setEnabled(not running)
        for pt1_input in (self.pt1_p_input, self.pt1_q_input):
            pt1_input.setEnabled(not running)
        self.disconnect_button.setEnabled(not running and self.modbus_service.is_connected)
        self._update_manual_write_buttons()
        self._update_remote_controls()
        self._update_time_copy_button()
        self._update_pause_button()
        self._apply_row_visuals()

    def _update_pause_button(self) -> None:
        if self.sequence_controller.is_paused:
            self.pause_button.setText("Fortsetzen")
            self.pause_button.setEnabled(True)
            return
        self.pause_button.setText("Pause")
        self.pause_button.setEnabled(self.sequence_controller.has_active_plan)

    def _update_manual_write_buttons(self) -> None:
        manual_write_enabled = self.modbus_service.is_connected and not self.sequence_controller.has_active_plan
        for input_field in self.start_value_inputs:
            input_field.setEnabled(manual_write_enabled)
        for button in self.send_value_buttons:
            button.setEnabled(manual_write_enabled)

    def _update_remote_controls(self) -> None:
        settings = self._current_remote_settings()
        remote_enabled = settings.enabled
        config_enabled = remote_enabled and not self.sequence_controller.has_active_plan
        self.remote_enabled_checkbox.setEnabled(
            not self.sequence_controller.has_active_plan and not self._remote_measurement_running
        )
        for widget in (
            self.remote_host_input,
            self.remote_port_input,
            self.remote_filename_input,
            self.remote_auto_checkbox,
        ):
            widget.setEnabled(config_enabled)
        self.remote_set_filename_button.setEnabled(config_enabled and bool(settings.filename))
        self.remote_start_button.setEnabled(config_enabled and bool(settings.filename))
        self.remote_stop_button.setEnabled(remote_enabled and self._remote_measurement_running)
        if not remote_enabled:
            self._remote_state_text = "Aus"
            self.remote_status_value.setText(self._remote_state_text)
            return
        self.remote_status_value.setText(self._remote_state_text)

    def _on_remote_settings_changed(self) -> None:
        if not self.remote_enabled_checkbox.isChecked():
            self._remote_measurement_running = False
            self._remote_state_text = "Aus"
        self._mark_dirty()
        self._update_remote_controls()

    def _update_time_copy_button(self) -> None:
        has_source_time = bool(self._item_text(0, self.COLUMN_DURATION))
        has_target_rows = any(self.row_checkboxes[row].isChecked() for row in range(1, self.test_table.rowCount()))
        self.copy_stage_time_button.setEnabled(
            has_source_time and has_target_rows and not self.sequence_controller.has_active_plan
        )

    def _on_remote_set_filename_clicked(self) -> None:
        settings = self._current_remote_settings()
        if not settings.enabled:
            self._show_error("Remote nicht aktiv", "Bitte die Remote-Verbindung zuerst aktivieren.")
            return
        try:
            self.scpi_service.set_filename(settings)
        except Exception as exc:
            self._remote_state_text = "Fehler"
            self.remote_status_value.setText(self._remote_state_text)
            self._show_error("Remote-Dateiname fehlgeschlagen", str(exc))
            return
        self._remote_state_text = "Bereit"
        self.remote_status_value.setText(self._remote_state_text)
        self.last_write_value.setText("Remote-Dateiname gesetzt")
        self._append_log(
            f"SCPI Dateiname gesetzt: {settings.filename} auf {settings.host}:{settings.port}."
        )
        self._update_remote_controls()

    def _on_remote_start_clicked(self) -> None:
        settings = self._current_remote_settings()
        if not settings.enabled:
            self._show_error("Remote nicht aktiv", "Bitte die Remote-Verbindung zuerst aktivieren.")
            return
        try:
            self._start_remote_measurement(settings, source_label="Manueller Start")
        except Exception as exc:
            self._remote_state_text = "Fehler"
            self.remote_status_value.setText(self._remote_state_text)
            self._show_error("Remote-Messung fehlgeschlagen", str(exc))
            return
        self.last_write_value.setText("Remote-Messung gestartet")
        self._update_remote_controls()

    def _on_remote_stop_clicked(self) -> None:
        if not self._remote_measurement_running:
            return
        self._stop_remote_measurement("Remote-Messung manuell gestoppt.", show_dialog_on_error=True)
        self.last_write_value.setText("Remote-Messung gestoppt")
        self._update_remote_controls()

    def _start_remote_measurement(self, settings: RemoteConnectionSettings, source_label: str) -> None:
        self.scpi_service.set_filename(settings)
        self.scpi_service.start_measurement(settings)
        self._remote_measurement_running = True
        self._remote_state_text = "Laeuft"
        self.remote_status_value.setText(self._remote_state_text)
        self._append_log(
            f"SCPI {source_label}: Dateiname gesetzt und Messung gestartet auf {settings.host}:{settings.port}."
        )
        self._update_remote_controls()

    def _stop_remote_measurement(self, success_message: str, show_dialog_on_error: bool) -> None:
        if not self._remote_measurement_running:
            self._update_remote_controls()
            return
        settings = self._current_remote_settings()
        try:
            self.scpi_service.stop_measurement(settings)
        except Exception as exc:
            self._remote_state_text = "Fehler"
            self.remote_status_value.setText(self._remote_state_text)
            self._append_log(f"SCPI Stop fehlgeschlagen: {exc}")
            if show_dialog_on_error:
                self._show_error("Remote-Stopp fehlgeschlagen", str(exc))
            return
        self._remote_measurement_running = False
        self._remote_state_text = "Bereit"
        self.remote_status_value.setText(self._remote_state_text)
        self._append_log(success_message)
        self._update_remote_controls()

    def _set_status(self, status: str, text: str) -> None:
        self.current_status_key = status
        self.current_status_text = text
        self.status_badge.setText(text)
        theme = self.THEMES.get(self.current_theme, self.THEMES["light"])
        status_styles = theme["status_styles"]
        self.status_badge.setStyleSheet(status_styles.get(status, status_styles[STATUS_DISCONNECTED]))

    def _append_log(self, message: str) -> None:
        self.log_panel.appendPlainText(message)
        scrollbar = self.log_panel.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def _on_stage_changed(self, stage_index: int, row_index: int) -> None:
        self.current_stage_value.setText(str(row_index + 1) if row_index >= 0 else "-")
        self._highlight_current_row(row_index)

    def _on_timing_changed(self, total_planned: int, total_remaining: int, stage_remaining: int) -> None:
        self.total_time_value.setText(self._format_seconds(total_planned))
        self.total_remaining_value.setText(self._format_seconds(total_remaining))
        self.stage_remaining_value.setText(self._format_seconds(stage_remaining))

    def _on_sequence_finished(self) -> None:
        self._stop_remote_measurement("Remote-Messung abgeschlossen.", show_dialog_on_error=False)
        self._set_running_widgets(False)
        self.last_write_value.setText("Testplan abgeschlossen")
        self._highlight_current_row(-1)

    def _on_sequence_error(self, message: str) -> None:
        self._stop_remote_measurement("Remote-Messung nach Testfehler gestoppt.", show_dialog_on_error=False)
        self._set_running_widgets(False)
        self.last_write_value.setText(message)
        self._show_error("Testlauf fehlgeschlagen", message)
        self._highlight_current_row(-1)

    def _on_stage_write_completed(self, stage_number: int, message: str) -> None:
        self.last_write_value.setText(f"Stufe {stage_number}: {message}")

    def _highlight_current_row(self, row_index: int) -> None:
        self.current_highlight_row = row_index
        self._apply_row_visuals()

    def _on_theme_changed(self) -> None:
        self._apply_theme(self.theme_combo.currentData())

    def _on_channel_label_changed(self) -> None:
        self._update_table_headers()
        self._mark_dirty()

    def _update_table_headers(self) -> None:
        headers = ["Aktiv", "Stufe"]
        for index, input_field in enumerate(self.channel_label_inputs, start=1):
            label = input_field.text().strip()
            headers.append(f"Sollwert {index}" if not label else f"Sollwert {index}\n{label}")
        headers.append("Zeit [s]")
        self.test_table.setHorizontalHeaderLabels(headers)

    def _format_seconds(self, total_seconds: int) -> str:
        seconds = max(0, int(total_seconds))
        return f"{seconds // 3600:02d}:{(seconds % 3600) // 60:02d}:{seconds % 60:02d}"

    def _show_error(self, title: str, message: str) -> None:
        QMessageBox.critical(self, title, message)

    def _show_retry_dialog(self, title: str, message: str) -> bool:
        answer = QMessageBox.question(
            self,
            title,
            f"{message}\n\nMoechtest du es erneut versuchen?",
            QMessageBox.Retry | QMessageBox.Cancel,
            QMessageBox.Retry,
        )
        return answer == QMessageBox.Retry

    def closeEvent(self, event: QCloseEvent) -> None:  # type: ignore[override]
        if not self._confirm_discard_unsaved_changes("Vor dem Schliessen speichern?", "Es gibt ungespeicherte Aenderungen."):
            event.ignore()
            return
        self.keepalive_timer.stop()
        self.sequence_controller.stop()
        self._stop_remote_measurement("Remote-Messung beim Schliessen gestoppt.", show_dialog_on_error=False)
        self.modbus_service.disconnect()
        super().closeEvent(event)

    def _mark_dirty(self) -> None:
        self._set_dirty(True)

    def _set_dirty(self, dirty: bool) -> None:
        if self._has_unsaved_changes == dirty:
            return
        self._has_unsaved_changes = dirty
        self._update_window_title()

    def _update_window_title(self) -> None:
        file_name = self.current_file_path.name if self.current_file_path else "Unbenannt"
        dirty_marker = " *" if self._has_unsaved_changes else ""
        self.setWindowTitle(f"{self.DEFAULT_WINDOW_TITLE} - {file_name}{dirty_marker}")

    def _confirm_discard_unsaved_changes(self, title: str, message: str) -> bool:
        if not self._has_unsaved_changes:
            return True
        answer = QMessageBox.question(
            self,
            title,
            f"{message}\n\nMoechtest du zuerst speichern?",
            QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel,
            QMessageBox.Save,
        )
        if answer == QMessageBox.Cancel:
            return False
        if answer == QMessageBox.Save:
            return self._save_as()
        return True


def run() -> int:
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = ModbusMainWindow()
    window.show()
    return app.exec_()

