from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QThread, Qt, Signal
from PySide6.QtWidgets import (
	QApplication,
	QButtonGroup,
	QCheckBox,
	QFileDialog,
	QFrame,
	QGroupBox,
	QHBoxLayout,
	QLabel,
	QLineEdit,
	QMessageBox,
	QPushButton,
	QRadioButton,
	QScrollArea,
	QVBoxLayout,
	QWidget,
)

from scanner import (
	FORMAT_EXTENSIONS,
	PALETTE,
	PDF_FIELD_ORDER,
	write_pdf_report,
	write_report,
)


class ScanWorker(QThread):
	completed = Signal(str, int)
	failed = Signal(str)

	def __init__(
		self,
		source_folder: Path,
		output_folder: Path,
		selected_formats: set[str] | None = None,
		include_extension: bool = False,
		mode: str = "image",
		selected_fields: set[str] | None = None,
	) -> None:
		super().__init__()
		self.source_folder = source_folder
		self.output_folder = output_folder
		self.selected_formats = selected_formats or set()
		self.include_extension = include_extension
		self.mode = mode
		self.selected_fields = selected_fields or set()

	def run(self) -> None:
		try:
			if self.mode == "pdf":
				report_path, image_count = write_pdf_report(
					self.source_folder,
					self.output_folder,
					self.selected_fields,
				)
			else:
				report_path, image_count = write_report(
					self.source_folder,
					self.output_folder,
					self.selected_formats,
					self.include_extension,
				)
		except Exception as error:
			self.failed.emit(str(error))
		else:
			self.completed.emit(str(report_path), image_count)


class PackshotLister(QWidget):
	def __init__(self) -> None:
		super().__init__()
		self.worker: ScanWorker | None = None
		self.setWindowTitle("Packshot Lister")
		self.setMinimumWidth(650)
		self.setStyleSheet(
			f"""
			QWidget {{
				background-color: {PALETTE['cream']};
				color: {PALETTE['blue']};
				font-family: 'Segoe UI';
				font-size: 10pt;
			}}
			QLabel#title {{ font-size: 24pt; font-weight: 600; }}
			QLabel[class="section"] {{ font-size: 11pt; font-weight: 600; }}
			QLineEdit {{
				background: white;
				border: 1px solid {PALETTE['sky']};
				border-radius: 4px;
				padding: 9px;
			}}
			QLineEdit:focus {{ border: 2px solid {PALETTE['blue']}; }}
			QPushButton {{
				background: {PALETTE['sky']};
				border: none;
				border-radius: 4px;
				color: {PALETTE['blue']};
				font-weight: 600;
				padding: 10px 14px;
			}}
			QPushButton:hover {{ background: {PALETTE['pink']}; }}
			QPushButton#scan {{ background: {PALETTE['blue']}; color: white; }}
			QPushButton#scan:hover {{ background: #344977; }}
			QPushButton:disabled {{ background: #AAB4CA; color: white; }}
			QRadioButton#modeChoice {{
				background: #89A482;
				border: none;
				border-radius: 4px;
				color: {PALETTE['blue']};
				font-weight: 600;
				padding: 10px 14px;
				spacing: 8px;
			}}
			QRadioButton#modeChoice:checked {{
				background: #89A482;
				color: {PALETTE['blue']};
			}}
			QRadioButton#modeChoice:hover {{ background: #89A482; }}
			QRadioButton#modeChoice::indicator {{
				width: 17px;
				height: 17px;
				border: 1px solid {PALETTE['blue']};
				border-radius: 9px;
				background: white;
			}}
			QRadioButton#modeChoice::indicator:checked {{
				border: 1px solid {PALETTE['blue']};
				background: {PALETTE['blue']};
			}}
			QCheckBox {{ spacing: 8px; }}
			QCheckBox::indicator {{ width: 17px; height: 17px; }}
			QCheckBox::indicator:unchecked {{
				background: white;
				border: 1px solid {PALETTE['blue']};
				border-radius: 3px;
			}}
			QCheckBox::indicator:checked {{
				background: {PALETTE['blue']};
				border: 1px solid {PALETTE['blue']};
				border-radius: 3px;
			}}
			QRadioButton {{ spacing: 8px; }}
			"""
		)
		self._build_ui()

	def _build_ui(self) -> None:
		window_layout = QVBoxLayout(self)
		window_layout.setContentsMargins(0, 0, 0, 0)

		scroll_area = QScrollArea()
		scroll_area.setWidgetResizable(True)
		scroll_area.setFrameShape(QFrame.Shape.NoFrame)
		window_layout.addWidget(scroll_area)

		content = QWidget()
		scroll_area.setWidget(content)
		layout = QVBoxLayout(content)
		layout.setContentsMargins(34, 28, 34, 24)
		layout.setSpacing(12)

		title = QLabel("Packshot Lister")
		title.setObjectName("title")
		layout.addWidget(title)

		accent = QFrame()
		accent.setFixedSize(76, 4)
		accent.setStyleSheet(f"background-color: {PALETTE['pink']}; border: none;")
		layout.addWidget(accent)
		layout.addSpacing(8)

		self.source_entry = self._add_folder_row(
			layout, "Source folder", "Select Folder", self._choose_source
		)
		self.output_entry = self._add_folder_row(
			layout, "Output location", "Output Location", self._choose_output
		)

		self.data_mode_group = QButtonGroup(self)
		self.image_mode_radio = QRadioButton("get data from image names")
		self.pdf_mode_radio = QRadioButton("get data from pdf")
		self.image_mode_radio.setObjectName("modeChoice")
		self.pdf_mode_radio.setObjectName("modeChoice")
		self.data_mode_group.addButton(self.image_mode_radio)
		self.data_mode_group.addButton(self.pdf_mode_radio)
		self.image_mode_radio.setChecked(True)

		self.image_mode_radio.toggled.connect(self._sync_mode_ui)
		self.pdf_mode_radio.toggled.connect(self._sync_mode_ui)

		self.format_group = QGroupBox()
		self.format_group.setCheckable(False)
		self.format_section = QWidget()
		self.format_section_layout = QVBoxLayout(self.format_section)
		self.format_section_layout.setSpacing(10)
		format_label = QLabel("Image formats to scan:")
		format_label.setProperty("class", "section")
		self.format_section_layout.addWidget(format_label)
		format_row = QHBoxLayout()
		self.format_checks: dict[str, QCheckBox] = {}
		for image_format in FORMAT_EXTENSIONS:
			checkbox = QCheckBox(image_format)
			checkbox.setChecked(image_format == "TIFF")
			self.format_checks[image_format] = checkbox
			format_row.addWidget(checkbox)
		format_row.addStretch(1)
		self.format_section_layout.addLayout(format_row)

		self.include_extension_check = QCheckBox("Include file extension")
		self.format_section_layout.addWidget(self.include_extension_check)
		self.format_section_layout.addStretch(1)
		self.format_group.setLayout(QVBoxLayout())
		self.format_group.layout().addWidget(self.format_section)
		layout.addWidget(self.image_mode_radio, 0, Qt.AlignmentFlag.AlignLeft)
		layout.addWidget(self.format_group)

		self.pdf_group = QGroupBox()
		self.pdf_group.setCheckable(False)
		self.pdf_field_section = QWidget()
		self.pdf_field_layout = QVBoxLayout(self.pdf_field_section)
		self.pdf_field_layout.setSpacing(10)
		self.pdf_field_label = QLabel("PDF fields to extract:")
		self.pdf_field_label.setProperty("class", "section")
		self.pdf_field_layout.addWidget(self.pdf_field_label)
		self.pdf_field_checks: dict[str, QCheckBox] = {}
		field_row = QVBoxLayout()
		for field_name in PDF_FIELD_ORDER:
			checkbox = QCheckBox(field_name)
			checkbox.setChecked(True)
			self.pdf_field_checks[field_name] = checkbox
			field_row.addWidget(checkbox)
		self.pdf_field_layout.addLayout(field_row)
		self.pdf_field_layout.addStretch(1)
		self.pdf_group.setLayout(QVBoxLayout())
		self.pdf_group.layout().addWidget(self.pdf_field_section)
		layout.addWidget(self.pdf_mode_radio, 0, Qt.AlignmentFlag.AlignLeft)
		layout.addWidget(self.pdf_group)

		self._sync_mode_ui()

		divider = QFrame()
		divider.setFrameShape(QFrame.Shape.HLine)
		divider.setStyleSheet(f"color: {PALETTE['sky']};")
		layout.addSpacing(14)
		layout.addWidget(divider)

		self.scan_button = QPushButton("Run Process")
		self.scan_button.setObjectName("scan")
		self.scan_button.clicked.connect(self._start_scan)
		layout.addWidget(self.scan_button)
		self.status_label = QLabel("Choose folders and formats to begin.")
		self.status_label.setWordWrap(True)
		layout.addWidget(self.status_label)

	def _add_folder_row(
		self,
		layout: QVBoxLayout,
		label_text: str,
		button_text: str,
		callback,
	) -> QLineEdit:
		label = QLabel(label_text)
		label.setProperty("class", "section")
		layout.addWidget(label)
		row = QHBoxLayout()
		button = QPushButton(button_text)
		button.clicked.connect(callback)
		entry = QLineEdit()
		row.addWidget(button)
		row.addWidget(entry, 1)
		layout.addLayout(row)
		return entry

	def _choose_source(self) -> None:
		selected = QFileDialog.getExistingDirectory(self, "Select the parent folder to scan")
		if selected:
			self.source_entry.setText(selected)

	def _choose_output(self) -> None:
		selected = QFileDialog.getExistingDirectory(
			self, "Select where the Excel report will be saved"
		)
		if selected:
			self.output_entry.setText(selected)

	def _sync_mode_ui(self) -> None:
		image_mode = self.image_mode_radio.isChecked()
		self.format_group.setEnabled(image_mode)
		self.format_section.setEnabled(image_mode)
		self.include_extension_check.setEnabled(image_mode)
		for checkbox in self.format_checks.values():
			checkbox.setEnabled(image_mode)
		self.pdf_group.setEnabled(not image_mode)
		self.pdf_field_section.setEnabled(not image_mode)
		for checkbox in self.pdf_field_checks.values():
			checkbox.setEnabled(not image_mode)

	def _start_scan(self) -> None:
		source_folder = Path(self.source_entry.text().strip()).expanduser()
		output_folder = Path(self.output_entry.text().strip()).expanduser()

		if not source_folder.is_dir():
			QMessageBox.critical(self, "Choose a folder", "Select an existing source folder to scan.")
			return
		if not output_folder.is_dir():
			QMessageBox.critical(
			self, "Choose an output location", "Select an existing folder for the Excel report."
		)
			return

		if self.image_mode_radio.isChecked():
			selected_formats = {
				image_format
				for image_format, checkbox in self.format_checks.items()
				if checkbox.isChecked()
			}
			if not selected_formats:
				QMessageBox.critical(self, "Choose image formats", "Select at least one format to scan.")
				return
			self.scan_button.setEnabled(False)
			self.status_label.setText("Scanning folders and creating the Excel report...")
			self.worker = ScanWorker(
				source_folder,
				output_folder,
				selected_formats=selected_formats,
				include_extension=self.include_extension_check.isChecked(),
				mode="image",
			)
		else:
			selected_fields = {
				field_name
				for field_name, checkbox in self.pdf_field_checks.items()
				if checkbox.isChecked()
			}
			if not selected_fields:
				QMessageBox.critical(self, "Choose PDF fields", "Select at least one PDF field to extract.")
				return
			self.scan_button.setEnabled(False)
			self.status_label.setText("Reading PDF data and creating the Excel report...")
			self.worker = ScanWorker(
				source_folder,
				output_folder,
				mode="pdf",
				selected_fields=selected_fields,
			)

		self.worker.completed.connect(self._scan_completed)
		self.worker.failed.connect(self._scan_failed)
		self.worker.start()

	def _scan_completed(self, report_path: str, image_count: int) -> None:
		self.scan_button.setEnabled(True)
		self.status_label.setText(f"Saved {image_count} row(s) to {Path(report_path).name}")
		QMessageBox.information(
			self,
			"Scan complete",
			f"Found {image_count} row(s).\n\nReport saved to:\n{report_path}",
		)

	def _scan_failed(self, error: str) -> None:
		self.scan_button.setEnabled(True)
		self.status_label.setText("Scan did not complete.")
		QMessageBox.critical(self, "Scan failed", error)

	def closeEvent(self, event) -> None:
		if self.worker is not None and self.worker.isRunning():
			self.worker.wait()
		event.accept()


def main() -> None:
	app = QApplication.instance() or QApplication([])
	window = PackshotLister()
	window.resize(760, 650)
	window.show()
	app.exec()
