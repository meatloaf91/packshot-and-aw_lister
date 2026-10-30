from __future__ import annotations

from pathlib import Path

from openpyxl.utils import get_column_letter
from PySide6.QtCore import QThread, Qt, Signal
from PySide6.QtWidgets import (
	QApplication,
	QButtonGroup,
	QCheckBox,
	QComboBox,
	QFileDialog,
	QFrame,
	QGroupBox,
	QGridLayout,
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
	read_naming_columns,
	rename_packshots,
	write_pdf_report,
	write_report,
)


class ScanWorker(QThread):
	completed = Signal(str, int, int)
	failed = Signal(str)

	def __init__(
		self,
		source_folder: Path,
		output_folder: Path,
		selected_formats: set[str] | None = None,
		include_extension: bool = False,
		mode: str = "image",
		selected_fields: set[str] | None = None,
		naming_file: Path | None = None,
		original_column: int | None = None,
		new_column: int | None = None,
		create_backup: bool = False,
		export_report: bool = False,
	) -> None:
		super().__init__()
		self.source_folder = source_folder
		self.output_folder = output_folder
		self.selected_formats = selected_formats or set()
		self.include_extension = include_extension
		self.mode = mode
		self.selected_fields = selected_fields or set()
		self.naming_file = naming_file
		self.original_column = original_column
		self.new_column = new_column
		self.create_backup = create_backup
		self.export_report = export_report

	def run(self) -> None:
		try:
			if self.mode == "pdf":
				report_path, image_count = write_pdf_report(
					self.source_folder,
					self.output_folder,
					self.selected_fields,
				)
			elif self.mode == "rename":
				if (
					self.naming_file is None
					or self.original_column is None
					or self.new_column is None
				):
					raise ValueError("Choose a naming workbook and both name columns.")
				rename_path, image_count, skipped_count = rename_packshots(
					self.source_folder,
					self.naming_file,
					self.original_column,
					self.new_column,
					self.create_backup,
					self.export_report,
				)
				report_path = rename_path
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
			self.completed.emit(
				str(report_path) if report_path else "",
				image_count,
				skipped_count if self.mode == "rename" else 0,
			)


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
			QLineEdit#folderPath {{
				background: transparent;
				border: none;
				border-bottom: 1px solid {PALETTE['sky']};
				border-radius: 0;
				padding: 6px 2px;
			}}
			QLineEdit#folderPath:focus {{
				border: none;
				border-bottom: 2px solid {PALETTE['blue']};
			}}
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
				color: white;
				font-weight: 600;
				padding: 10px 14px;
				spacing: 8px;
			}}
			QRadioButton#modeChoice:checked {{
				background: #89A482;
				color: white;
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

		self.data_mode_group = QButtonGroup(self)
		self.image_mode_radio = QRadioButton("get data from image names")
		self.pdf_mode_radio = QRadioButton("get data from pdf")
		self.rename_mode_radio = QRadioButton("rename packshots")
		self.image_mode_radio.setObjectName("modeChoice")
		self.pdf_mode_radio.setObjectName("modeChoice")
		self.rename_mode_radio.setObjectName("modeChoice")
		self.data_mode_group.addButton(self.image_mode_radio)
		self.data_mode_group.addButton(self.pdf_mode_radio)
		self.data_mode_group.addButton(self.rename_mode_radio)
		self.image_mode_radio.setChecked(True)

		self.image_mode_radio.toggled.connect(self._sync_mode_ui)
		self.pdf_mode_radio.toggled.connect(self._sync_mode_ui)
		self.rename_mode_radio.toggled.connect(self._sync_mode_ui)

		self.format_group = QGroupBox()
		self.format_group.setCheckable(False)
		self.format_section = QWidget()
		self.format_section_layout = QVBoxLayout(self.format_section)
		self.format_section_layout.setSpacing(10)
		self.image_source_entry = self._add_folder_row(
			self.format_section_layout, "Select Folder", self._choose_source
		)
		self.image_output_entry = self._add_folder_row(
			self.format_section_layout, "Output Location", self._choose_output
		)
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
		self.format_section_layout.addSpacing(40)
		self.image_run_button = self._create_run_button(self._start_image_scan)
		image_button_row = QHBoxLayout()
		image_button_row.addStretch(1)
		image_button_row.addWidget(self.image_run_button, 2)
		image_button_row.addStretch(1)
		self.format_section_layout.addLayout(image_button_row)
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
		self.pdf_source_entry = self._add_folder_row(
			self.pdf_field_layout, "Select Folder", self._choose_source
		)
		self.pdf_output_entry = self._add_folder_row(
			self.pdf_field_layout, "Output Location", self._choose_output
		)
		self.pdf_field_label = QLabel("PDF fields to extract:")
		self.pdf_field_label.setProperty("class", "section")
		self.pdf_field_layout.addWidget(self.pdf_field_label)
		self.pdf_field_checks: dict[str, QCheckBox] = {}
		field_grid = QGridLayout()
		for index, field_name in enumerate(PDF_FIELD_ORDER):
			checkbox = QCheckBox(field_name)
			checkbox.setChecked(True)
			self.pdf_field_checks[field_name] = checkbox
			field_grid.addWidget(checkbox, index // 4, index % 4)
		self.pdf_field_layout.addLayout(field_grid)
		self.pdf_field_layout.addSpacing(40)
		self.pdf_run_button = self._create_run_button(self._start_pdf_scan)
		pdf_button_row = QHBoxLayout()
		pdf_button_row.addStretch(1)
		pdf_button_row.addWidget(self.pdf_run_button, 2)
		pdf_button_row.addStretch(1)
		self.pdf_field_layout.addLayout(pdf_button_row)
		self.pdf_field_layout.addStretch(1)
		self.pdf_group.setLayout(QVBoxLayout())
		self.pdf_group.layout().addWidget(self.pdf_field_section)
		layout.addWidget(self.pdf_mode_radio, 0, Qt.AlignmentFlag.AlignLeft)
		layout.addWidget(self.pdf_group)

		self.rename_group = QGroupBox()
		self.rename_section = QWidget()
		self.rename_layout = QVBoxLayout(self.rename_section)
		self.rename_layout.setSpacing(10)
		self.rename_image_entry = self._add_folder_row(
			self.rename_layout, "Select Image Folder", self._choose_source
		)
		self.naming_file_entry = QLineEdit()
		self.naming_file_entry.setObjectName("folderPath")
		self.naming_file_entry.setReadOnly(True)
		self.naming_file_button = QPushButton("Select Naming File")
		self.naming_file_button.clicked.connect(self._choose_naming_file)
		naming_file_row = QHBoxLayout()
		naming_file_row.addWidget(self.naming_file_button)
		naming_file_row.addWidget(self.naming_file_entry, 1)
		self.rename_layout.addLayout(naming_file_row)

		self.original_name_combo = QComboBox()
		self.new_name_combo = QComboBox()
		self.original_name_combo.setEnabled(False)
		self.new_name_combo.setEnabled(False)
		self._add_combo_row(self.rename_layout, "original names", self.original_name_combo)
		self._add_combo_row(self.rename_layout, "new names", self.new_name_combo)

		self.create_backup_check = QCheckBox("Create a backup")
		self.create_backup_check.setChecked(True)
		self.export_report_check = QCheckBox("Export a report")
		self.export_report_check.setChecked(True)
		self.rename_layout.addWidget(self.create_backup_check)
		self.rename_layout.addWidget(self.export_report_check)
		self.rename_layout.addSpacing(40)
		self.rename_run_button = self._create_run_button(self._start_rename)
		rename_button_row = QHBoxLayout()
		rename_button_row.addStretch(1)
		rename_button_row.addWidget(self.rename_run_button, 2)
		rename_button_row.addStretch(1)
		self.rename_layout.addLayout(rename_button_row)
		self.rename_layout.addStretch(1)
		self.rename_group.setLayout(QVBoxLayout())
		self.rename_group.layout().addWidget(self.rename_section)
		layout.addWidget(self.rename_mode_radio, 0, Qt.AlignmentFlag.AlignLeft)
		layout.addWidget(self.rename_group)

		self._sync_mode_ui()

		self.status_label = QLabel("Choose folders and formats to begin.")
		self.status_label.setWordWrap(True)
		layout.addWidget(self.status_label)

	def _add_folder_row(
		self,
		layout: QVBoxLayout,
		button_text: str,
		callback,
	) -> QLineEdit:
		row = QHBoxLayout()
		button = QPushButton(button_text)
		entry = QLineEdit()
		entry.setObjectName("folderPath")
		button.clicked.connect(lambda checked=False: callback(entry))
		row.addWidget(button)
		row.addWidget(entry, 1)
		layout.addLayout(row)
		return entry

	def _add_combo_row(
		self,
		layout: QVBoxLayout,
		label_text: str,
		combo: QComboBox,
	) -> None:
		label = QLabel(label_text)
		label.setProperty("class", "section")
		layout.addWidget(label)
		layout.addWidget(combo)

	def _create_run_button(self, callback) -> QPushButton:
		button = QPushButton("Run Process")
		button.setObjectName("scan")
		button.clicked.connect(callback)
		return button

	def _choose_source(self, entry: QLineEdit) -> None:
		selected = QFileDialog.getExistingDirectory(self, "Select the parent folder to scan")
		if selected:
			entry.setText(selected)

	def _choose_output(self, entry: QLineEdit) -> None:
		selected = QFileDialog.getExistingDirectory(
			self, "Select where the Excel report will be saved"
		)
		if selected:
			entry.setText(selected)

	def _choose_naming_file(self) -> None:
		selected, _ = QFileDialog.getOpenFileName(
			self,
			"Select the naming workbook",
			"",
			"Excel workbooks (*.xlsx *.xlsm)",
		)
		if not selected:
			return
		try:
			columns = read_naming_columns(Path(selected))
		except Exception as error:
			QMessageBox.critical(self, "Unable to read naming workbook", str(error))
			return
		if not columns:
			QMessageBox.critical(
				self,
				"Naming workbook has no columns",
				"The active worksheet must have column names in its first row.",
			)
			return
		self.naming_file_entry.setText(selected)
		for combo in (self.original_name_combo, self.new_name_combo):
			combo.clear()
			for name, index in columns:
				combo.addItem(f"{name} (Column {get_column_letter(index + 1)})", index)
			combo.setEnabled(True)

	def _sync_mode_ui(self) -> None:
		image_mode = self.image_mode_radio.isChecked()
		pdf_mode = self.pdf_mode_radio.isChecked()
		rename_mode = self.rename_mode_radio.isChecked()
		self.format_group.setEnabled(image_mode)
		self.format_section.setEnabled(image_mode)
		self.include_extension_check.setEnabled(image_mode)
		for checkbox in self.format_checks.values():
			checkbox.setEnabled(image_mode)
		self.pdf_group.setEnabled(pdf_mode)
		self.pdf_field_section.setEnabled(pdf_mode)
		for checkbox in self.pdf_field_checks.values():
			checkbox.setEnabled(pdf_mode)
		self.rename_group.setEnabled(rename_mode)
		if self.worker is None or not self.worker.isRunning():
			self.image_run_button.setEnabled(image_mode)
			self.pdf_run_button.setEnabled(pdf_mode)
			self.rename_run_button.setEnabled(rename_mode)

	def _start_image_scan(self) -> None:
		selected_formats = {
			image_format
			for image_format, checkbox in self.format_checks.items()
			if checkbox.isChecked()
		}
		if not selected_formats:
			QMessageBox.critical(self, "Choose image formats", "Select at least one format to scan.")
			return
		self._start_scan(
			self.image_source_entry,
			self.image_output_entry,
			mode="image",
			selected_formats=selected_formats,
		)

	def _start_pdf_scan(self) -> None:
		selected_fields = {
			field_name
			for field_name, checkbox in self.pdf_field_checks.items()
			if checkbox.isChecked()
		}
		if not selected_fields:
			QMessageBox.critical(self, "Choose PDF fields", "Select at least one PDF field to extract.")
			return
		self._start_scan(
			self.pdf_source_entry,
			self.pdf_output_entry,
			mode="pdf",
			selected_fields=selected_fields,
		)

	def _start_rename(self) -> None:
		image_folder = Path(self.rename_image_entry.text().strip()).expanduser()
		naming_file = Path(self.naming_file_entry.text().strip()).expanduser()
		if not image_folder.is_dir():
			QMessageBox.critical(
				self, "Choose an image folder", "Select an existing folder containing packshots."
			)
			return
		if not naming_file.is_file():
			QMessageBox.critical(
				self, "Choose a naming file", "Select an existing Excel naming workbook."
			)
			return
		original_column = self.original_name_combo.currentData()
		new_column = self.new_name_combo.currentData()
		if original_column is None or new_column is None:
			QMessageBox.critical(
				self, "Choose name columns", "Select both the original names and new names columns."
			)
			return
		if original_column == new_column:
			QMessageBox.critical(
				self, "Choose different columns", "Original names and new names must use different columns."
			)
			return
		self._start_scan(
			image_folder,
			naming_file.parent,
			mode="rename",
			naming_file=naming_file,
			original_column=original_column,
			new_column=new_column,
			create_backup=self.create_backup_check.isChecked(),
			export_report=self.export_report_check.isChecked(),
		)

	def _start_scan(
		self,
		source_entry: QLineEdit | Path,
		output_entry: QLineEdit | Path,
		*,
		mode: str,
		selected_formats: set[str] | None = None,
		selected_fields: set[str] | None = None,
		naming_file: Path | None = None,
		original_column: int | None = None,
		new_column: int | None = None,
		create_backup: bool = False,
		export_report: bool = False,
	) -> None:
		source_folder = (
			Path(source_entry.text().strip()).expanduser()
			if isinstance(source_entry, QLineEdit)
			else source_entry
		)
		output_folder = (
			Path(output_entry.text().strip()).expanduser()
			if isinstance(output_entry, QLineEdit)
			else output_entry
		)

		if not source_folder.is_dir():
			QMessageBox.critical(self, "Choose a folder", "Select an existing source folder to scan.")
			return
		if not output_folder.is_dir():
			QMessageBox.critical(
			self, "Choose an output location", "Select an existing folder for the Excel report."
		)
			return

		self._set_all_run_buttons_enabled(False)
		if mode == "image":
			self.status_label.setText("Scanning folders and creating the Excel report...")
			self.worker = ScanWorker(
				source_folder,
				output_folder,
				selected_formats=selected_formats or set(),
				include_extension=self.include_extension_check.isChecked(),
				mode="image",
			)
		elif mode == "pdf":
			self.status_label.setText("Reading PDF data and creating the Excel report...")
			self.worker = ScanWorker(
				source_folder,
				output_folder,
				mode="pdf",
				selected_fields=selected_fields or set(),
			)
		else:
			self.status_label.setText("Renaming matching packshots...")
			self.worker = ScanWorker(
				source_folder,
				output_folder,
				mode="rename",
				naming_file=naming_file,
				original_column=original_column,
				new_column=new_column,
				create_backup=create_backup,
				export_report=export_report,
			)

		self.worker.completed.connect(self._scan_completed)
		self.worker.failed.connect(self._scan_failed)
		self.worker.start()

	def _scan_completed(self, report_path: str, image_count: int, skipped_count: int) -> None:
		self._set_run_buttons_for_mode()
		if self.rename_mode_radio.isChecked():
			skipped_message = (
				f"\nSkipped {skipped_count} matching file(s) due to an existing target "
				"or backup."
				if skipped_count
				else ""
			)
			if report_path:
				message = (
					f"Renamed {image_count} packshot(s).{skipped_message}\n\n"
					f"Report saved to:\n{report_path}"
				)
				self.status_label.setText(
					f"Renamed {image_count} packshot(s); "
					f"skipped {skipped_count}; report saved to {Path(report_path).name}"
				)
			else:
				message = f"Renamed {image_count} packshot(s).{skipped_message}"
				self.status_label.setText(message)
		else:
			message = f"Found {image_count} row(s).\n\nReport saved to:\n{report_path}"
			self.status_label.setText(f"Saved {image_count} row(s) to {Path(report_path).name}")
		QMessageBox.information(self, "Process complete", message)

	def _scan_failed(self, error: str) -> None:
		self._set_run_buttons_for_mode()
		self.status_label.setText("Scan did not complete.")
		QMessageBox.critical(self, "Scan failed", error)

	def _set_all_run_buttons_enabled(self, enabled: bool) -> None:
		for button in (self.image_run_button, self.pdf_run_button, self.rename_run_button):
			button.setEnabled(enabled)

	def _set_run_buttons_for_mode(self) -> None:
		self.image_run_button.setEnabled(self.image_mode_radio.isChecked())
		self.pdf_run_button.setEnabled(self.pdf_mode_radio.isChecked())
		self.rename_run_button.setEnabled(self.rename_mode_radio.isChecked())

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
