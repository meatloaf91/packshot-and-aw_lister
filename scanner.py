from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path

import fitz
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill


PALETTE = {
	"blue": "#425B9A",
	"sky": "#76C0EC",
	"cream": "#FFF6DC",
	"pink": "#FF95A5",
}

FORMAT_EXTENSIONS = {
	"TIFF": {".tif", ".tiff"},
	"PNG": {".png"},
	"JPG": {".jpg", ".jpeg"},
	"PSD": {".psd"},
	"PDF": {".pdf"},
}

PDF_FIELD_ORDER = [
	"IDH no.",
	"EAN",
	"Country",
	"Material no.",
	"Article Short Code",
	"SEAL ID",
	"Project No.",
	"File Name",
]

VIEW_PATTERN = re.compile(
	r"(?:^|_)(left_angle|right_angle|front|back|side|left|right|top|bottom|notino|dm|ing)$",
	re.IGNORECASE,
)
NUMERIC_TOKEN_PATTERN = re.compile(r"(?<![A-Za-z0-9])\d+(?![A-Za-z0-9])")
CAPACITY_PATTERN = re.compile(
	r"(?<![A-Za-z0-9])"
	r"(\d+(?:[.,]\d+)?)"
	r"(?:[_\s-]*)"
	r"(fl\.?[_\s-]*oz|millilit(?:er|re)s?|kilograms?|"
	r"centilit(?:er|re)s?|decilit(?:er|re)s?|milligrams?|micrograms?|"
	r"grams?|lit(?:er|re)s?|ounces?|gallons?|quarts?|pints?|"
	r"lbs?|oz|kg|mg|mcg|μg|µg|ml|cl|dl|ul|l|g|gal|qt|pt)"
	r"(?![A-Za-z0-9])",
	re.IGNORECASE,
)


def extract_codes_and_view(image_name: str) -> tuple[str, str, str]:
	"""Extract the first EAN and IDH numeric tokens and recognized view suffix."""
	stem = Path(image_name).stem
	numeric_tokens = NUMERIC_TOKEN_PATTERN.findall(stem)
	ean = next((token for token in numeric_tokens if len(token) >= 8), "")
	idh = next((token for token in numeric_tokens if 5 <= len(token) <= 7), "")
	view_match = VIEW_PATTERN.search(stem)
	return (
		ean,
		idh,
		view_match.group(1).lower() if view_match else "",
	)


def extract_capacity(image_name: str) -> str:
	"""Extract a number with a recognized volume or mass unit from an image name."""
	match = CAPACITY_PATTERN.search(Path(image_name).stem)
	if not match:
		return ""
	return f"{match.group(1)}{match.group(2).replace('_', '').replace(' ', '')}"


def _clean_pdf_value(raw_value: str) -> str:
	cleaned = raw_value.replace("\u00a0", " ")
	cleaned = re.sub(r"\s+", " ", cleaned)
	cleaned = cleaned.strip(" :;,-")
	return cleaned.strip()


def _extract_pdf_field(pdf_text: str, label: str) -> str:
	patterns = {
		"IDH no.": re.compile(r"IDH\s*no\.?\s*(?:[:\-]?\s*)?(\d{5,8})", re.IGNORECASE),
		"EAN": re.compile(
			r"EAN(?:\s*\(please check\))?\s*(?:[:\-]?\s*)?(\d{8,18})",
			re.IGNORECASE,
		),
		"Country": re.compile(
			r"Country\s*(?:[:\-]?\s*)?([A-Z0-9\(\)\-\s/]+?)(?=(?:\s*(?:Creation\s+date|Product\s+|Status\s+|Date\s+of\s+approval|$)))",
			re.IGNORECASE,
		),
		"Material no.": re.compile(r"Material\s*no\.?\s*(?:[:\-]?\s*)?([A-Z0-9\-]+)", re.IGNORECASE),
		"Article Short Code": re.compile(
			r"Article\s+Short\s+code\s*(?:[:\-]?\s*)?([A-Z0-9\-]+)",
			re.IGNORECASE,
		),
		"SEAL ID": re.compile(r"SEAL\s*ID\s*(?:[:\-]?\s*)?([A-Z0-9\-]+)", re.IGNORECASE),
		"Project No.": re.compile(
			r"Project\s*no\.?\s*(?:\(\s*HINT\s*\))?\s*(?:[:\-]?\s*)?([A-Z0-9\-]+)",
			re.IGNORECASE,
		),
		"File Name": re.compile(
			r"(?:LSD\s+FILE\s+NAME|File\s+Name)\s*(?:[:\-]?\s*)?([A-Za-z0-9_\-\.]+)",
			re.IGNORECASE,
		),
	}
	label_pattern = patterns.get(label)
	if label_pattern is None:
		return ""
	match = label_pattern.search(pdf_text)
	if not match:
		return ""
	value = match.group(1)
	if label == "EAN":
		value = re.sub(r"^0+", "", value)
	value = _clean_pdf_value(value)
	return value


def extract_pdf_data(pdf_path: Path) -> dict[str, str]:
	"""Read key-value data from the PDF info box and return selected fields."""
	text_chunks: list[str] = []
	try:
		with fitz.open(pdf_path) as document:
			for page in document:
				text_chunks.append(page.get_text("text", sort=True))
	except Exception:
		return {field: "" for field in PDF_FIELD_ORDER}

	pdf_text = " ".join(text_chunks)
	if not pdf_text.strip():
		return {field: "" for field in PDF_FIELD_ORDER}

	fields: dict[str, str] = {}
	for field in PDF_FIELD_ORDER:
		value = _extract_pdf_field(pdf_text, field)
		if not value and field == "File Name":
			value = pdf_path.name
		fields[field] = value
	return fields


def scan_images(
	source_folder: Path,
	selected_formats: set[str],
	include_extension: bool,
) -> tuple[list[list[str]], list[str]]:
	"""Return workbook rows and headers for matching images under source_folder."""
	suffixes = {
		suffix
		for image_format in selected_formats
		for suffix in FORMAT_EXTENSIONS[image_format]
	}
	matches: list[tuple[Path, tuple[str, ...]]] = []

	for image_path in source_folder.rglob("*"):
		if image_path.is_file() and image_path.suffix.lower() in suffixes:
			relative_path = image_path.relative_to(source_folder)
			matches.append((image_path, relative_path.parts[:-1]))

	matches.sort(key=lambda match: str(match[0]).casefold())
	deepest_folder = max((len(folders) for _, folders in matches), default=0)
	headers = [f"folder {level}" for level in range(1, deepest_folder + 2)]
	headers.extend(["ean", "idh", "capacity", "image_name", "image_format", "view"])

	rows: list[list[str]] = []
	for image_path, folders in matches:
		folder_columns = [source_folder.name, *folders]
		folder_columns.extend([""] * (deepest_folder + 1 - len(folder_columns)))
		ean, idh, view = extract_codes_and_view(image_path.stem)
		capacity = extract_capacity(image_path.stem)
		image_name = image_path.name if include_extension else image_path.stem
		image_format = next(
			name
			for name, extensions in FORMAT_EXTENSIONS.items()
			if image_path.suffix.lower() in extensions
		)
		rows.append([*folder_columns, ean, idh, capacity, image_name, image_format, view])

	return rows, headers


def scan_pdf_files(
	source_folder: Path,
	selected_fields: set[str],
) -> tuple[list[list[str]], list[str]]:
	"""Return pdf rows and headers for selected metadata fields."""
	pdf_matches: list[Path] = []
	for pdf_path in source_folder.rglob("*"):
		if pdf_path.is_file() and pdf_path.suffix.lower() == ".pdf":
			pdf_matches.append(pdf_path)

	pdf_matches.sort(key=lambda path: str(path).casefold())
	deepest_folder = max(
		(len(pdf_path.relative_to(source_folder).parts[:-1]) for pdf_path in pdf_matches),
		default=0,
	)
	headers = [f"folder {level}" for level in range(1, deepest_folder + 2)]
	headers.extend(field for field in PDF_FIELD_ORDER if field in selected_fields)

	rows: list[list[str]] = []
	for pdf_path in pdf_matches:
		relative_path = pdf_path.relative_to(source_folder)
		folders = relative_path.parts[:-1]
		folder_columns = [source_folder.name, *folders]
		folder_columns.extend([""] * (deepest_folder + 1 - len(folder_columns)))
		fields = extract_pdf_data(pdf_path)
		row: list[str] = [*folder_columns]
		for field in PDF_FIELD_ORDER:
			if field in selected_fields:
				row.append(fields.get(field, ""))
		rows.append(row)

	return rows, headers


def write_report(
	source_folder: Path,
	output_folder: Path,
	selected_formats: set[str],
	include_extension: bool,
) -> tuple[Path, int]:
	rows, headers = scan_images(source_folder, selected_formats, include_extension)
	timestamp = datetime.now().strftime("%Y_%m_%d_%H_%M_%S")
	report_path = output_folder / f"scanned_packshots_{timestamp}.xlsx"

	workbook = Workbook()
	worksheet = workbook.active
	worksheet.title = "report"
	worksheet.append(headers)
	for row in rows:
		worksheet.append(row)

	header_fill = PatternFill("solid", fgColor=PALETTE["blue"].lstrip("#"))
	for cell in worksheet[1]:
		cell.fill = header_fill
		cell.font = Font(color="FFFFFF", bold=True)
	worksheet.freeze_panes = "A2"
	worksheet.auto_filter.ref = worksheet.dimensions
	for column in worksheet.columns:
		column_letter = column[0].column_letter
		max_width = max(len(str(cell.value or "")) for cell in column)
		worksheet.column_dimensions[column_letter].width = min(max(max_width + 2, 14), 42)

	workbook.save(report_path)
	return report_path, len(rows)


def write_pdf_report(
	source_folder: Path,
	output_folder: Path,
	selected_fields: set[str],
) -> tuple[Path, int]:
	rows, headers = scan_pdf_files(source_folder, selected_fields)
	timestamp = datetime.now().strftime("%Y_%m_%d_%H_%M_%S")
	report_path = output_folder / f"pdf_packshots_{timestamp}.xlsx"

	workbook = Workbook()
	worksheet = workbook.active
	worksheet.title = "report"
	worksheet.append(headers)
	for row in rows:
		worksheet.append(row)

	header_fill = PatternFill("solid", fgColor=PALETTE["blue"].lstrip("#"))
	for cell in worksheet[1]:
		cell.fill = header_fill
		cell.font = Font(color="FFFFFF", bold=True)
	worksheet.freeze_panes = "A2"
	worksheet.auto_filter.ref = worksheet.dimensions
	for column in worksheet.columns:
		column_letter = column[0].column_letter
		max_width = max(len(str(cell.value or "")) for cell in column)
		worksheet.column_dimensions[column_letter].width = min(max(max_width + 2, 14), 42)

	workbook.save(report_path)
	return report_path, len(rows)
