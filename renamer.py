"""
Renamer Module

Core preview and rename logic. It corrects the filename serial prefix and can
add matched road type and distance when filename text is missing them.
"""

from __future__ import annotations

import os
import re
from dataclasses import dataclass
from typing import Dict, List, Tuple

from excel_reader import ExcelReader
from logger import RenameLogger
from matcher import FileMatcher


@dataclass
class FileRenameInfo:
    """Information about a file rename operation."""

    filename: str
    new_filename: str
    status: str
    excel_row: int
    matched_data: Dict
    reason: str
    score: float


class VBOXRenamer:
    """Core renaming logic for VBOX files."""

    SERIAL_FORMATS = {
        "decimal": "1.0",
        "integer": "1",
        "padded": "001",
        # Backward-compatible keys used by older UI builds.
        "format_1": "1.0",
        "format_2": "1",
        "format_3": "001",
    }

    def __init__(self, folder_path: str, excel_reader: ExcelReader):
        self.folder_path = folder_path
        self.excel_reader = excel_reader
        self.matcher = FileMatcher()
        self.logger = RenameLogger(folder_path)
        self.preview_data: List[FileRenameInfo] = []
        self.serial_format = "integer"

    def set_serial_format(self, format_key: str) -> None:
        if format_key in self.SERIAL_FORMATS:
            self.serial_format = format_key

    def get_files(self, extensions: List[str] = None, include_subfolders: bool = False) -> List[str]:
        """Get VBOX files as paths relative to the selected folder."""
        if extensions is None:
            extensions = [".vbo"]

        clean_extensions = []
        for extension in extensions:
            extension = extension.strip()
            if not extension:
                continue
            clean_extensions.append(extension if extension.startswith(".") else f".{extension}")

        files: List[str] = []
        if include_subfolders:
            for root, _, filenames in os.walk(self.folder_path):
                for filename in filenames:
                    if self._has_allowed_extension(filename, clean_extensions):
                        full_path = os.path.join(root, filename)
                        files.append(os.path.relpath(full_path, self.folder_path))
        else:
            for filename in os.listdir(self.folder_path):
                full_path = os.path.join(self.folder_path, filename)
                if os.path.isfile(full_path) and self._has_allowed_extension(filename, clean_extensions):
                    files.append(filename)

        return sorted(files, key=str.lower)

    def generate_preview(
        self, extensions: List[str] = None, include_subfolders: bool = False
    ) -> List[FileRenameInfo]:
        """Generate a dry-run preview without renaming files."""
        self.preview_data = []
        excel_data = self._with_route_order_metadata(self.excel_reader.get_data().to_dict("records"))
        files = self.get_files(extensions, include_subfolders)

        for filename in files:
            preview_info = self._preview_one_file(filename, excel_data)
            self.preview_data.append(preview_info)

        self._assign_continuous_file_serials()
        self.preview_data.sort(key=self._preview_order_key)
        return self.preview_data

    @staticmethod
    def _with_route_order_metadata(rows: List[Dict]) -> List[Dict]:
        """Keep Excel order for matching and assign file serials later."""
        ordered_rows: List[Dict] = []
        for index, row in enumerate(rows, start=1):
            ordered_row = dict(row)
            ordered_row["original_serial"] = ordered_row.get("serial", "")
            ordered_row["route_order"] = index
            ordered_rows.append(ordered_row)
        return ordered_rows

    def rename_files(self, rename_list: List[FileRenameInfo] = None) -> Tuple[int, str]:
        """Rename matched files after writing the planned mapping log."""
        if rename_list is None:
            rename_list = [item for item in self.preview_data if item.status == "matched"]

        self.logger.clear_logs()
        log_entries = []
        for file_info in rename_list:
            status = "planned" if file_info.status == "matched" else "skipped"
            log_entries.append(
                self.logger.add_entry(
                    old_name=file_info.filename,
                    new_name=file_info.new_filename,
                    status=status,
                    excel_row=file_info.excel_row,
                    matched_data=file_info.matched_data,
                    reason=file_info.reason,
                )
            )
        self.logger.save_logs()

        successful = 0
        errors: List[str] = []

        for index, file_info in enumerate(rename_list):
            entry = log_entries[index]
            if file_info.status != "matched":
                entry["status"] = "skipped"
                continue

            old_path = self._absolute_path(file_info.filename)
            new_path = self._absolute_path(file_info.new_filename)

            try:
                if self._same_path(old_path, new_path):
                    entry["status"] = "skipped"
                    entry["reason"] = "Filename already has the correct serial prefix"
                    continue

                if not os.path.exists(old_path):
                    entry["status"] = "error"
                    entry["reason"] = "Source file no longer exists"
                    errors.append(f"{file_info.filename}: source file no longer exists")
                    continue

                if os.path.exists(new_path):
                    entry["status"] = "conflict"
                    entry["reason"] = "Target filename already exists"
                    errors.append(f"{file_info.filename}: target already exists")
                    continue

                os.rename(old_path, new_path)
                entry["status"] = "success"
                entry["reason"] = "Successfully renamed"
                successful += 1
            except Exception as exc:
                entry["status"] = "error"
                entry["reason"] = str(exc)
                errors.append(f"{file_info.filename}: {exc}")

        self.logger.save_logs()

        summary = f"Renamed {successful}/{len(rename_list)} files"
        if errors:
            summary += f"\n{len(errors)} errors:\n" + "\n".join(errors)
        return successful, summary

    def undo_last_rename(self) -> Tuple[bool, str]:
        """Undo successful renames from rename_log.json."""
        success, message = self.logger.load_from_json(self.logger.json_log_path)
        if not success:
            return False, message

        undo_list = self.logger.get_undo_list()
        if not undo_list:
            return False, "No successful renames to undo"

        undone = 0
        errors: List[str] = []

        for current_name, original_name in undo_list:
            current_path = self._absolute_path(current_name)
            original_path = self._absolute_path(original_name)
            try:
                if not os.path.exists(current_path):
                    errors.append(f"{current_name}: file not found")
                    continue
                if os.path.exists(original_path):
                    errors.append(f"{original_name}: original target already exists")
                    continue
                os.rename(current_path, original_path)
                undone += 1
            except Exception as exc:
                errors.append(f"{current_name}: {exc}")

        summary = f"Undone {undone}/{len(undo_list)} renames"
        if errors:
            summary += f"\n{len(errors)} errors:\n" + "\n".join(errors)
        return undone > 0, summary

    def get_preview_summary(self) -> Dict[str, int]:
        summary = {
            "total": len(self.preview_data),
            "matched": 0,
            "duplicate": 0,
            "conflict": 0,
            "not_found": 0,
            "unchanged": 0,
        }
        for item in self.preview_data:
            summary[item.status] = summary.get(item.status, 0) + 1
        return summary

    def _preview_one_file(self, filename: str, excel_data: List[Dict]) -> FileRenameInfo:
        matches = self.matcher.find_all_matches_for_file(filename, excel_data)
        reliable_matches = [match for match in matches if match[3]]

        if not reliable_matches:
            reason = matches[0][4] if matches else "No matching Excel row found"
            score = matches[0][1] if matches else 0.0
            return FileRenameInfo(filename, filename, "not_found", -1, {}, reason, score)

        best_idx, best_score, matched_data, _, match_reason = reliable_matches[0]
        if len(reliable_matches) > 1 and abs(reliable_matches[0][1] - reliable_matches[1][1]) <= 0.03:
            return FileRenameInfo(
                filename,
                filename,
                "duplicate",
                best_idx,
                matched_data,
                "Multiple Excel rows match this file equally well",
                best_score,
            )

        return FileRenameInfo(filename, filename, "matched", best_idx, matched_data, match_reason, best_score)

    def _assign_continuous_file_serials(self) -> None:
        """Assign 1, 2, 3... to each matched file in Excel route order."""
        matched_items = [item for item in self.preview_data if item.status == "matched"]
        matched_items.sort(key=self._matched_file_order_key)

        new_names_seen: Dict[str, str] = {}
        for file_serial, item in enumerate(matched_items, start=1):
            matched_data = dict(item.matched_data)
            matched_data["serial"] = file_serial
            item.matched_data = matched_data

            new_filename = self._generate_new_filename(item.filename, matched_data)
            old_path = self._absolute_path(item.filename)
            new_path = self._absolute_path(new_filename)
            status = "matched"
            reason = item.reason

            target_key = self._path_key(new_filename)
            if self._same_path(old_path, new_path):
                status = "unchanged"
                reason = "Serial prefix is already correct"
            elif target_key in new_names_seen:
                status = "duplicate"
                reason = f"Duplicate target name also created by {new_names_seen[target_key]}"
            elif os.path.exists(new_path):
                status = "conflict"
                reason = "Target filename already exists"

            new_names_seen[target_key] = item.filename
            item.new_filename = new_filename
            item.status = status
            item.reason = reason

    def _matched_file_order_key(self, item: FileRenameInfo):
        route_order = item.matched_data.get("route_order", item.excel_row + 1)
        return (
            self._safe_int(route_order, item.excel_row + 1),
            self._part_order_key(item.filename),
            item.filename.lower(),
        )

    @staticmethod
    def _preview_order_key(item: FileRenameInfo):
        serial = VBOXRenamer._safe_int(item.matched_data.get("serial"), None)
        if serial is not None:
            return (0, serial, item.filename.lower())
        return (1, item.filename.lower())

    @staticmethod
    def _part_order_key(filename: str) -> Tuple[int, int, str]:
        basename = os.path.basename(filename)
        match = re.search(r"\bpart[\s_/-]*(\d+)(?:[\s_/-]*(\d+))?", basename, flags=re.IGNORECASE)
        if not match:
            return (999999, 999999, basename.lower())
        first = VBOXRenamer._safe_int(match.group(1), 999999)
        second = VBOXRenamer._safe_int(match.group(2), 999999)
        return (first, second, basename.lower())

    def _generate_new_filename(self, old_filename: str, matched_data: Dict) -> str:
        directory, basename = os.path.split(old_filename)
        _, remainder = self._extract_old_serial(basename)
        remainder = self._add_missing_road_type(remainder, matched_data.get("road_type"))
        remainder = self._add_missing_distance(remainder, matched_data.get("distance"))
        new_serial = matched_data.get("serial", "")
        formatted_serial = self._format_serial(new_serial)
        new_basename = f"{formatted_serial}_{remainder}"
        return os.path.join(directory, new_basename) if directory else new_basename

    def _add_missing_road_type(self, filename: str, road_type) -> str:
        road_type_text = self._format_road_type(road_type)
        if not road_type_text:
            return filename

        _, file_road, _, _, _ = self.matcher.extract_filename_parts(filename, str(road_type))
        if file_road and self.matcher.similarity_score(file_road, road_type) >= 0.8:
            return filename

        return f"{road_type_text}_{filename.lstrip('_- ')}"

    def _add_missing_distance(self, filename: str, distance) -> str:
        if self.matcher._extract_distance_from_text(filename) is not None:
            return filename

        distance_text = self._format_distance(distance)
        if not distance_text:
            return filename

        name_without_ext, extension = os.path.splitext(filename)
        distance_suffix = f"{distance_text}_KM"
        part_match = re.search(r"(\(?\s*part[\s_/-]*\d+(?:[\s_/-]*\d+)?\s*\)?)\s*$", name_without_ext, flags=re.IGNORECASE)
        if part_match:
            before_part = name_without_ext[: part_match.start()].rstrip(" _-")
            part_text = part_match.group(1).strip()
            return f"{before_part}__{distance_suffix}__{part_text}{extension}"

        clean_name = name_without_ext.rstrip(" _-")
        return f"{clean_name}__{distance_suffix}{extension}"

    @staticmethod
    def _extract_old_serial(filename: str) -> Tuple[str, str]:
        if "_" in filename:
            old_serial, remainder = filename.split("_", 1)
            if VBOXRenamer._looks_like_serial(old_serial):
                return old_serial, remainder
        return "", filename

    @staticmethod
    def _looks_like_serial(value: str) -> bool:
        value = str(value).strip()
        return bool(re.fullmatch(r"\d+(?:\.\d+)?", value))

    def _format_serial(self, value) -> str:
        value_text = str(value).strip()
        numeric_value = self._to_number(value_text)
        selected_format = self.SERIAL_FORMATS.get(self.serial_format, "1.0")

        if numeric_value is None:
            return value_text
        if selected_format == "001":
            return f"{int(round(numeric_value)):03d}"
        if selected_format == "1":
            return str(int(round(numeric_value)))
        return f"{numeric_value:.1f}"

    @classmethod
    def _format_distance(cls, value) -> str:
        if value in (None, ""):
            return ""

        value_text = str(value).strip()
        match = re.search(r"\d+(?:[.,]\d+)?", value_text)
        if not match:
            return ""

        numeric_value = cls._to_number(match.group(0).replace(",", "."))
        if numeric_value is None:
            return ""
        if numeric_value.is_integer():
            return str(int(numeric_value))
        return f"{numeric_value:g}"

    @staticmethod
    def _format_road_type(value) -> str:
        if value in (None, ""):
            return ""

        text = str(value).strip()
        if not text or text.lower() in {"nan", "none"}:
            return ""

        text = re.sub(r"[^A-Za-z0-9]+", "_", text)
        return text.strip("_")

    def _absolute_path(self, relative_filename: str) -> str:
        return os.path.abspath(os.path.join(self.folder_path, relative_filename))

    @staticmethod
    def _has_allowed_extension(filename: str, extensions: List[str]) -> bool:
        return any(filename.lower().endswith(extension.lower()) for extension in extensions)

    @staticmethod
    def _same_path(left: str, right: str) -> bool:
        return os.path.normcase(os.path.abspath(left)) == os.path.normcase(os.path.abspath(right))

    @staticmethod
    def _path_key(relative_filename: str) -> str:
        return os.path.normcase(os.path.normpath(relative_filename))

    @staticmethod
    def _to_number(value: str):
        try:
            return float(value)
        except (TypeError, ValueError):
            return None

    @staticmethod
    def _safe_int(value, default):
        try:
            if value is None or str(value).strip() == "":
                return default
            return int(float(str(value).strip()))
        except (TypeError, ValueError):
            return default
