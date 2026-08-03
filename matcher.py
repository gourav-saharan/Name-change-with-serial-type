"""
File Matcher Module

Matches VBOX filenames to Excel route rows while tolerating case, spacing,
underscores, hyphens, dots, abbreviations, and small spelling differences.

"""

from __future__ import annotations

import os
import re
from difflib import SequenceMatcher
from typing import Dict, List, Optional, Tuple


class FileMatcher:
    """Matches VBOX files with Excel route data."""

    def __init__(self):
        self.match_threshold = 0.74

    def normalize_text(self, text) -> str:
        """Normalize text for comparison."""
        if text is None:
            return ""

        normalized = str(text).lower()
        replacements = {
            "&": " and ",
            "madhya pradesh": " mp ",
            "m.p.": " mp ",
            "m p": " mp ",
            "maharashtra": " mh ",
            "maha rashtra": " mh ",
            "nashik": " nasik ",
            "devla": " devala ",
            "malegoan": " malegaon ",
            "julwaniya": " julwania ",
            "julvaniya": " julwania ",
            "palsner": " palasner ",
            "palsaner": " palasner ",
            "shirpoor": " shirpur ",
            "songeer": " songir ",
            "pimpalgaon": " pimplgaon ",
            "baswant": " basvant ",
        }
        for old, new in replacements.items():
            normalized = normalized.replace(old, new)

        normalized = re.sub(r"[_\-.(),/\\]+", " ", normalized)
        normalized = re.sub(r"[^a-z0-9 ]+", " ", normalized)
        normalized = " ".join(normalized.split())
        return normalized

    def similarity_score(self, left, right) -> float:
        """Calculate a robust similarity score between two strings."""
        norm_left = self.normalize_text(left)
        norm_right = self.normalize_text(right)
        if not norm_left or not norm_right:
            return 0.0
        if norm_left == norm_right:
            return 1.0
        if norm_left in norm_right or norm_right in norm_left:
            shorter = min(len(norm_left), len(norm_right))
            longer = max(len(norm_left), len(norm_right))
            return max(0.9, shorter / longer)

        sequence_score = SequenceMatcher(None, norm_left, norm_right).ratio()
        token_score = self._token_overlap_score(norm_left, norm_right)
        token_sort_score = self._token_sort_score(norm_left, norm_right)
        fuzzy_token_score = self._fuzzy_token_score(norm_left, norm_right)
        compact_score = SequenceMatcher(
            None, norm_left.replace(" ", ""), norm_right.replace(" ", "")
        ).ratio()
        return max(sequence_score, token_score, token_sort_score, fuzzy_token_score, compact_score)

    def extract_filename_parts(
        self, filename: str, expected_road_type: str = ""
    ) -> Tuple[Optional[str], Optional[str], Optional[str], Optional[str], Optional[float]]:
        """
        Extract serial, road type, start, end, and distance from a VBOX filename.

        Expected format:
        {serial}_{road_type}_{start}_TO_{end}_{distance}_KM_part_{part}.{ext}
        """
        base_name = os.path.basename(filename)
        name_without_ext, _ = os.path.splitext(base_name)
        serial = ""
        remainder = name_without_ext

        if "_" in name_without_ext:
            possible_serial, possible_remainder = name_without_ext.split("_", 1)
            if self._looks_like_serial(possible_serial):
                serial = possible_serial
                remainder = possible_remainder

        parsed = self._extract_from_text_pattern(remainder, expected_road_type)
        if parsed:
            road_type, start_point, end_point, distance = parsed
            if distance is None:
                distance = self._extract_distance_from_text(remainder)
            return serial, road_type, start_point, end_point, distance

        tokens = [token for token in remainder.split("_") if token]
        to_index = self._find_to_index(tokens)
        if to_index <= 0:
            return serial, None, None, None, None

        left_tokens = tokens[:to_index]
        right_tokens = tokens[to_index + 1 :]

        distance = self._extract_distance(tokens)
        end_tokens = self._trim_end_tokens(right_tokens)
        road_tokens, start_tokens = self._split_road_and_start(left_tokens, expected_road_type)

        road_type = "_".join(road_tokens)
        start_point = "_".join(start_tokens)
        end_point = "_".join(end_tokens)

        if not start_point or not end_point:
            return serial, road_type, None, None, distance

        return serial, road_type, start_point, end_point, distance

    def match_file_with_row(self, filename: str, excel_row: Dict) -> Tuple[bool, float, str]:
        """Match a file against one Excel row."""
        excel_start = excel_row.get("start_point", "")
        excel_end = excel_row.get("end_point", "")
        excel_road = excel_row.get("road_type", "")
        excel_distance = excel_row.get("distance")

        _, file_road, file_start, file_end, file_distance = self.extract_filename_parts(
            filename, str(excel_road)
        )

        if file_start is None or file_end is None:
            return False, 0.0, "Could not parse filename structure"

        start_score = self.similarity_score(file_start, excel_start)
        end_score = self.similarity_score(file_end, excel_end)
        road_score = self.similarity_score(file_road, excel_road) if file_road else 1.0

        distance_score = 1.0
        distance_reason = ""
        if file_distance is not None and excel_distance not in (None, ""):
            excel_distance_float = self._to_float(excel_distance)
            if excel_distance_float is not None:
                difference = abs(file_distance - excel_distance_float)
                distance_score = 1.0 if difference <= 0.5 else max(0.0, 1.0 - difference / 10.0)
                if difference > 0.5:
                    distance_reason = f"Distance differs ({file_distance:g} vs {excel_distance_float:g})"

        score = (start_score * 0.35) + (end_score * 0.35) + (road_score * 0.2) + (distance_score * 0.1)
        has_distance = file_distance is not None and excel_distance not in (None, "")
        location_threshold = self.match_threshold if has_distance else 0.78
        matched = (
            start_score >= location_threshold
            and end_score >= location_threshold
            and road_score >= 0.68
            and distance_score >= 0.85
            and score >= 0.80
        )

        if matched:
            if start_score < 0.9 or end_score < 0.9:
                return True, score, "Fuzzy matched with minor text difference"
            return True, score, "Matched"

        reasons = []
        if start_score < location_threshold:
            reasons.append(f"Start point mismatch ({start_score:.0%})")
        if end_score < location_threshold:
            reasons.append(f"End point mismatch ({end_score:.0%})")
        if road_score < 0.68:
            reasons.append(f"Road type mismatch ({road_score:.0%})")
        if distance_reason:
            reasons.append(distance_reason)

        return False, score, " | ".join(reasons) or "No reliable match"

    def _extract_from_text_pattern(
        self, text: str, expected_road_type: str = ""
    ) -> Optional[Tuple[str, str, str, Optional[float]]]:
        """
        Parse both modern and screenshot-style names:
        - Smooth_Highway_START_TO_END_48.6_KM_part_001
        - START__to__END__(part_1_3)
        - Highway_Road_START To END (Part_1_1)
        """
        distance = self._extract_distance_from_text(text)
        cleaned = self._remove_part_and_distance(text)
        cleaned = re.sub(r"[_]+", " ", cleaned)
        cleaned = re.sub(r"\s+", " ", cleaned).strip(" _-.,")

        pieces = re.split(r"\bto\b", cleaned, maxsplit=1, flags=re.IGNORECASE)
        if len(pieces) != 2:
            return None

        left_text = pieces[0].strip(" _-.,")
        right_text = pieces[1].strip(" _-.,")
        if not left_text or not right_text:
            return None

        road_type, start_point = self._split_road_and_start_text(left_text, expected_road_type)
        end_point = right_text.strip(" _-.,")
        if not start_point or not end_point:
            return None

        return road_type, start_point, end_point, distance

    def _split_road_and_start_text(self, left_text: str, expected_road_type: str) -> Tuple[str, str]:
        normalized_left = self.normalize_text(left_text)
        normalized_road = self.normalize_text(expected_road_type)

        if normalized_road and normalized_left.startswith(normalized_road):
            road_word_count = len(normalized_road.split())
            original_words = left_text.split()
            return " ".join(original_words[:road_word_count]), " ".join(original_words[road_word_count:])

        generic_prefixes = [
            "smooth highway",
            "rough highway",
            "hilly road",
            "city road",
            "highway road",
            "national highway",
            "state highway",
        ]
        for prefix in generic_prefixes:
            if normalized_left.startswith(prefix):
                word_count = len(prefix.split())
                original_words = left_text.split()
                return " ".join(original_words[:word_count]), " ".join(original_words[word_count:])

        return "", left_text

    @staticmethod
    def _remove_part_and_distance(text: str) -> str:
        cleaned = re.sub(
            r"\(?\s*part[\s_/-]*\d+(?:[\s_/-]*\d+)?\s*\)?",
            " ",
            text,
            flags=re.IGNORECASE,
        )
        cleaned = re.sub(
            r"(?<![\d.])\d+(?:[.,]\d+)?\s*[_ ]*km(?=$|[^a-z0-9])",
            " ",
            cleaned,
            flags=re.IGNORECASE,
        )
        return cleaned

    @staticmethod
    def _extract_distance_from_text(text: str) -> Optional[float]:
        match = re.search(
            r"(?<![\d.])(\d+(?:[.,]\d+)?)\s*[_ ]*km(?=$|[^a-z0-9])",
            text,
            flags=re.IGNORECASE,
        )
        if match:
            return FileMatcher._to_float(match.group(1))
        return None

    def find_all_matches_for_file(self, filename: str, excel_data: List[Dict]) -> List[Tuple[int, float, Dict, bool, str]]:
        """Find and sort candidate matches for one file."""
        matches: List[Tuple[int, float, Dict, bool, str]] = []
        for idx, row in enumerate(excel_data):
            matched, score, reason = self.match_file_with_row(filename, row)
            if matched or score >= 0.68:
                matches.append((idx, score, row, matched, reason))

        matches.sort(key=lambda item: item[1], reverse=True)
        return matches

    def _split_road_and_start(self, left_tokens: List[str], expected_road_type: str) -> Tuple[List[str], List[str]]:
        road_type_tokens = self.normalize_text(expected_road_type).split()
        normalized_left_tokens = [self.normalize_text(token) for token in left_tokens]

        if road_type_tokens:
            compact_left = " ".join(normalized_left_tokens)
            compact_road = " ".join(road_type_tokens)
            if compact_left.startswith(compact_road):
                road_count = len(road_type_tokens)
                return left_tokens[:road_count], left_tokens[road_count:]

        # Common VBOX names put road type first, usually as two words:
        # Smooth_Highway, Hilly_Road, Rough_Highway, City_Road.
        if len(left_tokens) >= 3:
            return left_tokens[:2], left_tokens[2:]
        if len(left_tokens) >= 2:
            return left_tokens[:1], left_tokens[1:]
        return [], left_tokens

    @staticmethod
    def _find_to_index(tokens: List[str]) -> int:
        for index, token in enumerate(tokens):
            if token.upper() == "TO":
                return index
        return -1

    @staticmethod
    def _trim_end_tokens(tokens: List[str]) -> List[str]:
        trimmed = []
        for index, token in enumerate(tokens):
            upper = token.upper()
            next_upper = tokens[index + 1].upper() if index + 1 < len(tokens) else ""
            if upper == "PART" or upper.startswith("PART"):
                break
            if upper == "KM":
                break
            if FileMatcher._to_float(token) is not None and next_upper == "KM":
                break
            trimmed.append(token)
        return trimmed

    @staticmethod
    def _extract_distance(tokens: List[str]) -> Optional[float]:
        for index, token in enumerate(tokens):
            if token.upper() == "" and index > 0:
                value = FileMatcher._to_float(tokens[index - 1])
                if value is not None:
                    return value
        joined = "_".join(tokens)
        match = re.search(r"(\d+(?:[.,]\d+)?)_?KM\b", joined, flags=re.IGNORECASE)
        if match:
            return FileMatcher._to_float(match.group(1))
        return None

    @staticmethod
    def _looks_like_serial(value: str) -> bool:
        return bool(re.fullmatch(r"\d+(?:\.\d+)?", str(value).strip()))

    @staticmethod
    def _to_float(value) -> Optional[float]:
        try:
            if value is None or str(value).strip() == "":
                return None
            return float(str(value).strip().replace(",", "."))
        except (TypeError, ValueError):
            return None

    @staticmethod
    def _token_overlap_score(left: str, right: str) -> float:
        left_tokens = set(left.split())
        right_tokens = set(right.split())
        if not left_tokens or not right_tokens:
            return 0.0
        overlap = len(left_tokens & right_tokens)
        return (2 * overlap) / (len(left_tokens) + len(right_tokens))

    @staticmethod
    def _token_sort_score(left: str, right: str) -> float:
        """Compare strings after sorting tokens, useful for word-order changes."""
        left_sorted = " ".join(sorted(left.split()))
        right_sorted = " ".join(sorted(right.split()))
        return SequenceMatcher(None, left_sorted, right_sorted).ratio()

    @staticmethod
    def _fuzzy_token_score(left: str, right: str) -> float:
        """Score tokens using best approximate token matches instead of exact matches."""
        left_tokens = left.split()
        right_tokens = right.split()
        if not left_tokens or not right_tokens:
            return 0.0

        shorter, longer = (left_tokens, right_tokens)
        if len(left_tokens) > len(right_tokens):
            shorter, longer = right_tokens, left_tokens

        scores = []
        for token in shorter:
            best = max(SequenceMatcher(None, token, candidate).ratio() for candidate in longer)
            scores.append(best)

        coverage_penalty = len(shorter) / len(longer)
        return (sum(scores) / len(scores)) * (0.85 + 0.15 * coverage_penalty)
