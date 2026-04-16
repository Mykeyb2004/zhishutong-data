from __future__ import annotations

import json
import re
from dataclasses import dataclass
from pathlib import Path


DEFAULT_RULES_PATH = Path(__file__).with_name("mapping_rules.json")


@dataclass
class MappingRules:
    question_id_regex: re.Pattern[str]
    question_header_patterns: list[re.Pattern[str]]
    subquestion_token_regex: re.Pattern[str]
    binary_flag_pairs: set[tuple[str, str]]
    termination_flag_pairs: set[tuple[str, str]]
    open_text_keywords: list[str]
    other_text_keywords: list[str]
    none_of_above_keywords: list[str]
    other_text_selected_prefixes: list[str]

    def is_question_header(self, header: str) -> bool:
        return any(pattern.search(header) for pattern in self.question_header_patterns)

    def extract_question_id(self, header: str) -> str:
        match = self.question_id_regex.search(header)
        return match.group(1) if match else header

    def is_open_text_header(self, header: str) -> bool:
        return any(keyword in header for keyword in self.open_text_keywords)

    def is_other_text_header(self, header: str) -> bool:
        return any(keyword in header for keyword in self.other_text_keywords)

    def is_none_of_above(self, text: str) -> bool:
        return any(keyword in text for keyword in self.none_of_above_keywords)

    def selected_other_prefix(self, value: str) -> bool:
        return any(value.startswith(prefix) for prefix in self.other_text_selected_prefixes)

    def extract_subquestion_tokens(self, header: str) -> list[str]:
        return self.subquestion_token_regex.findall(header)


def load_rules(path: str | Path | None = None) -> MappingRules:
    rules_path = Path(path) if path is not None else DEFAULT_RULES_PATH
    payload = json.loads(rules_path.read_text(encoding="utf-8"))
    return MappingRules(
        question_id_regex=re.compile(payload["question_id_regex"]),
        question_header_patterns=[re.compile(item) for item in payload["question_header_patterns"]],
        subquestion_token_regex=re.compile(payload["subquestion_token_regex"]),
        binary_flag_pairs={tuple(item) for item in payload["binary_flag_pairs"]},
        termination_flag_pairs={tuple(item) for item in payload["termination_flag_pairs"]},
        open_text_keywords=list(payload["open_text_keywords"]),
        other_text_keywords=list(payload["other_text_keywords"]),
        none_of_above_keywords=list(payload["none_of_above_keywords"]),
        other_text_selected_prefixes=list(payload["other_text_selected_prefixes"]),
    )
