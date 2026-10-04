import re
from dataclasses import dataclass
from pathlib import Path
from typing import Optional


def sanitize_filename(name: str, replacement: str = "_") -> str:
    name = name.strip()
    name = re.sub(r'[<>:"/\\|?*]+', replacement, name)
    name = re.sub(r"\s+", " ", name).strip()
    name = name.rstrip(". ").strip()
    return name or "export"


_TERM_RE = re.compile(r"\b(\d{2}[WS])\b", re.IGNORECASE)


@dataclass(frozen=True)
class GroupParts:
    prefix: str
    term: Optional[str]


def parse_group(group_value: str) -> GroupParts:
    s = (group_value or "").strip()
    if not s:
        return GroupParts(prefix="UNKNOWN", term=None)
    prefix = s.split()[0].upper()
    m = _TERM_RE.search(s)
    term = m.group(1).upper() if m else None
    return GroupParts(prefix=prefix, term=term)


def build_output_path(base_out_dir: str, group_value: str, filename: str) -> str:
    parts = parse_group(group_value)
    level1 = sanitize_filename(parts.prefix)
    level2 = sanitize_filename(f"{parts.prefix}{parts.term}") if parts.term else "NO_TERM"
    return str(Path(base_out_dir) / level1 / level2 / filename)
