"""見出し採番マネージャ。

Rust 版 `src/heading.rs` の Python 移植。
H1 → X / H2 → X.X / H3 → X.X.X / H4 → (N) / H5 → ① の形式で番号を生成し、
ユーザーが見出しテキストに既存番号を書いていた場合はそれを尊重してカウンタを同期する。
"""
from __future__ import annotations

from .ir import Inline, inline_to_plain_text


_CIRCLED = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"


class HeadingManager:
    def __init__(self) -> None:
        self.counters = [0, 0, 0, 0, 0]  # h1..h5

    def next_heading(self, level: int, content: list[Inline]) -> str:
        """指定レベルの番号を進めて文字列を返す。
        既に番号が含まれていればそれと同期して返す。"""
        plain_text = "".join(inline_to_plain_text(c) for c in content)
        existing = self._detect_existing_number(level, plain_text)
        if existing is not None:
            self._sync_counters(level, existing)
            return existing
        self._increment(level)
        return self._format_number(level)

    def strip_number(self, level: int, text: str) -> str:
        """既存番号があれば取り除いてタイトルのみ返す。"""
        trimmed = text.strip()
        if self._detect_existing_number(level, trimmed) is None:
            return trimmed
        if level in (1, 2, 3):
            parts = trimmed.split(" ", 1)
            if len(parts) == 2:
                return parts[1]
            return trimmed
        if level == 4:
            end = trimmed.find(")")
            if end != -1:
                return trimmed[end + 1:].lstrip()
            return trimmed
        if level == 5:
            return trimmed[1:].lstrip()
        return trimmed

    def current_h1_number(self) -> int:
        return self.counters[0]

    def _increment(self, level: int) -> None:
        idx = max(0, min(level - 1, 4))
        self.counters[idx] += 1
        for i in range(idx + 1, 5):
            self.counters[i] = 0

    def _format_number(self, level: int) -> str:
        c = self.counters
        if level == 1:
            return f"{c[0]}"
        if level == 2:
            return f"{c[0]}.{c[1]}"
        if level == 3:
            return f"{c[0]}.{c[1]}.{c[2]}"
        if level == 4:
            return f"({c[3]})"
        if level == 5:
            return _num_to_circled(c[4])
        return ""

    def _detect_existing_number(self, level: int, text: str) -> str | None:
        trimmed = text.strip()
        if not trimmed:
            return None
        if level == 1:
            first = trimmed.split()[0] if trimmed.split() else ""
            if first.isdigit():
                return first
            return None
        if level == 2:
            tokens = trimmed.split()
            if not tokens:
                return None
            first = tokens[0]
            parts = first.split(".")
            if len(parts) == 2 and all(p.isdigit() for p in parts):
                return first
            return None
        if level == 3:
            tokens = trimmed.split()
            if not tokens:
                return None
            first = tokens[0]
            parts = first.split(".")
            if len(parts) == 3 and all(p.isdigit() for p in parts):
                return first
            return None
        if level == 4:
            if trimmed.startswith("("):
                end = trimmed.find(")")
                if end != -1:
                    inner = trimmed[1:end]
                    if inner.isdigit():
                        return f"({inner})"
            return None
        if level == 5:
            if _is_circled_number(trimmed[0]):
                return trimmed[0]
            return None
        return None

    def _sync_counters(self, level: int, number: str) -> None:
        if level == 1:
            try:
                n = int(number)
            except ValueError:
                return
            self.counters[0] = n
            for i in range(1, 5):
                self.counters[i] = 0
        elif level == 2:
            parts = number.split(".")
            if len(parts) == 2:
                try:
                    a, b = int(parts[0]), int(parts[1])
                except ValueError:
                    return
                self.counters[0] = a
                self.counters[1] = b
                for i in range(2, 5):
                    self.counters[i] = 0
        elif level == 3:
            parts = number.split(".")
            if len(parts) == 3:
                try:
                    a, b, c = int(parts[0]), int(parts[1]), int(parts[2])
                except ValueError:
                    return
                self.counters[0] = a
                self.counters[1] = b
                self.counters[2] = c
                for i in range(3, 5):
                    self.counters[i] = 0
        elif level == 4:
            inner = number.lstrip("(").rstrip(")")
            try:
                n = int(inner)
            except ValueError:
                return
            self.counters[3] = n
            self.counters[4] = 0
        elif level == 5:
            ch = number[0] if number else "①"
            n = _circled_to_num(ch)
            if n is not None:
                self.counters[4] = n


def _num_to_circled(n: int) -> str:
    if 1 <= n <= 20:
        return _CIRCLED[n - 1]
    return f"({n})"


def _is_circled_number(c: str) -> bool:
    return "①" <= c <= "⑳"  # ① - ⑳


def _circled_to_num(c: str) -> int | None:
    if _is_circled_number(c):
        return ord(c) - 0x245F
    return None
