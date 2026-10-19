"""Deterministic, conservative Word-title line fitting.

Do not estimate widths from string lengths.  All lines are measured with a
known *wide* sans-serif face (DejaVu Sans Bold) and a generous margin to
accommodate Word font substitution and table-cell padding.  Explicit line
breaks are inserted; no individual line may exceed its allocation.

No font files are redistributed or embedded into Word documents.
"""
from __future__ import annotations

from dataclasses import dataclass
from functools import lru_cache
from math import inf

from PIL import ImageFont


@dataclass(frozen=True)
class FittedTitle:
    text: str
    font_size: float
    lines: tuple[str, ...]
    available_width_pt: float
    largest_line_width_pt: float


@lru_cache(maxsize=1)
def _font():
    # System-provided font. It has intentionally wider capital glyphs than
    # most Montserrat/Arial substitutes, preventing accidental Word wrapping.
    for face in ("DejaVuSans-Bold.ttf", "LiberationSans-Bold.ttf"):
        try:
            return ImageFont.truetype(face, 1000)
        except OSError:
            continue
    return None


def measured_width_pt(value: str, font_size: float) -> float:
    font = _font()
    if font is not None:
        return float(font.getlength(value)) / 1000.0 * float(font_size)
    # Last-resort conservative bound on platforms lacking DejaVu/Liberation.
    return sum(1.10 if c in "MW@QO" else .90 if c != " " else .42
               for c in value) * float(font_size)


def _split_unbreakable(words: list[str], font_size: float, max_width: float) -> list[str]:
    """Split unusually long identifiers rather than let one word overflow."""
    output = []
    for word in words:
        if measured_width_pt(word, font_size) <= max_width:
            output.append(word)
            continue
        buf = ""
        for char in word:
            candidate = buf + char
            if buf and measured_width_pt(candidate, font_size) > max_width:
                output.append(buf)
                buf = char
            else:
                buf = candidate
        if buf:
            output.append(buf)
    return output


def _balance_lines(words: list[str], line_count: int, font_size: float, max_width: float):
    """Optimal word-boundary line breaks for a *fixed* number of lines."""
    n = len(words)
    if n < line_count:
        return None
    memo: dict[tuple[int, int], tuple[float, list[str]]] = {}

    def solve(start: int, remaining: int):
        key = (start, remaining)
        if key in memo:
            return memo[key]
        if remaining == 1:
            line = " ".join(words[start:])
            if measured_width_pt(line, font_size) > max_width:
                return None
            # Penalize lonely short last lines, especially '& PLACES' or a
            # single word on its own. Prefer visually balanced two-line titles.
            width = measured_width_pt(line, font_size)
            score = ((max_width - width) / max_width) ** 2
            memo[key] = (score, [line])
            return memo[key]

        best = None
        for stop in range(start + 1, n - remaining + 2):
            line = " ".join(words[start:stop])
            width = measured_width_pt(line, font_size)
            if width > max_width:
                break
            if words[stop-1] in ("&", "AND", "OR", "OF", "THE"):
                continue
            rest = solve(stop, remaining - 1)
            if rest is None:
                continue
            score, rest_lines = rest
            line_penalty = ((max_width - width) / max_width) ** 2
            if len(line.split()) == 1:
                line_penalty += .12
            total = line_penalty + score
            if best is None or total < best[0]:
                best = total, [line] + rest_lines
        if best is not None:
            memo[key] = best
        return best

    return solve(0, line_count)


def fit_title(
    heading: str,
    *,
    cell_width_pt: float,
    preferred_size: float = 27.0,
    cell_padding_pt: float = 22.0,
    min_size: float = 11.0,
    max_lines: int = 10,
) -> FittedTitle:
    """Size and wrap any reasonably sized heading inside its original Word cell.

    Short titles prioritize one line; longer titles expand to two or more
    deliberately balanced lines. The cell height may grow; no silent clipping
    or automatic Word-wrapping is relied upon.
    """
    heading = " ".join((heading or "").split())
    if not heading:
        raise ValueError("Tiêu đề không được để trống.")
    if cell_width_pt <= 0:
        raise ValueError("Không xác định được chiều rộng ô tiêu đề Word.")

    # 0.84 accounts for different glyph metrics and inconsistent Word table
    # auto-fit behavior. It is more robust than the prior length heuristic.
    usable = min(cell_width_pt - cell_padding_pt, cell_width_pt * .84)
    if usable < 110:
        raise ValueError("Vùng hiển thị tiêu đề trong template quá hẹp.")

    preferred = max(float(min_size), float(preferred_size))
    words = heading.split()

    # An ordinary heading is better as one reasonably sized line than as a
    # single oversized line with a lonely final word wrapped by Word.
    for n_lines, floor in ((1, 18.0), (2, 17.0), (3, 15.0),
                           (4, 13.0), (5, 12.0), (6, min_size),
                           (7, min_size), (8, min_size), (9, min_size),
                           (10, min_size)):
        if n_lines > max_lines:
            break
        threshold = max(min_size, floor)
        sizes = [round(preferred - i * .5, 1)
                 for i in range(int((preferred - threshold) / .5) + 1)]
        for size in sizes:
            parts = _split_unbreakable(words, size, usable)
            if len(parts) < n_lines:
                continue
            candidate = _balance_lines(parts, n_lines, size, usable)
            if candidate is None:
                continue
            _, lines = candidate
            largest = max(measured_width_pt(line, size) for line in lines)
            return FittedTitle(
                text="\n".join(lines), font_size=size,
                lines=tuple(lines), available_width_pt=usable,
                largest_line_width_pt=largest,
            )
    raise ValueError(
        "Tiêu đề quá dài để dàn trang an toàn trong template hiện tại. "
        "Hãy rút ngắn tiêu đề hoặc sử dụng từ viết tắt."
    )
