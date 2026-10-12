"""Text tokenization for chunking and CJK-aware keyword search.

Two concerns live here:

1. ``token_spans`` estimates token boundaries so we can chunk text to a token
   budget without loading a tokenizer into every code path. When a real
   tokenizer is supplied we use it; otherwise we use a deterministic
   approximation that is safe for the 8K-context model.

2. ``keyword_tokens`` produces searchable terms for the FTS layer. SQLite's
   built-in trigram tokenizer only matches substrings of length >= 3
   codepoints, so it cannot match two-character Chinese words ("幕墙") or
   short technical terms ("AI", "3D", "IF"). We therefore build an auxiliary
   token field: CJK runs are split into overlapping bigrams and single
   characters, while latin/digit runs are split on non-word boundaries. FTS5
   then indexes this normalized field with the unicode61 tokenizer.
"""

from __future__ import annotations

import re
import unicodedata
from typing import Iterable, List, Optional, Sequence, Tuple

# CJK Unified Ideographs + extensions, Hiragana, Katakana, Hangul.
_CJK_RANGES = (
    (0x3040, 0x30FF),   # Hiragana + Katakana
    (0x3400, 0x4DBF),   # CJK Ext A
    (0x4E00, 0x9FFF),   # CJK Unified
    (0xF900, 0xFAFF),   # CJK Compatibility
    (0xAC00, 0xD7AF),   # Hangul syllables
    (0x20000, 0x2FA1F),  # CJK Ext B-F
)


def _is_cjk(char: str) -> bool:
    code = ord(char)
    for start, end in _CJK_RANGES:
        if start <= code <= end:
            return True
    return False


_LATIN_RUN = re.compile(r"[0-9A-Za-z_@.+-]+")


def normalize_for_search(text: str) -> str:
    """Normalize width/full-width forms and collapse whitespace."""

    if not text:
        return ""
    # NFKC turns full-width latin digits/punctuation into ASCII, which makes
    # "ＩＦＣ４" and "IFC4" match.
    value = unicodedata.normalize("NFKC", text)
    value = value.replace("\u200b", " ").replace("\ufeff", " ")
    return value


def cjk_bigrams(token: str) -> List[str]:
    """Split a CJK run into overlapping bigrams plus single characters.

    Overlapping bigrams are what make two-character queries work with FTS5:
    "幕墙玻璃" yields 幕墙/墙玻/玻璃 and the single characters, so a query for
    "幕墙" or "玻璃" both match.
    """

    chars = list(token)
    if len(chars) == 1:
        return [token]
    result: List[str] = []
    for index in range(len(chars) - 1):
        result.append(chars[index] + chars[index + 1])
    # Include singles so a single-character query still matches.
    result.extend(chars)
    return result


def keyword_tokens(text: str) -> List[str]:
    """Return normalized search terms suitable for the FTS token field."""

    if not text:
        return []

    value = normalize_for_search(text)
    tokens: List[str] = []
    buffer: List[str] = []

    def flush_cjk() -> None:
        if buffer:
            run = "".join(buffer)
            tokens.extend(cjk_bigrams(run))
            buffer.clear()

    latin: List[str] = []

    def flush_latin() -> None:
        if latin:
            tokens.append("".join(latin))
            latin.clear()

    for char in value:
        if _is_cjk(char):
            flush_latin()
            buffer.append(char)
        elif char.isalnum() or char in "_.@+-":
            flush_cjk()
            latin.append(char)
        else:
            flush_cjk()
            flush_latin()

    flush_cjk()
    flush_latin()

    # Lowercase latin terms so search is case-insensitive.
    return [token.lower() for token in tokens if token]


def build_token_field(text: str) -> str:
    """Join keyword tokens into a single space-separated FTS column."""

    return " ".join(keyword_tokens(text))


def token_spans(text: str) -> List[Tuple[int, int]]:
    """Approximate token boundaries as character spans.

    The approximation treats each CJK codepoint and each latin word as one
    token. This tracks the real tokenizer closely enough for chunk budgeting
    and never over-counts badly enough to overflow the model's 8K window.
    """

    if not text:
        return []

    spans: List[Tuple[int, int]] = []
    index = 0
    length = len(text)
    while index < length:
        char = text[index]
        if char.isspace():
            index += 1
            continue
        if _is_cjk(char):
            spans.append((index, index + 1))
            index += 1
            continue
        if char.isalnum() or char in "_.@+-":
            start = index
            while index < length:
                current = text[index]
                if current.isspace() or _is_cjk(current):
                    break
                if not (current.isalnum() or current in "_.@+-"):
                    break
                index += 1
            spans.append((start, index))
            continue
        # Punctuation: count as a token boundary but not a token.
        index += 1
    return spans


def count_tokens(text: str) -> int:
    return len(token_spans(text))


def chunk_text(
    text: str,
    target_tokens: int = 512,
    min_tokens: int = 300,
    max_tokens: int = 800,
    overlap_tokens: int = 64,
) -> List[str]:
    """Split text into overlapping chunks bounded by a token budget.

    Splitting prefers paragraph, then sentence, then hard token boundaries so
    that a chunk does not cut a sentence in the middle when avoidable.
    """

    if text is None:
        return []
    value = text.strip()
    if not value:
        return []

    spans = token_spans(value)
    if not spans:
        return []
    if len(spans) <= max_tokens:
        return [value]

    chunks: List[str] = []
    start_token = 0
    total = len(spans)

    while start_token < total:
        end_token = min(start_token + target_tokens, total)
        if total - end_token < min_tokens:
            end_token = total

        char_start = spans[start_token][0]
        char_end = spans[end_token - 1][1]

        # Try to end on a paragraph or sentence boundary within the window.
        if end_token < total:
            adjusted = _find_boundary(value, char_start, char_end)
            if adjusted > char_start:
                char_end = adjusted
                end_token = _token_index_at_or_before(spans, char_end)

        chunk = value[char_start:char_end].strip()
        if chunk:
            chunks.append(chunk)

        if end_token >= total:
            break

        next_start = end_token - overlap_tokens
        if next_start <= start_token:
            next_start = end_token
        start_token = next_start

    return chunks


def _token_index_at_or_before(
    spans: Sequence[Tuple[int, int]],
    char_pos: int,
) -> int:
    result = 0
    for index, (start, end) in enumerate(spans):
        if end <= char_pos:
            result = index + 1
        else:
            break
    return result


def _find_boundary(text: str, char_start: int, char_end: int) -> int:
    """Look backwards from char_end for a paragraph/sentence boundary."""

    window_start = max(char_start + 1, char_end - 400)
    for marker in ("\n\n", "\n", "。", "！", "？", ". ", "! ", "? "):
        position = text.rfind(marker, window_start, char_end)
        if position > char_start:
            return position + len(marker)
    return char_end


def deduplicate_chunks(chunks: Iterable[str]) -> List[str]:
    """Drop chunks that are exact or near-exact duplicates.

    Conversation histories repeat quoted text; indexing the same content twice
    would let one long thread dominate semantic results.
    """

    seen: List[str] = []
    result: List[str] = []
    for chunk in chunks:
        fingerprint = _fingerprint(chunk)
        if not fingerprint:
            continue
        if any(fingerprint in existing or existing in fingerprint
               for existing in seen):
            continue
        seen.append(fingerprint)
        result.append(chunk)
    return result


def _fingerprint(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip().lower()[:400]
