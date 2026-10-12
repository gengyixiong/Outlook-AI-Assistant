"""Email body cleaning shared by the indexing pipeline.

This mirrors ``Summary/EmailBodyCleaner.cs`` and extends it with HTML to text
conversion, signature trimming and prompt-injection neutralizing. The goal is
to index the message the user actually wrote, not the repeated quoted history
from earlier replies in the same thread.
"""

from __future__ import annotations

import re
from html.parser import HTMLParser
from typing import List

_QUOTE_MARKERS = (
    "-----Original Message-----",
    "-----原始邮件-----",
    "-----原始消息-----",
)

_HEADER_MARKERS = ("发件人:", "发件人：", "From:")

_HEADER_SIGNALS = (
    "\nSent:",
    "\n发送时间:",
    "\n发送时间：",
    "\n日期:",
    "\n日期：",
    "\nTo:",
    "\n收件人:",
    "\n收件人：",
    "\nSubject:",
    "\n主题:",
    "\n主题：",
    "\n抄送:",
    "\n抄送：",
)

# A line that is only an Outlook horizontal separator.
_SEPARATOR = re.compile(r"^_{10,}\s*$")
_SIGNATURE_MARKER = re.compile(r"^\s*--\s*$")


class _HtmlToText(HTMLParser):
    """Small, dependency-free HTML to text converter for MailItem.HTMLBody."""

    _BLOCK_TAGS = {
        "p", "div", "br", "tr", "li", "h1", "h2", "h3", "h4", "h5", "h6",
        "table", "blockquote", "section", "article",
    }
    _SKIP_TAGS = {"script", "style", "head", "title"}

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self._parts: List[str] = []
        self._skip_depth = 0

    def handle_starttag(self, tag: str, attrs) -> None:  # noqa: D401
        if tag in self._SKIP_TAGS:
            self._skip_depth += 1
        elif tag in self._BLOCK_TAGS:
            self._parts.append("\n")

    def handle_endtag(self, tag: str) -> None:
        if tag in self._SKIP_TAGS and self._skip_depth > 0:
            self._skip_depth -= 1
        elif tag in self._BLOCK_TAGS:
            self._parts.append("\n")

    def handle_data(self, data: str) -> None:
        if self._skip_depth == 0 and data:
            self._parts.append(data)

    def text(self) -> str:
        return "".join(self._parts)


def html_to_text(html: str) -> str:
    if not html:
        return ""
    parser = _HtmlToText()
    try:
        parser.feed(html)
        parser.close()
    except Exception:
        # Malformed HTML must never abort indexing.
        return re.sub(r"<[^>]+>", " ", html)
    value = parser.text()
    value = value.replace("\xa0", " ")
    value = re.sub(r"[ \t]+", " ", value)
    value = re.sub(r"\n{3,}", "\n\n", value)
    return value.strip()


def remove_quoted_history(body: str) -> str:
    """Cut the message at the first high-confidence quoted-reply boundary."""

    value = body or ""
    boundary = len(value)

    for marker in _QUOTE_MARKERS:
        position = _find_marker_at_line_start(value, marker)
        if position >= 0:
            boundary = min(boundary, position)

    header_position = _find_outlook_header_block(value)
    if header_position >= 0:
        boundary = min(boundary, header_position)

    separator_position = _find_separator_line(value)
    if separator_position >= 0:
        boundary = min(boundary, separator_position)

    return value[:boundary].strip()


def trim_signature(body: str) -> str:
    """Remove a trailing '--' signature block when it is clearly a footer."""

    if not body:
        return ""
    lines = body.splitlines()
    for index in range(len(lines) - 1, max(-1, len(lines) - 40), -1):
        if _SIGNATURE_MARKER.match(lines[index]):
            return "\n".join(lines[:index]).strip()
    return body.strip()


def neutralize_prompt_injection(text: str) -> str:
    """Mark instruction-like lines so the model never treats mail as commands.

    The add-in prompt already labels mail as untrusted. This is a defense in
    depth for any content that later flows through an LLM prompt.
    """

    if not text:
        return ""
    patterns = (
        re.compile(r"^\s*(system|assistant|user)\s*:", re.IGNORECASE),
        re.compile(r"ignore (all )?(previous|above) instructions", re.IGNORECASE),
        re.compile(r"忽略(以上|之前|前面)的?(所有)?(指令|指示|要求)"),
    )
    lines = text.splitlines()
    result = []
    for line in lines:
        if any(pattern.search(line) for pattern in patterns):
            result.append("[untrusted-content] " + line)
        else:
            result.append(line)
    return "\n".join(result)


def clean_email_body(raw_body: str, html_body: str = "") -> str:
    """Full cleaning pipeline used before chunking."""

    value = raw_body or ""
    if not value.strip() and html_body:
        value = html_to_text(html_body)

    value = remove_quoted_history(value)
    value = trim_signature(value)
    value = neutralize_prompt_injection(value)
    value = value.replace("\r\n", "\n").replace("\r", "\n")
    value = re.sub(r"\n{3,}", "\n\n", value)
    return value.strip()


def _find_marker_at_line_start(value: str, marker: str) -> int:
    search_from = 0
    while search_from < len(value):
        position = value.find(marker, search_from)
        if position < 0:
            return -1
        if position == 0 or value[position - 1] in "\r\n":
            return position
        search_from = position + len(marker)
    return -1


def _find_separator_line(value: str) -> int:
    offset = 0
    for line in value.splitlines(keepends=True):
        if _SEPARATOR.match(line.rstrip("\r\n")):
            return offset
        offset += len(line)
    return -1


def _find_outlook_header_block(value: str) -> int:
    for marker in _HEADER_MARKERS:
        search_from = 0
        while search_from < len(value):
            position = value.find(marker, search_from)
            if position < 0:
                break
            at_line_start = position == 0 or value[position - 1] in "\r\n"
            window = value[position:position + 800]
            if at_line_start and _count_header_signals(window) >= 2:
                return position
            search_from = position + len(marker)
    return -1


def _count_header_signals(value: str) -> int:
    return sum(1 for signal in _HEADER_SIGNALS if signal in value)
