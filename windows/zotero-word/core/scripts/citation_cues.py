"""Extract positional citation candidates; this module never resolves a paper.

DOI syntax, a matching name/year, and an article-looking URL are search clues.
They do not establish publication identity or support for a manuscript claim.
Offsets are Python character offsets into the exact text supplied by the caller.
"""
from __future__ import annotations

from html import unescape
import re
from urllib.parse import unquote, urlsplit


_DOI_START = re.compile(r"10\.\d{4,9}/", re.I)
_DOI_TOKEN = re.compile(r"(?<![A-Za-z0-9])(?:doi\s*:\s*)?10\.\d{4,9}/[\x21-\x7e]+", re.I)
_URL = re.compile(r"https?://[^\s<>\"“”‘’\u3000-\u303f\u3400-\u9fff\uff00-\uffef]+", re.I)
_YEAR = r"(?:18|19|20)\d{2}[a-z]?"
_YEARS = _YEAR + r"(?:\s*[,，/]\s*(?:" + _YEAR + r"|[a-z]))*"
_NAME = r"[A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÖØ-öø-ÿ'’.-]*"
_EN_AUTHORS = _NAME + r"(?:\s+(?:and|&)\s+" + _NAME + r")*(?:\s+et\s+al\.?|\s*等)?"
_CN_AUTHORS = r"[\u3400-\u9fff]{1,8}(?:等)?"
_AUTHORS = r"(?:" + _EN_AUTHORS + r"|" + _CN_AUTHORS + r")"
_NARRATIVE = re.compile(
    r"(?<![A-Za-zÀ-ž0-9])(?P<authors>" + _AUTHORS + r")\s*[（(]\s*(?P<years>" + _YEARS + r")\s*[)）]"
)
_PAREN = re.compile(r"[（(](?P<content>[^()（）\n]{1,700})[)）]")
_PAREN_ITEM = re.compile(r"\s*(?P<authors>" + _AUTHORS + r")\s*[,，]\s*(?P<years>" + _YEARS + r")\s*\Z")
_NUMERIC = re.compile(r"\[\s*\d+(?:\s*[,，–—-]\s*\d+)*\s*\]")
_CN_PREFIX = re.compile(r"^(?:(?:研究表明|结果表明|根据|依据|参照|参考|参见|据|如|见))+")


def _trim_end(value):
    """Drop prose punctuation and unmatched closing wrappers, not DOI parentheses."""
    previous = None
    pairs = (("(", ")"), ("[", "]"), ("{", "}"), ("<", ">"))
    while value != previous:
        previous = value
        value = value.rstrip(".,;:!?，。；：！？、\"'”’）】》")
        for opening, closing in pairs:
            while value.endswith(closing) and value.count(closing) > value.count(opening):
                value = value[:-1]
    return value


def _zotero_url(value):
    try:
        parsed = urlsplit(value)
    except ValueError:
        return False
    return (parsed.hostname or "").lower() in ("zotero.org", "www.zotero.org")


def _year_values(value):
    years = []
    for token in re.split(r"\s*[,，/]\s*", value):
        if re.fullmatch(_YEAR, token):
            years.append(token)
        elif years and re.fullmatch(r"[a-z]", token):
            years.append(years[-1][:4] + token)
    return years


def normalize_doi(value):
    """Return a syntax-normalized DOI or ''. No network validity is implied.

    Accepts a DOI field, DOI resolver URL, or publisher URL containing a DOI.
    Arbitrary sentences are intentionally not accepted as DOI fields; use
    extract_cues for those. URL query strings and fragments are not DOI suffixes.
    """
    if not isinstance(value, str):
        return ""
    s = unquote(unescape(value)).strip()
    s = s.lstrip("([{<（【《\"'“‘")
    s = re.sub(r"^doi\s*:\s*", "", s, flags=re.I)
    s = s.lstrip("([{<（【《\"'“‘")
    if re.match(r"^(?:dx\.)?doi\.org/", s, re.I):
        s = "https://" + s
    if re.match(r"^https?://", s, re.I):
        try:
            parsed = urlsplit(_trim_end(s))
        except ValueError:
            return ""
        if not parsed.hostname or _zotero_url(s):
            return ""
        path = unquote(parsed.path)
        match = _DOI_START.search(path)
        if not match:
            return ""
        s = path[match.start():]
        # ACS article pages append an article id and title after this DOI.
        # ACS DOI suffixes in these URLs are one path segment; resolver URLs
        # retain their complete path so ordinary slash-containing DOIs survive.
        if (parsed.hostname or "").lower().endswith("acs.org"):
            match = re.match(r"10\.1021/[^/]+", s, re.I)
            if match:
                s = match.group()
    s = _trim_end(s).strip().lower()
    if not re.fullmatch(r"10\.\d{4,9}/[\x21-\x7e]+", s):
        return ""
    if any(c in s for c in ('"', "`", "\\")):
        return ""
    return s


def extract_cues(text, links=None):
    """Find each author/year, number, DOI or URL clue at its exact position.

    ``links`` contains actual hyperlink targets found by the DOCX reader. Each
    link's start/end refers to its visible label in ``text``. DOI labels can be
    arbitrary (e.g. '该研究'). A visible URL and its hyperlink target produce one
    candidate, while materially different overlapping clues remain reviewable.
    """
    if not isinstance(text, str):
        raise TypeError("Citation text must be a string")
    result, occupied = [], []

    def overlap(start, end, spans):
        return any(start < b and end > a for a, b in spans)

    def add(kind, start, end, source="text", **extra):
        if not (isinstance(start, int) and isinstance(end, int) and 0 <= start < end <= len(text)):
            raise ValueError("Citation clue span is outside paragraph text")
        cue = {"kind": kind, "start": start, "end": end, "text": text[start:end],
               "source": source, "context_excerpt": text[max(0, start - 70):min(len(text), end + 70)], **extra}
        key = (kind, start, end, cue.get("DOI"), cue.get("target"))
        if not any((c["kind"], c["start"], c["end"], c.get("DOI"), c.get("target")) == key for c in result):
            result.append(cue)

    # Real hyperlink targets take priority over a link label that merely looks
    # like another URL. This preserves labels such as '全文' and '该研究'.
    for link in links or []:
        if not isinstance(link, dict):
            raise ValueError("Hyperlink evidence must be an object")
        start, end = link.get("start"), link.get("end")
        if not (isinstance(start, int) and isinstance(end, int) and 0 <= start < end <= len(text)):
            raise ValueError("Hyperlink span is outside paragraph text")
        target = str(link.get("target") or "").strip()
        if not re.match(r"^https?://", target, re.I) or _zotero_url(target):
            continue
        identifier = normalize_doi(target)
        add("doi" if identifier else "article_link", start, end,
            source=link.get("source", "hyperlink"), target=target,
            **({"DOI": identifier} if identifier else {}))
        visible_identifier = normalize_doi(text[start:end])
        if visible_identifier and visible_identifier != identifier:
            # A DOI-looking label can disagree with its actual destination.
            # Keep both identities for review instead of trusting either one.
            add("doi", start, end, source="text", DOI=visible_identifier,
                link_target_conflict=True)
        occupied.append((start, end))

    for match in _URL.finditer(text):
        value = _trim_end(match.group())
        start, end = match.start(), match.start() + len(value)
        if not value or overlap(start, end, occupied):
            continue
        occupied.append((start, end))
        if _zotero_url(value):
            continue
        identifier = normalize_doi(value)
        if identifier:
            label = re.search(r"\bdoi\s*:\s*\Z", text[:start], re.I)
            if label:
                start = label.start()
        add("doi" if identifier else "article_link", start, end, target=value,
            **({"DOI": identifier} if identifier else {}))

    for match in _DOI_TOKEN.finditer(text):
        if overlap(match.start(), match.end(), occupied):
            continue
        value = _trim_end(match.group())
        # ASCII quotes and markup delimiters bound pasted DOI text. Retain
        # angle brackets inside balanced legacy SICI DOI suffixes.
        value = re.split(r'["`\\]', value, maxsplit=1)[0]
        value = _trim_end(value)
        identifier = normalize_doi(value)
        if identifier:
            add("doi", match.start(), match.start() + len(value), DOI=identifier)
            occupied.append((match.start(), match.start() + len(value)))

    for match in _PAREN.finditer(text):
        if overlap(match.start(), match.end(), occupied):
            continue
        content = match.group("content")
        segments = list(re.finditer(r"[^;；]+", content))
        valid = [(segment, _PAREN_ITEM.fullmatch(segment.group())) for segment in segments]
        for segment, item in valid:
            if item is None:
                continue
            if len(segments) == 1:
                start, end = match.span()
            else:
                local = segment.group()
                start = match.start("content") + segment.start() + len(local) - len(local.lstrip())
                end = match.start("content") + segment.end() - len(local) + len(local.rstrip())
            add("author_year", start, end, authors=item.group("authors"),
                years=_year_values(item.group("years")),
                container_start=match.start(), container_end=match.end())
            occupied.append((start, end))

    for match in _NARRATIVE.finditer(text):
        start, end = match.span()
        if overlap(start, end, occupied):
            continue
        authors = match.group("authors")
        prefix = _CN_PREFIX.match(authors)
        if prefix and prefix.end() < len(authors):
            start += prefix.end()
            authors = authors[prefix.end():]
        add("author_year", start, end, authors=authors,
            years=_year_values(match.group("years")))
        occupied.append((start, end))

    for match in _NUMERIC.finditer(text):
        if not overlap(match.start(), match.end(), occupied):
            add("numeric", match.start(), match.end())

    return sorted(result, key=lambda c: (c["start"], c["end"], c["kind"]))
