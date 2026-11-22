"""Resolve exact missing paper identifiers using Windows system curl."""
from concurrent.futures import ThreadPoolExecutor
from html.parser import HTMLParser
import json
import os
from pathlib import Path
import re
import subprocess
from urllib.parse import quote, urlencode, urlsplit


def non_citation(cue):
    try:
        url = urlsplit(cue.get("target", ""))
    except ValueError:
        return None
    context = cue.get("context_excerpt", "").casefold()
    if (url.hostname in ("creativecommons.org", "www.creativecommons.org")
            and url.path.startswith(("/licenses/", "/publicdomain/"))
            and re.search(r"conditions of|distributed under|licensed under|under (?:the|a) (?:creative commons|cc )|授权条款|许可条款", context)):
        return "Creative Commons licence in a rights statement"
    if (re.fullmatch(r"word/(?:header|footer)\d+\.xml", cue["story"])
            and re.fullmatch(r"10\.\d{4,9}/x{3,}", cue.get("DOI", ""), re.I)):
        return "Template DOI placeholder in a header or footer"
    if (url.hostname in ("www.mdpi.com", "mdpi.com") and re.fullmatch(r"/article/(?:doi|10\.[^/]+/[^/]+)/s\d+/?", url.path, re.I)
            and re.search(r"(?:supporting|supplementary) (?:information|materials?).{0,45}download", context)):
        return "Supplementary material download in the manuscript's own availability statement"
    return None


def _download(url, accept):
    result = subprocess.run([str(Path(os.environ["SystemRoot"]) / "System32" / "curl.exe"), "--disable", "--fail", "--silent", "--show-error",
        "--globoff", "--location", "--proto", "=https", "--proto-redir", "=https", "--max-redirs", "3",
        "--connect-timeout", "5", "--max-time", "20", "--header", "Accept: " + accept,
        "--max-filesize", "8388608", url], capture_output=True, timeout=25)
    if result.returncode:
        raise ValueError(result.stderr.decode(errors="replace").strip() or "Metadata request failed")
    return result.stdout


def _get(endpoint, accept):
    return json.loads(_download("https://api.crossref.org/works" + endpoint, accept))


def _pii(url):
    try:
        parsed = urlsplit(url)
    except ValueError:
        return None
    if parsed.hostname in ("www.sciencedirect.com", "sciencedirect.com"):
        match = re.fullmatch(r"/science/article/(?:abs/)?pii/([A-Z0-9]{5,64})/?", parsed.path, re.I)
        if match:
            return match[1].upper()


class _ArticleHead(HTMLParser):
    def __init__(self):
        super().__init__()
        self.piis, self.dois = set(), set()
        self.in_head = False

    def handle_starttag(self, tag, attrs):
        if tag == "head":
            self.in_head = True
        if not self.in_head:
            return
        attrs = dict(attrs)
        key = (attrs.get("name") or attrs.get("property") or "").casefold()
        content = (attrs.get("content") or "").strip()
        if tag == "meta" and key in ("citation_doi", "prism.doi") and content:
            self.dois.add(content.casefold())
        if tag == "meta" and key == "citation_pii" and content:
            self.piis.add(content.upper())
        identity_url = None
        if tag == "link" and "canonical" in (attrs.get("rel") or "").casefold().split():
            identity_url = attrs.get("href")
        if tag == "meta" and key == "og:url":
            identity_url = content
        if identity_url and (identifier := _pii(identity_url)):
            self.piis.add(identifier)

    def handle_endtag(self, tag):
        if tag == "head":
            self.in_head = False


def _bibliographic_hint(pending, items, identifier):
    # Rank published library titles locally; never send manuscript prose to Crossref.
    contexts = " ".join(c.get("context_excerpt", "") for c in pending if _pii(c.get("target", "")) == identifier)
    terms = set(re.findall(r"[a-z0-9]{4,}", contexts.casefold()))
    ranks = []
    for item in items:
        title = item.get("title")
        if not isinstance(title, str) or not title.strip() or not item.get("DOI"):
            continue
        issns = item.get("ISSN") or []
        if isinstance(issns, str):
            issns = re.split(r"[,; ]+", issns)
        same_journal = any(str(x).replace("-", "").upper() == identifier[1:9] for x in issns)
        overlap = len(terms.intersection(re.findall(r"[a-z0-9]{4,}", title.casefold())))
        if same_journal or overlap > 1:
            ranks.append((same_journal, overlap, title))
    return max(ranks)[2][:240] if ranks else identifier


def _resolve_pii(value, hint):
    errors = []
    queries = [{"filter": "alternative-id:" + value, "rows": 3},
               {"query.bibliographic": hint, "rows": 32}]
    if re.match(r"^S[0-9]{7}[0-9X]", value):
        queries[1]["filter"] = "issn:" + value[1:5] + "-" + value[5:9]
    for query in queries:
        try:
            msg = _get("?" + urlencode(query), "application/json")["message"]
            hits = {str(row.get("DOI", "")).strip().casefold() for row in msg["items"]
                    if any(str(v).strip().upper() == value for v in row.get("alternative-id", []))}
        except (OSError, ValueError, TypeError, KeyError, AttributeError, subprocess.TimeoutExpired) as exc:
            errors.append("Crossref: " + str(exc))
            continue
        if len(hits) > 1:
            raise ValueError("Several Crossref records may have the same PII: " + value)
        if len(hits) == 1 and ("query.bibliographic" in query or msg["total-results"] == len(msg["items"])):
            doi = hits.pop()
            if re.fullmatch(r"10\.\d{4,9}/\S+", doi):
                return doi
        errors.append("Crossref: no matching alternative-id")
    try:
        page = _ArticleHead()
        page.feed(_download("https://www.sciencedirect.com/science/article/pii/" + value, "text/html").decode("utf-8", "replace"))
        page.close()
        if page.piis == {value} and len(page.dois) == 1:
            doi = next(iter(page.dois))
            if re.fullmatch(r"10\.\d{4,9}/\S+", doi):
                return doi
        errors.append("Publisher: no matching PII and DOI")
    except (OSError, ValueError, subprocess.TimeoutExpired) as exc:
        errors.append("Publisher: " + str(exc))
    raise ValueError("; ".join(errors))


def _lookup(key):
    kind, value, *hint = key
    try:
        if kind == "pii":
            return _resolve_pii(value, hint[0]), None
        data = _get("/" + quote(value, safe="/") + "/transform", "application/vnd.citationstyles.csl+json")
        if (str(data.get("DOI", "")).casefold() != value or not isinstance(data.get("title"), str)
                or not data["title"].strip() or not data.get("type") or not data.get("issued", {}).get("date-parts")
                or not (data.get("author") or data.get("editor"))):
            raise ValueError("Crossref returned a different DOI or incomplete metadata")
        return data, None
    except (OSError, ValueError, KeyError, TypeError, AttributeError, subprocess.TimeoutExpired) as exc:
        return None, str(exc)


def fetch_missing(pending, known, cache_path, source_hash, api, *, allow_import):
    cache = api.load(cache_path) if cache_path.exists() else {"source_sha256": source_hash, "links": {}, "records": {}}
    if cache["source_sha256"] != source_hash:
        raise ValueError("Reference metadata belongs to a different original")
    dois, links, problems = set(), {}, {}
    for cue in pending:
        if cue["candidates"] or cue.get("link_target_conflict") or cue["story"] != "word/document.xml":
            continue
        if cue.get("DOI"):
            dois.add(cue["DOI"].casefold())
        elif cue.get("target"):
            if identifier := _pii(cue["target"]):
                links[cue["target"]] = identifier
    if not links and not allow_import:
        return [], cache["links"], problems
    local = {}
    for item in known.values():
        identifier = _pii(item.get("URL") or item.get("url") or "")
        doi = api._cues.normalize_doi(item.get("DOI"))
        if identifier and doi:
            local.setdefault(identifier, set()).add(doi)
    missing = {}
    for url, identifier in links.items():
        candidates = local.get(identifier, set())
        if len(candidates) == 1:
            cache["links"][url] = next(iter(candidates))
        elif candidates:
            cache["links"].pop(url, None)
            problems[url] = "Multiple library DOIs have this publisher identifier"
        elif url not in cache["links"]:
            missing.setdefault(identifier, []).append(url)
    known_dois = {api._cues.normalize_doi(item.get("DOI")) for item in known.values()}
    with ThreadPoolExecutor(max_workers=2) as pool:
        queries = [("pii", key, _bibliographic_hint(pending, known.values(), key)) for key in missing]
        for identifier, (doi, error) in zip(missing, pool.map(_lookup, queries)):
            for url in missing[identifier]:
                if error or not api._cues.normalize_doi(doi):
                    problems[url] = error or "Invalid DOI for publisher identifier"
                else:
                    cache["links"][url] = doi
        if links:
            api.output(cache_path, cache)
        if not allow_import:
            return [], cache["links"], problems
        dois.update(cache["links"][url] for url in links if url in cache["links"])
        needed = sorted(dois - known_dois)
        missing = [doi for doi in needed if doi not in cache["records"]]
        for count, (doi, (item, error)) in enumerate(zip(missing, pool.map(_lookup, [("doi", d) for d in missing])), 1):
            if error:
                problems[doi] = error
            else:
                cache["records"][doi] = item
            if count % 10 == 0:
                api.output(cache_path, cache)
    api.output(cache_path, cache)
    return [cache["records"][d] for d in needed if d in cache["records"]], cache["links"], problems
