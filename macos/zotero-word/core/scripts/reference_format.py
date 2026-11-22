"""Read-only citation/References profile and native bibliography typography.

This module does not format ordinary manuscript text, write DOCX, create fields,
or operate Word/Zotero. Native Bibliography changes belong to the adapter.
"""
from __future__ import annotations
import hashlib
import importlib.util
import io
import json
import math
from pathlib import Path
import re
import sys
import xml.etree.ElementTree as ET
from zipfile import ZipFile, BadZipFile
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
PROFILE = "zotero-word.reference-format-profile.v1"
BIBLIOGRAPHY_VERIFY = "zotero-word.bibliography-format-verification.v1"
BIBLIOGRAPHY_RULES = {"font_family", "font_size_pt", "line_spacing", "space_before_pt", "space_after_pt", "left_indent_pt", "right_indent_pt", "first_line_indent_pt"}
BOOLS = {"keep_next": "keepNext", "keep_lines": "keepLines", "page_break_before": "pageBreakBefore", "widow_control": "widowControl"}
FONT = {"font_family", "font_size_pt"}
PARA = BIBLIOGRAPHY_RULES - FONT
REVISION = {"ins", "del", "moveFrom", "moveTo", "moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd", "cellIns", "cellDel", "cellMerge", "numberingChange"}
_revision_spec = importlib.util.spec_from_file_location("_zotero_reference_revision_history", Path(__file__).with_name("revision_history.py"))
_revisions = importlib.util.module_from_spec(_revision_spec)
_revision_spec.loader.exec_module(_revisions)
def digest(value):
    if not isinstance(value, bytes):
        value = json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode()
    return hashlib.sha256(value).hexdigest()


def require(ok, message):
    if not ok:
        raise ValueError(message)


def _object(value):
    return value if isinstance(value, dict) else json.loads(Path(value).read_text(encoding="utf-8"))


def _number(value, minimum, maximum):
    return type(value) in (int, float) and math.isfinite(value) and minimum <= value <= maximum


def _validate_rule(name, rule):
    require(isinstance(rule, dict) and set(rule) == {"value", "basis", "evidence"}, f"Rule {name} needs only value, basis and evidence")
    require(rule["basis"] in ("required", "workflow_default"), f"Rule {name} basis must be required or workflow_default")
    require(isinstance(rule["evidence"], str) and rule["evidence"].strip(), f"Rule {name} needs actual source/default evidence")
    value = rule["value"]
    if name == "font_family":
        valid = isinstance(value, str) and 0 < len(value.strip()) <= 120 and not any(ord(c) < 32 for c in value)
    elif name == "font_size_pt":
        valid = _number(value, 4, 72) and value * 2 == round(value * 2)
    elif name in ("space_before_pt", "space_after_pt"):
        valid = _number(value, 0, 144)
    elif name in ("left_indent_pt", "right_indent_pt"):
        valid = _number(value, 0, 720)
    elif name == "first_line_indent_pt":
        valid = _number(value, -720, 720)
    elif name == "citation_boundary":
        valid = value in ("superscript_no_gap", "style_managed")
    elif name == "line_spacing":
        valid = isinstance(value, dict) and set(value) == {"mode", "value"}
        if valid:
            valid = ((value["mode"] == "multiple" and _number(value["value"], .5, 5)) or
                     (value["mode"] in ("exact", "at_least") and _number(value["value"], 4, 144)))
    else:
        valid = False
    require(valid, f"Unsupported or invalid reference typography value for {name}")


def validate_profile(profile):
    p = _object(profile)
    require(isinstance(p, dict) and p.get("schema") == PROFILE, "Use a citation/References profile: " + PROFILE)
    require(set(p) <= {"schema", "journal", "article_type", "submission_stage", "citation_boundary", "bibliography", "csl"}, "Profile must not contain ordinary manuscript roles, page settings or unknown keys")
    for name in ("journal", "article_type"):
        require(isinstance(p.get(name), str) and p[name].strip(), f"Reference profile needs {name}")
    require(p.get("submission_stage") in ("initial", "revision", "final"), "Stage must be initial, revision or final")
    require(isinstance(p.get("citation_boundary"), dict), "Reference profile needs a declared citation boundary policy")
    _validate_rule("citation_boundary", p["citation_boundary"])
    require(isinstance(p.get("bibliography"), dict) and set(p["bibliography"]) <= BIBLIOGRAPHY_RULES, "Unsupported References typography rules")
    for name, rule in p["bibliography"].items():
        _validate_rule(name, rule)
    if "csl" in p:
        spec = importlib.util.spec_from_file_location("_zotero_reference_csl_config", Path(__file__).with_name("csl_quality.py"))
        module = importlib.util.module_from_spec(spec)
        previous = sys.dont_write_bytecode
        try:
            sys.dont_write_bytecode = True
            spec.loader.exec_module(module)
        finally:
            sys.dont_write_bytecode = previous
        module.validate_config(p["csl"])
    return p


def _xml(raw):
    require(b"<!DOCTYPE" not in raw.upper(), "DOCTYPE is unsupported")
    return ET.fromstring(raw)


def _revision_node(node):
    return _revisions.is_revision(node)


def _historical_node(node):
    return _revisions.is_historical(node)


def _guard_native_result_revisions(root, part):
    """Allow unrelated history, never a current Zotero field touching it."""
    stack = []
    def visit(node, tracked=False, revised_run=False):
        if _historical_node(node):
            # A wholly historical field is outside the active field stack and
            # is deliberately ignored. A deletion within a current field must
            # not turn its surviving fragments into an apparently valid cite.
            for field in stack:
                field["tracked"] = True
                field["raw_parts"].append("".join(e.text or "" for e in node.iter() if e.tag in (W + "instrText", W + "delInstrText")))
            return
        tracked = tracked or _revision_node(node)
        if node.tag == W + "r":
            properties = node.find(W + "rPr")
            revised_run = properties is not None and any(_revision_node(e) for e in properties.iter())
        if tracked or revised_run:
            for field in stack:
                field["tracked"] = True
        if node.tag == W + "fldSimple" and re.search(r"\bZOTERO_", node.get(W + "instr", ""), re.I):
            require(not (tracked or revised_run), "Native Zotero field touches tracked revision: " + part)
        if node.tag == W + "fldChar":
            kind = node.get(W + "fldCharType")
            if kind == "begin":
                stack.append({"parts": [], "raw_parts": [], "phase": "code", "tracked": tracked or revised_run})
            elif kind == "separate":
                require(stack and stack[-1]["phase"] == "code", "Malformed current-view native field separator: " + part)
                stack[-1]["phase"] = "result"
            elif kind == "end":
                require(stack, "Unbalanced current-view native field end: " + part)
                field = stack.pop()
                code = "".join(field["parts"])
                if re.search(r"\bZOTERO_", code, re.I) or re.search(r"\bZOTERO_", "".join(field["raw_parts"]), re.I):
                    require(not field["tracked"], "Native Zotero field touches tracked revision: " + part)
            else:
                require(False, "Unknown current-view Word field marker: " + part)
        elif node.tag in (W + "instrText", W + "delInstrText"):
            require(stack and stack[-1]["phase"] == "code", "Unbound current-view native field instruction: " + part)
            stack[-1]["parts"].append(node.text or "")
            stack[-1]["raw_parts"].append(node.text or "")
            if node.tag == W + "delInstrText":
                stack[-1]["tracked"] = True
        elif node.tag == W + "delText":
            for field in stack:
                field["tracked"] = True
        for child in _children(node):
            visit(child, tracked, revised_run)
    visit(root)
    require(not stack, "Unclosed current-view native field: " + part)


def _native_current_view(root):
    """Prune history on an in-memory parse only; retain active ins/moveTo."""
    for parent in root.iter():
        for child in list(parent):
            if _historical_node(child):
                parent.remove(child)
    return root


def _package(path, *, native_results_readonly=False):
    raw = Path(path).read_bytes()
    with ZipFile(io.BytesIO(raw)) as archive:
        infos = archive.infolist()
        require(len({i.filename for i in infos}) == len(infos), "Duplicate ZIP part names")
        require(sum(i.file_size for i in infos) <= 256 * 1024 * 1024, "DOCX exceeds bounded package size")
        parts = {i.filename: archive.read(i) for i in infos}
    require("word/document.xml" in parts, "Missing main Word document part")
    root = _xml(parts["word/document.xml"])
    # The default remains a strict gate for every formatting writer. The sole
    # opt-in below projects separate in-memory parts for native-result readers;
    # raw bytes, infos and the actual document file remain untouched.
    for name, data in parts.items():
        if name.startswith("word/") and name.endswith(".xml"):
            parsed = _xml(data)
            if native_results_readonly:
                _guard_native_result_revisions(parsed, name)
                _native_current_view(parsed)
                parts[name] = ET.tostring(parsed, encoding="utf-8")
                if name == "word/document.xml":
                    root = parsed
            else:
                require(not any(_revision_node(e) for e in parsed.iter()), "Tracked revisions require a separately reviewed workflow")
    return raw, infos, parts, root


def _children(e):
    if e.tag == MC + "AlternateContent":
        selected = e.find(MC + "Choice")
        if selected is None:
            selected = e.find(MC + "Fallback")
        return list(selected) if selected is not None else []
    return list(e)


def _elements(e):
    yield e
    for child in _children(e):
        yield from _elements(child)


def paragraph_text(p):
    values = []
    for e in _elements(p):
        if e.tag == W + "t":
            values.append(e.text or "")
        elif e.tag in (W + "tab", W + "br"):
            values.append(" " if e.tag == W + "tab" else "\n")
    return "".join(values)


def _paragraphs(root):
    return [e for e in _elements(root) if e.tag == W + "p"]


def _canonical(e):
    # Prefix names, XML declaration and attribute order are not semantic changes.
    return [e.tag, sorted(e.attrib.items()), e.text or "", e.tail or "", [_canonical(c) for c in e]]


def _field_runs(root):
    protected, depth = set(), 0
    def visit(e, inherited=False):
        nonlocal depth
        inherited = inherited or e.tag == W + "fldSimple" or e.tag in (M + "oMath", M + "oMathPara")
        if e.tag == W + "r":
            markers = list(e.iter(W + "fldChar"))
            inside = inherited or depth > 0 or bool(markers) or any(x.tag == W + "instrText" for x in e.iter())
            if inside:
                protected.add(e)
            for marker in markers:
                kind = marker.get(W + "fldCharType")
                if kind == "begin": depth += 1
                elif kind == "end":
                    require(depth > 0, "Unbalanced native Word field")
                    depth -= 1
                elif kind == "separate": require(depth > 0, "Unbalanced native field separator")
                else: require(False, "Unknown Word field marker")
        for c in _children(e):
            visit(c, inherited)
    visit(root)
    require(depth == 0, "Unclosed native Word field")
    return protected


def _special_font_runs(root, styles_raw=None):
    """Conservatively retain legacy symbol/private-use glyph font semantics."""
    def fonts(node):
        return set(node.attrib.values()) if node is not None else set()
    def symbol(names):
        return any(re.sub(r"[\s_-]", "", name).casefold().startswith(("symbol", "wingdings", "webdings", "mtextra")) for name in names)
    styles, cache, default_fonts, default_style = {}, {}, set(), None
    if styles_raw:
        style_root = _xml(styles_raw)
        default_fonts = fonts(style_root.find(W+"docDefaults/"+W+"rPrDefault/"+W+"rPr/"+W+"rFonts"))
        for s in style_root.findall(W+"style"):
            styles[s.get(W+"styleId")] = s
            if s.get(W+"type") == "paragraph" and s.get(W+"default") in ("1", "true", "on"):
                default_style = s.get(W+"styleId")
    def inherited(style_id, visiting=None):
        if style_id in cache: return cache[style_id]
        visiting = set(visiting or ())
        if style_id not in styles or style_id in visiting: return set()
        visiting.add(style_id); s=styles[style_id]; parent=s.find(W+"basedOn")
        value=fonts(s.find(W+"rPr/"+W+"rFonts")) | inherited(parent.get(W+"val") if parent is not None else None, visiting)
        cache[style_id]=value
        return value
    result={}
    for paragraph in _paragraphs(root):
        style = paragraph.find(W+"pPr/"+W+"pStyle")
        pf=default_fonts | inherited(style.get(W+"val") if style is not None else default_style)
        for run in paragraph.iter(W+"r"):
            character=run.find(W+"rPr/"+W+"rStyle")
            names=pf | fonts(run.find(W+"rPr/"+W+"rFonts")) | inherited(character.get(W+"val") if character is not None else None)
            private=any(0xE000 <= ord(c) <= 0xF8FF or 0xF0000 <= ord(c) <= 0xFFFFD or 0x100000 <= ord(c) <= 0x10FFFD for t in run.iter(W+"t") for c in (t.text or ""))
            if symbol(names) or private:
                result[run]="Legacy symbol/private-use glyph font; substitution could change scientific meaning"
    return result


def _font_key(value):
    """Keep literal observations; compare the standard localized DengXian alias."""
    if not isinstance(value, str):
        return None
    key = value.strip().casefold()
    return "dengxian" if key == "等线" else key


def _effective_typography(root, bindings, profile, parts, native_result_runs=None):
    """Resolve actual relevant attributes, without assuming an unknown theme/font.

    Actual w:rStyle uses its own basedOn chain. A w:link alone is not inheritance;
    it is never used to invent character properties for an unstyled run.
    """
    style_root=_xml(parts["word/styles.xml"]) if "word/styles.xml" in parts else ET.Element(W+"styles")
    styles={s.get(W+"styleId"):s for s in style_root.findall(W+"style")}
    default_para=next((s.get(W+"styleId") for s in styles.values() if s.get(W+"type")=="paragraph" and s.get(W+"default") in ("1","true","on")),None)
    default_char=next((s.get(W+"styleId") for s in styles.values() if s.get(W+"type")=="character" and s.get(W+"default") in ("1","true","on")),None)
    defaults=style_root.find(W+"docDefaults")
    dp=defaults.find(W+"pPrDefault/"+W+"pPr") if defaults is not None else None
    dr=defaults.find(W+"rPrDefault/"+W+"rPr") if defaults is not None else None
    def chain(style_id,seen=None):
        if style_id is None:return []
        require(style_id in styles,"Current typography refers to a missing style")
        seen=set(seen or ());require(style_id not in seen,"Cyclic style inheritance cannot be resolved")
        seen.add(style_id);s=styles[style_id];base=s.find(W+"basedOn")
        return chain(base.get(W+"val") if base is not None else None,seen)+[s]
    def attrs(layers,local):
        result={}
        for layer in layers:
            if layer is not None:
                child=layer.find(W+local)
                if child is not None:
                    if local in set(BOOLS.values())|{'cs','rtl','contextualSpacing'}:result['val']=child.get(W+'val','1')
                    else:result.update({k.removeprefix(W):v for k,v in child.attrib.items()})
        return result
    def number(value):
        if isinstance(value,str) and re.fullmatch(r'-?\d+',value):return int(value)
        return None
    def boolean(value,default):
        if value is None:return default
        if value in ('1','true','on'):return True
        if value in ('0','false','off'):return False
        return None
    # Theme fonts can be used only if their explicit typeface is known. Empty
    # locale-dependent East Asian/complex slots remain unresolved, not guessed.
    A='{http://schemas.openxmlformats.org/drawingml/2006/main}'
    themes=[]
    for part,data in parts.items():
        if re.fullmatch(r'word/theme/theme\d+\.xml',part):themes.append(_xml(data))
    def theme_font(value):
        if len(themes)!=1:return None
        family='majorFont' if value.startswith('major') else 'minorFont' if value.startswith('minor') else None
        slot='ea' if value.endswith('EastAsia') else 'cs' if value.endswith('Bidi') else 'latin'
        if not family:return None
        e=themes[0].find('.//'+A+'fontScheme/'+A+family+'/'+A+slot)
        return e.get('typeface') if e is not None and e.get('typeface') else None
    def font(layers,slot):
        current=None;theme_key={'ascii':'asciiTheme','hAnsi':'hAnsiTheme','eastAsia':'eastAsiaTheme','cs':'cstheme'}[slot]
        for layer in layers:
            rf=layer.find(W+'rFonts') if layer is not None else None
            if rf is None:continue
            if rf.get(W+slot) is not None:current=rf.get(W+slot)
            theme=rf.get(W+theme_key)
            if theme is None and slot=='cs':theme=rf.get(W+'csTheme')
            if theme is not None:current=theme_font(theme)
        return current
    protected=_field_runs(root);special=_special_font_runs(root,parts.get('word/styles.xml'))
    drift,resolved=[],[]
    def compare(scope,name,observed,expected,rule):
        if name=='font_family':ok=isinstance(observed,dict) and all(isinstance(v,str) and _font_key(v)==_font_key(expected) for v in observed.values()) and bool(observed)
        else:ok=observed==expected
        record={**scope,'property':name,'observed':observed,'expected':expected,'basis':rule['basis'],'evidence':rule['evidence']}
        (resolved if ok else drift).append(record)
    for item,paragraph in bindings:
        scope={'story':'word/document.xml','paragraph':item['paragraph'],'role':item['role']}
        rules=profile['roles'][item['role']];ps=paragraph.find(W+'pPr/'+W+'pStyle')
        paragraph_chain=chain(ps.get(W+'val') if ps is not None else default_para)
        pl=[dp]+[s.find(W+'pPr') for s in paragraph_chain]+[paragraph.find(W+'pPr')]
        for name in sorted(set(rules)&PARA):
            expected=rules[name]['value']
            if name in BOOLS:observed=boolean(attrs(pl,BOOLS[name]).get('val'),name=='widow_control')
            elif name=='alignment':observed=attrs(pl,'jc').get('val','left')
            elif name=='line_spacing':
                a=attrs(pl,'spacing');mode=a.get('lineRule','auto');amount=number(a.get('line','240'))
                observed={'mode':{'auto':'multiple','exact':'exact','atLeast':'at_least'}.get(mode),'value':amount/(240 if mode=='auto' else 20) if amount is not None else None}
            elif name in ('space_before_pt','space_after_pt'):
                a=attrs(pl,'spacing');key='before' if name=='space_before_pt' else 'after'
                amount=number(a.get(key,'0'))
                contextual=boolean(attrs(pl,'contextualSpacing').get('val'),False)
                # Same-style neighboring paragraphs can suppress nominal positive
                # spacing. Without fully resolving that layout context, do not
                # report the nominal value as an observed effective gap.
                observed=amount/20 if amount is not None and boolean(a.get(key+'Autospacing'),False) is False and number(a.get(key+'Lines','0'))==0 and not (expected>0 and contextual is not False) else None
            else:
                a=attrs(pl,'ind')
                if name=='first_line_indent_pt':
                    x,y=number(a.get('firstLine','0')),number(a.get('hanging','0'))
                    # firstLine and hanging are mutually exclusive settings;
                    # subtracting both produces a fictitious net indentation.
                    # This also catches conflicts inherited across style/direct
                    # layers. Keep ambiguous combinations HOLD rather than
                    # guessing Word's winning setting or converted glyph width.
                    observed=((x-y)/20) if x is not None and y is not None and x>=0 and y>=0 and not (x and y) and number(a.get('firstLineChars','0'))==0 and number(a.get('hangingChars','0'))==0 else None
                else:
                    key='left' if name=='left_indent_pt' else 'right';modern='start' if key=='left' else 'end'
                    amount=number(a.get(modern,a.get(key,'0')))
                    observed=amount/20 if amount is not None and number(a.get(modern+'Chars',a.get(key+'Chars','0')))==0 else None
            compare(scope,name,observed,expected,rules[name])
        for run_index,run in enumerate(paragraph.iter(W+'r'),1):
            if ((run in protected and run not in (native_result_runs or ())) or run in special or not any(e.tag==W+'t' for e in run.iter())):continue
            if native_result_runs is not None and run not in native_result_runs:continue
            rs=run.find(W+'rPr/'+W+'rStyle');character_chain=chain(rs.get(W+'val') if rs is not None else default_char)
            rl=[dr]+[s.find(W+'rPr') for s in paragraph_chain]+[s.find(W+'rPr') for s in character_chain]+[run.find(W+'rPr')]
            text=''.join(t.text or '' for t in run.iter(W+'t'))
            complex_flag=boolean(attrs(rl,'cs').get('val'),False) or boolean(attrs(rl,'rtl').get('val'),False)
            complex_text=any(0x590<=ord(c)<=0x8ff or 0xfb1d<=ord(c)<=0xfdff or 0xfe70<=ord(c)<=0xfeff for c in text)
            scripts=set()
            if complex_flag:scripts.add('cs')
            else:
                for c in text:
                    code=ord(c)
                    scripts.add('ascii' if code<128 else 'cs' if (0x590<=code<=0x8ff or 0xfb1d<=code<=0xfdff or 0xfe70<=code<=0xfeff) else 'eastAsia' if (0x2e80<=code<=0x9fff or 0xac00<=code<=0xd7af or 0xf900<=code<=0xfaff or 0x20000<=code<=0x3134f) else 'hAnsi')
            if 'font_family' in rules:
                compare({**scope,'run':run_index},'font_family',{slot:font(rl,slot) for slot in sorted(scripts)},rules['font_family']['value'],rules['font_family'])
            if 'font_size_pt' in rules:
                # Normal Latin uses sz; a stale inherited szCs must not fail it.
                # Complex-script runs require their separately resolved szCs.
                local='szCs' if complex_flag or complex_text else 'sz'
                amount=number(attrs(rl,local).get('val'))
                if amount is None and local=='szCs':amount=number(attrs(rl,'sz').get('val'))
                compare({**scope,'run':run_index},'font_size_pt',amount/2 if amount is not None else None,rules['font_size_pt']['value'],rules['font_size_pt'])
    return drift,resolved


def _inspect_native_document(document):
    """Lazy standalone fallback; public completion paths pass their in-call parse."""
    spec = importlib.util.spec_from_file_location("_zotero_word_bibliography_inspection", Path(__file__).with_name("zotero_docx.py"))
    module = importlib.util.module_from_spec(spec)
    previous = sys.dont_write_bytecode
    try:
        sys.dont_write_bytecode = True
        spec.loader.exec_module(module)
    finally:
        sys.dont_write_bytecode = previous
    return module.inspect_docx(document)


def _bibliography_scope(root):
    """Bind only actual complex-field result runs, including cross-paragraph bibs."""
    paragraphs = _paragraphs(root)
    indices = {p: i for i, p in enumerate(paragraphs, 1)}
    parents = {c: p for p in root.iter() for c in p}
    stack, fields, visible = [], [], {}
    def visit(node, paragraph=None, run=None):
        if node.tag == W + "p": paragraph = node
        if node.tag == W + "r": run = node
        if node.tag == W + "fldSimple" and stack:
            stack[-1]["nested"] = True
        if node.tag == W + "fldChar":
            kind = node.get(W + "fldCharType")
            if kind == "begin":
                if stack: stack[-1]["nested"] = True
                stack.append({"phase": "code", "code_parts": [], "text_parts": [], "runs": set(), "paragraphs": {}, "nested": bool(stack)})
            elif kind == "separate":
                require(stack and stack[-1]["phase"] == "code", "Malformed native field separator")
                stack[-1]["phase"] = "result"
            elif kind == "end":
                require(stack, "Unbalanced native Word field end")
                field = stack.pop(); field["code"] = "".join(field.pop("code_parts")); fields.append(field)
            else:
                require(False, "Unknown native field marker")
        elif node.tag == W + "instrText":
            require(stack and stack[-1]["phase"] == "code", "Unbound or mixed native field instruction")
            stack[-1]["code_parts"].append(node.text or "")
        elif node.tag in (W + "t", W + "tab", W + "br"):
            if not stack or stack[-1]["phase"] == "result":
                text = (node.text or "") if node.tag == W + "t" else " " if node.tag == W + "tab" else "\n"
                owner = stack[-1] if stack else None
                if paragraph is not None:
                    visible.setdefault(paragraph, []).append((owner, text))
                if owner is not None:
                    require(paragraph is not None and run is not None, "Native result text is not bound to a main-story paragraph/run")
                    owner["paragraphs"].setdefault(paragraph, []).append(text)
                    if node.tag == W + "t":
                        owner["text_parts"].append(text); owner["runs"].add(run)
        for child in _children(node): visit(child, paragraph, run)
    visit(root)
    require(not stack, "Unclosed native field")
    candidates = [f for f in fields if re.match(r"^\s*(?:ADDIN\s+)?ZOTERO_BIBL\b", f["code"], re.I)]
    require(len(candidates) == 1, "Bibliography typography needs exactly one actual main-story native bibliography")
    field = candidates[0]
    require(field["phase"] == "result" and not field["nested"], "Nested or incomplete bibliography fields are unsupported")
    marker = re.match(r"^\s*(?:ADDIN\s+)?ZOTERO_BIBL\s*(?=\{)", field["code"], re.I)
    require(marker is not None, "Malformed native bibliography code")
    data, end = json.JSONDecoder().raw_decode(field["code"][marker.end():])
    require(isinstance(data, dict) and field["code"][marker.end()+end:].strip().upper() in ("", "CSL_BIBLIOGRAPHY"), "Invalid native bibliography metadata")
    result_paragraphs = [p for p in paragraphs if p in field["paragraphs"] and "".join(field["paragraphs"][p]).strip()]
    require(result_paragraphs and "".join(field["text_parts"]).strip(), "Native bibliography has no visible result entries")
    for paragraph in result_paragraphs:
        require(not any(owner is not field and text.strip() for owner, text in visible.get(paragraph, [])), "Bibliography result paragraph mixes ordinary body or unrelated field text")
        ancestor, tags = paragraph, []
        while ancestor in parents:
            ancestor = parents[ancestor]; tags.append(ancestor.tag)
        require(not any(t in (W+"tc", W+"txbxContent", W+"sdt", W+"fldSimple", MC+"Choice", MC+"Fallback", MC+"AlternateContent") for t in tags), "Bibliography result is in an unsupported/ambiguous container")
        require(not any(e.tag in (M+"oMath", M+"oMathPara", W+"drawing", W+"pict", W+"object") for e in paragraph.iter()), "Bibliography result contains unsupported graphical/equation content")
    allowed = {r for p in result_paragraphs for r in p.iter(W+"r") if r in field["runs"]}
    require(all(not any(e.tag in (W+"fldChar", W+"instrText") for e in r.iter()) for r in allowed), "Bibliography result shares a run with native field markers/code")
    return field, [(indices[p], p) for p in result_paragraphs], allowed, data, fields


def verify_bibliography_typography(document, profile, inspection=None, *, read_size=None):
    """Read-only actual native-bibliography font, size, spacing and indentation.

    This never formats result runs or asserts a Refresh. The adapter must change
    the Bibliography style/CSL/native paragraph settings and save/reopen first.
    """
    report = {"schema": BIBLIOGRAPHY_VERIFY, "result": "HOLD", "issues": [], "changes": [], "resolved_properties": [],
              "bibliography_paragraphs": [], "checked_run_count": 0,
              "scope": "Read-only effective typography of exactly one actual native main-story bibliography result; no field writes, native Refresh or visual review is implied."}
    try:
        p = validate_profile(profile)
        require(set(p.get("bibliography", {})) == BIBLIOGRAPHY_RULES, "Native bibliography requires all eight typography rules")
        raw, _, parts, root = _package(document, native_results_readonly=True)
        report.update(document_sha256=digest(raw), target_document_sha256=digest(raw), profile_sha256=digest(p))
        observed = inspection if inspection is not None else _inspect_native_document(document)
        require(isinstance(observed, dict) and observed.get("schema") == "zotero-word.inspection.v1"
                and observed.get("document_sha256") == digest(raw)
                and observed.get("document") == str(Path(document).resolve()), "Bibliography inspection does not bind the current saved document")
        require(not observed.get("issues"), "Native document inspection has unresolved issues")
        bibliographies = observed.get("bibliographies", [])
        require(len(bibliographies) == 1 and observed.get("bibliography_count") == 1, "Need exactly one inspected native bibliography")
        require(bibliographies[0].get("context", {}).get("story") == "word/document.xml", "Bibliography story is unsupported")
        field, paragraphs, allowed, metadata, all_fields = _bibliography_scope(root)
        require(bibliographies[0].get("field_data") == metadata and bibliographies[0].get("display") == "".join(field["text_parts"]), "Inspected bibliography does not match the actual native field")
        special = _special_font_runs(root, parts.get("word/styles.xml"))
        require(not any(r in special for r in allowed), "Bibliography symbol/private-use glyph fonts require native manual review")
        subprofile = {"roles": {"references": p["bibliography"]}, "page": {}}
        bindings = [({"story": "word/document.xml", "paragraph": index, "role": "references"}, paragraph) for index, paragraph in paragraphs]
        drift, resolved = _effective_typography(root, bindings, subprofile, parts, native_result_runs=allowed)
        checked = [r for r in resolved + drift if r["property"] == "font_family"]
        require(len(checked) == len(allowed), "Some bibliography result runs could not be safely checked")
        # Word's Font.Size covers ordinary text. Never infer a complex-script size from it.
        ordinary = {(r["paragraph"], r["run"]) for r in checked if "cs" not in r["observed"]}
        missing = [r for r in drift if r["property"] == "font_size_pt" and r["observed"] is None
                   and (r["paragraph"], r["run"]) in ordinary]
        if read_size and missing and len(missing) == len(drift):
            text = "\n".join("".join(field["paragraphs"][p]) for _, p in paragraphs)
            try:
                word_size = read_size(document, raw, text)
            except (ValueError, OSError) as exc:
                word_size = None
                report["issues"].append(str(exc))
            if word_size is not None:
                for record in missing:
                    record.update(observed=word_size, observed_source="Word bibliography result range")
                resolved.extend(r for r in missing if r["observed"] == r["expected"])
                drift = [r for r in drift if r["property"] != "font_size_pt" or r["observed"] != r["expected"]]
        identities = [{"citationID": c.get("citationID"), "uris": sorted({u for item in c.get("citationItems", []) for u in item.get("uris", [])})} for c in observed.get("citations", [])]
        report.update(changes=drift, resolved_properties=resolved, bibliography_paragraphs=[{"story":"word/document.xml", "paragraph":index, "text_sha256":hashlib.sha256(paragraph_text(paragraph).encode()).hexdigest()} for index,paragraph in paragraphs],
                      checked_run_count=len(checked), checked_paragraph_count=len(paragraphs), effective_inheritance_verified=True,
                      native_field_code_sha256=digest(field["code"].encode()), bibliography_result_sha256=digest("".join(field["text_parts"]).encode()),
                      native_identity_sha256=digest(identities), citation_identities=identities,
                      all_native_field_codes_sha256=digest([f["code"] for f in all_fields if re.search(r"\bZOTERO_(?:ITEM|BIBL)\b", f["code"], re.I)]),
                      native_refresh_performed=False, visual_review_required=True)
        require(not drift, f"{len(drift)} bibliography typography properties differ from the declared final rules")
        require(digest(Path(document).read_bytes()) == digest(raw), "Document changed during bibliography typography verification")
        report["result"] = "PASS"
    except (ValueError, TypeError, KeyError, AttributeError, OSError, BadZipFile, ET.ParseError, json.JSONDecodeError) as exc:
        report["issues"].append(str(exc))
    return report


if __name__ == "__main__":
    raise SystemExit("Read-only helper; use zotero_docx.py bibliography-quality --reference-profile PROFILE")
