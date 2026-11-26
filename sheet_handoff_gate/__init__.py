"""Local, bounded inspection of selected SpreadsheetML package signals."""
import json
import io
import os
import stat
import struct
import zipfile
import zlib
from xml.parsers import expat

MAX_FILE = 8 * 1024 * 1024
MAX_EXPANDED = 64 * 1024 * 1024
MAX_XML = 1024 * 1024
MAX_XML_TOTAL = 4 * MAX_XML
MAX_PARTS = 1024
MAX_BATCH = 20
SML = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
DOC_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
OFFICE_REL = DOC_REL + "officeDocument"
SHEET_REL_TYPES = {DOC_REL + name: name for name in ("worksheet", "chartsheet", "dialogsheet")}
SHEET_REL_TYPES.update({"http://schemas.microsoft.com/office/2006/relationships/" + name: name
                        for name in ("xlMacrosheet", "xlIntlMacrosheet")})
METRICS = ("hidden_sheets", "very_hidden_sheets", "external_relationships",
           "external_link_parts", "macro_parts", "embedded_parts",
           "connection_parts", "pivot_cache_parts")
DEFAULT_POLICY = {key: 0 for key in METRICS}

class InputError(ValueError):
    """Unsupported, malformed or over-budget input."""

def open_regular(path, limit):
    """Return a bounded local bytes snapshot; refuse special files and final links."""
    def identity(info):
        return (info.st_dev, info.st_ino, info.st_size, info.st_mode,
                info.st_mtime_ns, info.st_ctime_ns)
    before = os.lstat(path)
    if (not stat.S_ISREG(before.st_mode) or before.st_size > limit or
            getattr(before, "st_file_attributes", 0) & getattr(stat, "FILE_ATTRIBUTE_REPARSE_POINT", 0)):
        raise InputError("regular bounded file required")
    fd = os.open(path, os.O_RDONLY | getattr(os, "O_NONBLOCK", 0) |
                 getattr(os, "O_NOFOLLOW", 0) | getattr(os, "O_BINARY", 0))
    with os.fdopen(fd, "rb") as f:
        opened = os.fstat(f.fileno())
        if identity(opened) != identity(before):
            raise InputError("file changed before read")
        data = f.read(opened.st_size + 1)
        final = os.fstat(f.fileno())
        named = os.lstat(path)
        if len(data) != opened.st_size or identity(final) != identity(opened) or identity(named) != identity(opened):
            raise InputError("file changed during read")
    return io.BytesIO(data)


def load_policy(path=None):
    if path is None:
        return dict(DEFAULT_POLICY)
    def unique(pairs):
        result = {}
        for key, value in pairs:
            if key in result:
                raise InputError("duplicate policy field")
            result[key] = value
        return result
    with open_regular(path, 16384) as f:
        raw = f.read(16385)
    if len(raw) > 16384:
        raise InputError("policy limit")
    p = json.loads(raw.decode("utf-8"), object_pairs_hook=unique)
    if not isinstance(p, dict) or set(p) != {"version", "maximum"} or type(p["version"]) is not int or p["version"] != 1:
        raise InputError("policy format")
    limits = p["maximum"]
    if not isinstance(limits, dict) or set(limits) != set(METRICS):
        raise InputError("policy must declare every metric")
    if any(type(v) is not int or not 0 <= v <= MAX_PARTS for v in limits.values()):
        raise InputError("invalid policy limits")
    return limits

def zip_preflight(f):
    # Bound central directory allocation before constructing ZipFile.
    size = f.seek(0, os.SEEK_END)
    if size < 22 or size > MAX_FILE:
        raise InputError("archive size")
    f.seek(max(0, size - 65557))
    tail = f.read(65557)
    i = tail.rfind(b"PK\x05\x06")
    if i < 0 or len(tail) - i < 22:
        raise InputError("missing zip directory")
    _, disk, cd_disk, per_disk, count, cd_size, cd_offset, comment = struct.unpack_from("<4s4H2LH", tail, i)
    eocd = size - len(tail) + i
    if (disk or cd_disk or count != per_disk or not 1 <= count <= MAX_PARTS
            or cd_size > 512 * 1024 or cd_offset + cd_size != eocd
            or i + 22 + comment != len(tail)):
        raise InputError("unsupported zip directory")
    return count

def inspect_xml(raw, root_name, on_start):
    parser = expat.ParserCreate(namespace_separator=" ")
    stack = []
    nodes = 0
    def reject(*_args):
        raise InputError("DTD or entity declaration")
    def start(name, attrs):
        nonlocal nodes
        nodes += 1
        if nodes > 20000 or len(stack) >= 32:
            raise InputError("XML complexity limit")
        if not stack and name != root_name:
            raise InputError("unexpected XML root")
        stack.append(name)
        on_start(tuple(stack), attrs)
    def end(_name):
        stack.pop()
    parser.StartElementHandler = start
    parser.EndElementHandler = end
    parser.StartDoctypeDeclHandler = reject
    parser.EntityDeclHandler = reject
    parser.ExternalEntityRefHandler = reject
    parser.Parse(raw, True)

def inspect_workbook(path):
    counts = {key: 0 for key in METRICS}
    with open_regular(path, MAX_FILE) as f:
        expected = zip_preflight(f)
        with zipfile.ZipFile(f) as archive:
            parts = archive.infolist()
            if len(parts) != expected or sum(p.file_size for p in parts) > MAX_EXPANDED:
                raise InputError("archive limits")
            names = {}
            for part in parts:
                name = part.orig_filename
                normalized = name.rstrip("/")
                if (name != part.filename or not normalized or "\\" in name
                        or any(ord(c) < 32 for c in name)
                        or any(s in ("", ".", "..") for s in normalized.split("/"))
                        or ":" in name or "%" in name
                        or normalized.lower() in names or part.flag_bits & 1
                        or part.compress_type not in (zipfile.ZIP_STORED, zipfile.ZIP_DEFLATED)):
                    raise InputError("unsupported archive member")
                names[normalized.lower()] = part
            required = ("[content_types].xml", "_rels/.rels", "xl/workbook.xml", "xl/_rels/workbook.xml.rels")
            if any(k not in names or names[k].is_dir() for k in required):
                raise InputError("standard workbook layout required")
            remaining = MAX_XML_TOTAL
            def read_xml(part):
                nonlocal remaining
                if part.file_size > min(MAX_XML, remaining):
                    raise InputError("XML size limit")
                with archive.open(part) as stream:
                    raw = stream.read(min(MAX_XML, remaining) + 1)
                if len(raw) != part.file_size or len(raw) > remaining:
                    raise InputError("XML size mismatch")
                remaining -= len(raw)
                return raw
            sheet_count = 0
            sheet_ids = set()
            sheet_targets = {}
            def sheet(stack, attrs):
                nonlocal sheet_count
                if stack == (SML+" workbook", SML+" sheets", SML+" sheet"):
                    state = attrs.get("state", "visible")
                    if state not in ("visible", "hidden", "veryHidden"):
                        raise InputError("unsupported sheet state")
                    rid = attrs.get("http://schemas.openxmlformats.org/officeDocument/2006/relationships id")
                    if not rid or rid in sheet_ids:
                        raise InputError("missing or duplicate sheet relationship")
                    sheet_ids.add(rid)
                    sheet_count += 1
                    counts["hidden_sheets"] += state == "hidden"
                    counts["very_hidden_sheets"] += state == "veryHidden"
            inspect_xml(read_xml(names["xl/workbook.xml"]), SML+" workbook", sheet)
            if not 1 <= sheet_count <= MAX_PARTS:
                raise InputError("sheet count")
            content_type_ok = False
            declared_macro = False
            content_keys = set()
            def content_type(stack, attrs):
                nonlocal content_type_ok, declared_macro
                if len(stack) == 1:
                    return
                if len(stack) != 2 or stack[1] not in (CT+" Override", CT+" Default"):
                    raise InputError("unexpected content-type element")
                typ = attrs.get("ContentType", "").lower()
                is_override = stack[1] == CT+" Override"
                key = attrs.get("PartName" if is_override else "Extension", "")
                if not key or not typ or (is_override and not key.startswith("/")):
                    raise InputError("missing or invalid content-type field")
                identity = (is_override, key.lower())
                if identity in content_keys:
                    raise InputError("duplicate content-type field")
                content_keys.add(identity)
                if "macroenabled" in typ or "vbaproject" in typ or "macrosheet" in typ:
                    declared_macro = True
                if is_override and key == "/xl/workbook.xml":
                    content_type_ok = typ in (
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                        "application/vnd.ms-excel.sheet.macroenabled.main+xml")
            inspect_xml(read_xml(names["[content_types].xml"]), CT+" Types", content_type)
            if not content_type_ok:
                raise InputError("unsupported workbook content type")
            office_documents = 0
            macro_relationship = False
            for name, part in names.items():
                if part.is_dir():
                    continue
                counts["external_link_parts"] += name.startswith("xl/externallinks/") and not name.endswith(".rels")
                counts["macro_parts"] += (name.endswith("/vbaproject.bin") or name.startswith("xl/macrosheets/")) and not name.endswith(".rels")
                counts["embedded_parts"] += name.startswith(("xl/embeddings/", "xl/activex/")) and not name.endswith(".rels")
                counts["connection_parts"] += name == "xl/connections.xml" or name.startswith("xl/querytables/") and not name.endswith(".rels")
                counts["pivot_cache_parts"] += name.startswith("xl/pivotcache/") and not name.endswith(".rels")
                if not name.endswith(".rels"):
                    continue
                relationship_ids = set()
                def relationship(stack, attrs):
                    nonlocal office_documents, macro_relationship
                    if len(stack) == 1:
                        return
                    if stack != (REL+" Relationships", REL+" Relationship"):
                        raise InputError("unexpected relationship element")
                    rid = attrs.get("Id")
                    if not rid or rid in relationship_ids or not attrs.get("Type") or not attrs.get("Target"):
                        raise InputError("missing or duplicate relationship fields")
                    relationship_ids.add(rid)
                    mode = attrs.get("TargetMode", "Internal")
                    if mode not in ("Internal", "External"):
                        raise InputError("unknown relationship mode")
                    counts["external_relationships"] += mode == "External"
                    if name == "xl/_rels/workbook.xml.rels" and attrs.get("Id") in sheet_ids:
                        rid = attrs["Id"]
                        target = attrs.get("Target", "")
                        if target.startswith("/xl/"):
                            target = target[1:]
                        else:
                            target = "xl/" + target
                        typ = SHEET_REL_TYPES.get(attrs["Type"])
                        if (rid in sheet_targets or mode != "Internal"
                                or typ is None
                                or target.lower() not in names or names[target.lower()].is_dir()):
                            raise InputError("unsupported sheet relationship")
                        sheet_targets[rid] = target.lower()
                        if typ in ("xlMacrosheet","xlIntlMacrosheet"):
                            macro_relationship = True
                    if name == "_rels/.rels" and attrs.get("Type") == OFFICE_REL:
                        if mode != "Internal" or attrs.get("Target") not in ("xl/workbook.xml", "/xl/workbook.xml"):
                            raise InputError("nonstandard workbook target")
                        office_documents += 1
                inspect_xml(read_xml(part), REL+" Relationships", relationship)
            if set(sheet_targets) != sheet_ids or len(set(sheet_targets.values())) != len(sheet_ids):
                raise InputError("missing or duplicate sheet target")
            if office_documents != 1:
                raise InputError("missing or ambiguous workbook relationship")
            if (declared_macro or macro_relationship) and counts["macro_parts"] == 0:
                counts["macro_parts"] = 1
    return {"sheets": sheet_count, **counts}

def audit_batch(paths, limits):
    if not 1 <= len(paths) <= MAX_BATCH:
        raise InputError("batch limit")
    results = []
    for index, path in enumerate(paths, 1):
        try:
            counts = inspect_workbook(path)
            violations = [key for key in METRICS if counts[key] > limits[key]]
            results.append({"index": index, "status": "fail" if violations else "pass",
                            "counts": counts, "violations": violations})
        except (InputError, OSError, ValueError, zipfile.BadZipFile, expat.ExpatError,
                RuntimeError, NotImplementedError, EOFError, OverflowError, zlib.error):
            results.append({"index": index, "status": "uninspected", "reason": "unsupported_or_invalid_input"})
    summary = {status: sum(r["status"] == status for r in results)
               for status in ("pass", "fail", "uninspected")}
    return {"policy_version": 1, "maximum": limits, "summary": summary, "files": results}
