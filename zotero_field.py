#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
zotero_field.py — pembangun payload field sitasi Zotero
======================================================
Bagian yang berbeda dari versi Mendeley:

    Mendeley :  ADDIN CSL_CITATION {json}
    Zotero   :  ADDIN ZOTERO_ITEM CSL_CITATION {json}

Perbedaan lain:
    * citationID unik 8 karakter per kelompok sitasi
    * properties: formattedCitation / plainCitation / noteIndex
    * citationItems[].uris menunjuk ke http://zotero.org/users/<id>/items/<key>
    * dokumen WAJIB punya properti ZOTERO_PREF_1 di docProps/custom.xml
"""

from __future__ import annotations

import hashlib
import json
import random
import string

import lxml.etree as ET

CSL_SCHEMA = ("https://github.com/citation-style-language/schema/"
              "raw/master/csl-citation.json")

ITEM_PREFIX = "ADDIN ZOTERO_ITEM CSL_CITATION "
BIBL_PREFIX = "ADDIN ZOTERO_BIBL "
BIBL_SUFFIX = " CSL_BIBLIOGRAPHY"

# Peta tipe item Zotero -> tipe CSL
CSL_TYPES = {
    "journalArticle": "article-journal",
    "magazineArticle": "article-magazine",
    "newspaperArticle": "article-newspaper",
    "conferencePaper": "paper-conference",
    "book": "book",
    "bookSection": "chapter",
    "thesis": "thesis",
    "report": "report",
    "preprint": "article",
    "manuscript": "manuscript",
    "webpage": "webpage",
    "blogPost": "post-weblog",
    "forumPost": "post",
    "letter": "personal_communication",
    "email": "personal_communication",
    "interview": "interview",
    "film": "motion_picture",
    "videoRecording": "motion_picture",
    "audioRecording": "song",
    "artwork": "graphic",
    "map": "map",
    "computerProgram": "software",
    "presentation": "speech",
    "dataset": "dataset",
    "patent": "patent",
    "encyclopediaArticle": "entry-encyclopedia",
    "dictionaryEntry": "entry-dictionary",
    "document": "article",
    "standard": "standard",
}


def as_csl_type(zotero_type, fallback="article-journal"):
    return CSL_TYPES.get(zotero_type or "", fallback)


def zotero_uri(key, user_id="", library_id=1, mode="auto"):
    """Bangun URI item Zotero.

    mode:
        user  -> http://zotero.org/users/<userID>/items/<key>
        local -> http://zotero.org/users/local/<libraryID>/items/<key>
        auto  -> 'user' kalau user_id diketahui, kalau tidak 'local'
    """
    key = str(key or "").strip()
    if mode == "local" or (mode == "auto" and not str(user_id or "").strip()):
        return f"http://zotero.org/users/local/{library_id}/items/{key}"
    return f"http://zotero.org/users/{user_id}/items/{key}"


def make_citation_id(n=8):
    pool = string.ascii_letters + string.digits
    rng = random.SystemRandom()
    return "".join(rng.choice(pool) for _ in range(n))


def stable_pseudo_id(key):
    """ID numerik deterministik kalau itemID asli tidak tersedia."""
    if not key:
        return 0
    digest = hashlib.blake2b(str(key).encode("utf-8"), digest_size=4).digest()
    return int.from_bytes(digest, "big") % 2_000_000_000 + 1


def build_item_data(ref, csl_type=None, item_id=None):
    """ref (dari Excel) -> dict itemData gaya CSL-JSON seperti yang ditulis Zotero."""
    item = {}
    if item_id:
        item["id"] = int(item_id)
    item["type"] = csl_type or ref.get("csl_type") or "article-journal"
    item["title"] = ref.get("title") or ""
    if ref.get("journal"):
        item["container-title"] = ref["journal"]
    for key in ("volume", "issue", "page"):
        if ref.get(key):
            item[key] = str(ref[key])
    if ref.get("publisher"):
        item["publisher"] = ref["publisher"]
    if ref.get("doi"):
        item["DOI"] = str(ref["doi"])
    if ref.get("authors"):
        item["author"] = ref["authors"]
    if ref.get("year"):
        item["issued"] = {"date-parts": [[str(ref["year"])]]}
    return {k: v for k, v in item.items() if v not in ("", None, [])}


def citation_field_code(entries, display, citation_id=None):
    """entries: list dict {key, uri, item}.  Menghasilkan kode field lengkap."""
    citation_id = citation_id or make_citation_id()
    citation_items = []
    for e in entries:
        item = dict(e.get("item") or {})
        if "id" not in item:
            item["id"] = stable_pseudo_id(e.get("key") or item.get("title"))
        cit = {"id": item["id"], "uris": [e["uri"]], "itemData": item}
        if e.get("locator"):
            cit["locator"] = str(e["locator"])
        if e.get("label"):
            cit["label"] = str(e["label"])
        citation_items.append(cit)

    payload = {
        "citationID": citation_id,
        "properties": {
            "formattedCitation": display,
            "plainCitation": display,
            "noteIndex": 0,
        },
        "citationItems": citation_items,
        "schema": CSL_SCHEMA,
    }
    return ITEM_PREFIX + json.dumps(payload, ensure_ascii=False, separators=(",", ":"))


def bibliography_field_code(custom=None, uncited=None, omitted=None):
    payload = {
        "uncited": uncited or [],
        "omitted": omitted or [],
        "custom": custom or [],
    }
    return BIBL_PREFIX + json.dumps(payload, ensure_ascii=False,
                                    separators=(",", ":")) + BIBL_SUFFIX


def build_prefs(user_id="", library_id=1, style_uri=None, locale="en-US"):
    """Isi properti ZOTERO_PREF_1.

    CATATAN: struktur ini masih kandidat. Kalibrasi lewat Fase 0 (lihat FLOW.md §6)
    dengan membandingkan dokumen Word yang dibuat langsung oleh plugin Zotero.
    """
    style_uri = style_uri or "http://www.zotero.org/styles/apa"
    url = (f"http://zotero.org/users/{user_id}/" if str(user_id or "").strip()
           else f"http://zotero.org/users/local/{library_id}/")
    return {
        "data": {
            "mode": "field",
            "prefsVersion": "2",
            "fieldType": "Field",
        },
        "url": url,
        "prefs": {
            "citationStyle": style_uri,
            "locale": locale,
        },
        "version": 3,
    }


# --------------------------------------------------------------------------
# Injeksi properti dokumen (docProps/custom.xml)
# --------------------------------------------------------------------------
CP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CUSTOM_CT = "application/vnd.openxmlformats-officedocument.custom-properties+xml"
CUSTOM_REL = ("http://schemas.openxmlformats.org/officeDocument/2006/"
              "relationships/custom-properties")
CUSTOM_PART = "docProps/custom.xml"
FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"


def _serialize(root):
    return ET.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _ensure_content_type(members):
    name = "[Content_Types].xml"
    if name not in members:
        return None
    root = ET.fromstring(members[name])
    target = "/" + CUSTOM_PART
    for ov in root.findall(f"{{{CT_NS}}}Override"):
        if ov.get("PartName") == target:
            if ov.get("ContentType") != CUSTOM_CT:
                ov.set("ContentType", CUSTOM_CT)
            return _serialize(root)
    ET.SubElement(root, f"{{{CT_NS}}}Override",
                  PartName=target, ContentType=CUSTOM_CT)
    return _serialize(root)


def _ensure_relationship(members):
    name = "_rels/.rels"
    if name in members:
        root = ET.fromstring(members[name])
    else:
        root = ET.Element(f"{{{REL_NS}}}Relationships")
    for rel in root.findall(f"{{{REL_NS}}}Relationship"):
        if rel.get("Type") == CUSTOM_REL:
            return _serialize(root)
    ids = {r.get("Id") for r in root.findall(f"{{{REL_NS}}}Relationship")}
    new_id, i = "rIdCustomProps", 1
    while new_id in ids:
        i += 1
        new_id = f"rIdCustomProps{i}"
    ET.SubElement(root, f"{{{REL_NS}}}Relationship",
                  Id=new_id, Type=CUSTOM_REL, Target=CUSTOM_PART)
    return _serialize(root)


def _set_custom_property(members, prop_name, value_json):
    name = CUSTOM_PART
    if name in members:
        root = ET.fromstring(members[name])
    else:
        root = ET.Element(f"{{{CP_NS}}}Properties", nsmap={None: CP_NS, "vt": VT_NS})
    props = root.findall(f"{{{CP_NS}}}property")
    pids = []
    for p in props:
        try:
            pids.append(int(p.get("pid") or 0))
        except ValueError:
            pass
        if p.get("name") == prop_name:
            for child in list(p):
                p.remove(child)
            lpw = ET.SubElement(p, f"{{{VT_NS}}}lpwstr")
            lpw.text = value_json
            return _serialize(root)
    pid = (max(pids) if pids else 1) + 1
    prop = ET.SubElement(root, f"{{{CP_NS}}}property",
                         fmtid=FMTID, pid=str(pid), name=prop_name)
    lpw = ET.SubElement(prop, f"{{{VT_NS}}}lpwstr")
    lpw.text = value_json
    return _serialize(root)


def inject_doc_prefs(members, prefs_obj, prop_name="ZOTERO_PREF_1"):
    """Kembalikan dict part yang berubah: Content_Types, _rels/.rels, docProps/custom.xml."""
    changed = {}
    value = json.dumps(prefs_obj, ensure_ascii=False, separators=(",", ":"))
    for fn in (_ensure_content_type, _ensure_relationship):
        res = fn(members)
        if res is not None:
            part = "[Content_Types].xml" if fn is _ensure_content_type else "_rels/.rels"
            if members.get(part) != res:
                changed[part] = res
                members[part] = res
    custom = _set_custom_property(members, prop_name, value)
    if custom is not None:
        changed[CUSTOM_PART] = custom
        members[CUSTOM_PART] = custom
    return changed
