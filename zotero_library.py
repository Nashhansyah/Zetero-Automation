#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
zotero_library.py — lapisan akses library Zotero
================================================
Satu kontrak, tiga sumber data:

    localapi  Zotero desktop, http://localhost:23119/api/   (default, tanpa auth)
    sqlite    baca langsung zotero.sqlite lewat salinan temporer
    webapi    api.zotero.org (butuh API key, untuk library group / lintas mesin)
    fixture   file JSON lokal — untuk pengujian tanpa Zotero

Semua adapter mengembalikan list[Ref] yang sudah dinormalisasi, sehingga
pipeline di zotero_cite.py tidak peduli sumbernya.
"""

from __future__ import annotations

import difflib
import json
import os
import re
import shutil
import sqlite3
import tempfile
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass, field
from pathlib import Path

from zotero_field import as_csl_type, zotero_uri

LOCAL_BASE = "http://127.0.0.1:23119/api"
WEB_BASE = "https://api.zotero.org"
UA = "zotero-automation/0.1 (word-citation-converter)"


# --------------------------------------------------------------------------
# Normalisasi
# --------------------------------------------------------------------------
def norm_title(s) -> str:
    return re.sub(r"[^a-z0-9]", "", str(s or "").lower())


def norm_doi(s) -> str:
    d = str(s or "").strip().lower()
    for pre in ("https://doi.org/", "http://doi.org/", "https://dx.doi.org/",
                "http://dx.doi.org/", "https://www.doi.org/", "doi:"):
        d = d.removeprefix(pre)
    d = d.split()[0] if d.split() else ""
    return d.rstrip(".,;)")


def parse_year(v):
    m = re.search(r"(?:19|20)\d{2}", str(v or ""))
    return m.group(0) if m else None


def csl_name(family, given="", literal=None):
    if literal:
        return {"literal": literal}
    return {"family": family or "", "given": given or ""}


@dataclass
class Ref:
    key: str
    title: str = ""
    authors: list = field(default_factory=list)
    year: str = ""
    journal: str = ""
    volume: str = ""
    issue: str = ""
    page: str = ""
    doi: str = ""
    type_name: str = ""
    csl_type: str = "article-journal"
    item_id: int | None = None
    library_id: int = 1
    user_id: str = ""
    uri: str = ""

    def label(self):
        return f"{self.key} | {self.title[:70]}"


# --------------------------------------------------------------------------
# Indeks pencocokan (sama dengan versi Mendeley)
# --------------------------------------------------------------------------
class Library:
    def __init__(self, docs: list[Ref]):
        self.docs = docs
        self.by_doi, self.by_title = {}, {}
        for d in docs:
            if d.doi:
                self.by_doi.setdefault(d.doi, d)
            if d.title:
                self.by_title.setdefault(norm_title(d.title), d)

    def __len__(self):
        return len(self.docs)

    def lookup(self, doi, title, year=None):
        doi = norm_doi(doi)
        if doi and doi in self.by_doi:
            return self.by_doi[doi], "DOI"
        nt = norm_title(title)
        if nt and nt in self.by_title:
            return self.by_title[nt], "Judul (Persis)"
        if nt:
            best, score = None, 0.0
            for key, d in self.by_title.items():
                if not key or abs(len(key) - len(nt)) > 40:
                    continue
                r = difflib.SequenceMatcher(None, nt, key).ratio()
                if r > score:
                    best, score = d, r
            if best is not None and score >= 0.90:
                return best, f"Judul (Fuzzy {int(score * 100)}%)"
        return None, ""


# --------------------------------------------------------------------------
# Utilitas lingkungan
# --------------------------------------------------------------------------
def default_zotero_dir() -> Path:
    return Path(os.environ.get("USERPROFILE") or Path.home()) / "Zotero"


def find_db(zotero_dir=None):
    base = Path(zotero_dir) if zotero_dir else default_zotero_dir()
    db = base / "zotero.sqlite"
    return db if db.is_file() else None


def detect_user_id(zotero_dir=None) -> str:
    """Cari ID akun Zotero dari prefs.js (cepat, tanpa menyentuh DB)."""
    base = Path(zotero_dir) if zotero_dir else default_zotero_dir()
    prefs = base / "prefs.js"
    if prefs.is_file():
        try:
            txt = prefs.read_text(encoding="utf-8", errors="ignore")
            m = re.search(
                r'user_pref\("extensions\.zotero\.account\.userID",\s*"?(\d+)"?\)',
                txt)
            if m:
                return m.group(1)
        except OSError:
            pass
    db = find_db(zotero_dir)
    if db:
        try:
            con = sqlite3.connect(f"file:{db}?mode=ro&immutable=1", uri=True, timeout=2)
            try:
                row = con.execute(
                    "select value from settings where setting='account' and key='userID'"
                ).fetchone()
                if row and str(row[0]).isdigit():
                    return str(row[0])
            finally:
                con.close()
        except sqlite3.Error:
            pass
    return ""


# --------------------------------------------------------------------------
# Adapter dasar
# --------------------------------------------------------------------------
class AdapterError(RuntimeError):
    pass


class BaseAdapter:
    name = "base"

    def __init__(self, cfg: dict):
        self.cfg = cfg or {}
        self.user_id = str(self.cfg.get("user_id") or "").strip()
        try:
            self.library_id = int(self.cfg.get("library_id") or 1)
        except (TypeError, ValueError):
            self.library_id = 1
        self.uri_mode = (self.cfg.get("uri_mode") or "auto").lower()
        self.zotero_dir = self.cfg.get("zotero_data_dir") or None

    # --- kontrak ---
    def available(self) -> bool:
        return True

    def ping(self) -> bool:
        return False

    def library(self) -> list[Ref]:
        raise NotImplementedError

    # --- util ---
    def ensure_user_id(self):
        if not self.user_id:
            self.user_id = detect_user_id(self.zotero_dir)

    def uri_for(self, key) -> str:
        return zotero_uri(key, self.user_id, self.library_id, self.uri_mode)

    @property
    def info(self) -> dict:
        return {"adapter": self.name, "user_id": self.user_id or "-",
                "library_id": self.library_id, "uri_mode": self.uri_mode}


# --------------------------------------------------------------------------
# Adapter 1: Zotero Local API
# --------------------------------------------------------------------------
class LocalApiAdapter(BaseAdapter):
    name = "localapi"

    def __init__(self, cfg):
        super().__init__(cfg)
        self.base = (self.cfg.get("local_api_base") or LOCAL_BASE).rstrip("/")
        self.timeout = float(self.cfg.get("timeout") or 8)

    def _get(self, path):
        url = self.base + path
        req = urllib.request.Request(url, headers={
            "Accept": "application/json",
            "User-Agent": UA,
            "Zotero-API-Version": "3",
        })
        try:
            with urllib.request.urlopen(req, timeout=self.timeout) as r:
                return json.loads(r.read().decode("utf-8"))
        except urllib.error.HTTPError as exc:
            if exc.code == 403:
                raise AdapterError(
                    "Local API Zotero menolak (403). Aktifkan: Settings -> Advanced -> "
                    "\"Allow other applications on this computer to communicate with Zotero\"."
                ) from exc
            raise AdapterError(f"Local API error {exc.code}: {path}") from exc
        except (urllib.error.URLError, OSError) as exc:
            raise AdapterError(
                f"Tidak bisa menghubungi Zotero di {self.base}. Pastikan Zotero berjalan. ({exc})"
            ) from exc

    def ping(self):
        try:
            self._get("/users/0/items/top?format=json&limit=1")
            return True
        except AdapterError:
            return False

    def library(self):
        self.ensure_user_id()
        rows = []
        start = 0
        while True:
            page = self._get(f"/users/0/items/top?format=json&limit=100&start={start}")
            if not isinstance(page, list):
                break
            rows.extend(page)
            if len(page) < 100:
                break
            start += 100

        out = []
        for obj in rows:
            data = obj.get("data") or {}
            if (data.get("itemType") or "") in ("attachment", "note", "annotation"):
                continue
            out.append(self._to_ref(data, obj.get("library") or {}))
        return out

    def _to_ref(self, d, lib):
        lib_id = lib.get("id") or self.library_id
        key = d.get("key") or ""
        authors = []
        for c in d.get("creators") or []:
            if c.get("creatorType") not in ("author", "editor", None):
                continue
            if c.get("name"):
                authors.append(csl_name("", "", literal=c["name"]))
            elif c.get("lastName") or c.get("firstName"):
                authors.append(csl_name(c.get("lastName", ""), c.get("firstName", "")))
        if not authors:
            for c in d.get("creators") or []:
                if c.get("lastName"):
                    authors.append(csl_name(c.get("lastName", ""), c.get("firstName", "")))
        journal = (d.get("publicationTitle") or d.get("proceedingsTitle")
                   or d.get("bookTitle") or d.get("websiteTitle")
                   or d.get("blogTitle") or d.get("forumTitle")
                   or d.get("repository") or "")
        return Ref(
            key=key,
            title=d.get("title") or "",
            authors=authors,
            year=parse_year(d.get("date") or d.get("dateDecided") or ""),
            journal=journal,
            volume=str(d.get("volume") or ""),
            issue=str(d.get("issue") or ""),
            page=str(d.get("pages") or ""),
            doi=norm_doi(d.get("DOI") or ""),
            type_name=d.get("itemType") or "",
            csl_type=as_csl_type(d.get("itemType")),
            item_id=None,
            library_id=int(lib_id),
            user_id=self.user_id,
            uri=zotero_uri(key, self.user_id, lib_id, self.uri_mode),
        )


# --------------------------------------------------------------------------
# Adapter 2: SQLite langsung
# --------------------------------------------------------------------------
class SqliteAdapter(BaseAdapter):
    name = "sqlite"

    def available(self):
        return find_db(self.zotero_dir) is not None

    def ping(self):
        return self.available()

    def _read_items(self, con):
        rows = con.execute(
            """
            SELECT i.itemID, i.key, i.libraryID, it.typeName
            FROM items i
            JOIN itemTypes it ON it.itemTypeID = i.itemTypeID
            WHERE i.itemID NOT IN (SELECT itemID FROM deletedItems)
              AND it.typeName NOT IN ('attachment','note','annotation')
            """
        ).fetchall()

        fields = {}
        for item_id, name, value in con.execute(
            """
            SELECT d.itemID, f.fieldName, v.value
            FROM itemData d
            JOIN fields f ON f.fieldID = d.fieldID
            JOIN itemDataValues v ON v.valueID = d.valueID
            """
        ):
            fields.setdefault(item_id, {})[name] = value

        creators = {}
        for item_id, first, last, mode, ctype, _order in con.execute(
            """
            SELECT ic.itemID, c.firstName, c.lastName, c.fieldMode, ct.creatorType, ic.orderIndex
            FROM itemCreators ic
            JOIN creators c ON c.creatorID = ic.creatorID
            JOIN creatorTypes ct ON ct.creatorTypeID = ic.creatorTypeID
            ORDER BY ic.itemID, ic.orderIndex
            """
        ):
            creators.setdefault(item_id, []).append((first, last, mode, ctype))
        return rows, fields, creators

    def library(self):
        self.ensure_user_id()
        db = find_db(self.zotero_dir)
        if not db:
            raise AdapterError(f"zotero.sqlite tidak ditemukan di {self.zotero_dir or default_zotero_dir()}")

        tmpdir = Path(tempfile.mkdtemp(prefix="zotero-read-"))
        tmp_db = tmpdir / "zotero.sqlite"
        try:
            for suf in ("", "-wal", "-shm"):
                src = Path(str(db) + suf)
                if src.exists():
                    try:
                        shutil.copy2(src, str(tmp_db) + suf)
                    except OSError:
                        pass
            try:
                con = sqlite3.connect(tmp_db)
                rows, fields, creators = self._read_items(con)
                con.close()
            except sqlite3.Error as exc:
                raise AdapterError(f"Gagal membaca zotero.sqlite: {exc}") from exc

            out = []
            for item_id, key, lib_id, type_name in rows:
                f = fields.get(item_id, {})
                authors = []
                for first, last, mode, ctype in creators.get(item_id, []):
                    if ctype not in ("author", "editor"):
                        continue
                    if mode == 1:
                        authors.append(csl_name("", "", literal=last or first or ""))
                    else:
                        authors.append(csl_name(last or "", first or ""))
                journal = (f.get("publicationTitle") or f.get("proceedingsTitle")
                           or f.get("bookTitle") or f.get("websiteTitle")
                           or f.get("blogTitle") or f.get("forumTitle") or "")
                out.append(Ref(
                    key=key,
                    title=f.get("title") or "",
                    authors=authors,
                    year=parse_year(f.get("date")),
                    journal=journal,
                    volume=str(f.get("volume") or ""),
                    issue=str(f.get("issue") or ""),
                    page=str(f.get("pages") or ""),
                    doi=norm_doi(f.get("DOI") or ""),
                    type_name=type_name,
                    csl_type=as_csl_type(type_name),
                    item_id=item_id,
                    library_id=lib_id,
                    user_id=self.user_id,
                    uri=zotero_uri(key, self.user_id, lib_id, self.uri_mode),
                ))
            return out
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)


# --------------------------------------------------------------------------
# Adapter 3: Web API
# --------------------------------------------------------------------------
class WebApiAdapter(BaseAdapter):
    name = "webapi"

    def __init__(self, cfg):
        super().__init__(cfg)
        self.api_key = (self.cfg.get("api_key")
                        or os.environ.get("ZOTERO_API_KEY") or "")
        self.group = self.cfg.get("group_id") or None

    def available(self):
        return bool(self.api_key) and bool(self.user_id or self.group)

    def ping(self):
        return self.available()

    def _get(self, path):
        req = urllib.request.Request(WEB_BASE + path, headers={
            "Accept": "application/json",
            "User-Agent": UA,
            "Zotero-API-Key": self.api_key,
            "Zotero-API-Version": "3",
        })
        with urllib.request.urlopen(req, timeout=20) as r:
            return json.loads(r.read().decode("utf-8"))

    def library(self):
        self.ensure_user_id()
        prefix = (f"/groups/{self.group}" if self.group
                  else f"/users/{self.user_id}")
        rows, start = [], 0
        while True:
            page = self._get(f"{prefix}/items/top?format=json&limit=100&start={start}")
            if not isinstance(page, list):
                break
            rows.extend(page)
            if len(page) < 100:
                break
            start += 100
        out = []
        for obj in rows:
            d = obj.get("data") or {}
            if (d.get("itemType") or "") in ("attachment", "note", "annotation"):
                continue
            tmp = LocalApiAdapter.__new__(LocalApiAdapter)
            tmp.user_id = self.user_id
            tmp.library_id = self.library_id
            tmp.uri_mode = self.uri_mode
            out.append(tmp._to_ref(d, obj.get("library") or {}))
        return out


# --------------------------------------------------------------------------
# Adapter 4: Fixture JSON (untuk uji coba tanpa Zotero)
# --------------------------------------------------------------------------
class FixtureAdapter(BaseAdapter):
    name = "fixture"

    def __init__(self, cfg):
        super().__init__(cfg)
        self.path = Path(self.cfg.get("fixture") or "samples/library_fixture.json")

    def available(self):
        return self.path.is_file()

    def ping(self):
        return self.available()

    def library(self):
        self.ensure_user_id()
        raw = json.loads(self.path.read_text(encoding="utf-8"))
        rows = raw.get("items") if isinstance(raw, dict) else raw
        out = []
        for d in rows or []:
            authors = []
            for a in d.get("authors") or []:
                if isinstance(a, str):
                    authors.append(csl_name(a, ""))
                else:
                    authors.append(csl_name(a.get("family", ""), a.get("given", ""),
                                            literal=a.get("literal")))
            key = str(d.get("key") or "")
            lib_id = int(d.get("libraryID") or self.library_id)
            out.append(Ref(
                key=key,
                title=d.get("title") or "",
                authors=authors,
                year=str(d.get("year") or "") or parse_year(d.get("date")),
                journal=d.get("journal") or "",
                volume=str(d.get("volume") or ""),
                issue=str(d.get("issue") or ""),
                page=str(d.get("page") or d.get("pages") or ""),
                doi=norm_doi(d.get("doi") or ""),
                type_name=d.get("itemType") or "journalArticle",
                csl_type=as_csl_type(d.get("itemType")),
                item_id=d.get("itemID"),
                library_id=lib_id,
                user_id=self.user_id,
                uri=zotero_uri(key, self.user_id, lib_id, self.uri_mode),
            ))
        return out


# --------------------------------------------------------------------------
# Factory
# --------------------------------------------------------------------------
ADAPTERS = {
    "localapi": LocalApiAdapter,
    "sqlite": SqliteAdapter,
    "webapi": WebApiAdapter,
    "fixture": FixtureAdapter,
}


def get_adapter(cfg: dict) -> BaseAdapter:
    mode = (cfg.get("adapter") or "auto").lower()
    if mode != "auto":
        cls = ADAPTERS.get(mode)
        if not cls:
            raise AdapterError(f"Adapter tidak dikenal: {mode}. Pilihan: {', '.join(ADAPTERS)}")
        return cls(cfg)

    # Catatan: adapter fixture sengaja TIDAK ikut dalam rantai "auto".
    # Kalau ikut, mesin tanpa Zotero akan diam-diam memakai library contoh
    # dan hasilnya menyesatkan. Fixture hanya aktif kalau diminta eksplisit.
    notes = []
    local = LocalApiAdapter(cfg)
    if local.ping():
        return local
    notes.append("localapi tidak aktif")

    sql = SqliteAdapter(cfg)
    if sql.available():
        return sql
    notes.append("zotero.sqlite tidak ditemukan")

    raise AdapterError(
        "Tidak ada sumber library yang bisa dipakai (" + "; ".join(notes) + ").\n"
        "  - Jalankan Zotero dan aktifkan Settings -> Advanced -> izin komunikasi, atau\n"
        "  - tutup Zotero lalu pakai --adapter sqlite, atau\n"
        "  - untuk uji coba tanpa Zotero: --adapter fixture --fixture <path>."
    )
