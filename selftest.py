#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
selftest.py — uji end-to-end tanpa perlu Zotero terpasang
=========================================================
1. Membuat naskah Word + Excel + library fixture di folder samples/
2. Menjalankan zotero_cite.py dengan adapter fixture
3. Memeriksa hasil: field Zotero valid, ZOTERO_PREF_1 ada, laporan benar

Jalankan:  python selftest.py
"""

from __future__ import annotations

import csv
import json
import subprocess
import sys
import zipfile
from pathlib import Path

import lxml.etree as ET
import openpyxl
from docx import Document
from docx.shared import Pt

ROOT = Path(__file__).resolve().parent
SAMPLES = ROOT / "samples"
W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
CP_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/custom-properties}"
VT_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}"
CT_NS = "{http://schemas.openxmlformats.org/package/2006/content-types}"
REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"

DOCX = SAMPLES / "naskah_uji.docx"
XLSX = SAMPLES / "referensi_uji.xlsx"
FIXTURE = SAMPLES / "library_fixture.json"
OUT = SAMPLES / "manuscript - Zotero.docx"
REPORT = SAMPLES / "laporan_sitasi.xlsx"
RIS = SAMPLES / "referensi_belum_ada.ris"
CSV = SAMPLES / "mapping_report.csv"

USER_ID = "1234567"

ROWS = [
    # id, title, doi, authors, year, journal
    ("NR001", "Deep Learning for Medical Imaging", "10.1000/demo001",
     "Smith, John; Doe, Jane", "2021", "Nature Machine Intelligence"),
    ("NR002", "Transformer Models in Natural Language Processing", "10.1000/demo002",
     "Lee, Min; Park, Soo", "2022", "Journal of Artificial Intelligence"),
    ("NR003", "Graph Neural Networks: A Survey", "10.1000/demo003",
     "Brown, Alice", "2020", "IEEE Transactions on Pattern Analysis"),
    ("NR004", "Self-Supervised Learning for Vision", "10.1000/demo004",
     "Garcia, Luis", "2023", "CVPR Proceedings"),
    ("NR005", "Quantum Computing for Molecular Chemistry", "10.1000/demo005",
     "Kim, Hana", "2019", "Science"),
    ("NR010", "A Reference Never Cited in the Manuscript", "10.1000/demo010",
     "Unused, Author", "2018", "Unused Journal"),
]

LIBRARY = {
    "userID": USER_ID,
    "items": [
        {"key": "ABCD0001", "itemID": 101, "itemType": "journalArticle",
         "title": "Deep Learning for Medical Imaging", "doi": "10.1000/demo001",
         "authors": [{"family": "Smith", "given": "John"}, {"family": "Doe", "given": "Jane"}],
         "year": "2021", "journal": "Nature Machine Intelligence",
         "volume": "3", "issue": "2", "page": "101-115"},
        {"key": "ABCD0002", "itemID": 102, "itemType": "journalArticle",
         "title": "Transformer Models in Natural Language Processing", "doi": "10.1000/demo002",
         "authors": [{"family": "Lee", "given": "Min"}, {"family": "Park", "given": "Soo"}],
         "year": "2022", "journal": "Journal of Artificial Intelligence",
         "volume": "12", "issue": "4", "page": "55-72"},
        {"key": "ABCD0003", "itemID": 103, "itemType": "journalArticle",
         "title": "Graph Neural Networks: A Survey", "doi": "10.1000/demo003",
         "authors": [{"family": "Brown", "given": "Alice"}],
         "year": "2020", "journal": "IEEE Transactions on Pattern Analysis",
         "volume": "42", "issue": "8", "page": "2001-2020"},
        {"key": "ABCD0004", "itemID": 104, "itemType": "conferencePaper",
         "title": "Self-Supervised Learning for Vision", "doi": "10.1000/demo004",
         "authors": [{"family": "Garcia", "given": "Luis"}],
         "year": "2023", "journal": "CVPR Proceedings",
         "page": "3300-3312"},
    ],
}


def build_samples():
    SAMPLES.mkdir(exist_ok=True)

    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    doc.add_heading("Naskah Uji Otomatisasi Sitasi Zotero", level=1)
    p = doc.add_paragraph()
    p.add_run("Pendahuluan penelitian ini mengacu pada studi terdahulu (NR0")
    p.add_run("01; NR002).")          # sengaja dipotong antar-run
    doc.add_paragraph(
        "Beberapa penelitian lain [NR003-NR004] juga mendukung temuan tersebut.")
    doc.add_paragraph(
        "Namun ada klaim yang belum terverifikasi (NR005).")
    doc.add_paragraph(
        "Sitasi yang tidak ada di Excel: (NR099).")
    doc.add_paragraph(
        "Paragraf ini tidak memiliki placeholder sitasi sama sekali.")
    doc.save(DOCX)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data Referensi"
    ws.append(["Study_ID", "Judul", "DOI", "Penulis", "Tahun", "Jurnal"])
    for row in ROWS:
        ws.append(list(row))
    wb.save(XLSX)

    FIXTURE.write_text(json.dumps(LIBRARY, ensure_ascii=False, indent=2), encoding="utf-8")


def run_pipeline():
    cmd = [
        sys.executable, "zotero_cite.py",
        "--folder", str(SAMPLES),
        "--docx", str(DOCX),
        "--workbook", str(XLSX),
        "--out", str(OUT),
        "--adapter", "fixture",
        "--fixture", str(FIXTURE),
        "--user-id", USER_ID,
        "--style", "author-year",
        "--no-enrich",
    ]
    proc = subprocess.run(cmd, cwd=ROOT, capture_output=True, text=True,
                          encoding="utf-8", errors="replace")
    return proc


def field_payloads(docx_path):
    with zipfile.ZipFile(docx_path) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    out = []
    for node in root.iter(W_NS + "instrText"):
        code = "".join(node.itertext())
        if code.startswith("ADDIN ZOTERO_ITEM CSL_CITATION "):
            out.append(json.loads(code[len("ADDIN ZOTERO_ITEM CSL_CITATION "):]))
    return out


def main():
    checks = []

    def check(name, ok, detail=""):
        checks.append((name, bool(ok), detail))
        mark = "PASS" if ok else "FAIL"
        line = f"[{mark}] {name}"
        if detail and not ok:
            line += f"  -> {detail}"
        print(line)

    print("=" * 74)
    print("  SELF-TEST: ZOTERO CITATION AUTOMATION")
    print("=" * 74)

    build_samples()
    print(f"  Sampel dibuat : {SAMPLES}")
    proc = run_pipeline()
    if proc.returncode != 0:
        print("\n--- STDOUT ---\n" + proc.stdout)
        print("\n--- STDERR ---\n" + proc.stderr)
        check("pipeline selesai tanpa error", False, f"exit={proc.returncode}")
        summarize(checks)
        return 1
    check("pipeline selesai tanpa error", True)

    # --- output dasar ---
    check("docx hasil dibuat", OUT.is_file(), str(OUT))
    check("laporan Excel dibuat", REPORT.is_file(), str(REPORT))
    check("file RIS dibuat", RIS.is_file(), str(RIS))
    check("mapping_report.csv dibuat", CSV.is_file(), str(CSV))

    # --- field Zotero ---
    payloads = field_payloads(OUT)
    check("jumlah kelompok field Zotero == 2", len(payloads) == 2, f"ditemukan {len(payloads)}")

    if payloads:
        p1 = payloads[0]
        check("citationID panjang 8", len(p1.get("citationID", "")) == 8,
              str(p1.get("citationID")))
        check("citationID unik antar field",
              len(payloads) == len({p.get("citationID") for p in payloads}))
        check("schema CSL citation benar",
              p1.get("schema", "").endswith("csl-citation.json"))
        check("properties.formattedCitation ada",
              bool(p1.get("properties", {}).get("formattedCitation")))

        cits = p1.get("citationItems") or []
        check("field pertama berisi 2 item", len(cits) == 2, f"ditemukan {len(cits)}")
        if cits:
            uri = (cits[0].get("uris") or [""])[0]
            check("URI memakai userID dari config",
                  uri == f"http://zotero.org/users/{USER_ID}/items/ABCD0001", uri)
            data = cits[0].get("itemData") or {}
            check("itemData.id = itemID asli", data.get("id") == 101, str(data.get("id")))
            check("itemData.type = article-journal", data.get("type") == "article-journal")
            check("itemData.DOI terisi", data.get("DOI") == "10.1000/demo001")
            check("itemData.author 2 nama", len(data.get("author") or []) == 2)
            check("itemData.issued tahun 2021",
                  (data.get("issued") or {}).get("date-parts") == [["2021"]])

    # --- teks tampilan ---
    rendered = "\n".join(p.text for p in Document(OUT).paragraphs)
    check("display sitasi pertama benar",
          "(Smith & Doe, 2021; Lee & Park, 2022)" in rendered, rendered[:200])
    check("display grup kedua benar", "(Brown, 2020; Garcia, 2023)" in rendered, rendered[:200])
    check("placeholder yang berhasil hilang", "NR001" not in rendered and "NR003" not in rendered)
    check("placeholder gagal dipertahankan sebagai teks",
          "(NR005)" in rendered and "(NR099)" in rendered)

    # --- docx masih valid ---
    try:
        doc = Document(OUT)
        ok = len(doc.paragraphs) >= 5
    except Exception as exc:
        ok = False
    check("docx hasil bisa dibuka python-docx", ok)

    # --- properti dokumen ---
    with zipfile.ZipFile(OUT) as z:
        names = set(z.namelist())
        check("part docProps/custom.xml ada", "docProps/custom.xml" in names)
        custom = ET.fromstring(z.read("docProps/custom.xml")) if "docProps/custom.xml" in names else None
        ct = ET.fromstring(z.read("[Content_Types].xml"))
        rels = ET.fromstring(z.read("_rels/.rels"))

    prefs_val = None
    if custom is not None:
        for prop in custom.findall(CP_NS + "property"):
            if prop.get("name") == "ZOTERO_PREF_1":
                prefs_val = (prop.findtext(VT_NS + "lpwstr") or "")
    check("ZOTERO_PREF_1 ada", bool(prefs_val))
    if prefs_val:
        try:
            prefs = json.loads(prefs_val)
            ok = prefs.get("url", "").endswith(f"/users/{USER_ID}/")
            detail = prefs.get("url", "")
        except Exception:
            ok, detail = False, prefs_val[:80]
        check("prefs.url memakai userID", ok, detail)

    overrides = [o.get("PartName") for o in ct.findall(CT_NS + "Override")]
    check("Content_Types punya Override custom-properties",
          "/docProps/custom.xml" in overrides, str(overrides))
    rel_targets = [r.get("Target") for r in rels.findall(REL_NS + "Relationship")]
    check("_rels/.rels menunjuk docProps/custom.xml",
          "docProps/custom.xml" in rel_targets, str(rel_targets))

    # --- laporan ---
    wb = openpyxl.load_workbook(REPORT, data_only=True)
    check("laporan punya 4 sheet",
          wb.sheetnames == ["Ringkasan", "Detail Sitasi Word", "Perlu Tindakan", "Excel Belum Digunakan"],
          str(wb.sheetnames))

    detail = wb["Detail Sitasi Word"]
    status = {}
    for row in detail.iter_rows(min_row=2, values_only=True):
        if row[1]:
            status[str(row[1])] = str(row[2])
    expected = {
        "NR001": "BERHASIL", "NR002": "BERHASIL",
        "NR003": "BERHASIL", "NR004": "BERHASIL",
        "NR005": "GAGAL: TIDAK ADA DI ZOTERO",
        "NR099": "GAGAL: TIDAK ADA DI EXCEL",
    }
    for sid, want in expected.items():
        check(f"status {sid} = {want}", status.get(sid) == want, str(status.get(sid)))

    ring = wb["Ringkasan"]
    metrics = {}
    for row in ring.iter_rows(min_row=12, values_only=True):
        if row[0]:
            metrics[str(row[0])] = row[1]
    check("ringkasan: 4 sitasi berhasil",
          metrics.get("Sitasi BERHASIL (Ada di Zotero)") == 4,
          str(metrics.get("Sitasi BERHASIL (Ada di Zotero)")))
    check("ringkasan: 1 gagal Zotero",
          metrics.get("Sitasi GAGAL (Tidak Ada di Zotero)") == 1)
    check("ringkasan: 1 gagal Excel",
          metrics.get("Sitasi GAGAL (Tidak Terdaftar di Excel)") == 1)
    check("ringkasan: 1 referensi Excel tidak disitasi",
          metrics.get("Referensi di Excel yang Tidak Disitasi di Word") == 1)

    # --- RIS ---
    ris_text = RIS.read_text(encoding="utf-8")
    check("RIS berisi referensi NR005",
          "Quantum Computing for Molecular Chemistry" in ris_text)
    check("RIS tidak berisi referensi yang sudah ada",
          "Deep Learning for Medical Imaging" not in ris_text)

    # --- CSV ---
    with open(CSV, encoding="utf-8-sig") as f:
        rows = list(csv.reader(f))
    check("CSV berisi 6 baris data", len(rows) == 7, f"{len(rows) - 1} baris")

    return summarize(checks)


def summarize(checks):
    total = len(checks)
    passed = sum(1 for _, ok, _ in checks if ok)
    print("\n" + "=" * 74)
    print(f"  HASIL: {passed}/{total} pemeriksaan lulus")
    failed = [c for c in checks if not c[1]]
    if failed:
        print("  GAGAL:")
        for name, _, detail in failed:
            print(f"    - {name} {detail}")
    print("=" * 74)
    return 0 if not failed else 1


if __name__ == "__main__":
    sys.exit(main())
