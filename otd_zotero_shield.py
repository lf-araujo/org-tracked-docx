#!/usr/bin/env python3
"""Shield/unshield Zotero Google-Docs hyperlink fields in a .docx.

Used by org-tracked-docx-zotero.el to roundtrip a docx through pandoc without
losing Zotero citation hyperlinks.  The Zotero "Switch Word Processor"
integration for Google Docs stores citations as Word HYPERLINK fields
pointing at https://www.zotero.org/google-docs/?<code> URLs.  Pandoc's
docx<->markdown roundtrip silently drops these.

shield   docx -> shielded.docx + sidecar.json
         Each Zotero field group <w:r>...fldChar begin..HYPERLINK..end...</w:r>
         is replaced with a single <w:r><w:t>U+2770 ZOT NNNN U+2771</w:t></w:r>
         placeholder.  The original XML chunks are saved in sidecar.json
         keyed by the four-digit token id.

unshield docx + sidecar.json -> final.docx
         Each placeholder in the post-pandoc docx is located in its
         enclosing run and replaced with the saved field-group XML, splitting
         surrounding text on either side into separate runs that inherit the
         original run's <w:rPr>.
"""

import argparse
import json
import os
import re
import shutil
import sys
import zipfile
from xml.sax.saxutils import escape as xml_escape


PLACEHOLDER_OPEN = "❰"   # heavy left-pointing angle bracket
PLACEHOLDER_CLOSE = "❱"  # heavy right-pointing angle bracket


def placeholder(key: str) -> str:
    return f"{PLACEHOLDER_OPEN}ZOT{key}{PLACEHOLDER_CLOSE}"


PLACEHOLDER_RE = re.compile(
    re.escape(PLACEHOLDER_OPEN) + r"ZOT(\d{4})" + re.escape(PLACEHOLDER_CLOSE)
)

# A complete field-character run group: from a <w:r> with fldChar begin through
# the <w:r> with fldChar end.  Non-greedy on inner <w:r> so we capture the
# minimum span containing both markers.
FIELD_RE = re.compile(
    r'<w:r\b[^>]*>(?:(?!<w:r\b).)*?<w:fldChar w:fldCharType="begin"\s*/>'
    r".*?"
    r'<w:fldChar w:fldCharType="end"\s*/>\s*</w:r>',
    re.DOTALL,
)


def _unzip(src_docx: str, work: str) -> None:
    shutil.rmtree(work, ignore_errors=True)
    os.makedirs(work)
    with zipfile.ZipFile(src_docx) as z:
        z.extractall(work)


def _rezip(work: str, dst_docx: str) -> None:
    if os.path.exists(dst_docx):
        os.remove(dst_docx)
    with zipfile.ZipFile(dst_docx, "w", zipfile.ZIP_DEFLATED) as zout:
        for root, _, files in os.walk(work):
            for name in files:
                p = os.path.join(root, name)
                zout.write(p, os.path.relpath(p, work))


def shield(src_docx: str, out_docx: str, sidecar_json: str) -> dict:
    work = out_docx + ".work"
    _unzip(src_docx, work)
    xml_path = os.path.join(work, "word", "document.xml")
    xml = open(xml_path, encoding="utf-8").read()

    saved: dict[str, str] = {}
    counter = [0]

    def sub(match: re.Match) -> str:
        chunk = match.group(0)
        if "zotero.org/google-docs/" not in chunk:
            return chunk
        counter[0] += 1
        key = f"{counter[0]:04d}"
        saved[key] = chunk
        return (
            f'<w:r><w:t xml:space="preserve">{placeholder(key)}</w:t></w:r>'
        )

    new_xml, _ = FIELD_RE.subn(sub, xml)
    open(xml_path, "w", encoding="utf-8").write(new_xml)
    _rezip(work, out_docx)
    shutil.rmtree(work, ignore_errors=True)

    with open(sidecar_json, "w", encoding="utf-8") as f:
        json.dump(
            {"source_docx": os.path.abspath(src_docx), "fields": saved},
            f,
            ensure_ascii=False,
        )
    return {"shielded": len(saved), "sidecar": sidecar_json, "out": out_docx}


def unshield(src_docx: str, sidecar_json: str, out_docx: str) -> dict:
    saved = json.load(open(sidecar_json, encoding="utf-8"))["fields"]
    work = out_docx + ".work"
    _unzip(src_docx, work)
    xml_path = os.path.join(work, "word", "document.xml")
    xml = open(xml_path, encoding="utf-8").read()

    restored = 0
    warnings: list[str] = []

    for key, chunk in saved.items():
        token = placeholder(key)
        if token not in xml:
            warnings.append(f"missing token {key}")
            continue
        run_re = re.compile(
            r"<w:r\b[^>]*>(?:(?!<w:r\b).)*?"
            + re.escape(token)
            + r".*?</w:r>",
            re.DOTALL,
        )
        m = run_re.search(xml)
        if not m:
            warnings.append(f"no enclosing run for {key}")
            continue
        run = m.group(0)
        rpr_match = re.search(r"<w:rPr>.*?</w:rPr>", run, re.DOTALL)
        rpr = rpr_match.group(0) if rpr_match else ""
        t_match = None
        for tm in re.finditer(r"<w:t(\s[^>]*)?>([^<]*)</w:t>", run):
            if token in tm.group(2):
                t_match = tm
                break
        if not t_match:
            warnings.append(f"token split across <w:t> for {key}")
            continue
        before, after = t_match.group(2).split(token, 1)
        repl = ""
        if before:
            repl += (
                f"<w:r>{rpr}<w:t xml:space=\"preserve\">"
                f"{xml_escape(before)}</w:t></w:r>"
            )
        repl += chunk
        if after:
            repl += (
                f"<w:r>{rpr}<w:t xml:space=\"preserve\">"
                f"{xml_escape(after)}</w:t></w:r>"
            )
        xml = xml[: m.start()] + repl + xml[m.end():]
        restored += 1

    leftover = PLACEHOLDER_RE.findall(xml)
    open(xml_path, "w", encoding="utf-8").write(xml)
    _rezip(work, out_docx)
    shutil.rmtree(work, ignore_errors=True)

    return {
        "restored": restored,
        "expected": len(saved),
        "leftover": len(leftover),
        "warnings": warnings,
        "out": out_docx,
    }


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    sub = ap.add_subparsers(dest="cmd", required=True)

    s = sub.add_parser("shield")
    s.add_argument("src_docx")
    s.add_argument("out_docx")
    s.add_argument("sidecar_json")

    u = sub.add_parser("unshield")
    u.add_argument("src_docx")
    u.add_argument("sidecar_json")
    u.add_argument("out_docx")

    args = ap.parse_args()
    if args.cmd == "shield":
        result = shield(args.src_docx, args.out_docx, args.sidecar_json)
    else:
        result = unshield(args.src_docx, args.sidecar_json, args.out_docx)
    print(json.dumps(result, ensure_ascii=False))
    return 0 if not result.get("warnings") else 0


if __name__ == "__main__":
    sys.exit(main())
