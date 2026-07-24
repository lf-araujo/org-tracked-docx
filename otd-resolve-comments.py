#!/usr/bin/env python3
"""Mark specific comments as "Resolved" in a docx's native reviewing pane.

Word's own resolved-state mechanism is word/commentsExtended.xml: a
<w15:commentEx w15:paraId="XXXXXXXX" w15:done="1"/> entry, keyed by the
w14:paraId attribute on the <w:p> inside that comment in word/comments.xml.
Pandoc's docx writer does not emit w14:paraId at all, so this script:

  1. Injects a synthetic w14:paraId on every comment paragraph in
     word/comments.xml (declaring the w14 namespace on the root element
     if it isn't already there).
  2. Writes word/commentsExtended.xml with a w15:done="1" entry for each
     comment id in DONE_IDS.
  3. Registers the new part in [Content_Types].xml and
     word/_rels/document.xml.rels, since an unregistered part is exactly
     the kind of thing Word flags as "unreadable content" and offers to
     repair/discard on open.

Used as a post-processing stage in otd-export (org-tracked-docx.el),
mirroring otd-table-borders.py's direct-XML-patch approach.

Usage: otd-resolve-comments.py DOCX_PATH DONE_IDS
  DONE_IDS is a comma-separated list of comment id="N" values (may be
  empty, in which case comments.xml still gets paraIds but no comment is
  marked done).
"""
import sys
import re
import zipfile
import shutil
import tempfile
import os


def inject_para_ids(comments_xml):
    """Return (new_xml, {comment_id: paraId}) with a w14:paraId added to
    each comment's first <w:p>, and the w14 namespace declared on the
    root <w:comments> element if missing."""
    if 'xmlns:w14=' not in comments_xml.split('>', 1)[0]:
        comments_xml = comments_xml.replace(
            '<w:comments ',
            '<w:comments xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ',
            1,
        )

    para_ids = {}
    counter = [0x10000001]

    def add_id(m):
        cid = m.group(1)
        pid = "%08X" % counter[0]
        counter[0] += 1
        para_ids[cid] = pid
        comment_body = m.group(0)
        # Add w14:paraId to the comment's first <w:p ...> (whether or not
        # it already has other attributes).
        def stamp_first_p(pm):
            tag = pm.group(0)
            if 'w14:paraId' in tag:
                return tag
            if tag.endswith('/>'):
                return tag[:-2] + ' w14:paraId="%s"/>' % pid
            return tag[:-1] + ' w14:paraId="%s">' % pid
        return re.sub(r'<w:p\b[^>]*>', stamp_first_p, comment_body, count=1)

    new_xml = re.sub(
        r'<w:comment\s+[^>]*w:id="(\d+)"[^>]*>.*?</w:comment>',
        add_id, comments_xml, flags=re.DOTALL,
    )
    return new_xml, para_ids


def build_comments_extended(para_ids, done_ids):
    entries = []
    for cid in done_ids:
        pid = para_ids.get(cid)
        if pid is None:
            continue
        entries.append('<w15:commentEx w15:paraId="%s" w15:done="1"/>' % pid)
    body = "".join(entries)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">'
        + body
        + '</w15:commentsEx>'
    )


def patch_content_types(xml):
    if 'commentsExtended' in xml:
        return xml
    override = (
        '<Override PartName="/word/commentsExtended.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.'
        'wordprocessingml.commentsExtended+xml"/>'
    )
    return xml.replace('</Types>', override + '</Types>')


def patch_document_rels(xml):
    if 'commentsExtended.xml' in xml:
        return xml
    ids = [int(m) for m in re.findall(r'Id="rId(\d+)"', xml)]
    next_id = max(ids, default=0) + 1
    rel = (
        '<Relationship Id="rId%d" '
        'Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" '
        'Target="commentsExtended.xml"/>' % next_id
    )
    return xml.replace('</Relationships>', rel + '</Relationships>')


def patch_docx(path, done_ids):
    tmp_fd, tmp_path = tempfile.mkstemp(suffix=".docx")
    os.close(tmp_fd)
    with zipfile.ZipFile(path) as zin:
        names = set(zin.namelist())
        if "word/comments.xml" not in names:
            # Nothing to resolve; leave the file untouched.
            shutil.copy(path, tmp_path)
            shutil.move(tmp_path, path)
            return
        comments_xml = zin.read("word/comments.xml").decode("utf-8")
        new_comments_xml, para_ids = inject_para_ids(comments_xml)
        comments_ext_xml = build_comments_extended(para_ids, done_ids)
        content_types = zin.read("[Content_Types].xml").decode("utf-8")
        new_content_types = patch_content_types(content_types)
        doc_rels = zin.read("word/_rels/document.xml.rels").decode("utf-8")
        new_doc_rels = patch_document_rels(doc_rels)

        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            wrote_ext = False
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/comments.xml":
                    data = new_comments_xml.encode("utf-8")
                elif item.filename == "[Content_Types].xml":
                    data = new_content_types.encode("utf-8")
                elif item.filename == "word/_rels/document.xml.rels":
                    data = new_doc_rels.encode("utf-8")
                zout.writestr(item, data)
            if "word/commentsExtended.xml" not in names:
                zout.writestr("word/commentsExtended.xml",
                               comments_ext_xml.encode("utf-8"))
                wrote_ext = True
    shutil.move(tmp_path, path)
    return wrote_ext


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("usage: otd-resolve-comments.py DOCX_PATH DONE_IDS", file=sys.stderr)
        sys.exit(1)
    docx_path = sys.argv[1]
    done_ids_arg = [s for s in sys.argv[2].split(",") if s]
    patch_docx(docx_path, done_ids_arg)
