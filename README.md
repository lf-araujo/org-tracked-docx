# org-tracked-docx

Round-trip a Word `.docx` through Emacs org-mode **preserving tracked changes
and comments**, using `pandoc --track-changes=all` and
[CriticMarkup](http://criticmarkup.com/spec.php).

Author manuscripts in org, export to `.docx`, send to co-authors who edit with
Track Changes on, then re-import the returned file: insertions, deletions,
substitutions, and comments come back as font-locked CriticMarkup tokens
(`{++ins++}`, `{--del--}`, `{~~old~>new~~}`, `{>>comment<<}`) that you can
accept/reject and merge against the canonical org source.

## Files

| File | Role |
|------|------|
| `org-tracked-docx.el` | Main package: `otd-import`, `otd-export`, accept/reject/merge, `otd-criticmarkup-mode`. |
| `org-tracked-docx-zotero.el` | Optional Zotero layer: shields Zotero (Google-Docs) citation fields across the pandoc round-trip. |
| `otd_zotero_shield.py` | Helper invoked by the Zotero layer to shield/unshield Zotero HYPERLINK fields in the `.docx`. |
| `otd-zotero-unshield.lua` | Pandoc Lua filter that restores shielded Zotero fields on export. |

## Requirements

- Emacs 27.1+
- [pandoc](https://pandoc.org/)
- `unzip` / `zip` on `PATH`
- Python 3 + a Lua-enabled pandoc (for the optional Zotero layer)

## Usage

```elisp
(load "/path/to/org-tracked-docx/org-tracked-docx.el")
(require 'org-tracked-docx)        ; if added to load-path
```

- `M-x otd-import RET reviewed.docx RET` — import a tracked docx as CriticMarkup org.
- `M-x otd-export RET manuscript.org RET` — export org back to docx (embeds the
  canonical org source as a customXml part for lossless re-merge).
- `M-x otd-accept-all` / `M-x otd-reject-all` — flatten tracked markup.
- `M-x otd-diff-against` — word-diff the imported file against the canonical org.

## Notes

`otd--preserve-leading-spaces` rewrites leading spaces in
`xml:space="preserve"` runs to NBSP so Word table indentation survives pandoc.
It anchors on the literal attribute (not a greedy regexp) because
`word/document.xml` is a single line and can contain very long bracket-free
spans (e.g. base64 Mendeley citation blobs in `<w:tag w:val="...">`), which
overflow Emacs' regexp matcher stack under any unbounded greedy quantifier.
