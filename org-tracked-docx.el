;;; org-tracked-docx.el --- Round-trip Word docx with track changes  -*- lexical-binding: t; -*-

;; Author: drafted with Claude
;; Keywords: org, pandoc, docx, track changes, CriticMarkup
;; URL: project-local
;; Package-Requires: ((emacs "27.1"))

;;; Commentary:
;;
;; A tiny wrapper around `pandoc --track-changes=all' that imports a Word
;; .docx into org-mode preserving tracked changes and comments as
;; CriticMarkup.  The mode highlights insertions/deletions/substitutions and
;; provides accept-/reject-all helpers.
;;
;; Workflow:
;;   1. Author the manuscript in org (e.g. 2026-LGCM2.org).
;;   2. Export to docx with `M-x org-pandoc-export-to-docx'.
;;   3. Send the docx to a co-author.  They turn on Track Changes (Review
;;      -> Track Changes in Word; Edit -> Track Changes -> Record in
;;      LibreOffice), edit, and return the docx.
;;   4. `M-x otd-import' on the returned docx.  The output buffer holds
;;      the manuscript with `{++inserted++}', `{--deleted--}',
;;      `{~~old~>new~~}', and `{>>comment<<}' markers, font-lock-coloured.
;;   5. To merge: open the canonical .org and `M-x otd-diff-against' to
;;      word-diff against the imported tracked file; or `M-x
;;      otd-accept-all' / `otd-reject-all' to flatten the tracked file
;;      and then diff or copy paragraphs across.
;;
;; CriticMarkup spec: <http://criticmarkup.com/spec.php>
;; Pandoc track-changes docs:
;;   <https://pandoc.org/MANUAL.html#option--track-changes>

;;; Code:

(require 'subr-x)
(require 'cl-lib)
(require 'ansi-color)

(defgroup org-tracked-docx nil
  "Import Word docx with tracked changes as CriticMarkup-decorated org."
  :group 'org)

(defcustom otd-pandoc-program "pandoc"
  "Path to the pandoc executable."
  :type 'string :group 'org-tracked-docx)

(defcustom otd-extract-media-suffix "media"
  "Suffix appended to docx basename for the extracted-media dir."
  :type 'string :group 'org-tracked-docx)

(defcustom otd-output-suffix "-tracked"
  "Suffix appended to docx basename for the imported .org file."
  :type 'string :group 'org-tracked-docx)

(defcustom otd-export-author user-full-name
  "Author name written into pandoc track-change spans on `otd-export'.
Surfaces in Word as the reviewer who made each tracked edit/comment."
  :type 'string :group 'org-tracked-docx)

(defcustom otd-embed-source t
  "When non-nil, `otd-export' embeds the source .org file inside the
output .docx as a customXml part.  Word and LibreOffice preserve
customXml across saves and tracked edits, so when the reviewed docx
returns it still carries the canonical org text (with cite keys,
metadata headers, etc.) which the docx body cannot represent.
`otd-import' detects the part automatically and uses it for merge."
  :type 'boolean :group 'org-tracked-docx)

(defcustom otd-embed-custom-property nil
  "When non-nil, `otd-export' ALSO embeds the org source as a single
docProps/custom.xml custom property, as a fallback for WPS Office
\(which strips customXml/ parts on save but keeps custom properties).

Off by default: a single custom property holding an entire manuscript's
org source as one base64 string is well outside what docProps/custom.xml
properties are meant to hold, and is a confirmed cause of Word's \"found
unreadable content\" repair prompt.  Only enable this if you specifically
need a reviewed docx to survive a round-trip through WPS Office, and
accept the risk of Word flagging the file for repair on open."
  :type 'boolean :group 'org-tracked-docx)

(defcustom otd-import-auto-merge t
  "When non-nil and the imported docx carries an embedded org source
\(see `otd-embed-source'), `otd-import' runs `otd--merge-content' to
produce a merged file: canonical structure (cite keys, metadata)
plus CriticMarkup tokens placed at matching paragraphs.

When nil or no embedded source is present, the lossy pandoc roundtrip
is used as-is."
  :type 'boolean :group 'org-tracked-docx)

(defcustom otd-import-backend 'json
  "Import pipeline used by `otd-import'.

`json' (default) is `otd-import-json': a structured walk over pandoc's
JSON AST that replaces tracked-change spans with CriticMarkup inside the
tree, then renders json->org.  It has no markdown intermediate and no
textual span-matching, so it preserves comments containing brackets,
URLs and citations that the `markdown' backend can silently drop, and it
emits no pandoc escaping artifacts.

`markdown' is the original `otd--import-markdown': docx->markdown, regex
span rewriting, markdown->org.  Kept as a fallback.

The JSON backend lives in the optional `org-tracked-docx-json' feature;
if it is not loaded, `otd-import' uses the markdown backend regardless
of this setting."
  :type '(choice (const :tag "JSON AST (robust, default)" json)
                 (const :tag "Markdown (legacy)" markdown))
  :group 'org-tracked-docx)

(defcustom otd-fix-table-borders t
  "When non-nil, `otd-export' rewrites every table's borders directly
\(top + bottom rule, plus a rule under the header row; no vertical or
inter-row lines) after pandoc generates the docx, via the
`otd-table-borders-script' Python script, instead of relying solely on
the reference-doc's \"Table\" table-style and its `firstRow' conditional
formatting.

Confirmed (by directly viewing the exported docx) to render more
reliably than depending on the reference-doc's table-style borders
alone, so this defaults on."
  :type 'boolean :group 'org-tracked-docx)

(defcustom otd-table-borders-script
  (expand-file-name "otd-table-borders.py"
                     (file-name-directory (or load-file-name buffer-file-name)))
  "Path to the Python script `otd-export' runs (see `otd-fix-table-borders')
to rewrite table borders directly in the generated docx."
  :type 'string :group 'org-tracked-docx)

(defcustom otd-python-program "python3"
  "Python interpreter used to run `otd-table-borders-script'."
  :type 'string :group 'org-tracked-docx)

(defcustom otd-resolve-done-comments t
  "When non-nil, `otd-export' marks every `[DONE]'-tagged CriticMarkup
comment as actually Resolved in Word's reviewing pane, via
`otd-resolve-comments-script', instead of leaving the literal `[DONE] '
text as part of the comment body."
  :type 'boolean :group 'org-tracked-docx)

(defcustom otd-resolve-comments-script
  (expand-file-name "otd-resolve-comments.py"
                     (file-name-directory (or load-file-name buffer-file-name)))
  "Path to the Python script `otd-export' runs (see
`otd-resolve-done-comments') to mark `[DONE]' comments Resolved."
  :type 'string :group 'org-tracked-docx)

(defcustom otd-tangle-before-export t
  "When non-nil, `otd-export' tangles embedded resource blocks first.

A manuscript can embed its own bibliography, CSL style, and Word
reference-doc as source blocks (see the \"Embedded resources\" section
convention) instead of pointing `#+bibliography:'/`#+PANDOC_OPTIONS:'
at external paths.  This makes the .org file fully self-contained and
portable across machines: whatever is embedded in the document is
always what gets used, with no dependency on any other file existing
at some particular path on whatever machine runs the export.

Plain-text resources (bibliography, CSL) use ordinary Org Babel
`:tangle PATH' header args and are written via `org-babel-tangle-file'.
Binary resources (e.g. a .docx reference template) cannot round-trip
through a text buffer as raw bytes, so they use a custom `:tangle-binary
PATH' header instead: the block body is expected to be base64 text,
which `otd--tangle-binary-blocks' decodes and writes with no-conversion
coding.  See `otd--tangle-embedded-resources'."
  :type 'boolean :group 'org-tracked-docx)

;;;; --- tangle embedded resources -----------------------------------------

(defun otd--tangle-binary-blocks (org-file dir)
  "Decode any `:tangle-binary PATH' source blocks in ORG-FILE to disk.

Each such block's body must be base64 text (whitespace ignored); it is
decoded and written as raw bytes to PATH, resolved relative to DIR.
This is the binary-file counterpart to Org Babel's own `:tangle', which
only ever writes its blocks out as text."
  (with-temp-buffer
    (insert-file-contents org-file)
    (goto-char (point-min))
    (while (re-search-forward
            "^[ \t]*#\\+begin_src[ \t]+\\S-+.*:tangle-binary[ \t]+\\(\\S-+\\(?:[ \t]\\S-+\\)*?\\)\\(?:[ \t]:\\|[ \t]*$\\)"
            nil t)
      (let* ((target (expand-file-name (match-string 1) dir))
             (body-start (progn (forward-line 1) (point)))
             (body-end (progn
                         (re-search-forward "^[ \t]*#\\+end_src" nil t)
                         (match-beginning 0)))
             (b64 (buffer-substring-no-properties body-start body-end))
             (bytes (base64-decode-string (replace-regexp-in-string "[ \t\n\r]" "" b64))))
        (make-directory (file-name-directory target) t)
        (let ((coding-system-for-write 'no-conversion))
          (write-region bytes nil target nil 'silent))
        (message "otd: tangled binary resource -> %s" target)))))

(defun otd--tangle-embedded-resources (org-file dir)
  "Regenerate any embedded-resource files ORG-FILE tangles to.

Runs ordinary `org-babel-tangle-file' for plain-text `:tangle' blocks
(bibliography, CSL, ...), then `otd--tangle-binary-blocks' for any
`:tangle-binary' blocks (e.g. a Word reference-doc), so pandoc always
sees fresh files regenerated from whatever is embedded in ORG-FILE --
never a stale copy left over from a previous export, and never a
missing file because some external path doesn't exist on this
machine."
  (require 'ob-tangle)
  (let ((default-directory dir))
    (org-babel-tangle-file org-file))
  (otd--tangle-binary-blocks org-file dir))

;;;; --- import ----------------------------------------------------------

;;;###autoload
(defun otd-import (docx &optional output)
  "Import DOCX to CriticMarkup org, dispatching on `otd-import-backend'.

The default `json' backend (`otd-import-json') is more robust than the
original `markdown' backend (`otd--import-markdown'): it preserves
comments containing brackets, URLs and citations that the markdown path
can silently drop, and emits no pandoc escaping artifacts.  When the
JSON module is not loaded, falls back to the markdown backend.

With prefix arg, prompt for OUTPUT path; otherwise the output is placed
alongside DOCX with `otd-output-suffix' appended."
  (interactive
   (list (read-file-name "Docx file: " nil nil t nil
                         (lambda (f) (or (file-directory-p f)
                                         (string-match-p "\\.docx\\'" f))))
         (when current-prefix-arg (read-file-name "Output .org: "))))
  (if (and (eq otd-import-backend 'json) (fboundp 'otd-import-json))
      (otd-import-json docx output)
    (otd--import-markdown docx output)))

(defun otd--import-markdown (docx &optional output)
  "Convert DOCX to org-mode preserving tracked changes AND comments as CriticMarkup.
With prefix arg, prompt for OUTPUT path; otherwise the output is placed
alongside DOCX with `otd-output-suffix' appended.  After conversion the
output buffer is opened with `otd-criticmarkup-mode' enabled.

Pipeline:
  pandoc docx -> markdown (--track-changes=all preserves ins/del/comment as
                           pandoc-attribute spans),
  rewrite spans to CriticMarkup tokens,
  pandoc markdown -> org (CriticMarkup tokens pass through as plain text)."
  (interactive
   (list (read-file-name "Docx file: " nil nil t nil
                         (lambda (f) (or (file-directory-p f)
                                         (string-match-p "\\.docx\\'" f))))
         (when current-prefix-arg (read-file-name "Output .org: "))))
  (unless (file-exists-p docx) (user-error "File not found: %s" docx))
  (let* ((dir        (file-name-directory (expand-file-name docx)))
         (base       (file-name-sans-extension (file-name-nondirectory docx)))
         (out        (or output
                         (expand-file-name (concat base otd-output-suffix ".org") dir)))
         (media      (expand-file-name (concat base "-" otd-extract-media-suffix) dir))
         (md-tmp     (make-temp-file "otd-" nil ".md"))
         (docx-fixed (otd--preserve-leading-spaces (expand-file-name docx)))
         (grid-alist nil))
    (unwind-protect
        (progn
          ;; Stage 1: docx -> markdown with track-change/comment spans.
          ;; (`docx-fixed' has leading-space cell content rewritten to NBSP
          ;; so the indentation survives pandoc's docx AST parser.)
          (with-temp-buffer
            (let ((exit (apply #'call-process otd-pandoc-program nil t nil
                               (list "-f" "docx" "-t" "markdown"
                                     "--wrap=none"
                                     "--track-changes=all"
                                     (format "--extract-media=%s" media)
                                     docx-fixed
                                     "-o" md-tmp))))
              (unless (zerop exit)
                (error "pandoc docx->md exit %s:\n%s" exit (buffer-string)))))
          ;; Stage 2: rewrite pandoc spans to CriticMarkup.  Pull the
          ;; resolved (`done') comment ids out of the docx first so
          ;; rewrite-spans can tag those with ` [DONE] ' after the
          ;; author prefix.  Pandoc strips this signal, so we read
          ;; the source xml directly.
          (let* ((done-set (otd--extract-done-comment-ids
                            (expand-file-name docx)))
                 (counts (otd--rewrite-spans md-tmp done-set)))
            ;; Stage 2b: shield grid_tables (multi-line cells) with sentinels
            ;; so pandoc's md->org cannot collapse them into pipe-table rows.
            (setq grid-alist (otd--shield-grid-tables md-tmp))
            ;; Stage 2c: stash CriticMarkup tokens as opaque alphanumeric
            ;; sentinels so pandoc's md->org pass cannot mangle them.
            ;; Without this `{~~old~>new~~}` is parsed as strikethrough +
            ;; subscript and gets corrupted to `{+old~>new+}` / `{+old_{>new}~}`
            ;; (we saw 23/23 substitutions destroyed before this stash); the
            ;; same risk applies to any `{++...++}` containing a `^...^`
            ;; superscript run, hence the bracket-balanced sentinel rather
            ;; than a regex disable of the offending extensions, which would
            ;; also clobber legitimate inline super/subscript content.
            (let ((cm-alist (otd--stash-criticmarkup md-tmp)))
              ;; Stage 3: markdown (with sentinels) -> org.
              ;; Disable `smart' (markdown-smart): without this pandoc
              ;; rewrites straight `"' to curly Unicode `LEFT/RIGHT
              ;; DOUBLE QUOTATION MARK' inside surviving attribute
              ;; spans, leaving artifacts like `]{.comment-end id=U+201C
              ;; 128 U+201D}' that the rewrite-spans pass can no longer
              ;; recognize on a re-import.
              (with-temp-buffer
                (let ((exit (apply #'call-process otd-pandoc-program nil t nil
                                   (list "-f" "markdown-smart" "-t" "org"
                                         "--wrap=none" md-tmp "-o" out))))
                  (unless (zerop exit)
                    (error "pandoc md->org exit %s:\n%s" exit (buffer-string)))))
              ;; Stage 3a: restore CriticMarkup tokens in the org output.
              (otd--unstash-criticmarkup out cm-alist))
            ;; Stage 3b: restore shielded grid_tables as raw markdown blocks
            ;; (`#+BEGIN_EXPORT markdown' survives the org->md->docx round-trip
            ;; in `otd-export', so the multi-line cell structure is preserved).
            (otd--restore-grid-tables out grid-alist)
            ;; Stage 4: sanitize the org file so the next round-trip back
            ;; to docx renders super/subscripts (see `otd--postfix-org').
            (otd--postfix-org out)
            ;; Stage 5: if the docx carries an embedded org source (sent
            ;; out by an earlier `otd-export'), extract it and merge so
            ;; the resulting tracked file keeps cite keys / metadata.
            (let* ((embedded (otd--extract-org-source (expand-file-name docx)))
                   merge-info)
              (when embedded
                (let ((src (expand-file-name (concat base "-source.org") dir)))
                  (with-temp-file src (insert embedded))
                  (when otd-import-auto-merge
                    (let* ((tracked (with-temp-buffer
                                      (insert-file-contents out)
                                      (buffer-string)))
                           (result (otd--merge-content embedded tracked)))
                      (setq merge-info (cdr result))
                      (with-temp-file out (insert (car result)))))))
              (find-file out)
              (otd-criticmarkup-mode 1)
              (message
               (concat "Imported %s -> %s  (++%d --%d ~~%d ==%d >>%d ✓%d  grid:%d)"
                       (when embedded
                         (format "  source-found%s"
                                 (if merge-info
                                     (format ":merged %d/%d para, %d cite-keys%s"
                                             (plist-get merge-info :merged-paragraphs)
                                             (plist-get merge-info :cm-paragraphs)
                                             (plist-get merge-info :cite-keys)
                                             (let ((n (plist-get merge-info
                                                                 :reanchored-comments)))
                                               (if (and n (> n 0))
                                                   (format ", %d cmt-grafted" n) "")))
                                   ""))))
               (file-name-nondirectory docx)
               (file-name-nondirectory out)
               (plist-get counts :ins) (plist-get counts :del)
               (plist-get counts :sub) (plist-get counts :hi)
               (plist-get counts :cmt) (or (plist-get counts :done) 0)
               (length grid-alist)))))
      (when (file-exists-p md-tmp)     (delete-file md-tmp))
      (when (file-exists-p docx-fixed) (delete-file docx-fixed)))
    out))

(defun otd--extract-done-comment-ids (docx-path)
  "Return a hash table (string -> t) of comment ids resolved (`done')
in DOCX-PATH's Word reviewing pane.

Word stores the per-paragraph resolved state in `word/commentsExtended.xml'
as `<w15:commentEx w15:paraId=... w15:done=\"1\"/>'.  Each `paraId' refers
to a `w14:paraId' attribute on a `<w:p>' inside `word/comments.xml',
where a single comment may carry several paragraphs (top-level + replies).
This walker treats a comment as resolved when ANY of its paragraphs is
marked done, matching Word's UI behavior where clicking `Resolve' on
the conversation flips every reply to done together.

Returns an empty hash table if either XML part is missing (Word versions
< 2013 do not write `commentsExtended.xml')."
  (let ((done (make-hash-table :test 'equal))
        (tmpdir (make-temp-file "otd-done-" t)))
    (unwind-protect
        (when (zerop (call-process
                      "unzip" nil nil nil "-qo" docx-path
                      "word/comments.xml" "word/commentsExtended.xml"
                      "-d" tmpdir))
          (let ((paraid-done (make-hash-table :test 'equal))
                (ext (expand-file-name "word/commentsExtended.xml" tmpdir))
                (com (expand-file-name "word/comments.xml" tmpdir)))
            ;; paraId -> done flag
            (when (file-exists-p ext)
              (with-temp-buffer
                (insert-file-contents ext)
                (goto-char (point-min))
                (while (re-search-forward
                        "<w15:commentEx\\(?:[[:space:]][^>]*\\)?[[:space:]]w15:paraId=\"\\([0-9A-Fa-f]+\\)\"\\(?:[[:space:]][^>]*\\)?[[:space:]]w15:done=\"\\([01]\\)\""
                        nil t)
                  (when (string= (match-string 2) "1")
                    (puthash (upcase (match-string 1)) t paraid-done)))))
            ;; Walk comments.xml: any <w:p w14:paraId="X"> inside a
            ;; <w:comment w:id="N"> flagged done => N is done.
            (when (and (file-exists-p com) (> (hash-table-count paraid-done) 0))
              (with-temp-buffer
                (insert-file-contents com)
                (goto-char (point-min))
                (let (cur-id)
                  (while (re-search-forward
                          "<w:comment[[:space:]][^>]*w:id=\"\\([0-9]+\\)\"\\|<w:p[[:space:]][^>]*w14:paraId=\"\\([0-9A-Fa-f]+\\)\""
                          nil t)
                    (cond
                     ((match-string 1) (setq cur-id (match-string 1)))
                     ((and cur-id
                           (gethash (upcase (match-string 2)) paraid-done))
                      (puthash cur-id t done)))))))))
      (delete-directory tmpdir t))
    done))

(defun otd--rewrite-spans (md-file &optional done-set)
  "Rewrite pandoc track-change/comment spans in MD-FILE to CriticMarkup.
When DONE-SET (a hash table mapping comment-id strings to t) lists a
comment as resolved in Word, prefix the emitted comment text with
`[DONE] ' after the author tag, so the round-trip is:

  Word `done=1' (.docx)  ->  {>>[Author] [DONE] text<<} (.org)

Return a plist with counts of each marker emitted."
  (let ((n-ins 0) (n-del 0) (n-sub 0) (n-hi 0) (n-cmt 0) (n-done 0))
    (with-temp-buffer
      (insert-file-contents md-file)
      ;; Pre-pass 0: downgrade `# Lead paragraph...' lines that pandoc's
      ;; docx reader misclassified as ATX headings.  When the docx's
      ;; first body paragraph (e.g. the Abstract opener) carries a
      ;; heading-like style, pandoc emits it as `# <300-char paragraph>'
      ;; which would otherwise become a section heading on md->org and
      ;; eat the abstract body entirely.  Heuristic: strip CriticMarkup
      ;; attribute spans (`[text]{.attr=...}', `[]{.attr...}'), then if
      ;; what remains exceeds 80 chars of running prose treat it as a
      ;; misclassified paragraph and remove the leading `# '.  This keeps
      ;; genuinely short headings (`# Abstract', `# Results') intact
      ;; even when they carry inline reviewer comments.
      (goto-char (point-min))
      (while (re-search-forward
              "^#[[:space:]]+\\(.+\\)$" nil t)
        (let* ((line (match-string 1))
               ;; Strip attribute span tails like `]{.comment-end id="N"}',
               ;; `]{.insertion author="..." date="..."}', `]{.mark}'.
               (stripped
                (replace-regexp-in-string
                 "\\][[:space:]]*{\\.[A-Za-z0-9_-]+[^}]*}" "" line))
               ;; Strip the leading `[' of any opener like `[CTEXT]{...}'.
               (stripped (replace-regexp-in-string "\\[" "" stripped)))
          (when (> (length (string-trim stripped)) 80)
            (replace-match line t t))))
      ;; Pre-pass 1: flatten pandoc's nested closing spans for overlapping
      ;; reviewer comments.  For two comments on the same range pandoc emits
      ;; `[[]{.comment-end id="INNER"}]{.comment-end id="OUTER"}'; for three
      ;; (which does happen on heavily-reviewed paragraphs) it emits
      ;; `[[[]{end:130}]{end:129}]{end:128}'.  The pattern below accepts one
      ;; OR MORE inner empty `[]{.comment-end id="N"}' spans between the
      ;; outer brackets, so each iteration peels one wrapping layer
      ;; regardless of depth; the loop runs until no nested form remains.
      (goto-char (point-min))
      (let ((flat t))
        (while flat
          (setq flat nil)
          (goto-char (point-min))
          (while (re-search-forward
                  "\\[\\(\\(?:\\[\\]{\\.comment-end[[:space:]]+id=\"[0-9]+\"}\\)+\\)\\]\\({\\.comment-end[[:space:]]+id=\"[0-9]+\"}\\)"
                  nil t)
            (replace-match "\\1[]\\2" t)
            (setq flat t))))
      ;; Comments: [CTEXT]{.comment-start id="N"...}RANGE[]{.comment-end id="N"}
      ;; -> {==RANGE==}{>>CTEXT<<}.  Rewrite innermost-first so an outer
      ;; comment's lazy range cannot swallow an inner comment's spans (the
      ;; common case where reviewers leave overlapping comments on the same
      ;; paragraph).  Each iteration finds the comment-start whose matching
      ;; comment-end is closest, replaces that pair, and re-scans.
      ;;
      ;; CTEXT itself can contain `]' (a reviewer's comment quoting a
      ;; citation or markdown link, e.g. `[text](url)'), so this cannot
      ;; search forward for the opening `[' with a `[^]]*' class the way
      ;; the old version did -- that stops at CTEXT's own first `]' and
      ;; the whole match then fails to find `{.comment-start' next,
      ;; silently dropping the comment: its start marker vanishes, `]'
      ;; and everything after briefly gets swallowed into pandoc's own
      ;; markdown-link/attribute-span parsing on the next pandoc pass, and
      ;; the comment body itself leaks into the docx as literal prose.
      ;; Instead, search forward for the unambiguous fixed closing tag
      ;; `]{.comment-start id="N"...}' first (never itself contains `]'),
      ;; then scan backward from that `]' counting bracket depth (same
      ;; technique as the insertion/deletion scanner below) to find its
      ;; true matching `[', however many nested `[...]' constructs CTEXT
      ;; itself contains.
      (let ((more t))
        (while more
          (setq more nil)
          (let ((best-dist most-positive-fixnum)
                (best nil))
            (save-excursion
              (goto-char (point-min))
              (while (re-search-forward
                      "\\]{\\.comment-start[[:space:]]+id=\"\\([0-9]+\\)\"\\([^}]*\\)}"
                      nil t)
                (let* ((id    (match-string 1))
                       (attrs (match-string 2))
                       (rbpos (match-beginning 0)) ; position of the closing `]'
                       ;; Capture positions BEFORE string-match clobbers the
                       ;; outer re-search-forward's match data.
                       (after (match-end 0))
                       (open  (let ((depth 1) (p (1- rbpos)))
                                (while (and (> p (point-min)) (> depth 0))
                                  (let ((c (char-after p)))
                                    (cond ((eq c ?\]) (cl-incf depth))
                                          ((eq c ?\[) (cl-decf depth)
                                           (when (zerop depth) (setq open p)))))
                                  (setq p (1- p)))
                                open))
                       (start (or open rbpos))
                       (ctext (if open
                                  (buffer-substring-no-properties (1+ open) rbpos)
                                ""))
                       (author (and (string-match
                                     "author=\"\\([^\"]*\\)\"" attrs)
                                    (match-string 1 attrs))))
                  (save-excursion
                    (when (re-search-forward
                           (concat "\\[\\]{\\.comment-end[[:space:]]+id=\""
                                   id "\"}")
                           nil t)
                      (let ((dist (- (match-beginning 0) after)))
                        (when (< dist best-dist)
                          (setq best-dist dist
                                best (list start (match-end 0)
                                           ctext author after
                                           (match-beginning 0)
                                           id))))))))) ; carry id into outer scope
            (when best
              (setq more t)
              (let* ((start-pos   (nth 0 best))
                     (end-pos     (nth 1 best))
                     (ctext       (nth 2 best))
                     (author      (nth 3 best))
                     (range-start (nth 4 best))
                     (range-end   (nth 5 best))
                     (id          (nth 6 best))
                     (range-text  (buffer-substring-no-properties
                                   range-start range-end))
                     ;; Encode docx author into the CriticMarkup comment text as
                     ;; `[Author Name] …' prefix, so `otd-export' can round-trip
                     ;; authorship back into Word's `<w:comment w:author="…">'.
                     ;; When the docx flagged this comment as resolved in
                     ;; Word's reviewing pane, inject ` [DONE]' immediately
                     ;; after the author tag.  Reading: `[Author] [DONE] body'.
                     (resolved (and done-set (gethash id done-set)))
                     (ctext-tagged
                      (cond
                       ((and author (not (string-empty-p author)) resolved)
                        (cl-incf n-done)
                        (format "[%s] [DONE] %s" author ctext))
                       ((and author (not (string-empty-p author)))
                        (format "[%s] %s" author ctext))
                       (resolved
                        (cl-incf n-done)
                        (format "[DONE] %s" ctext))
                       (t ctext))))
                (delete-region start-pos end-pos)
                (goto-char start-pos)
                (insert (format "{==%s==}{>>%s<<}" range-text ctext-tagged))
                (cl-incf n-hi) (cl-incf n-cmt))))))
      ;; Paragraph-level insertions/deletions (inserted/deleted whole paragraphs)
      (goto-char (point-min))
      (while (re-search-forward "\\[\\([^]]*\\)\\]{\\.paragraph-insertion[^}]*}" nil t)
        (replace-match "{++\\1++}" t) (cl-incf n-ins))
      (goto-char (point-min))
      (while (re-search-forward "\\[\\([^]]*\\)\\]{\\.paragraph-deletion[^}]*}" nil t)
        (replace-match "{--\\1--}" t) (cl-incf n-del))
      ;; Inline insertions/deletions: [text]{.insertion ...} / [text]{.deletion ...}
      ;; Reviewer insertions can wrap text containing `]` characters (inserted
      ;; citations like `[@key]`, nested attribute spans, etc.); a plain
      ;; `[^]]*' regex drops every such hit.  Walk the buffer instead with a
      ;; balanced-bracket scan: locate each `]{.insertion' / `]{.deletion'
      ;; marker, then count brackets backwards to find the matching `['.
      ;; Counter is a 1-element list so the cl-flet helper can mutate it
      ;; without depending on dynamic binding of the let-bound n-ins/n-del.
      (let ((ins-box (list 0))
            (del-box (list 0)))
        (cl-flet
            ((rewrite-span (class open-marker close-marker box)
               (goto-char (point-min))
               (let ((tag-re (concat (regexp-quote (concat "]{." class))
                                     "[^}]*}")))
                 (while (re-search-forward tag-re nil t)
                   (let* ((mend  (match-end 0))
                          (rbpos (match-beginning 0)) ; pos of `]'
                          (depth 1)
                          (p     (1- rbpos))          ; scan back from before `]'
                          (open  nil))
                     (while (and (> p (point-min)) (> depth 0))
                       (let ((c (char-after p)))
                         (cond ((eq c ?\]) (cl-incf depth))
                               ((eq c ?\[) (cl-decf depth)
                                (when (zerop depth) (setq open p)))))
                       (setq p (1- p)))
                     (when open
                       (let ((text (buffer-substring-no-properties
                                    (1+ open) rbpos)))
                         (delete-region open mend)
                         (goto-char open)
                         (insert open-marker text close-marker)
                         (setcar box (1+ (car box))))))))))
          (rewrite-span "insertion" "{++" "++}" ins-box)
          (rewrite-span "deletion"  "{--" "--}" del-box))
        (cl-incf n-ins (car ins-box))
        (cl-incf n-del (car del-box)))
      ;; Adjacent {--del--}{++ins++} pairs -> single {~~old~>new~~} substitution
      (goto-char (point-min))
      (while (re-search-forward
              "{--\\([^{}]*?\\)--}\\(?:[[:space:]]*\\){\\+\\+\\([^{}]*?\\)\\+\\+}"
              nil t)
        (replace-match "{~~\\1~>\\2~~}" t)
        (cl-decf n-ins) (cl-decf n-del) (cl-incf n-sub))
      ;; Preserve docx anchor spans (e.g. `[]{#citeproc_bib_item_1 .anchor}'
      ;; on each bibliography entry) by stripping the `.anchor' class.
      ;; Pandoc's md->org writer drops a Span with [class] but preserves
      ;; one with just an id, emitting it as `<<NAME>>'.  That round-trips
      ;; back to a Word bookmark on docx export, so the existing in-text
      ;; `[[#citeproc_bib_item_N][N]]' hyperlinks still resolve.
      (goto-char (point-min))
      (while (re-search-forward
              "\\[\\]{\\(#[A-Za-z][A-Za-z0-9_-]*\\)[[:space:]]+\\.anchor}"
              nil t)
        (replace-match "[]{\\1}" t))
      (write-file md-file))
    (list :ins n-ins :del n-del :sub n-sub :hi n-hi :cmt n-cmt :done n-done)))

(defun otd--count (re)
  "Count buffer-wide matches of RE."
  (save-excursion
    (goto-char (point-min))
    (let ((n 0)) (while (re-search-forward re nil t) (cl-incf n)) n)))

(defun otd--stash-criticmarkup (md-file)
  "Replace every CriticMarkup token in MD-FILE with an opaque sentinel.
Return an alist mapping each sentinel string to the original token
text.  Stage 3 (`pandoc md->org') would otherwise parse `~~..~~' as
strikethrough and `~..~' as subscript, mangling `{~~old~>new~~}'
substitutions and any `{++..++}'/`{--..--}' containing such characters.
Sentinels are alphanumeric (`OTDIMP00001ZEND') and survive pandoc
untouched; `otd--unstash-criticmarkup' restores them in the org output."
  (let ((alist nil)
        (id 0))
    (with-temp-buffer
      (insert-file-contents md-file)
      (cl-flet
          ((stash-region (start end)
             (cl-incf id)
             (let* ((token (buffer-substring-no-properties start end))
                    (tag   (format "OTDIMP%05dZEND" id)))
               (push (cons tag token) alist)
               (delete-region start end)
               (goto-char start)
               (insert tag))))
        ;; Order matters: handle substitutions and highlight+comment pairs
        ;; before bare insertions/deletions so the outer brackets are
        ;; stashed first.  Each pattern is anchored to its full closing
        ;; delimiter so inner `{++..++}' tokens are NOT prematurely stashed
        ;; inside a `{==..==}{>>..<<}' highlight (the comment text often
        ;; contains other markup that should travel inside the sentinel).
        (goto-char (point-min))
        (while (re-search-forward
                "{==\\(?:[^{}]\\|{[^{}]*}\\)*?==}{>>\\(?:[^<>]\\|<[^<>]*>\\)*?<<}"
                nil t)
          (stash-region (match-beginning 0) (match-end 0)))
        (goto-char (point-min))
        (while (re-search-forward "{~~[^{}]*?~~}" nil t)
          (stash-region (match-beginning 0) (match-end 0)))
        (goto-char (point-min))
        (while (re-search-forward "{\\+\\+[^{}]*?\\+\\+}" nil t)
          (stash-region (match-beginning 0) (match-end 0)))
        (goto-char (point-min))
        (while (re-search-forward "{--[^{}]*?--}" nil t)
          (stash-region (match-beginning 0) (match-end 0)))
        ;; Orphan highlights (no following comment) and orphan comments.
        (goto-char (point-min))
        (while (re-search-forward "{==[^{}]*?==}" nil t)
          (stash-region (match-beginning 0) (match-end 0)))
        (goto-char (point-min))
        (while (re-search-forward "{>>\\(?:[^<>]\\|<[^<>]*>\\)*?<<}" nil t)
          (stash-region (match-beginning 0) (match-end 0))))
      (write-file md-file))
    alist))

(defun otd--unstash-criticmarkup (org-file alist)
  "Replace OTDIMP sentinels in ORG-FILE with the original CriticMarkup
tokens from ALIST (produced by `otd--stash-criticmarkup')."
  (when alist
    (with-temp-buffer
      (insert-file-contents org-file)
      (dolist (entry alist)
        (let ((tag (car entry))
              (val (cdr entry)))
          (goto-char (point-min))
          (while (search-forward tag nil t)
            (replace-match val t t))))
      (write-file org-file))))

(defun otd--preserve-leading-spaces (docx-in)
  "Return a fresh .docx copy of DOCX-IN with leading regular spaces
in `<w:t xml:space=\"preserve\">' runs replaced by NBSP (U+00A0).

Pandoc's docx reader trims leading whitespace from cell text during
AST parsing, dropping the visual indentation that signals hierarchy
in a Word table (e.g. `  Age 68' nested under `Age, mean (SD), years').
NBSP is preserved through the entire AST -> markdown -> org -> docx
pipeline and renders identically in Word.

Caller is responsible for `delete-file' on the returned path."
  (let* ((tmpdir  (make-temp-file "otd-docx-" t))
         (out     (make-temp-file "otd-docx-fixed-" nil ".docx"))
         (doc-xml (expand-file-name "word/document.xml" tmpdir)))
    (delete-file out)  ; zip refuses to write into an existing non-zip file
    (condition-case err
        (progn
          (unless (zerop (call-process "unzip" nil nil nil
                                       "-q" docx-in "-d" tmpdir))
            (error "unzip failed on %s" docx-in))
          (with-temp-buffer
            (insert-file-contents doc-xml)
            (goto-char (point-min))
            (let ((nbsp (string ? )))
              ;; Anchor on the literal `xml:space="preserve"' rather than a
              ;; greedy `<w:t[^>]*...>' regexp.  `document.xml' is a single line
              ;; and can contain very long bracket-free spans (e.g. a base64
              ;; Mendeley citation blob in `<w:tag w:val="...">', ~450k chars):
              ;; any unbounded greedy quantifier scanned across such a span
              ;; overflows Emacs' regexp matcher stack ("Stack overflow in
              ;; regexp matcher").  A literal `search-forward' plus bounded
              ;; `looking-at' does no backtracking and is immune to line length.
              (while (search-forward "xml:space=\"preserve\"" nil t)
                (when (and (save-excursion
                             (let ((lt (search-backward "<" nil t)))
                               (and lt (looking-at "<w:t[ />]"))))
                           (looking-at "[^<>]*>"))   ; remainder of start tag
                  (goto-char (match-end 0))          ; just past the `>'
                  (when (looking-at "[ \t]\\{2,\\}")
                    (let ((n (- (match-end 0) (match-beginning 0))))
                      (replace-match
                       (apply #'concat (make-list n nbsp)) t t))))))
            (write-region (point-min) (point-max) doc-xml nil 'silent))
          (let ((default-directory (file-name-as-directory tmpdir)))
            (unless (zerop (call-process "zip" nil nil nil
                                         "-q" "-r" out "."))
              (error "zip failed building %s" out)))
          out)
      (error
       (delete-directory tmpdir t)
       (when (file-exists-p out) (delete-file out))
       (signal (car err) (cdr err))))
    (delete-directory tmpdir t)
    out))

(defun otd--shield-grid-tables (md-file)
  "Replace each pandoc grid_table block in MD-FILE with a unique
alphanumeric sentinel.  Return an alist mapping each sentinel to
the original grid_table text (including its trailing newline).

Pandoc's md->org collapses grid_tables (multi-line cell capable)
into org pipe tables (single-line cells), merging adjacent cell
lines into single strings.  Shielding keeps them out of the org
reader; the caller restores them as raw markdown export blocks."
  (let ((alist nil) (id 0))
    (with-temp-buffer
      (insert-file-contents md-file)
      (goto-char (point-min))
      (while (re-search-forward "^\\+\\(?:[-=]+\\+\\)+[[:space:]]*$" nil t)
        (let ((start (match-beginning 0)))
          (goto-char start)
          (forward-line 1)
          (while (and (not (eobp))
                      (looking-at "^[+|]"))
            (forward-line 1))
          (let* ((end  (point))
                 (text (buffer-substring-no-properties start end))
                 (n    (cl-incf id))
                 (tag  (format "OTDGRIDTBLPHX%05dXEND" n)))
            (delete-region start end)
            (goto-char start)
            (insert tag "\n")
            (push (cons tag text) alist))))
      (write-region (point-min) (point-max) md-file nil 'silent))
    alist))

(defun otd--restore-grid-tables (org-file alist)
  "Replace each sentinel from ALIST in ORG-FILE with a `#+BEGIN_EXPORT
markdown' block holding the original grid_table.  The block survives
`otd-export's two-step org->md->docx pipeline as raw markdown, which
pandoc's docx writer renders with multi-line cells intact."
  (when alist
    (with-temp-buffer
      (insert-file-contents org-file)
      (dolist (pair alist)
        (goto-char (point-min))
        (when (search-forward (car pair) nil t)
          (replace-match
           (concat "#+BEGIN_EXPORT markdown\n"
                   (cdr pair)
                   (unless (string-suffix-p "\n" (cdr pair)) "\n")
                   "#+END_EXPORT")
           t t)))
      (write-region (point-min) (point-max) org-file nil 'silent))))

(defun otd--parse-author-block (text)
  "Parse `#+AFFIL:'/`#+AUTHOR_LIST:'/`#+AUTHOR_GROUP:' lines out of TEXT.

Returns (AFFILS AUTHORS GROUP), where AFFILS is an alist of
(KEY . DESCRIPTION) in declaration order, AUTHORS is a list of
\(NAME AFFIL-KEYS CORRESPONDING-P), and GROUP is the trailing
consortium/group-credit string or nil.

Line syntax:
  #+AFFIL: key :: institution, address ...
  #+AUTHOR_LIST: Name :: key1, key2 :: corresponding
  #+AUTHOR_GROUP: for the ... Study Group

The `::' fields are used (rather than commas) because affiliation
text routinely contains commas.  `corresponding' is the only
recognized flag in the third field; absence of a third field means
not corresponding."
  (let (affils authors group)
    (dolist (line (split-string text "\n"))
      (cond
       ((string-match "^#\\+AFFIL:[ \t]*\\([^ \t]+\\)[ \t]*::[ \t]*\\(.*\\)$" line)
        ;; Capture into locals immediately: `match-string' reads global
        ;; match-data, which a nested regex call (e.g. `split-string'
        ;; below) would otherwise clobber before we get to read it.
        (let ((key (match-string 1 line))
              (desc (match-string 2 line)))
          (push (cons key desc) affils)))
       ((string-match "^#\\+AUTHOR_LIST:[ \t]*\\(.*?\\)[ \t]*::[ \t]*\\(.*?\\)\\(?:[ \t]*::[ \t]*\\(.*\\)\\)?$" line)
        (let ((name (match-string 1 line))
              (keys-str (match-string 2 line))
              (flags-str (match-string 3 line)))
          (push (list name
                      (split-string keys-str "[ \t]*,[ \t]*")
                      (and flags-str (string-match-p "corresponding" flags-str)))
                authors)))
       ((string-match "^#\\+AUTHOR_GROUP:[ \t]*\\(.*\\)$" line)
        (setq group (match-string 1 line)))))
    (list (nreverse affils) (nreverse authors) group)))

(defun otd--generate-author-block (text)
  "Generate the flat markdown author/affiliation block from TEXT.

Affiliation superscript letters (a, b, c, ...) are assigned by
`#+AFFIL:' declaration order -- the same order the numbered
affiliation list is printed in -- so letter N always matches list
item N, matching the journal's existing convention."
  (cl-destructuring-bind (affils authors group) (otd--parse-author-block text)
    (let* ((letters (let ((i -1))
                      (mapcar (lambda (a) (cl-incf i) (cons (car a) (string (+ ?a i)))) affils)))
           (author-strs
            (mapcar (lambda (au)
                      (let* ((name (nth 0 au))
                             (keys (nth 1 au))
                             (corresponding (nth 2 au))
                             (sups (append (mapcar (lambda (k) (cdr (assoc k letters))) keys)
                                           (when corresponding '("*")))))
                        ;; `^text^' (pandoc markdown superscript), not
                        ;; org's `^{text}' -- this block is emitted inside
                        ;; a #+begin_export markdown span, which bypasses
                        ;; org's own superscript handling entirely and is
                        ;; read by pandoc's *markdown* reader instead, so
                        ;; it needs pandoc's own superscript syntax or the
                        ;; braces render as literal, unformatted text.
                        (format "%s^%s^" name (string-join sups ","))))
                    authors))
           (author-line (concat (string-join author-strs ", ")
                                 (when group (concat ", " group))))
           (affil-lines
            (let ((n 0))
              (mapcar (lambda (a)
                        (cl-incf n)
                        (format "%d. %s" n (cdr a)))
                      affils)))
           (has-corresponding (cl-some (lambda (au) (nth 2 au)) authors)))
      (concat "#+begin_export markdown\n"
              author-line "\n\n"
              (string-join affil-lines "\n\n")
              "\n\n"
              (when has-corresponding "*Corresponding author.\n")
              "#+end_export"))))

(defun otd--expand-author-block ()
  "In the current buffer, replace `#+AFFIL:'/`#+AUTHOR_LIST:'/
`#+AUTHOR_GROUP:' header lines with a generated `#+begin_export
markdown' author/affiliation block.

This is a pre-pandoc textual rewrite of the raw org source -- the
same approach `otd--read-pandoc-options' uses for `#+PANDOC_OPTIONS:'
-- because pandoc's org reader cannot carry structured per-author
affiliation data through its native metadata: repeated `#+AUTHOR:'
lines merge into one string, and arbitrary custom `#+KEY:' lines are
silently dropped.  Doing nothing if no such lines are present keeps
this backward-compatible with a hand-typed `#+begin_export markdown'
block."
  (let ((line-re "^#\\+\\(?:AFFIL\\|AUTHOR_LIST\\|AUTHOR_GROUP\\):.*\n?")
        (text (buffer-string))
        regions)
    (when (string-match-p "^#\\+\\(?:AFFIL\\|AUTHOR_LIST\\|AUTHOR_GROUP\\):" text)
      (let ((block (otd--generate-author-block text)))
        (goto-char (point-min))
        (while (re-search-forward line-re nil t)
          (push (cons (match-beginning 0) (match-end 0)) regions))
        (setq regions (nreverse regions))
        (when regions
          (let ((first-start (caar regions)))
            (dolist (r (reverse regions))
              (delete-region (car r) (cdr r)))
            (goto-char first-start)
            (insert block "\n\n")))))))

(defun otd--read-pandoc-options (org-file)
  "Return a list of pandoc CLI args derived from headers in ORG-FILE.

Each `#+PANDOC_OPTIONS: key:value' line becomes `--key=value' (or
`--key' for boolean t/true/yes).  This mirrors the convention used
by ox-pandoc, so an existing canonical .org file works unchanged.

Each `#+bibliography: path' line becomes `--bibliography=path',
unless an equivalent `#+PANDOC_OPTIONS: bibliography:path' is also
present (the explicit form wins).  This means a stock org-cite
setup (#+bibliography + #+PANDOC_OPTIONS: citeproc:t + csl) is
enough to get citations resolved into the docx output."
  (let (args bib-pandoc bib-org)
    (with-temp-buffer
      (insert-file-contents org-file)
      ;; #+PANDOC_OPTIONS: key:value
      (goto-char (point-min))
      (while (re-search-forward
              "^#\\+PANDOC_OPTIONS:[[:space:]]+\\([A-Za-z][A-Za-z0-9_-]*\\)[[:space:]]*:\\(.*\\)$"
              nil t)
        (let ((key (match-string 1))
              (val (string-trim (match-string 2))))
          (when (string= key "bibliography")
            (push val bib-pandoc))
          (cond
           ((member val '("t" "true" "yes" "TRUE" "T" "Yes"))
            (push (concat "--" key) args))
           (t
            (push (concat "--" key "=" val) args)))))
      ;; #+bibliography: path
      (goto-char (point-min))
      (while (re-search-forward
              "^#\\+bibliography:[[:space:]]+\\(.*\\)$"
              nil t)
        (push (string-trim (match-string 1)) bib-org)))
    (dolist (b (nreverse bib-org))
      (unless (member b bib-pandoc)
        (push (concat "--bibliography=" b) args)))
    (nreverse args)))

(defun otd--decode-image-paths (md-file)
  "URL-decode `file://' URIs in markdown image links in MD-FILE.

Pandoc emits image refs as `![alt](file:///path/with%20encoded%20spaces)'
when reading from org.  The docx writer fails to embed such images
because it tries to open the literal percent-encoded path on disk.
Decoding here lets pandoc find the media file and embed it natively."
  (require 'url-util)
  (with-temp-buffer
    (insert-file-contents md-file)
    (goto-char (point-min))
    (while (re-search-forward
            "!\\[\\([^]]*\\)\\](file://\\([^)]+\\))"
            nil t)
      ;; `url-unhex-string' performs its own regex matching and clobbers
      ;; the match data from the outer search, so `replace-match' would
      ;; otherwise splice the new text into whatever sub-region url-util
      ;; last matched -- leaving the original `file://' URL untouched and
      ;; causing an infinite loop.  Isolate it with `save-match-data'.
      (let* ((alt  (match-string 1))
             (raw  (match-string 2))
             (path (save-match-data (url-unhex-string raw))))
        (replace-match (format "![%s](%s)" alt path) t t)))
    (write-region (point-min) (point-max) md-file nil 'silent)))

(defun otd--postfix-fixups ()
  "Apply pandoc-org-reader fixups in the current buffer.
Pandoc treats `^{...}'/`_{...}' as super/subscript only when glued to
an alphanumeric character.  Several patterns produced during the docx
-> md -> org import otherwise survive the next docx export as the
literal four characters `^{N}':

  - whitespace before the marker:  `text ^{41}',
  - punctuation before the marker:  `(VETSA)^{28--30}', `/umx/^{40}',
    `et al.^{47}', `]^{5,25}',
  - Zotero-broken citation links wrapping the marker:
    `[[https://www.zotero.org/google-docs/?broken=...][^{25}]]'
    (the `[' before `^{' is matched by the punctuation rule below,
    so the hyperlink is preserved and the superscript still attaches).

Whitespace is stripped; punctuation gets a zero-width space inserted
between it and the marker (pandoc treats U+200B as a non-space char,
so the marker attaches without visibly changing the text)."
  ;; Strip whitespace between a non-space char and `^{' or `_{'.
  (goto-char (point-min))
  (while (re-search-forward
          "\\([^[:space:]]\\)[[:space:]]+\\(\\^\\|_\\){"
          nil t)
    (replace-match "\\1\\2{"))
  ;; Insert U+200B between punctuation and `^{'/`_{' so pandoc's org
  ;; reader sees a non-space, non-alphanumeric attachment point.  The
  ;; ZWSP itself is excluded from the negative class so this is safe to
  ;; run repeatedly.
  (let ((zwsp (string ?​)))
    (goto-char (point-min))
    (while (re-search-forward
            (concat "\\([^[:alnum:][:space:]" zwsp "]\\)\\(\\^\\|_\\){")
            nil t)
      (replace-match (concat "\\1" zwsp "\\2{"))))
  ;; Drop pandoc-emitted definition-list duplicate captions for
  ;; pandoc-crossref targets.  When a docx that was originally exported
  ;; with `[cite:@fig:LABEL]' references is re-imported, pandoc's
  ;; org writer emits the rendered caption as a `term : definition'
  ;; pair right after the real `#+name: / #+caption:` block, e.g.:
  ;;
  ;;     [cite:@fig:rmzrdz]: Within-pair MZ and DZ correlations...
  ;;
  ;; These duplicate captions are NOT picked up by pandoc-crossref on
  ;; the next export and otherwise sit alongside the real figure
  ;; caption as orphaned prose, often misattributing one figure's
  ;; caption to the next figure block.  Remove every such line
  ;; (single- or multi-paragraph) emitted at column 0.
  (goto-char (point-min))
  (while (re-search-forward
          "^\\[cite:@\\(?:fig\\|tbl\\|sec\\|eq\\|lst\\):[A-Za-z][A-Za-z0-9_:.-]*\\]:[ \t].*\\(?:\n[^\n[:space:]].*\\)*\n*"
          nil t)
    (replace-match "" t t)))

(defconst otd--image-ext-re
  "\\.\\(?:png\\|jpe?g\\|gif\\|pdf\\|svg\\|emf\\|eps\\|tiff?\\|bmp\\|webp\\)\\'"
  "Regexp (used with `case-fold-search') matching image/figure file
extensions in an org link target.")

(defun otd--url-scheme-p (path)
  "Return non-nil if PATH begins with a URL scheme like `http:' (not a
bare filesystem path)."
  (string-match-p "\\`[A-Za-z][A-Za-z0-9+.-]*:" path))

(defun otd--map-image-links (fn)
  "Rewrite each org image-link target in the current buffer through FN.
Handles both `[[TARGET]]' and `[[TARGET][DESC]]'.  FN receives the bare
target path (a leading `file:' prefix stripped) and returns the new full
target string (prefix included), or nil to leave the link unchanged.
Only links whose target carries an image extension (see
`otd--image-ext-re') are considered."
  (goto-char (point-min))
  (while (re-search-forward
          "\\[\\[\\([^][\n]+\\)\\(\\(?:\\]\\[[^][\n]*\\)?\\)\\]\\]" nil t)
    (let* ((raw  (match-string 1))           ; full target, e.g. "file:fig.png"
           (rest (match-string 2))           ; "" or "][DESC"
           ;; `string-match' clobbers the outer search's match data that
           ;; `replace-match' below relies on; isolate it.  (`string-match-p'
           ;; and the path functions in FN preserve match data.)
           (path (save-match-data
                   (if (string-match "\\`file:" raw)
                       (substring raw (match-end 0)) raw)))
           (case-fold-search t))
      (when (string-match-p otd--image-ext-re path)
        (let ((new (funcall fn path)))
          (when (and new (not (string= new raw)))
            (replace-match (concat "[[" new rest "]]") t t)))))))

(defun otd--fileify-image-links ()
  "Ensure every local (non-URL) image link in the current buffer carries a
`file:' prefix.  Pandoc's org reader only emits an Image for a `file:'
link (or an absolute path); a bare relative `[[fig-media/media/f1.png]]'
is otherwise dropped as a `.spurious-link' text span and never embeds.
The path itself is left unchanged (relative stays relative); the docx
stage's `--resource-path' then resolves and embeds it.  Idempotent."
  (otd--map-image-links
   (lambda (path)
     (unless (otd--url-scheme-p path)
       (concat "file:" path)))))

(defun otd--relativize-image-links (dir)
  "Make absolute image-link targets under DIR relative to DIR, and give
every local image link a `file:' prefix, so an imported org file stays
portable across machines (the docx importer's `--extract-media' writes
machine-specific absolute paths) and still re-exports as an embedded
image.  Targets outside DIR and URLs are left untouched."
  (setq dir (file-name-as-directory (expand-file-name dir)))
  (otd--map-image-links
   (lambda (path)
     (cond
      ((otd--url-scheme-p path) nil)
      ((file-name-absolute-p path)
       (let ((abs (expand-file-name path)))
         (when (string-prefix-p dir abs)
           (concat "file:" (file-relative-name abs dir)))))
      (t (concat "file:" path))))))

(defun otd--postfix-org (org-file)
  "Apply `otd--postfix-fixups' to ORG-FILE on disk and relativize its
image links (see `otd--relativize-image-links') so the imported file is
portable."
  (with-temp-buffer
    (insert-file-contents org-file)
    (otd--postfix-fixups)
    (otd--relativize-image-links (file-name-directory (expand-file-name org-file)))
    (write-region (point-min) (point-max) org-file nil 'silent)))

;;;###autoload
(defun otd-postfix-buffer ()
  "Apply org-reader fixups to the current buffer in place.
Use this once on a `-tracked.org' file imported before the fixups
existed so a re-export to docx preserves super/subscripts."
  (interactive)
  (save-excursion (otd--postfix-fixups))
  (message "otd-postfix-buffer: superscript adjacency normalised"))

;;;; --- highlighting ----------------------------------------------------

;; Every face below starts from `:inherit default' before adding its own
;; override.  This is load-bearing, not decorative: org-mode's own emphasis
;; fontification (`=verbatim=', `~code~', `+strikethrough+', ...) treats
;; `{' and `}' as valid boundary characters by default (see
;; `org-emphasis-regexp-components'), so almost every CriticMarkup token
;; -- `{==...==}', `{--...--}', etc. -- accidentally also satisfies org's
;; own emphasis syntax and gets fontified a second time underneath ours.
;; Without an explicit `:inherit default' reset, any attribute OUR face
;; doesn't set (e.g. foreground, if we only set background) falls through
;; to org's face instead of the buffer's normal text -- which is why the
;; highlighted-range text was showing up in org-verbatim's magenta.
(defface otd-cm-marker
  '((t :inherit default :height 0.75 :weight normal :slant normal :foreground "gray55"))
  "Face for raw CriticMarkup delimiter punctuation ({++, ++}, {--,
--}, {~~, ~>, ~~}, {>>, <<}, {==, ==}).  Shrunk and dimmed relative
to body text so the delimiters recede visually instead of competing
with the substantive text for attention -- the actual inserted/
deleted/comment text stays at the buffer's normal font height.")

(defface otd-insert
  '((t :inherit default :foreground "#2e7d32" :underline t :weight normal :slant normal))
  "Face for {++inserted++} text.")

(defface otd-delete
  '((t :inherit default :foreground "#c62828" :strike-through t :weight normal :slant normal))
  "Face for {--deleted--} text.")

(defface otd-substitute-old
  '((t :inherit default :foreground "#c62828" :strike-through t :weight normal :slant normal))
  "Face for the old/replaced half of a {~~old~>new~~} substitution.")

(defface otd-substitute-new
  '((t :inherit default :foreground "#2e7d32" :underline t :weight normal :slant normal))
  "Face for the new/replacement half of a {~~old~>new~~} substitution.")

(defface otd-comment
  '((t :inherit default :foreground "#455a64" :weight normal :slant normal))
  "Face for {>>comment<<} reviewer comment text.")

(defface otd-comment-author
  '((t :inherit default :foreground "#1565c0" :weight normal :slant normal))
  "Face for the [Author] tag inside a reviewer comment.")

(defface otd-comment-done
  '((t :inherit default :foreground "#2e7d32" :strike-through t :weight normal :slant normal))
  "Face for the [DONE] tag marking a resolved comment.")

(defface otd-highlight
  '((t :inherit default :background "#fff9c4" :weight normal :slant normal))
  "Face for {==highlighted==} text -- a soft background wash, no
foreground/weight change, so it layers cleanly under an attached
comment's own coloring.")

;; Each pattern is split into its own delimiter/content capture groups
;; so `otd-cm-marker' can be applied to the punctuation alone (shrunk +
;; dimmed) while the substantive text keeps the buffer's normal height
;; and gets a face that is color-coded (not size-coded) by change type.
;; Pandoc may stamp {++[author]++} / {--[author]--} attributions; the
;; optional [author]/[DONE] groups use `t' (laxmatch) since they don't
;; always participate in the match.
(defvar otd--keywords
  '(;; Guard against a false-positive collision with org's own `<<target>>'
    ;; radio-link syntax.  A comment closes with `<<}' and the next one
    ;; opens with `{>>'; since paragraphs in an `--wrap=none' export are
    ;; long single lines, two comments sharing a paragraph put a bare
    ;; `<<' ... `>>' pair around all the plain prose between them, which
    ;; org(-modern) then reads as a radio target and fontifies (org-modern
    ;; shrinks/dims it) -- the "text right after a comment looks smaller
    ;; and gray" symptom.  This rule only matches OUR OWN literal `<<}'/
    ;; `{>>' delimiter pair (not bare `<<'/`>>'), so it can't suppress an
    ;; unrelated, genuine org radio target elsewhere in the document.
    ;;
    ;; It runs FIRST (lowest priority) and unconditionally defaults the
    ;; whole span, including any highlight/insert/delete tokens nested
    ;; inside it; the rules below run AFTER and prepend their own face
    ;; for their own matched characters, so they still win over this
    ;; blanket default wherever they overlap it.
    ("<<}\\([^<>]*?\\){>>" (1 'default prepend))
    ("\\({\\+\\+\\)\\(\\[[^]]+\\]\\)?\\([^{}]*?\\)\\(\\+\\+}\\)"
     (1 'otd-cm-marker prepend)
     (2 'otd-comment-author prepend t)
     (3 'otd-insert prepend)
     (4 'otd-cm-marker prepend))
    ("\\({--\\)\\(\\[[^]]+\\]\\)?\\([^{}]*?\\)\\(--}\\)"
     (1 'otd-cm-marker prepend)
     (2 'otd-comment-author prepend t)
     (3 'otd-delete prepend)
     (4 'otd-cm-marker prepend))
    ("\\({~~\\)\\([^~{}]*?\\)\\(~>\\)\\([^~{}]*?\\)\\(~~}\\)"
     (1 'otd-cm-marker prepend)
     (2 'otd-substitute-old prepend)
     (3 'otd-cm-marker prepend)
     (4 'otd-substitute-new prepend)
     (5 'otd-cm-marker prepend))
    ("\\({>>\\)\\(\\[[^]]+\\]\\)?[ \t]*\\(\\[DONE\\]\\)?[ \t]*\\([^{}<]*?\\)\\(<<}\\)"
     (1 'otd-cm-marker prepend)
     (2 'otd-comment-author prepend t)
     (3 'otd-comment-done prepend t)
     (4 'otd-comment prepend)
     (5 'otd-cm-marker prepend))
    ("\\({==\\)\\(.*?\\)\\(==}\\)"
     (1 'otd-cm-marker prepend)
     (2 'otd-highlight prepend)
     (3 'otd-cm-marker prepend))
    ;; Reassert `org-cite' on any `[cite:...]' unconditionally, as the
    ;; last (highest-priority) rule.  A `[cite:@key]' that happens to
    ;; fall inside the false `<<...>>' radio-target span above (see the
    ;; neutralizer rule) loses its citation face entirely, not just its
    ;; priority -- org's own citation fontification and the spurious
    ;; org-modern radio-target rule are both part of org's *own* keyword
    ;; pass, so whichever of THOSE two runs later wins before any of our
    ;; rules even get a turn.  Reapplying `org-cite' here sidesteps that
    ;; tug-of-war entirely rather than trying to referee it.
    ("\\(\\[cite:[^][]*\\]\\)" (1 'org-cite prepend)))
  "Font-lock keywords for CriticMarkup tokens.")

(defvar otd-criticmarkup-mode-map
  (let ((m (make-sparse-keymap)))
    (define-key m (kbd "M-n")   #'otd-cm-next)
    (define-key m (kbd "M-p")   #'otd-cm-prev)
    (define-key m (kbd "C-c a") #'otd-cm-accept-at-point)
    (define-key m (kbd "C-c r") #'otd-cm-reject-at-point)
    (define-key m (kbd "C-c d") #'otd-cm-mark-done-at-point)
    (define-key m (kbd "C-c D") #'otd-cm-unmark-done-at-point)
    (define-key m (kbd "C-c A") #'otd-accept-all)
    (define-key m (kbd "C-c R") #'otd-reject-all)
    (define-key m (kbd "C-c X") #'otd-cm-remove-done)
    m)
  "Keymap for `otd-criticmarkup-mode'.")

(defun otd--cm-header-button (label tip cmd)
  "Return a mouse-clickable LABEL for the CriticMarkup header line.
TIP is shown as tooltip and CMD is the command bound to mouse-1."
  (let ((map (make-sparse-keymap)))
    (define-key map [header-line mouse-1]
      (lambda () (interactive)
        (with-current-buffer (window-buffer (selected-window))
          (call-interactively cmd))))
    (propertize (format " %s " label)
                'mouse-face 'mode-line-highlight
                'help-echo tip
                'local-map map)))

(defvar otd--cm-header-line-format
  '(:eval
    (concat
     (otd--cm-header-button "◀"    "Previous CriticMarkup (M-p)"      #'otd-cm-prev)
     (otd--cm-header-button "▶"    "Next CriticMarkup (M-n)"          #'otd-cm-next)
     "│"
     (otd--cm-header-button "✓"    "Accept at point (C-c a)"          #'otd-cm-accept-at-point)
     (otd--cm-header-button "✗"    "Reject at point (C-c r)"          #'otd-cm-reject-at-point)
     "│"
     (otd--cm-header-button "✓✓"   "Accept ALL (C-c A)"               #'otd-accept-all)
     (otd--cm-header-button "✗✗"   "Reject ALL (C-c R)"               #'otd-reject-all)
     "│"
     (otd--cm-header-button "done" "Mark comment DONE (C-c d)"        #'otd-cm-mark-done-at-point)
     (otd--cm-header-button "undo done" "Unmark DONE (C-c D)"         #'otd-cm-unmark-done-at-point)
     (otd--cm-header-button "🗑 done" "Remove all DONE comments (C-c X)" #'otd-cm-remove-done)))
  "Header-line spec used while `otd-criticmarkup-mode' is on.")

;;;###autoload
(define-minor-mode otd-criticmarkup-mode
  "Minor mode that fontifies CriticMarkup tokens and exposes a
review toolbar in the header line.

\\{otd-criticmarkup-mode-map}

The header-line buttons mirror the keymap commands:
  ◀ / ▶          `otd-cm-prev'/`otd-cm-next'      (M-p / M-n)
  ✓ / ✗          `otd-cm-accept-at-point' / `otd-cm-reject-at-point' (C-c a / C-c r)
  ✓✓ / ✗✗        `otd-accept-all' / `otd-reject-all'  (C-c A / C-c R)
  done / undo    `otd-cm-mark-done-at-point' / `otd-cm-unmark-done-at-point' (C-c d / C-c D)
  🗑 done        `otd-cm-remove-done'                (C-c X)"
  :lighter " CM"
  :keymap otd-criticmarkup-mode-map
  (cond
   (otd-criticmarkup-mode
    (font-lock-add-keywords nil otd--keywords 'append)
    ;; Stash the user's previous header-line so we can restore on disable.
    (setq-local otd--cm-saved-header-line header-line-format)
    (setq header-line-format otd--cm-header-line-format))
   (t
    (font-lock-remove-keywords nil otd--keywords)
    (setq header-line-format
          (and (boundp 'otd--cm-saved-header-line)
               otd--cm-saved-header-line))
    (kill-local-variable 'otd--cm-saved-header-line)))
  (font-lock-flush))

;;;; --- accept / reject -------------------------------------------------

(defun otd--replace (re repl)
  "Replace every match of RE with REPL across the buffer."
  (save-excursion
    (goto-char (point-min))
    (while (re-search-forward re nil t)
      (replace-match repl t nil))))

;;;###autoload
(defun otd-accept-all ()
  "Accept all CriticMarkup changes in the current buffer.
Insertions are kept, deletions are dropped, substitutions take the new
value, comments are dropped, highlights are kept as plain text."
  (interactive)
  (otd--replace "{\\+\\+\\(?:\\[[^]]+\\]\\)?\\([^{}]*?\\)\\+\\+}" "\\1")
  (otd--replace "{--\\(?:\\[[^]]+\\]\\)?\\([^{}]*?\\)--}"         "")
  (otd--replace "{~~\\([^~{}]*?\\)~>\\([^~{}]*?\\)~~}"            "\\2")
  (otd--replace "{>>[^{}<]*<<}"                                   "")
  (otd--replace "{==\\(.*?\\)==}"                               "\\1"))

;;;###autoload
(defun otd-reject-all ()
  "Reject all CriticMarkup changes in the current buffer (revert to original)."
  (interactive)
  (otd--replace "{\\+\\+\\(?:\\[[^]]+\\]\\)?\\([^{}]*?\\)\\+\\+}" "")
  (otd--replace "{--\\(?:\\[[^]]+\\]\\)?\\([^{}]*?\\)--}"         "\\1")
  (otd--replace "{~~\\([^~{}]*?\\)~>\\([^~{}]*?\\)~~}"            "\\1")
  (otd--replace "{>>[^{}<]*<<}"                                   "")
  (otd--replace "{==\\(.*?\\)==}"                               "\\1"))

;;;; --- review workflow: navigate, accept/reject at point, manage done -

(defconst otd--cm-token-re
  ;; Matches any single CriticMarkup token.  Group 1 captures the
  ;; opening sentinel so callers know which kind.  Order matters in
  ;; the alternation because `{==' would otherwise be eaten by the
  ;; earlier `{--' alternative on a string like `{==}'.
  (concat
   "\\("
   "{\\+\\+\\(?:\\[[^]]+\\]\\)?[^{}]*?\\+\\+}"      ; insertion
   "\\|{--\\(?:\\[[^]]+\\]\\)?[^{}]*?--}"            ; deletion
   "\\|{~~[^~{}]*?~>[^~{}]*?~~}"                     ; substitution
   "\\|{==.*?==}"                                  ; highlight
   "\\|{>>[^{}<]*<<}"                                ; comment
   "\\)")
  "Single-pass regex matching any one CriticMarkup token.")

(defun otd--cm-token-kind (text)
  "Return one of the symbols `ins', `del', `sub', `hi', `cmt'
describing the CriticMarkup TEXT, or nil if TEXT is not a token."
  (cond
   ((string-prefix-p "{++" text) 'ins)
   ((string-prefix-p "{--" text) 'del)
   ((string-prefix-p "{~~" text) 'sub)
   ((string-prefix-p "{==" text) 'hi)
   ((string-prefix-p "{>>" text) 'cmt)))

(defun otd--cm-token-at-point ()
  "Return (KIND BEG END TEXT) for the CriticMarkup token containing
point or, if point is between tokens, the next one on the same line.
Return nil when no token is reachable on the current line.

A `hi' token that is immediately followed by one or more `cmt'
tokens is returned as the whole `{==..==}{>>..<<}...' run; that
matches what users perceive as a single reviewer comment.

Walks the line left-to-right rather than searching backward from
EOL, because a line can carry many tokens and only one of them
contains point."
  (save-excursion
    (let* ((orig (point))
           (bol (line-beginning-position))
           (eol (line-end-position))
           (hit nil))
      (goto-char bol)
      (while (and (not hit)
                  (re-search-forward otd--cm-token-re eol t))
        (let ((beg (match-beginning 0))
              (end (match-end 0)))
          ;; Extend `hi' across any trailing comment chain so accept/
          ;; reject treats `{==X==}{>>cmt<<}' as one logical edit.
          (when (eq (otd--cm-token-kind
                     (buffer-substring-no-properties beg end))
                    'hi)
            (goto-char end)
            (while (looking-at "{>>[^{}<]*<<}")
              (setq end (match-end 0))
              (goto-char end)))
          (cond
           ;; Point is inside this token (or right at its closer).
           ((and (<= beg orig) (<= orig end))
            (setq hit (list beg end)))
           ;; Point is BEFORE this token: it's the next-on-line.
           ((< orig beg)
            (setq hit (list beg end))))
          (goto-char end)))
      (when hit
        (let* ((beg  (nth 0 hit))
               (end  (nth 1 hit))
               (text (buffer-substring-no-properties beg end)))
          (list (otd--cm-token-kind text) beg end text))))))

;;;###autoload
(defun otd-cm-next ()
  "Move point to the next CriticMarkup token in the buffer."
  (interactive)
  (let ((start (point)))
    ;; Skip past the token at point so we move forward predictably.
    (when (looking-at otd--cm-token-re)
      (goto-char (match-end 0)))
    (if (re-search-forward otd--cm-token-re nil t)
        (goto-char (match-beginning 0))
      (goto-char start)
      (message "otd-cm-next: no more CriticMarkup tokens"))))

;;;###autoload
(defun otd-cm-prev ()
  "Move point to the previous CriticMarkup token in the buffer."
  (interactive)
  (let ((start (point)))
    (if (re-search-backward otd--cm-token-re nil t)
        nil   ; point is already at match-beginning
      (goto-char start)
      (message "otd-cm-prev: no previous CriticMarkup tokens"))))

(defun otd--cm-apply (action)
  "Apply ACTION (`accept' or `reject') to the CriticMarkup token at point.

`accept' takes coauthor edits as proposed (keep insertions, drop
deletions, take the new side of substitutions, strip highlight
markers, drop comments).  `reject' reverts (drop insertions, keep
deletions, take the old side of substitutions, strip highlight
markers, drop comments)."
  (let ((tok (otd--cm-token-at-point)))
    (unless tok (user-error "No CriticMarkup token at point"))
    (let* ((kind (nth 0 tok))
           (beg  (nth 1 tok))
           (end  (nth 2 tok))
           (text (nth 3 tok))
           (new
            (pcase (cons kind action)
              (`(ins . accept)
               (and (string-match
                     "\\`{\\+\\+\\(?:\\[[^]]+\\]\\)?\\([^{}]*?\\)\\+\\+}\\'"
                     text)
                    (match-string 1 text)))
              (`(ins . reject) "")
              (`(del . accept) "")
              (`(del . reject)
               (and (string-match
                     "\\`{--\\(?:\\[[^]]+\\]\\)?\\([^{}]*?\\)--}\\'"
                     text)
                    (match-string 1 text)))
              (`(sub . accept)
               (and (string-match
                     "\\`{~~[^~{}]*?~>\\([^~{}]*?\\)~~}\\'" text)
                    (match-string 1 text)))
              (`(sub . reject)
               (and (string-match
                     "\\`{~~\\([^~{}]*?\\)~>[^~{}]*?~~}\\'" text)
                    (match-string 1 text)))
              ((or `(hi . accept) `(hi . reject))
               ;; Strip the `{==...==}' wrapper and drop every trailing
               ;; `{>>..<<}' so the visible prose is what remains.
               (when (string-match
                      "\\`{==\\(.*?\\)==}\\(?:{>>[^{}<]*<<}\\)*\\'" text)
                 (match-string 1 text)))
              ((or `(cmt . accept) `(cmt . reject)) ""))))
      (unless new
        (user-error "Could not parse %s token" kind))
      (save-excursion
        (goto-char beg)
        (delete-region beg end)
        (insert new))
      (goto-char beg))))

;;;###autoload
(defun otd-cm-accept-at-point ()
  "Accept the CriticMarkup change at point."
  (interactive)
  (otd--cm-apply 'accept))

;;;###autoload
(defun otd-cm-reject-at-point ()
  "Reject the CriticMarkup change at point."
  (interactive)
  (otd--cm-apply 'reject))

(defun otd--cm-comment-at-point ()
  "Return (BEG END BODY DONE-P) for the `{>>..<<}' comment containing
point.  Walks the line forward so a line carrying multiple
stacked comments resolves to the one actually under point."
  (save-excursion
    (let ((orig (point))
          (bol (line-beginning-position))
          (eol (line-end-position))
          (hit nil))
      (goto-char bol)
      (while (and (not hit)
                  (re-search-forward "{>>\\([^{}<]*\\)<<}" eol t))
        (let ((beg  (match-beginning 0))
              (end  (match-end 0))
              (body (match-string 1)))
          (when (and (<= beg orig) (<= orig end))
            (setq hit (list beg end body
                            (and (string-match-p
                                  "\\`\\(\\[[^]]+\\][[:space:]]+\\)?\\[DONE\\]"
                                  body)
                                 t))))))
      hit)))

;;;###autoload
(defun otd-cm-mark-done-at-point ()
  "Mark the comment at point as resolved by inserting `[DONE]'
immediately after the author tag (or at the start when no author).
No-op if the comment already carries `[DONE]'."
  (interactive)
  (let ((c (otd--cm-comment-at-point)))
    (unless c (user-error "No reviewer comment at point"))
    (when (nth 3 c)
      (user-error "Comment is already marked DONE"))
    (let* ((beg  (nth 0 c))
           (body (nth 2 c))
           (new-body
            (cond
             ((string-match "\\`\\(\\[[^]]+\\]\\)[[:space:]]+\\(.*\\)\\'" body)
              (format "%s [DONE] %s"
                      (match-string 1 body) (match-string 2 body)))
             (t (concat "[DONE] " body)))))
      (save-excursion
        (goto-char beg)
        (delete-region beg (nth 1 c))
        (insert (format "{>>%s<<}" new-body)))
      (message "Marked DONE: %s" (or (string-trim new-body) "")))))

;;;###autoload
(defun otd-cm-unmark-done-at-point ()
  "Remove the `[DONE]' marker from the comment at point."
  (interactive)
  (let ((c (otd--cm-comment-at-point)))
    (unless c (user-error "No reviewer comment at point"))
    (unless (nth 3 c) (user-error "Comment is not marked DONE"))
    (let* ((beg  (nth 0 c))
           (end  (nth 1 c))
           (body (nth 2 c))
           (new (replace-regexp-in-string
                 "\\(\\[[^]]+\\][[:space:]]+\\)?\\[DONE\\][[:space:]]*"
                 (lambda (m)
                   (if (string-match "\\`\\(\\[[^]]+\\]\\)" m)
                       (concat (match-string 1 m) " ")
                     ""))
                 body t)))
      (save-excursion
        (goto-char beg)
        (delete-region beg end)
        (insert (format "{>>%s<<}" new))))))

;;;###autoload
(defun otd-cm-remove-done ()
  "Delete every comment that carries the `[DONE]' marker.

An orphan highlight (`{==text==}' with no remaining `{>>..<<}'
attached) has its markers stripped so the visible prose stays
intact; otherwise other reviewers' comments in the same stack
are preserved."
  (interactive)
  (let ((removed 0)
        (orphans 0))
    (save-excursion
      (goto-char (point-min))
      (while (re-search-forward
              "{>>\\(\\[[^]]+\\][[:space:]]+\\)?\\[DONE\\][^<]*<<}"
              nil t)
        (replace-match "" t t)
        (cl-incf removed)))
    ;; Clean up highlights that now stand alone (no `{>>..<<}' chain).
    ;; Emacs regex has no `(?!...)' lookahead, so check by hand.
    (save-excursion
      (goto-char (point-min))
      (while (re-search-forward "{==\\(.*?\\)==}" nil t)
        (let ((beg (match-beginning 0))
              (end (match-end 0))
              (inner (match-string 1)))
          (unless (looking-at "{>>")
            (delete-region beg end)
            (goto-char beg)
            (insert inner)
            (cl-incf orphans)))))
    (message "otd-cm-remove-done: removed %d comment%s%s"
             removed (if (= removed 1) "" "s")
             (if (> orphans 0)
                 (format ", stripped %d orphan highlight%s"
                         orphans (if (= orphans 1) "" "s"))
               ""))))

;;;; --- compare against canonical org -----------------------------------

;;;###autoload
(defun otd-diff-against (canonical)
  "Show a coloured word-diff between CANONICAL .org and current buffer.
Run after `otd-import' on the docx returned by a co-author to see Word
edits aligned to the org master."
  (interactive
   (list (read-file-name "Canonical .org: " nil nil t nil
                         (lambda (f) (or (file-directory-p f)
                                         (string-match-p "\\.org\\'" f))))))
  (let ((current (buffer-file-name)))
    (unless current (user-error "Buffer is not visiting a file"))
    (with-current-buffer (get-buffer-create "*otd-diff*")
      (let ((inhibit-read-only t)) (erase-buffer))
      (call-process "git" nil t nil
                    "--no-pager" "diff" "--no-index" "--no-color"
                    "--word-diff=color" "--word-diff-regex=[A-Za-z0-9]+|[^[:space:]]"
                    (expand-file-name canonical) current)
      (goto-char (point-min))
      (ansi-color-apply-on-region (point-min) (point-max))
      (special-mode)
      (display-buffer (current-buffer)))))

;;;; --- embed / extract / merge (round-trip via customXml part) ---------

(defconst otd--cx-item-name      "item-otd-source.xml")
(defconst otd--cx-itemprops-name "itemProps-otd-source.xml")
(defconst otd--cx-namespace      "urn:org-tracked-docx:source")
(defconst otd--cxprop-name       "OrgTrackedSource"
  "Name of the docProps/custom.xml fallback property holding the
embedded org source, base64-encoded.  Word, LibreOffice, AND WPS
Office preserve custom document properties across saves; some of
them (notably WPS) strip the customXml/ part on save, so we write
to both locations and read whichever survived.")

(defun otd--xml-cdata-escape (s)
  "Wrap S in a CDATA section, escaping any inner `]]>'."
  (concat "<![CDATA["
          (replace-regexp-in-string "]]>" "]]]]><![CDATA[>" s t t)
          "]]>"))

(defun otd--embed-org-source (docx-path org-path)
  "Embed ORG-PATH content inside DOCX-PATH as a customXml part.
Properly registers the part in [Content_Types].xml and the document
relationships so Word/LibreOffice preserve it across edits.  Counterpart
to `otd--extract-org-source'."
  (let* ((tmpdir (make-temp-file "otd-embed-" t))
         (org-content (with-temp-buffer
                        (insert-file-contents org-path)
                        (buffer-string))))
    (condition-case err
        (progn
          (unless (zerop (call-process "unzip" nil nil nil
                                       "-q" docx-path "-d" tmpdir))
            (error "unzip failed on %s" docx-path))
          (let ((cx-dir (expand-file-name "customXml" tmpdir)))
            (make-directory cx-dir t)
            (make-directory (expand-file-name "_rels" cx-dir) t)
            (with-temp-file (expand-file-name otd--cx-item-name cx-dir)
              (insert "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
                      "<orgTrackedSource xmlns=\"" otd--cx-namespace "\">"
                      (otd--xml-cdata-escape org-content)
                      "</orgTrackedSource>\n"))
            (with-temp-file (expand-file-name otd--cx-itemprops-name cx-dir)
              (insert "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
                      "<ds:datastoreItem ds:itemID=\"{ORG-TRACKED-DOCX-SOURCE}\""
                      " xmlns:ds=\"http://schemas.openxmlformats.org/officeDocument/2006/customXml\">"
                      "<ds:schemaRefs><ds:schemaRef ds:uri=\""
                      otd--cx-namespace "\"/></ds:schemaRefs>"
                      "</ds:datastoreItem>\n"))
            (with-temp-file (expand-file-name (concat "_rels/" otd--cx-item-name ".rels")
                                              cx-dir)
              (insert "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
                      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                      "<Relationship Id=\"rId1\""
                      " Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps\""
                      " Target=\"" otd--cx-itemprops-name "\"/>"
                      "</Relationships>\n")))
          ;; [Content_Types].xml — register the new parts.
          (let ((ct (expand-file-name "[Content_Types].xml" tmpdir)))
            (with-temp-buffer
              (insert-file-contents ct)
              (goto-char (point-min))
              (when (re-search-forward "</Types>" nil t)
                (goto-char (match-beginning 0))
                (insert "<Override PartName=\"/customXml/" otd--cx-item-name
                        "\" ContentType=\"application/xml\"/>"
                        "<Override PartName=\"/customXml/" otd--cx-itemprops-name
                        "\" ContentType=\"application/vnd.openxmlformats-officedocument.customXmlProperties+xml\"/>"))
              (write-region (point-min) (point-max) ct nil 'silent)))
          ;; word/_rels/document.xml.rels — point document at the customXml part.
          (let ((rels (expand-file-name "word/_rels/document.xml.rels" tmpdir))
                (max-id 0))
            (with-temp-buffer
              (insert-file-contents rels)
              (goto-char (point-min))
              (while (re-search-forward "Id=\"rId\\([0-9]+\\)\"" nil t)
                (setq max-id (max max-id (string-to-number (match-string 1)))))
              (goto-char (point-min))
              (when (re-search-forward "</Relationships>" nil t)
                (goto-char (match-beginning 0))
                (insert (format "<Relationship Id=\"rId%d\"" (1+ max-id))
                        " Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml\""
                        " Target=\"../customXml/" otd--cx-item-name "\"/>"))
              (write-region (point-min) (point-max) rels nil 'silent)))
          ;; Belt-and-braces: also embed in docProps/custom.xml as a
          ;; base64 property.  WPS Office strips customXml/ on save but
          ;; keeps custom document properties; this redundant copy is
          ;; what lets the round-trip survive a WPS edit.  Off by
          ;; default (see `otd-embed-custom-property'): a single custom
          ;; property holding the entire org source as one base64 string
          ;; (150KB+ for a real manuscript) is well outside what
          ;; docProps/custom.xml properties are meant to hold, and is a
          ;; confirmed, reproducible cause of Word's "found unreadable
          ;; content" repair prompt on this document.
          (when otd-embed-custom-property
            (otd--embed-custom-property tmpdir org-content))
          ;; Repackage.
          (delete-file docx-path)
          (let ((default-directory (file-name-as-directory tmpdir)))
            (unless (zerop (call-process "zip" nil nil nil "-q" "-r" docx-path "."))
              (error "zip failed building %s" docx-path))))
      (error
       (delete-directory tmpdir t)
       (signal (car err) (cdr err))))
    (delete-directory tmpdir t)
    docx-path))

(defun otd--embed-custom-property (tmpdir org-content)
  "Embed ORG-CONTENT (base64-encoded) as `OrgTrackedSource' property
in docProps/custom.xml inside TMPDIR (an extracted docx tree).
Creates and registers custom.xml if pandoc didn't already emit it."
  (let* ((cxp (expand-file-name "docProps/custom.xml" tmpdir))
         (b64 (base64-encode-string
               (encode-coding-string org-content 'utf-8) t)))
    ;; If pandoc didn't write docProps/custom.xml, create it now and
    ;; register it in [Content_Types].xml + _rels/.rels.
    (unless (file-exists-p cxp)
      (make-directory (expand-file-name "docProps" tmpdir) t)
      (with-temp-file cxp
        (insert "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
                "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/custom-properties\""
                " xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">"
                "</Properties>\n"))
      ;; [Content_Types].xml override
      (let ((ct (expand-file-name "[Content_Types].xml" tmpdir)))
        (when (file-exists-p ct)
          (with-temp-buffer
            (insert-file-contents ct)
            (goto-char (point-min))
            (unless (search-forward "/docProps/custom.xml" nil t)
              (goto-char (point-min))
              (when (re-search-forward "</Types>" nil t)
                (goto-char (match-beginning 0))
                (insert "<Override PartName=\"/docProps/custom.xml\""
                        " ContentType=\"application/vnd.openxmlformats-officedocument.custom-properties+xml\"/>")))
            (write-region (point-min) (point-max) ct nil 'silent))))
      ;; Top-level _rels/.rels relationship
      (let ((rels (expand-file-name "_rels/.rels" tmpdir))
            (max-id 0))
        (when (file-exists-p rels)
          (with-temp-buffer
            (insert-file-contents rels)
            (goto-char (point-min))
            (unless (search-forward "docProps/custom.xml" nil t)
              (goto-char (point-min))
              (while (re-search-forward "Id=\"rId\\([0-9]+\\)\"" nil t)
                (setq max-id (max max-id (string-to-number (match-string 1)))))
              (goto-char (point-min))
              (when (re-search-forward "</Relationships>" nil t)
                (goto-char (match-beginning 0))
                (insert (format "<Relationship Id=\"rId%d\"" (1+ max-id))
                        " Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties\""
                        " Target=\"docProps/custom.xml\"/>")))
            (write-region (point-min) (point-max) rels nil 'silent)))))
    (with-temp-buffer
      (insert-file-contents cxp)
      ;; Drop any stale OrgTrackedSource property from a previous embed.
      (goto-char (point-min))
      (while (re-search-forward
              (concat "<property[^>]*name=\""
                      (regexp-quote otd--cxprop-name)
                      "\"[^>]*>[^<]*\\(?:<[^>]+>[^<]*</[^>]+>\\)*</property>")
              nil t)
        (replace-match "" t t))
      ;; Compute a fresh pid (must be >= 2 and unique per ECMA-376).
      (let ((max-pid 1))
        (goto-char (point-min))
        (while (re-search-forward "pid=\"\\([0-9]+\\)\"" nil t)
          (setq max-pid (max max-pid (string-to-number (match-string 1)))))
        (goto-char (point-min))
        (when (re-search-forward "</Properties>" nil t)
          (goto-char (match-beginning 0))
          (insert (format "<property fmtid=\"{D5CDD505-2E9C-101B-9397-08002B2CF9AE}\" pid=\"%d\" name=\"%s\"><vt:lpwstr>%s</vt:lpwstr></property>"
                          (1+ max-pid) otd--cxprop-name b64))))
      (write-region (point-min) (point-max) cxp nil 'silent))))

(defun otd--extract-org-source (docx-path)
  "Return the embedded org source from DOCX-PATH or nil.
Tries the customXml/ part first (preserved by Word and LibreOffice),
falls back to the docProps/custom.xml `OrgTrackedSource' property
\(preserved by WPS Office and other tools that strip customXml/).

Normalizes CRLF/CR to LF: Word rewrites embedded customXml parts with
CRLF line endings whenever it re-saves the docx (e.g. after a
coauthor's review round), even though `otd--embed-org-source' only
ever wrote LF. The stray CR before each newline shifts every
`otd--split-paragraphs' paragraph boundary check and breaks the
`^:PROPERTIES:$'/`^:END:$' drawer regexes in `otd--merge-content',
which then fail to peel the property drawer off the first body
paragraph -- corrupting that paragraph's alignment fingerprint and
cascading into dropped or misplaced content for everything merged
after it (we saw the Abstract paragraph itself replaced by the
document's byline block this way)."
  (let ((source (or (otd--extract-org-source-customxml docx-path)
                     (otd--extract-org-source-property  docx-path))))
    (and source (replace-regexp-in-string "\r\n?" "\n" source))))

(defun otd--extract-org-source-property (docx-path)
  "Extract source from docProps/custom.xml OrgTrackedSource property."
  (with-temp-buffer
    (when (zerop (call-process "unzip" nil t nil "-p"
                               docx-path "docProps/custom.xml"))
      (goto-char (point-min))
      (when (re-search-forward
             (concat "<property[^>]*name=\""
                     (regexp-quote otd--cxprop-name)
                     "\"[^>]*>[[:space:]]*<vt:lpwstr>\\([^<]*\\)</vt:lpwstr>")
             nil t)
        (let ((b64 (match-string 1)))
          (when (and b64 (not (string-empty-p b64)))
            (decode-coding-string (base64-decode-string b64) 'utf-8)))))))

(defun otd--extract-org-source-customxml (docx-path)
  "Extract source from a customXml/ part by content (LibreOffice
renames our `item-otd-source.xml' to `item1.xml' on save, so we
identify the part by its `<orgTrackedSource>' element)."
  (let ((entries
         (with-temp-buffer
           (call-process "unzip" nil t nil "-Z1" docx-path)
           (split-string (buffer-string) "\n" t)))
        result)
    (dolist (entry entries)
      (when (and (not result)
                 (string-match-p "^customXml/[^/]+\\.xml$" entry))
        (with-temp-buffer
          (when (zerop (call-process "unzip" nil t nil "-p" docx-path entry))
            (goto-char (point-min))
            (when (re-search-forward
                   (concat "<orgTrackedSource[^>]*xmlns=\""
                           (regexp-quote otd--cx-namespace)
                           "\"[^>]*>[[:space:]]*<!\\[CDATA\\[")
                   nil t)
              (let ((start (point)))
                (goto-char (point-max))
                (when (re-search-backward
                       "\\]\\]>[[:space:]]*</orgTrackedSource>"
                       nil t)
                  (let ((raw (buffer-substring-no-properties
                              start (match-beginning 0))))
                    (setq result
                          (replace-regexp-in-string
                           "]]]]><!\\[CDATA\\[>" "]]>" raw t t))))))))))
    result))

;;;###autoload
(defun otd-extract-source (docx)
  "Extract the embedded org source from DOCX and write it next to it
as `<basename>-source.org'.  Errors if no embedded source is found."
  (interactive
   (list (read-file-name "Docx file: " nil nil t nil
                         (lambda (f) (or (file-directory-p f)
                                         (string-match-p "\\.docx\\'" f))))))
  (let* ((dir (file-name-directory (expand-file-name docx)))
         (base (file-name-sans-extension (file-name-nondirectory docx)))
         (out (expand-file-name (concat base "-source.org") dir))
         (src (otd--extract-org-source (expand-file-name docx))))
    (unless src
      (user-error "No embedded org source in %s" (file-name-nondirectory docx)))
    (with-temp-file out (insert src))
    (find-file out)
    (message "Extracted source -> %s" (file-name-nondirectory out))
    out))

;;;; --- merge (canonical org + lossy tracked org) -----------------------

(defun otd--strip-criticmarkup (text)
  "Return TEXT with all CriticMarkup tokens flattened to the rejected
state (insertions removed, deletions kept, substitutions take old,
comments removed, highlights stripped of markers).

Runs the token substitutions to a fixpoint rather than a single pass.
When multiple reviewers comment on the same or overlapping span,
CriticMarkup nests, e.g. `{=={==text{====}{>>c1<<}==}{>>c2<<}==}{>>c3<<}'.
A single pass over `{==\\(.*?\\)==}' only unwinds the innermost
`==}{>>...<<}' it meets and leaves outer `{=='/`==}'/`{>>...<<}'
fragments behind in the result (e.g. `text{=={====}==}'); that
leftover then corrupts the normalized fingerprint `otd--normalize-for-
match' builds from it, which in turn desyncs `otd--align-paragraphs'
from that paragraph on and cascades into dropped/misplaced content
for everything after it in the merge. Looping until the string stops
changing peels one nesting level per pass, so any depth of nested
comments fully flattens."
  (let ((s text) (prev nil))
    (while (not (equal s prev))
      (setq prev s)
      (setq s (replace-regexp-in-string "{\\+\\+\\(?:\\[[^]]+\\]\\)?[^{}]*?\\+\\+}" "" s))
      (setq s (replace-regexp-in-string "{--\\(?:\\[[^]]+\\]\\)?\\([^{}]*?\\)--}" "\\1" s))
      (setq s (replace-regexp-in-string "{~~\\([^~{}]*?\\)~>[^~{}]*?~~}" "\\1" s))
      (setq s (replace-regexp-in-string "{==\\([^{}]*?\\)==}{>>[^<]*?<<}" "\\1" s))
      (setq s (replace-regexp-in-string "{==\\([^{}]*?\\)==}" "\\1" s))
      (setq s (replace-regexp-in-string "{>>[^<]*?<<}" "" s)))
    s))

(defun otd--has-criticmarkup-p (text)
  "Return non-nil if TEXT contains any CriticMarkup token."
  (string-match-p "{\\(\\+\\+\\|--\\|~~\\|==\\|>>\\)" text))

(defconst otd--xref-rendered-regexp
  (concat "\\b\\(?:"
          "[Ff]ig\\(?:ure\\|\\.\\)\\|"
          "[Ss]ec\\(?:tion\\|\\.\\)\\|"
          "[Ee]q\\(?:uation\\|\\.\\)\\|"
          "[Tt]bl\\.\\|[Tt]able\\|"
          "[Ll]st\\.\\|[Ll]isting"
          "\\)[\u00a0[:space:]]*[0-9]+\\(?:[.\u2013-][0-9]+\\)?\\b")
  "Match a pandoc-crossref-rendered cross-ref like `fig. 1', `Sec. 2',
`Equation 3', `tbl. 4', `Table 4-2', etc.  Used to collapse every
rendered cross-ref to the same fingerprint as canonical's
`[cite:@type:label]' (which the citation rule already collapses to
`[CITE]') so paragraph alignment in the merge survives a docx
roundtrip through pandoc-crossref.")

(defun otd--normalize-for-match (text)
  "Normalize TEXT for paragraph matching across canonical and tracked.
Strips CriticMarkup, collapses every variant of citation AND cross-ref rendering
\(citeproc anchor, org-cite, pandoc-org `cite/t', bare `[@key]') to a
fixed `[CITE]' token, and collapses whitespace."
  (let ((s (otd--strip-criticmarkup text)))
    (setq s (replace-regexp-in-string
             "\\[\\[#citeproc_bib_item_[0-9]+\\]\\[[0-9]+\\]\\]"
             "[CITE]" s))
    (setq s (replace-regexp-in-string
             "\\[\\[cite\\(?:/[a-z]+\\)?:[^]]+\\]\\]" "[CITE]" s))
    (setq s (replace-regexp-in-string
             "\\[cite:[^]]+\\]" "[CITE]" s))
    (setq s (replace-regexp-in-string
             "\\[@[A-Za-z0-9_;,@[:space:]-]+\\]" "[CITE]" s))
    ;; Anchored cross-ref (linkReferences=true): `sec. [[#sec:foo][2]]'.
    (setq s (replace-regexp-in-string
             "\\(?:\\b\\(?:[Ff]ig\\(?:ure\\|\\.\\)\\|[Ss]ec\\(?:tion\\|\\.\\)\\|[Ee]q\\(?:uation\\|\\.\\)\\|[Tt]bl\\.\\|[Tt]able\\|[Ll]st\\.\\|[Ll]isting\\)[\u00a0 ]+\\)?\\[\\[#\\(?:fig\\|sec\\|eq\\|tbl\\|lst\\):[^]]+\\]\\[[0-9]+\\]\\]"
             "[CITE]" s))
    ;; pandoc-crossref-rendered cross-refs: `fig. 1', `Section 2', etc.
    (setq s (replace-regexp-in-string otd--xref-rendered-regexp "[CITE]" s))
    ;; NBSP (U+00A0) appears in tracked text because `otd--preserve-leading-
    ;; spaces' converts table-cell leading spaces; the canonical has regular
    ;; spaces.  Treat them as equivalent for matching.
    (setq s (replace-regexp-in-string "[ \t\n ]+" " " s))
    (string-trim s)))

(defun otd--build-cite-map (org-content)
  "Walk ORG-CONTENT for `[cite:@key]' entries (including multi-key
forms like `[cite:@a;@b]') and return an alist (key . N) ordered by
first occurrence, mirroring pandoc-citeproc's numbering."
  (let ((map nil) (n 0))
    (with-temp-buffer
      (insert org-content)
      (goto-char (point-min))
      (while (re-search-forward "\\[cite:\\([^]]+\\)\\]" nil t)
        (let ((body (match-string 1)))
          (dolist (k (split-string body "[;,[:space:]]+" t))
            (when (string-prefix-p "@" k)
              (let ((key (substring k 1)))
                (unless (assoc key map)
                  (cl-incf n)
                  (push (cons key n) map))))))))
    (nreverse map)))

(defun otd--substitute-cite-keys (text cite-map)
  "Convert lossy citation forms in TEXT back to canonical `[cite:@key]'.
Handles every form pandoc emits when reading docx and writing org:
- citeproc-anchored: `[[#citeproc_bib_item_N][N]]' (uses CITE-MAP for N->key)
- pandoc-org reconstruction: `[[cite/t:@key]]' / `[[cite:@key]]'
- bare: `[@key]'
- Vancouver-style numeric superscript: `^{N}', `^{N,M}', `^{N--M}',
  `^{N-M}', mixed `^{1,2,5--7}'.  The numeric form is what citeproc
  produces for numbered styles (e.g. Nature, Vancouver) and is the
  default body-of-text form on a docx round-trip; without this branch
  every cite collapses to a bare `^{N}' and the next export ships
  just `1.', `2.', etc. with no keys."
  (let ((s text)
        ;; Reverse map: position N (string) -> key, for the numeric-
        ;; superscript pass.  EXCLUDE pandoc-crossref cross-refs
        ;; (`fig:LABEL', `tbl:LABEL', `sec:LABEL', `eq:LABEL',
        ;; `lst:LABEL') because citeproc numbers only bibliographic
        ;; references; cross-refs are numbered separately by
        ;; pandoc-crossref ("Figure 1", "Table 2", ...) and get
        ;; round-tripped by `otd--substitute-xrefs'.  Renumber the
        ;; remaining keys 1..N so the reverse map matches the
        ;; citeproc-rendered superscripts in the docx body.
        (n->key (let ((h (make-hash-table :test 'equal))
                      (n 0))
                  (dolist (e cite-map h)
                    (let ((key (car e)))
                      (unless (string-match-p
                               "\\`\\(?:fig\\|tbl\\|sec\\|eq\\|lst\\):"
                               key)
                        (cl-incf n)
                        (puthash (number-to-string n) key h)))))))
    (dolist (entry cite-map)
      (let* ((key (car entry))
             (n   (cdr entry))
             (re  (format "\\[\\[#citeproc_bib_item_%d\\]\\[%d\\]\\]" n n)))
        (setq s (replace-regexp-in-string
                 re (format "[cite:@%s]" key) s t t))))
    (setq s (replace-regexp-in-string
             "\\[\\[cite\\(?:/[a-z]+\\)?:\\([^]]+\\)\\]\\]"
             "[cite:\\1]" s t))
    ;; Numeric superscript form: `^{N}', `^{N,M}', `^{N--M}'.  Org
    ;; uses U+2013 (EN DASH) for ranges via pandoc's `--' shortcut,
    ;; but the raw ASCII `--' also occurs.  The character class
    ;; accepts both.  `save-match-data' is mandatory because the
    ;; inner `string-match'/`split-string' calls clobber the outer
    ;; regex's match data, which `replace-regexp-in-string' relies
    ;; on after the function returns -- without it, only the first
    ;; number in each superscript is consumed.
    (setq s
          (replace-regexp-in-string
           "\\^{\\([0-9,;[:space:]–-]+\\)}"
           (lambda (m)
             (save-match-data
               (let* ((body (substring m 2 (1- (length m))))
                      (nums nil))
                 (dolist (chunk (split-string body "[,;[:space:]]+" t))
                   (cond
                    ((string-match
                      "\\`\\([0-9]+\\)[–-]\\{1,2\\}\\([0-9]+\\)\\'"
                      chunk)
                     (let ((a (string-to-number (match-string 1 chunk)))
                           (b (string-to-number (match-string 2 chunk))))
                       (when (> a b) (cl-rotatef a b))
                       (dotimes (i (1+ (- b a)))
                         (push (number-to-string (+ a i)) nums))))
                    ((string-match-p "\\`[0-9]+\\'" chunk)
                     (push chunk nums))))
                 (let ((keys (delq nil
                                   (mapcar (lambda (n) (gethash n n->key))
                                           (nreverse nums)))))
                   (if keys
                       (concat "[cite:"
                               (mapconcat (lambda (k) (concat "@" k)) keys ";")
                               "]")
                     m)))))
           s t))
    (setq s (replace-regexp-in-string
             "\\[\\(@[A-Za-z0-9_;,@[:space:]-]+\\)\\]"
             "[cite:\\1]" s t))
    s))

(defun otd--build-xref-map (canonical)
  "Walk CANONICAL for pandoc-crossref target definitions and return a
plist mapping each cross-ref type to an alist of (LABEL . N), where
N is the document-order number pandoc-crossref assigns at export.

Detects:
- :fig — `#+NAME: fig:LABEL'
- :tbl — `#+NAME: tbl:LABEL'
- :lst — `#+NAME: lst:LABEL'
- :sec — `:CUSTOM_ID: sec:LABEL' under any heading
- :eq  — `{#eq:LABEL}' anywhere (typically right after `$$ ... $$')

Used by `otd--substitute-xrefs' to rewrite `fig. 1', `Section 2',
etc. back to `[cite:@fig:LABEL]' form during merge."
  (let ((figs nil) (secs nil) (eqs nil) (tbls nil) (lsts nil)
        (n-fig 0) (n-sec 0) (n-eq 0) (n-tbl 0) (n-lst 0))
    (with-temp-buffer
      (insert canonical)
      (goto-char (point-min))
      (while (re-search-forward
              "^#\\+NAME:[ \t]*\\(fig\\|tbl\\|lst\\):\\([A-Za-z][A-Za-z0-9_:.-]*\\)"
              nil t)
        (let ((type  (downcase (match-string 1)))
              (label (match-string 2)))
          (cond
           ((string= type "fig") (cl-incf n-fig) (push (cons label n-fig) figs))
           ((string= type "tbl") (cl-incf n-tbl) (push (cons label n-tbl) tbls))
           ((string= type "lst") (cl-incf n-lst) (push (cons label n-lst) lsts)))))
      (goto-char (point-min))
      (while (re-search-forward
              "^:CUSTOM_ID:[ \t]*sec:\\([A-Za-z][A-Za-z0-9_:.-]*\\)"
              nil t)
        (cl-incf n-sec) (push (cons (match-string 1) n-sec) secs))
      (goto-char (point-min))
      (while (re-search-forward "{#eq:\\([A-Za-z][A-Za-z0-9_:.-]*\\)}" nil t)
        (cl-incf n-eq) (push (cons (match-string 1) n-eq) eqs)))
    (list :fig (nreverse figs) :sec (nreverse secs)
          :eq  (nreverse eqs)  :tbl (nreverse tbls) :lst (nreverse lsts))))

(defun otd--substitute-xrefs (text xref-map)
  "Replace pandoc-crossref-rendered cross-refs in TEXT with the canonical
`[cite:@TYPE:LABEL]' form, using XREF-MAP from `otd--build-xref-map'.

Recognises both abbreviated (`fig. 1', `Sec. 2') and long forms
\(`Figure 1', `Section 2', `Equation 3', `Table 4', `Listing 5'),
case-insensitively, with either a regular space or a non-breaking
space between the prefix and the number (pandoc-crossref's default
template uses NBSP)."
  (let ((s text))
    ;; Anchored form (linkReferences=true): pandoc emits
    ;; `sec. [[#sec:foo][2]]' which carries the LABEL in the link target,
    ;; so XREF-MAP isn't needed.  Strip the optional prefix word too --
    ;; the cite key already names the type.
    (setq s (replace-regexp-in-string
             "\\(?:\\b\\(?:[Ff]ig\\(?:ure\\|\\.\\)\\|[Ss]ec\\(?:tion\\|\\.\\)\\|[Ee]q\\(?:uation\\|\\.\\)\\|[Tt]bl\\.\\|[Tt]able\\|[Ll]st\\.\\|[Ll]isting\\)[\u00a0 ]+\\)?\\[\\[#\\(\\(?:fig\\|sec\\|eq\\|tbl\\|lst\\):[^]]+\\)\\]\\[[0-9]+\\]\\]"
             "[cite:@\\1]" s))
    (cl-flet
        ((sub (type-key prefix patterns)
           (dolist (entry (plist-get xref-map type-key))
             (let* ((label (car entry))
                    (n     (cdr entry))
                    (rep   (format "[cite:@%s:%s]" prefix label)))
               (dolist (pat patterns)
                 (setq s (replace-regexp-in-string
                          (format pat n) rep s t t)))))))
      (sub :fig "fig"
           '("\\b[Ff]igure[  ]+%d\\b"
             "\\b[Ff]ig\\.[  ]*%d\\b"))
      (sub :sec "sec"
           '("\\b[Ss]ection[  ]+%d\\b"
             "\\b[Ss]ec\\.[  ]*%d\\b"))
      (sub :eq  "eq"
           '("\\b[Ee]quation[  ]+%d\\b"
             "\\b[Ee]q\\.[  ]*%d\\b"))
      (sub :tbl "tbl"
           '("\\b[Tt]able[  ]+%d\\b"
             "\\b[Tt]bl\\.[  ]*%d\\b"))
      (sub :lst "lst"
           '("\\b[Ll]isting[  ]+%d\\b"
             "\\b[Ll]st\\.[  ]*%d\\b")))
    s))

(defun otd--split-paragraphs (content)
  "Split CONTENT into paragraphs for merge alignment.

A paragraph is a maximal run of consecutive non-blank lines that are
not org headings (`*+ '), not property drawers (`:PROPERTIES:'..`:END:'),
not export blocks (`#+BEGIN_EXPORT'..`#+END_EXPORT'), not src blocks
(`#+BEGIN_SRC'..`#+END_SRC'), and not other keyword lines (`#+...').
Headings become their own one-line paragraphs; drawers, export/src
blocks, and keyword lines are skipped from matching (they are emitted
by pandoc inconsistently between canonical and lossy versions and
would otherwise prevent paragraph-level alignment)."
  (let ((paras nil)
        (current nil)
        (in-drawer nil)
        (in-export nil)
        (in-src nil))
    (dolist (line (split-string content "\n"))
      (cond
       ((string-match-p "^[ \t]*:PROPERTIES:[ \t]*$" line)
        (setq in-drawer t))
       ((and in-drawer (string-match-p "^[ \t]*:END:[ \t]*$" line))
        (setq in-drawer nil))
       (in-drawer)
       ;; Org keywords are case-insensitive: `#+BEGIN_EXPORT' and
       ;; `#+begin_export' are equivalent.  Pandoc's org writer emits
       ;; lowercase, so a case-sensitive match on the uppercase form
       ;; missed the canonical's frontmatter export block entirely and
       ;; surfaced the author/affiliation lines as `body paragraphs',
       ;; which then misaligned every following paragraph by ~6.
       ((let ((case-fold-search t))
          (string-match-p "^#\\+BEGIN_EXPORT" line))
        (setq in-export t))
       ((and in-export
             (let ((case-fold-search t))
               (string-match-p "^#\\+END_EXPORT" line)))
        (setq in-export nil))
       (in-export)
       ;; Same rationale for `#+begin_src'/`#+end_src' asset blocks
       ;; (bibtex/CSL/reference-doc): Word never shows the reviewer
       ;; the raw asset, only citeproc's rendered numbered reference
       ;; list, so those lines can never fingerprint-match a bibtex
       ;; entry and must not be counted as alignable body paragraphs.
       ((let ((case-fold-search t))
          (string-match-p "^#\\+BEGIN_SRC" line))
        (setq in-src t))
       ((and in-src
             (let ((case-fold-search t))
               (string-match-p "^#\\+END_SRC" line)))
        (setq in-src nil))
       (in-src)
       ((string-match-p "^#\\+" line))     ;; skip keyword lines
       ((string-empty-p (string-trim line))
        (when current
          (push (mapconcat #'identity (nreverse current) "\n") paras)
          (setq current nil)))
       ((string-match-p "^\\*+[ \t]" line)
        (when current
          (push (mapconcat #'identity (nreverse current) "\n") paras)
          (setq current nil))
        (push line paras))
       (t (push line current))))
    (when current
      (push (mapconcat #'identity (nreverse current) "\n") paras))
    (nreverse paras)))

(defun otd--parse-comment-author (text)
  "Parse a `[Author Name] rest…' prefix out of TEXT.
Returns (AUTHOR . REST) when the prefix is present, otherwise (nil . TEXT).
`otd--rewrite-spans' adds the prefix on import so `otd-export' can route
the original reviewer name back to Word's `<w:comment w:author=…>'.
Allows one level of nested brackets in the name (e.g. `[Chandra Reynolds
[2]]'), since a bracket-exclusion class alone can't match those and
would otherwise leave the whole `[Name [2]]' tag as unparsed comment
text, hiding a `[DONE]' marker one level too deep for
`otd--strip-done-prefix' to find at the very start of the body."
  (if (string-match "\\`\\[\\(\\(?:[^][]\\|\\[[^][]*\\]\\)+\\)\\][ \t]+\\(\\(?:.\\|\n\\)*\\)\\'" text)
      (cons (match-string 1 text) (match-string 2 text))
    (cons nil text)))

(defun otd--scan-annotations (text)
  "Return a list of (RANGE . CMT) pairs from TEXT.
Balance-aware scan: walks `{==' / `==}' delimiters as a bracket stack
so overlapping/nested `{==…==}{>>…<<}' annotations yield correct
inner and outer pairs (regex with lazy `.*?' splits them wrong)."
  (let ((pos 0)
        (len (length text))
        (stack nil)
        (pairs nil))
    (while (< pos len)
      (let ((op (string-search "{==" text pos))
            (cl (string-search "==}" text pos)))
        (cond
         ((and op (or (null cl) (< op cl)))
          (push (+ op 3) stack)
          (setq pos (+ op 3)))
         (cl
          (let ((start (pop stack)))
            (if (and start
                     (<= (+ cl 6) len)
                     (string= (substring text (+ cl 3) (+ cl 6)) "{>>"))
                (let ((cmt-end (string-search "<<}" text (+ cl 6))))
                  (if cmt-end
                      (progn
                        (push (cons (substring text start cl)
                                    (substring text (+ cl 6) cmt-end))
                              pairs)
                        (setq pos (+ cmt-end 3)))
                    (setq pos (+ cl 3))))
              (setq pos (+ cl 3)))))
         (t (setq pos len)))))
    (nreverse pairs)))

(defun otd--anchor-all-positions (needle haystack &optional min-len)
  "Return all unannotated start positions of NEEDLE in HAYSTACK, in order.
Returns nil if NEEDLE is shorter than MIN-LEN (default 3) or empty
after trimming.  A position is unannotated when it is not the content
of an existing `{==…==}{>>…<<}' wrap (rough check on the surrounding
chars)."
  (let ((min (or min-len 3))
        (nlen (length needle)))
    (when (and (>= nlen min)
               (not (string-empty-p (string-trim needle))))
      (let ((start 0) (positions nil))
        (while (let ((p (string-search needle haystack start)))
                 (when p
                   (let ((before (if (>= p 3) (substring haystack (- p 3) p) ""))
                         (tail   (substring haystack (+ p nlen)
                                            (min (length haystack)
                                                 (+ p nlen 6)))))
                     (unless (or (string= before "{==")
                                 (string-prefix-p "==}{>>" tail))
                       (push p positions)))
                   (setq start (1+ p))
                   p)))
        (nreverse positions)))))

(defun otd--anchor-unique-position (needle haystack &optional min-len)
  "Return the unique unannotated start position of NEEDLE in HAYSTACK.
Returns nil if NEEDLE is shorter than MIN-LEN (default 3), empty after
trimming, or appears zero or 2+ times outside existing
`{==…==}{>>…<<}' annotations."
  (let ((positions (otd--anchor-all-positions needle haystack min-len)))
    (when (= 1 (length positions))
      (car positions))))

(defun otd--anchor-candidates (range xref-map)
  "Return a list of variant strings to try as anchors for RANGE.
Mechanical substitutions covering the common pandoc-roundtrip
divergences between tracked and canonical: em-dash, pandoc-crossref
rendering (`Figure N' / `Table N' / `Section N'), and stripped-prefix
caption text."
  (let ((cands nil))
    ;; em-dash variants (pandoc roundtrip may swap one for the other)
    (push (replace-regexp-in-string "---" "—" range t t) cands)
    (push (replace-regexp-in-string "—" "---" range t t) cands)
    ;; cross-refs: substitute "Figure N" / "Table N" / "Section N" -> [cite:@type:label]
    (cl-flet
        ((sub (type-key prefix patterns)
           (dolist (entry (plist-get xref-map type-key))
             (let* ((label (car entry))
                    (n     (cdr entry))
                    (rep   (format "[cite:@%s:%s]" prefix label)))
               (dolist (pat patterns)
                 (push (replace-regexp-in-string (format pat n) rep range t)
                       cands))))))
      (sub :fig "fig"
           '("\\b[Ff]igure[  ]+%d\\b"
             "\\b[Ff]ig\\.[  ]*%d\\b"))
      (sub :tbl "tbl"
           '("\\b[Tt]able[  ]+%d\\b"
             "\\b[Tt]bl\\.[  ]*%d\\b"))
      (sub :sec "sec"
           '("\\b[Ss]ection[  ]+%d\\b"
             "\\b[Ss]ec\\.[  ]*%d\\b")))
    ;; cross-ref STRIP: caption-style ranges like "Table 6: <text>" or
    ;; "Figure 3: <text>" — pandoc-crossref renders the prefix, canonical
    ;; has only the caption text after the colon.
    (push (replace-regexp-in-string
           "^[Tt]able[  ]+[0-9]+:[ \t ]*" "" range)
          cands)
    (push (replace-regexp-in-string
           "^[Ff]igure[  ]+[0-9]+:[ \t ]*" "" range)
          cands)
    ;; emphasis-marker variants: ranges wrapping a phrase in markdown
    ;; emphasis (`*Importance*', `**bold**') or org emphasis (`/italic/',
    ;; `_under_', `=verb=', `~code~') won't substring-match canonical text
    ;; that uses a different marker or none at all.  Try with surrounding
    ;; markers stripped and with all internal markers stripped.
    (push (replace-regexp-in-string "\\`[*/_=~]+\\|[*/_=~]+\\'" "" range) cands)
    (push (replace-regexp-in-string "[*/_=~]" "" range) cands)
    (delete-dups (delq nil cands))))

(defun otd--reanchor-comments (merged tracked &optional xref-map)
  "Splice reviewer comments from TRACKED into MERGED.
For each `{==RANGE==}{>>CMT<<}' in TRACKED not already in MERGED, try
to anchor it by substring search in MERGED.  Tiered strategy:
exact -> mechanical variants from `otd--anchor-candidates' (em-dash,
cross-ref) -> case-insensitive (for capitalization mismatches like a
lowercase domain name in the tracked source vs Title Case in canonical).
On each strategy, splice only if the candidate appears exactly once
and is not already inside an existing annotation.  Returns
(NEW-MERGED . N-GRAFTED)."
  (let* ((doc-order (otd--scan-annotations tracked))
         ;; Per (stripped-range) document-order index of each annotation.
         ;; Used by tier 5 to pair tracked-Kth occurrence with canonical-Kth
         ;; when a range matches multiple positions in canonical.
         (ordinal (let ((seen (make-hash-table :test 'equal))
                        (map  (make-hash-table :test 'eq)))
                    (dolist (p doc-order)
                      (let* ((r (string-trim (otd--strip-criticmarkup (car p))))
                             (k (gethash r seen 0)))
                        (puthash p k map)
                        (puthash r (1+ k) seen)))
                    map))
         ;; Splice longest range first so an outer comment's substring is
         ;; placed before a shorter inner overlap can break the outer text
         ;; by inserting `{==…==}{>>…<<}' markers in the middle of it.
         (annotations (sort (copy-sequence doc-order)
                            (lambda (a b)
                              (> (length (otd--strip-criticmarkup (car a)))
                                 (length (otd--strip-criticmarkup (car b)))))))
         (n 0)
         (out merged))
    (dolist (pair annotations)
      (let* ((range-raw (car pair))
             (range (string-trim (otd--strip-criticmarkup range-raw)))
             (cmt   (cdr pair))
             (wrap0 (format "{==%s==}{>>%s<<}" range cmt))
             (rlen  (length range))
             (placed nil))
        (unless (or (< rlen 3)
                    (string-search wrap0 out)
                    ;; either variant of the wrap (em-dash flipped) already present
                    (string-search (replace-regexp-in-string "---" "—" wrap0 t t) out))
          ;; Tier 1: exact
          (let ((p (otd--anchor-unique-position range out)))
            (when p
              (setq out (concat (substring out 0 p) wrap0
                                (substring out (+ p rlen))))
              (cl-incf n)
              (setq placed t)))
          ;; Tier 2: mechanical variants
          (unless placed
            (dolist (cand (otd--anchor-candidates range xref-map))
              (unless placed
                (let* ((clen (length cand))
                       (p (otd--anchor-unique-position cand out)))
                  (when p
                    (let ((wrap (format "{==%s==}{>>%s<<}" cand cmt)))
                      (setq out (concat (substring out 0 p) wrap
                                        (substring out (+ p clen))))
                      (cl-incf n)
                      (setq placed t)))))))
          ;; Tier 3: case-insensitive (requires a longer range to avoid false hits)
          (unless placed
            (when (>= rlen 8)
              (let* ((lc-range (downcase range))
                     (lc-out   (downcase out))
                     (p (otd--anchor-unique-position lc-range lc-out 8)))
                (when p
                  ;; Splice using the actual canonical-case substring at that pos.
                  (let* ((canon (substring out p (+ p rlen)))
                         (wrap (format "{==%s==}{>>%s<<}" canon cmt)))
                    (setq out (concat (substring out 0 p) wrap
                                      (substring out (+ p rlen))))
                    (cl-incf n)
                    (setq placed t))))))
          ;; Tier 4: progressive prefix truncation, applied to the raw range
          ;; AND to each variant from `otd--anchor-candidates' (so e.g.
          ;; "Table 6: <long caption>" can be reduced to a unique prefix of
          ;; the stripped "<long caption>" form).  When canonical has been
          ;; reworded near the tail of the commented sentence, the head
          ;; still matches uniquely.  Tries 75/60/45/30/20/15 % prefixes;
          ;; floor 16 chars to limit false-positive risk.
          (unless placed
            (when (>= rlen 24)
              (let ((all-cands (cons range
                                     (otd--anchor-candidates range xref-map))))
                (cl-loop
                 for cand in all-cands
                 while (not placed)
                 do
                 (let ((clen (length cand)))
                   (when (>= clen 24)
                     (cl-loop
                      for frac in '(0.75 0.6 0.45 0.3 0.2 0.15)
                      for tlen = (max 16 (floor (* frac clen)))
                      while (not placed)
                      do
                      (let* ((trunc (substring cand 0 (min clen tlen)))
                             (p (otd--anchor-unique-position trunc out 16)))
                        (when p
                          (let ((wrap (format "{==%s==}{>>%s<<}"
                                              trunc cmt)))
                            (setq out (concat (substring out 0 p) wrap
                                              (substring out (+ p tlen))))
                            (cl-incf n)
                            (setq placed t)))))))))))
          ;; Tier 5: ordinal-position resolution for multi-match.  When a
          ;; range appears multiple times in canonical and the same range
          ;; is anchored by multiple tracked annotations (or only one but
          ;; canonical has few occurrences), pair the K-th tracked
          ;; occurrence with the K-th canonical occurrence in document
          ;; order.  Floor 4 chars; skip if canonical has > 5 occurrences
          ;; (too ambiguous to anchor safely).
          (unless placed
            (when (>= rlen 4)
              (let ((positions (otd--anchor-all-positions range out 4))
                    (k (gethash pair ordinal)))
                (when (and positions
                           k
                           (<= (length positions) 5)
                           (< k (length positions)))
                  (let* ((p (nth k positions))
                         (wrap (format "{==%s==}{>>%s<<}" range cmt)))
                    (setq out (concat (substring out 0 p) wrap
                                      (substring out (+ p rlen))))
                    (cl-incf n)
                    (setq placed t)))))))))
    (cons out n)))

(defun otd--align-paragraphs (canon-paras tracked-paras)
  "Two-pointer align CANON-PARAS to TRACKED-PARAS by normalized fingerprint.
Returns a vector of length (length CANON-PARAS); index I holds the
tracked-paragraph string aligned to canonical paragraph I, or nil when
no aligned tracked paragraph exists (canonical extends past tracked).

Why positional + fingerprint, not fingerprint-only: a coauthor who
ACCEPTS tracked changes in Word strips the w:ins/w:del runs, so the
imported tracked paragraph has the new wording but no CriticMarkup -
its fingerprint no longer matches canonical, and a fingerprint-only
lookup falls back to canonical (losing the accepted edits).
Two-pointer alignment keeps canonical and tracked in step paragraph
by paragraph, lookahead handles inserted/deleted paragraphs on either
side, and the positional fallback inside a matched run carries
accepted edits through into the merge."
  (let* ((nc (length canon-paras))
         (nt (length tracked-paras))
         (align (make-vector nc nil))
         (tracked-vec (vconcat tracked-paras))
         (canon-fp-vec (vconcat (mapcar #'otd--normalize-for-match canon-paras)))
         (tracked-fp-vec (vconcat (mapcar #'otd--normalize-for-match tracked-paras)))
         ;; Lookahead window: how far we will scan ahead on either side
         ;; to find a matching fingerprint before falling back to a
         ;; positional assumption.  Default 12 covers the common pain
         ;; point where pandoc unwraps a frontmatter author block into
         ;; ~6 separate paragraphs that the canonical (org source) keeps
         ;; in a single `#+begin_export ... #+end_export' block; without
         ;; the wide lookahead those extra tracked paragraphs would
         ;; drift the abstract body into the Results section.
         (la 12)
         (i 0) (j 0))
    (cl-labels
        ;; Prefix-equality on normalized fingerprints, not strict
        ;; equality.  Coauthors who accepted a single edit (e.g. one
        ;; letter `characterising' -> `characterizing') would otherwise
        ;; produce a fingerprint mismatch on every match attempt, so
        ;; the alignment falls back to the positional case and the
        ;; abstract body lands wherever the author block sat in the
        ;; tracked sequence.  40 chars is enough prose to be unique
        ;; across all body paragraphs in a typical manuscript yet
        ;; tolerates the long tail of small edits Chandra usually makes.
        ((fp= (a b)
           (let* ((pref 40)
                  (na (length a)) (nb (length b))
                  (mn (min pref na nb)))
             (and (>= mn 8)
                  (string= (substring a 0 mn)
                           (substring b 0 mn)))))
         (scan (vec start fp end)
           (let ((k start) (found nil))
             (while (and (not found) (< k end))
               (when (fp= (aref vec k) fp) (setq found k))
               (cl-incf k))
             found)))
      (while (and (< i nc) (< j nt))
        (let* ((cfp (aref canon-fp-vec i))
               (tfp (aref tracked-fp-vec j))
               (la-end-t (min nt (+ j la)))
               (la-end-c (min nc (+ i la))))
          (cond
           ;; Prefix-equality match: paragraphs correspond, advance.
           ((fp= cfp tfp)
            (aset align i (aref tracked-vec j))
            (cl-incf i) (cl-incf j))
           ;; Tracked has extra paragraphs (unwrapped frontmatter,
           ;; bibliography section, etc.): scan ahead for the canonical
           ;; fingerprint and skip the gap.
           ((scan tracked-fp-vec (1+ j) cfp la-end-t)
            (setq j (scan tracked-fp-vec (1+ j) cfp la-end-t)))
           ;; Canonical has extra paragraphs (deleted in tracked): scan
           ;; ahead in canonical for the tracked fingerprint.
           ((scan canon-fp-vec (1+ i) tfp la-end-c)
            (setq i (scan canon-fp-vec (1+ i) tfp la-end-c)))
           ;; Neither side finds a match in window: assume they
           ;; correspond and the tracked paragraph carries accepted
           ;; edits.  Use tracked.
           (t
            (aset align i (aref tracked-vec j))
            (cl-incf i) (cl-incf j))))))
    align))

(defun otd--merge-content (canonical tracked)
  "Return merged content combining CANONICAL (cite keys, metadata,
property drawers, etc.) with CriticMarkup tokens AND accepted edits
from TRACKED.

Walks CANONICAL line-by-line, preserving its full structure.  Body
paragraphs are aligned to tracked paragraphs by positional + fingerprint
alignment (see `otd--align-paragraphs'); the tracked text takes the
canonical paragraph's place when the alignment is non-nil, with
citation and cross-ref forms back-substituted to `[cite:@key]'.  After
the walk, any tracked comments whose anchor was missed by paragraph
alignment are re-anchored by `otd--reanchor-comments' (substring
splice for exactly-once matches).

Tracked paragraphs that align to no canonical paragraph (e.g. the docx
bibliography section, which canonical does not contain because it is
regenerated by citeproc on export) are dropped.

`#+begin_src'..`#+end_src' blocks (bibtex/CSL/reference-doc assets
tangled out on export) are passed through from CANONICAL verbatim,
same as property drawers and export blocks: Word only ever shows the
reader a citeproc-rendered numbered reference list, never the
underlying bibtex, so there is no tracked-side text that could ever
fingerprint-match a bibtex entry. Without this passthrough those
numbered-reference-list paragraphs get positionally misaligned onto
the bibtex/CSL/reference-doc paragraphs one-for-one, corrupting all
three src blocks.

Returns (MERGED-STRING . PLIST) where PLIST has keys
:merged-paragraphs, :cm-paragraphs, :cite-keys, :reanchored-comments."
  (let* ((cite-map (otd--build-cite-map canonical))
         (xref-map (otd--build-xref-map canonical))
         (canon-paras (otd--split-paragraphs canonical))
         (tracked-paras (otd--split-paragraphs tracked))
         (canon-body-paras (cl-remove-if
                            (lambda (p) (string-match-p "^\\*+[ \t]" p))
                            canon-paras))
         (tracked-body-paras (cl-remove-if
                              (lambda (p) (string-match-p "^\\*+[ \t]" p))
                              tracked-paras))
         (align (otd--align-paragraphs canon-body-paras tracked-body-paras))
         (n-cm-tracked (cl-count-if #'otd--has-criticmarkup-p tracked-body-paras))
         (n-merged 0)
         (body-idx 0)
         (output nil)
         (current nil)
         (in-drawer nil)
         (in-export nil)
         (in-src nil))
    (cl-labels
        ((flush ()
           (when current
             (let* ((para (mapconcat #'identity (nreverse current) "\n"))
                    (replacement (and (< body-idx (length align))
                                      (aref align body-idx))))
               (push (if replacement
                         (progn (cl-incf n-merged)
                                (otd--substitute-xrefs
                                 (otd--substitute-cite-keys replacement
                                                            cite-map)
                                 xref-map))
                       para)
                     output)
               (cl-incf body-idx))
             (setq current nil))))
      (dolist (line (split-string canonical "\n"))
        (cond
         ((string-match-p "^[ \t]*:PROPERTIES:[ \t]*$" line)
          (flush) (setq in-drawer t) (push line output))
         ((and in-drawer (string-match-p "^[ \t]*:END:[ \t]*$" line))
          (setq in-drawer nil) (push line output))
         (in-drawer (push line output))
         ;; Case-insensitive to match pandoc's lowercase `#+begin_export'.
         ((let ((case-fold-search t))
            (string-match-p "^#\\+BEGIN_EXPORT" line))
          (flush) (setq in-export t) (push line output))
         ((and in-export
               (let ((case-fold-search t))
                 (string-match-p "^#\\+END_EXPORT" line)))
          (setq in-export nil) (push line output))
         (in-export (push line output))
         ;; `#+begin_src'/`#+end_src' asset blocks: verbatim passthrough,
         ;; same rationale as export blocks (see docstring).
         ((let ((case-fold-search t))
            (string-match-p "^#\\+BEGIN_SRC" line))
          (flush) (setq in-src t) (push line output))
         ((and in-src
               (let ((case-fold-search t))
                 (string-match-p "^#\\+END_SRC" line)))
          (setq in-src nil) (push line output))
         (in-src (push line output))
         ((or (string-match-p "^#\\+" line)
              (string-match-p "^\\*+[ \t]" line)
              (string-empty-p (string-trim line)))
          (flush) (push line output))
         (t (push line current))))
      (flush))
    (let* ((merged-str (mapconcat #'identity (nreverse output) "\n"))
           (re (otd--reanchor-comments merged-str tracked xref-map)))
      (cons (car re)
            (list :merged-paragraphs n-merged
                  :cm-paragraphs n-cm-tracked
                  :cite-keys (length cite-map)
                  :reanchored-comments (cdr re))))))

;;;; --- export (CriticMarkup -> native Word tracked changes & comments) -

;;;###autoload
(defun otd-export (org-file &optional output)
  "Export ORG-FILE to docx, converting CriticMarkup back to native
Word tracked changes and comments.
With prefix arg, prompt for OUTPUT path; otherwise the docx goes
alongside ORG-FILE with the .docx extension.

Pipeline:
  1. Replace each CriticMarkup token in a copy of the org file with a
     unique alphanumeric placeholder (so pandoc cannot mangle it).
  2. pandoc org -> markdown on the placeholder-laden copy.
  3. In the markdown, swap each placeholder for a pandoc native span:
       {++text++}            -> [text]{.insertion}
       {--text--}            -> [text]{.deletion}
       {~~old~>new~~}        -> [old]{.deletion}[new]{.insertion}
       {==range==}{>>c<<}    -> [c]{.comment-start id=N}range[]{.comment-end id=N}
       {==range==} (orphan)  -> range  (highlights are dropped)
       {>>c<<} (orphan)      -> [c]{.comment-start id=N}[]{.comment-end id=N}
  4. pandoc markdown -> docx.  Pandoc emits real w:ins / w:del runs and
     a real word/comments.xml from those spans.
Author defaults to `otd-export-author'."
  (interactive
   (list (or (and (eq major-mode 'org-mode) (buffer-file-name))
             (read-file-name "Org file: " nil nil t nil
                             (lambda (f) (or (file-directory-p f)
                                             (string-match-p "\\.org\\'" f)))))
         (when current-prefix-arg (read-file-name "Output .docx: "))))
  (unless (file-exists-p org-file) (user-error "File not found: %s" org-file))
  (let* ((dir     (file-name-directory (expand-file-name org-file)))
         (base    (file-name-sans-extension (file-name-nondirectory org-file)))
         (out     (or output (expand-file-name (concat base ".docx") dir)))
         (org-tmp (make-temp-file "otd-out-" nil ".org"))
         (md-tmp  (make-temp-file "otd-out-" nil ".md"))
         (alist   nil)
         (counts  nil)
         (done-ids nil))
    (unwind-protect
        (progn
          ;; Stage 0: regenerate any embedded bibliography/CSL/reference-doc
          ;; resources from their canonical `:tangle'/`:tangle-binary' source
          ;; blocks in ORG-FILE, so the export always reflects whatever is
          ;; embedded in the document itself rather than a possibly-stale or
          ;; possibly-missing external file.  See `otd-tangle-before-export'.
          (when otd-tangle-before-export
            (otd--tangle-embedded-resources org-file dir))
          ;; Stage 1: fixups + CriticMarkup -> placeholders
          (with-temp-buffer
            (insert-file-contents org-file)
            ;; Expand `#+AFFIL:'/`#+AUTHOR_LIST:'/`#+AUTHOR_GROUP:' header
            ;; lines (if present) into the flat markdown author block
            ;; pandoc actually needs.  See `otd--expand-author-block'.
            (otd--expand-author-block)
            (otd--postfix-fixups)
            ;; Give bare relative image links a `file:' prefix so pandoc's org
            ;; reader emits them as Images (a bare relative link is otherwise
            ;; dropped as a spurious link).  The path stays relative; the
            ;; md->docx stage's `--resource-path=dir' resolves and embeds it.
            (otd--fileify-image-links)
            (let ((res (otd--mark-criticmarkup)))
              (setq alist    (nth 0 res)
                    counts   (nth 1 res)
                    done-ids (nth 2 res)))
            (write-region (point-min) (point-max) org-tmp nil 'silent))
          ;; Stage 2: org -> markdown (placeholders pass through unchanged).
          ;; `--standalone' is required to preserve `#+title:', `#+author:'
          ;; etc. as YAML frontmatter; without it pandoc's markdown writer
          ;; drops document metadata, which then never reaches docx.
          ;; `--resource-path=dir' is required because ORG-TMP lives in a
          ;; temp directory, not alongside ORG-FILE: any relative image
          ;; link (e.g. `[[media/fig1.png]]') is only valid relative to
          ;; ORG-FILE's own directory, so without this pandoc's org reader
          ;; can't resolve it as a file and silently degrades the link to
          ;; a `.spurious-link' text span instead of a real image --
          ;; which in turn means pandoc-crossref never sees a figure to
          ;; number, and every `[@fig:...]' reference to it comes back
          ;; "Undefined cross-reference".
          (with-temp-buffer
            (let ((exit (apply #'call-process otd-pandoc-program nil t nil
                               (list "-f" "org" "-t" "markdown"
                                     "--standalone" "--wrap=none"
                                     (concat "--resource-path=" dir)
                                     org-tmp "-o" md-tmp))))
              (unless (zerop exit)
                (error "pandoc org->md exit %s:\n%s" exit (buffer-string)))))
          ;; Stage 3a: decode `file://' image URIs so pandoc's docx
          ;; writer can find and embed the referenced media.
          (otd--decode-image-paths md-tmp)
          ;; Stage 3b: placeholders -> pandoc native spans
          (otd--unmark-criticmarkup md-tmp alist)
          ;; Stage 4: markdown -> docx (real comments + tracked changes).
          ;; User-supplied `#+PANDOC_OPTIONS:' and `#+bibliography:'
          ;; headers are forwarded so citeproc, CSL, reference-doc, etc.
          ;; behave the same as M-x org-pandoc-export-to-docx would.
          ;; `--resource-path=dir' again, for the same reason as stage 2:
          ;; MD-TMP lives in a temp directory, so any relative image path
          ;; pandoc's org reader left unresolved (or re-relativized) must
          ;; still be resolved against ORG-FILE's directory when the docx
          ;; writer goes to actually embed the image bytes.
          (with-temp-buffer
            (let* ((user-args (otd--read-pandoc-options org-file))
                   (pandoc-args (append (list "-f" "markdown" "-t" "docx"
                                              "--wrap=none"
                                              (concat "--resource-path=" dir)
                                              md-tmp "-o" out)
                                        user-args))
                   (exit (apply #'call-process otd-pandoc-program nil t nil
                                pandoc-args)))
              (unless (zerop exit)
                (error "pandoc md->docx exit %s:\n%s" exit (buffer-string)))))
          ;; Stage 5: embed the canonical org source as a customXml part
          ;; so a future `otd-import' on the reviewed docx can recover
          ;; cite keys / metadata that the docx body cannot represent.
          (when otd-embed-source
            (otd--embed-org-source out org-file))
          ;; Stage 6: rewrite table borders directly (see `otd-fix-table-borders').
          (when otd-fix-table-borders
            (with-temp-buffer
              (let ((exit (call-process otd-python-program nil t nil
                                        otd-table-borders-script out)))
                (unless (zerop exit)
                  (error "otd-table-borders.py exit %s:\n%s" exit (buffer-string))))))
          ;; Stage 7: mark `[DONE]' comments Resolved in Word's reviewing
          ;; pane (see `otd-resolve-done-comments').  Always run this even
          ;; when DONE-IDS is empty: it still stamps every comment with a
          ;; w14:paraId, which pandoc never emits and a future otd-export
          ;; run needs already present to append further resolved-comment
          ;; entries idempotently.
          (when otd-resolve-done-comments
            (with-temp-buffer
              (let ((exit (call-process otd-python-program nil t nil
                                        otd-resolve-comments-script out
                                        (string-join done-ids ","))))
                (unless (zerop exit)
                  (error "otd-resolve-comments.py exit %s:\n%s" exit (buffer-string))))))
          (message "Exported %s -> %s  (++%d --%d ~~%d ==%d >>%d)"
                   (file-name-nondirectory org-file)
                   (file-name-nondirectory out)
                   (plist-get counts :ins) (plist-get counts :del)
                   (plist-get counts :sub) (plist-get counts :hi)
                   (plist-get counts :cmt)))
      (when (file-exists-p org-tmp) (delete-file org-tmp))
      (when (file-exists-p md-tmp)  (delete-file md-tmp)))
    out))

(defun otd--strip-done-prefix (ctext)
  "If CTEXT (a comment body, author-prefix already stripped) begins
with a `[DONE]' marker, return (T . REST-WITHOUT-MARKER); otherwise
return (nil . CTEXT) unchanged.  Used so a `[DONE]'-tagged comment
exports to Word as an actually-resolved comment (see
`otd--resolve-comments-script') instead of carrying the literal
`[DONE] ' text into the comment body."
  (if (string-match "\\`\\[DONE\\][ \t]*" ctext)
      (cons t (substring ctext (match-end 0)))
    (cons nil ctext)))

(defun otd--mark-criticmarkup ()
  "Replace CriticMarkup tokens in current buffer with sentinel placeholders.
Return (ALIST COUNTS DONE-IDS) where ALIST maps each placeholder string
to the pandoc-span replacement to splice in once we are out of the org
reader, COUNTS is a plist of token counts, and DONE-IDS is a list of
pandoc comment-id strings (matching the `id=\"N\"' in the comment-start/
comment-end spans) whose CriticMarkup body carried a `[DONE]' marker --
consumed by `otd--resolve-comments-script' to mark those comments
resolved in the generated docx's `commentsExtended.xml'.

Runs the full 6-pass sweep to a fixpoint rather than once. When
reviewers comment on the same/overlapping span, CriticMarkup nests,
e.g. `{=={====}{>>c1<<}==}{>>c2<<}'. A single sweep's highlight+comment
regex used `.*?' (any character, brace-blind), so on a nested span it
paired the OUTER `{==' with the FIRST `==}{>>' it could find -- which
sits inside the inner `{====}' token -- swallowing the inner token's
own `{==' into the outer match's captured range and leaking it as
literal comment-anchored text in the exported docx body, while the
now-orphaned outer `==}' had no token left to belong to and leaked too
(this is what showed up as stray `{==' / `==}' fragments in Word).
`[^{}]*?' makes the range brace-aware so a sweep only matches complete,
non-nested tokens; looping the whole 6-pass sweep until nothing changes
then lets each pass peel one nesting level (converting an inner token
to a plain-text placeholder) so the next pass's sweep can cleanly match
what is now an unnested outer token, however deep the original nesting."
  (let ((alist nil)
        (id 0)
        (n-ins 0) (n-del 0) (n-sub 0) (n-hi 0) (n-cmt 0)
        (done-ids nil)
        (author (or otd-export-author ""))
        (date   (format-time-string "%Y-%m-%dT%H:%M:%SZ" nil t))
        (changed t))
    (cl-flet
        ((stash (replacement)
           (cl-incf id)
           (let ((tag (format "OTDPHX%05dXEND" id)))
             (push (cons tag replacement) alist)
             tag)))
      (while changed
        (setq changed nil)
        ;; {==range==}{>>comment<<} -> comment-start..comment-end pair
        ;; If COMMENT begins with `[Author Name] ' (the prefix `otd--rewrite-spans'
        ;; encodes on import), peel it off and use it as the docx author so the
        ;; original reviewer name round-trips back to Word.
        (goto-char (point-min))
        (while (re-search-forward "{==\\([^{}]*?\\)==}{>>\\([^<]*\\)<<}" nil t)
          (setq changed t)
          (let* ((range (match-string 1))
                 (ctext-raw (match-string 2))
                 ;; save-match-data: `otd--parse-comment-author' calls
                 ;; `string-match' which clobbers the outer regex's match data
                 ;; before `replace-match' uses it.
                 (parsed (save-match-data
                           (otd--parse-comment-author ctext-raw)))
                 (use-author (or (car parsed) author))
                 (done-parsed (save-match-data (otd--strip-done-prefix (cdr parsed))))
                 (ctext (cdr done-parsed))
                 (n     (1+ id)))
            (when (car done-parsed) (push (number-to-string n) done-ids))
            (replace-match
             (stash (format "[%s]{.comment-start id=\"%d\" author=\"%s\" date=\"%s\"}%s[]{.comment-end id=\"%d\"}"
                            ctext n use-author date range n))
             t t))
          (cl-incf n-cmt) (cl-incf n-hi))
        ;; {~~old~>new~~} -> [old]{.deletion}[new]{.insertion}
        (goto-char (point-min))
        (while (re-search-forward "{~~\\([^~{}]*?\\)~>\\([^~{}]*?\\)~~}" nil t)
          (setq changed t)
          (let ((old (match-string 1)) (new (match-string 2)))
            (replace-match
             (stash (format "[%s]{.deletion author=\"%s\" date=\"%s\"}[%s]{.insertion author=\"%s\" date=\"%s\"}"
                            old author date new author date))
             t t))
          (cl-incf n-sub))
        ;; {++text++} -> [text]{.insertion}
        (goto-char (point-min))
        (while (re-search-forward "{\\+\\+\\(?:\\[[^]]+\\]\\)?\\([^{}]*?\\)\\+\\+}" nil t)
          (setq changed t)
          (let ((s (match-string 1)))
            (replace-match
             (stash (format "[%s]{.insertion author=\"%s\" date=\"%s\"}"
                            s author date))
             t t))
          (cl-incf n-ins))
        ;; {--text--} -> [text]{.deletion}
        (goto-char (point-min))
        (while (re-search-forward "{--\\(?:\\[[^]]+\\]\\)?\\([^{}]*?\\)--}" nil t)
          (setq changed t)
          (let ((s (match-string 1)))
            (replace-match
             (stash (format "[%s]{.deletion author=\"%s\" date=\"%s\"}"
                            s author date))
             t t))
          (cl-incf n-del))
        ;; Orphan {==range==} (no following comment): drop markers, keep text.
        (goto-char (point-min))
        (while (re-search-forward "{==\\([^{}]*?\\)==}" nil t)
          (setq changed t)
          (replace-match (stash (match-string 1)) t t)
          (cl-incf n-hi))
        ;; Orphan {>>comment<<}: zero-range comment so it still appears in Word.
        (goto-char (point-min))
        (while (re-search-forward "{>>\\([^<]*\\)<<}" nil t)
          (setq changed t)
          (let* ((ctext-raw (match-string 1))
                 (parsed (save-match-data
                           (otd--parse-comment-author ctext-raw)))
                 (use-author (or (car parsed) author))
                 (done-parsed (save-match-data (otd--strip-done-prefix (cdr parsed))))
                 (ctext (cdr done-parsed))
                 (n     (1+ id)))
            (when (car done-parsed) (push (number-to-string n) done-ids))
            (replace-match
             (stash (format "[%s]{.comment-start id=\"%d\" author=\"%s\" date=\"%s\"}[]{.comment-end id=\"%d\"}"
                            ctext n use-author date n))
             t t))
          (cl-incf n-cmt))))
    (list alist
          (list :ins n-ins :del n-del :sub n-sub :hi n-hi :cmt n-cmt)
          done-ids)))

(defun otd--unmark-criticmarkup (md-file alist)
  "Replace each ALIST placeholder with its pandoc-span value in MD-FILE."
  (with-temp-buffer
    (insert-file-contents md-file)
    (dolist (pair alist)
      (goto-char (point-min))
      (while (search-forward (car pair) nil t)
        (replace-match (cdr pair) t t)))
    (write-region (point-min) (point-max) md-file nil 'silent)))

;;;; --- auto-import on find-file ---------------------------------------

(defcustom otd-auto-import-reuse t
  "When non-nil, `otd-auto-import-mode' reuses an existing import.
If the `otd-output-suffix' .org that `otd-import' produces for a docx
already exists and is at least as new as the docx, opening the docx
visits that .org instead of re-running pandoc -- which would overwrite
any edits made in the imported file.  When the docx is newer than the
existing .org (a freshly returned review), it is re-imported."
  :type 'boolean :group 'org-tracked-docx)

(defun otd--auto-import-target (docx)
  "Return the .org path `otd-import' writes for DOCX (no conversion)."
  (let ((dir  (file-name-directory (expand-file-name docx)))
        (base (file-name-sans-extension (file-name-nondirectory docx))))
    (expand-file-name (concat base otd-output-suffix ".org") dir)))

(defun otd--auto-import-noselect (orig filename &rest args)
  "Around-advice for `find-file-noselect': import docx via `otd-import'.
When FILENAME is an existing `*.docx', convert it and return the buffer
visiting the imported .org instead of a buffer of raw docx bytes.  Any
error falls back to ORIG so the file still opens."
  (let ((file (and (stringp filename) (expand-file-name filename))))
    (if (and file
             (string-match-p "\\.docx\\'" file)
             (file-regular-p file))
        (condition-case err
            (let* ((out   (otd--auto-import-target file))
                   (reuse (and otd-auto-import-reuse
                               (file-exists-p out)
                               (not (file-newer-than-file-p file out)))))
              ;; `otd-import' already visits OUT; calling ORIG on it just
              ;; returns that live buffer.  In the reuse branch ORIG opens
              ;; the previously imported .org directly.
              (apply orig (if reuse out (otd-import file)) args))
          (error
           (message "otd auto-import failed (%s); opening docx raw"
                    (error-message-string err))
           (apply orig filename args)))
      (apply orig filename args))))

;;;###autoload
(define-minor-mode otd-auto-import-mode
  "Global minor mode: opening a Word .docx auto-imports it with `otd-import'.
When enabled, visiting any `*.docx' file (via \\[find-file], dired, a
desktop restore, etc.) runs the tracked-changes import pipeline and
shows the resulting CriticMarkup org buffer instead of raw binary.
See `otd-auto-import-reuse' for re-import vs. reuse behaviour."
  :global t :group 'org-tracked-docx
  (if otd-auto-import-mode
      (advice-add 'find-file-noselect :around #'otd--auto-import-noselect)
    (advice-remove 'find-file-noselect #'otd--auto-import-noselect)))

(provide 'org-tracked-docx)

;; Load the JSON-AST import backend (the default for `otd-import') when it
;; sits alongside this file.  Soft by design: if it is absent or fails to
;; load, `otd-import' transparently falls back to the markdown backend.
;; This runs after `provide' above, so the backend's own
;; `(require 'org-tracked-docx)' is already satisfied (no load cycle).
(condition-case nil
    (let* ((here (or load-file-name buffer-file-name))
           (sib  (and here (expand-file-name "org-tracked-docx-json"
                                             (file-name-directory here)))))
      (when (and sib (or (file-exists-p (concat sib ".elc"))
                         (file-exists-p (concat sib ".el"))))
        (load sib nil t)))
  (error nil))

;;; org-tracked-docx.el ends here
