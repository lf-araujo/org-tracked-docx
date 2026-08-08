;;; org-tracked-docx-json.el --- JSON-AST import backend -*- lexical-binding: t; -*-

;; A more robust import path for org-tracked-docx.
;;
;; The original importer (`otd--import-markdown') runs
;;
;;     docx --pandoc--> markdown --regex-rewrite--> CriticMarkup-md
;;          --stash sentinels--> pandoc md->org --unstash--> org
;;
;; and converts tracked changes to CriticMarkup with ~27 textual regex
;; operations over the *markdown* intermediate (`otd--rewrite-spans',
;; `otd--stash-criticmarkup').  That layer is where the recurring drop/leak/
;; corruption bugs live (comment text containing `]', nested/overlapping
;; comments, `~>' reparsed as strikethrough, escaping leaks like `\[').
;;
;; This backend instead does
;;
;;     docx --pandoc--> JSON AST --structured walk--> AST'  --pandoc--> org
;;
;; replacing tracked-change Span nodes with literal CriticMarkup delimiter
;; `Str' nodes *inside the AST*, then letting pandoc render json->org.  No
;; markdown intermediate, no span-matching regexes, no sentinel stashing.
;;
;; The walk is generic: `otd--j-walk' descends the whole AST structurally and
;; runs the sequence-aware `otd--j-map-inlines' on every inline list it finds,
;; wherever it occurs -- top-level paragraphs, headings, list items, block
;; quotes, table cells, footnote bodies.  There is no per-block-type code and
;; no index juggling, so new pandoc block shapes are handled automatically.
;;
;; It reuses the stock helpers for everything downstream of the org file
;; (`otd--postfix-org', embedded-source merge) and for the resolved-comment
;; OOXML read (`otd--extract-done-comment-ids'), so the review UX, accept/
;; reject and export contract are unchanged.
;;
;; Behavioural note vs the markdown backend: a Word table cell containing
;; several paragraphs renders as a native multi-row org table (pandoc's
;; json->org writer splits it) rather than a raw `#+BEGIN_EXPORT markdown'
;; grid block.  This keeps the org buffer free of embedded raw markdown.

;;; Code:

(require 'org-tracked-docx)
(require 'json)
(require 'cl-lib)

;;;; --- dynamic context for the walk ---------------------------------------
;; Bound around a transform so the generic walker's helpers stay single-arg.

(defvar otd--j-done-set nil
  "Hash of resolved comment ids (string -> t) for the current import.")
(defvar otd--j-counts nil
  "Hash tallying markers emitted during the current import.")

(defun otd--j-bump (key)
  "Increment tally KEY in `otd--j-counts'."
  (when otd--j-counts
    (puthash key (1+ (gethash key otd--j-counts 0)) otd--j-counts)))

;;;; --- tiny AST accessors -------------------------------------------------
;;
;; Nodes are alists with symbol keys `t' and `c' (from `json-parse' with
;; :object-type 'alist).  Arrays are vectors (:array-type 'array).  A Span's
;; `c' is [ [id classes kvs] [inlines] ]; classes/kvs are vectors.

(defconst otd--j-inline-tags
  '("Str" "Emph" "Underline" "Strong" "Strikeout" "Superscript" "Subscript"
    "SmallCaps" "Quoted" "Cite" "Code" "Space" "SoftBreak" "LineBreak"
    "Math" "RawInline" "Link" "Image" "Note" "Span")
  "Pandoc Inline constructor names, used to recognise an inline list.")

(defsubst otd--j-tag (node) (and (consp node) (alist-get 't node)))
(defsubst otd--j-c   (node) (alist-get 'c node))

(defun otd--j-str (s)
  "A pandoc `Str' inline node carrying literal text S."
  (list (cons 't "Str") (cons 'c s)))

(defun otd--j-span-attr    (node) (aref (otd--j-c node) 0))
(defun otd--j-span-inlines (node) (aref (otd--j-c node) 1))

(defun otd--j-attr-kv (attr key)
  "Value of KEY in pandoc attr triple ATTR's key/value vector, or nil."
  (let ((res nil))
    (cl-loop for pair across (aref attr 2)
             when (equal (aref pair 0) key) do (setq res (aref pair 1)))
    res))

(defun otd--j-span-classes (node)
  "List of class strings on Span NODE."
  (append (aref (otd--j-span-attr node) 1) nil))

(defun otd--j-change-kind (node)
  "Return the tracked-change kind of NODE as a symbol, or nil.
One of `insertion' `deletion' `paragraph-insertion'
`paragraph-deletion' `comment-start' `comment-end'."
  (when (equal (otd--j-tag node) "Span")
    (let ((classes (otd--j-span-classes node)))
      (cond ((member "insertion" classes)           'insertion)
            ((member "deletion" classes)             'deletion)
            ((member "paragraph-insertion" classes)  'paragraph-insertion)
            ((member "paragraph-deletion" classes)   'paragraph-deletion)
            ((member "comment-start" classes)        'comment-start)
            ((member "comment-end" classes)          'comment-end)))))

(defun otd--j-space-p (node)
  "Non-nil if NODE is inter-word whitespace."
  (member (otd--j-tag node) '("Space" "SoftBreak")))

(defun otd--j-anchor-span-p (node)
  "Non-nil if NODE is a bookmark anchor Span (e.g. bibliography item)."
  (and (equal (otd--j-tag node) "Span")
       (member "anchor" (otd--j-span-classes node))))

;;;; --- comment body -> plain text -----------------------------------------

(defun otd--j-inlines->text (inlines)
  "Flatten INLINES (a list) to plain text, for `{>>..<<}' bodies and
heading-length heuristics."
  (mapconcat
   (lambda (n)
     (pcase (otd--j-tag n)
       ("Str" (otd--j-c n))
       ((or "Space" "SoftBreak" "LineBreak") " ")
       ("Code" (aref (otd--j-c n) 1))
       ((or "Emph" "Strong" "Strikeout" "Superscript"
            "Subscript" "SmallCaps" "Underline")
        (otd--j-inlines->text (append (otd--j-c n) nil)))
       ((or "Quoted" "Link" "Image" "Cite" "Span")
        (otd--j-inlines->text (append (aref (otd--j-c n) 1) nil)))
       (_ "")))
   inlines ""))

(defun otd--j-comment-text (id author body)
  "CriticMarkup comment body: `[AUTHOR] [DONE] text', matching the
markdown backend.  Consults `otd--j-done-set' for the resolved flag."
  (let* ((txt  (otd--j-inlines->text body))
         (done (and otd--j-done-set (gethash id otd--j-done-set))))
    (cond
     ((and author (not (string-empty-p author)) done)
      (format "[%s] [DONE] %s" author txt))
     ((and author (not (string-empty-p author)))
      (format "[%s] %s" author txt))
     (done (format "[DONE] %s" txt))
     (t txt))))

;;;; --- the generic walk ---------------------------------------------------

(defun otd--j-inline-list-p (vec)
  "Non-nil if VEC is a non-empty vector of pandoc inline nodes."
  (and (vectorp vec) (> (length vec) 0)
       (cl-every (lambda (e) (member (otd--j-tag e) otd--j-inline-tags)) vec)))

(defun otd--j-walk (x)
  "Structurally rewrite AST fragment X, running `otd--j-map-inlines'
on every inline list found anywhere inside it.  Handles arbitrary
pandoc block/inline nesting (tables, footnotes, lists, divs) with no
per-type code."
  (cond
   ((vectorp x)
    (if (otd--j-inline-list-p x)
        (vconcat (otd--j-map-inlines (append x nil)))
      (vconcat (mapcar #'otd--j-walk x))))
   ((consp x)                             ; an object/node alist -> recurse values
    (mapcar (lambda (pair) (cons (car pair) (otd--j-walk (cdr pair)))) x))
   (t x)))                                ; scalar / string / number / :null

(defun otd--j-map-inlines (inlines)
  "Rewrite INLINES (a list of inline nodes) into a new list, replacing
tracked-change Spans with literal CriticMarkup delimiter `Str' nodes.
Ordinary inline nodes are recursed into via `otd--j-walk', so nested
inline lists (link labels, footnote bodies) are still processed."
  (let ((out '()))
    (cl-flet ((emit (n) (push n out))
              (splice (lst) (dolist (n lst) (push n out))))
      (while inlines
        (let* ((node (car inlines))
               (kind (otd--j-change-kind node)))
          (cond
           ;; substitution: deletion, then optional space, then insertion
           ((and (eq kind 'deletion)
                 (let ((a (cadr inlines)) (b (caddr inlines)))
                   (or (eq (otd--j-change-kind a) 'insertion)
                       (and (otd--j-space-p a)
                            (eq (otd--j-change-kind b) 'insertion)))))
            (let* ((spaced (not (eq (otd--j-change-kind (cadr inlines)) 'insertion)))
                   (ins    (if spaced (caddr inlines) (cadr inlines)))
                   (tail   (if spaced (cdddr inlines) (cddr inlines))))
              (otd--j-bump :sub)
              (emit (otd--j-str "{~~"))
              (splice (otd--j-map-inlines (append (otd--j-span-inlines node) nil)))
              (emit (otd--j-str "~>"))
              (splice (otd--j-map-inlines (append (otd--j-span-inlines ins) nil)))
              (emit (otd--j-str "~~}"))
              (setq inlines tail)))
           ;; insertion
           ((memq kind '(insertion paragraph-insertion))
            (otd--j-bump :ins)
            (emit (otd--j-str "{++"))
            (splice (otd--j-map-inlines (append (otd--j-span-inlines node) nil)))
            (emit (otd--j-str "++}"))
            (setq inlines (cdr inlines)))
           ;; deletion
           ((memq kind '(deletion paragraph-deletion))
            (otd--j-bump :del)
            (emit (otd--j-str "{--"))
            (splice (otd--j-map-inlines (append (otd--j-span-inlines node) nil)))
            (emit (otd--j-str "--}"))
            (setq inlines (cdr inlines)))
           ;; comment: [body]{.comment-start id=N} RANGE []{.comment-end id=N}
           ((eq kind 'comment-start)
            (let* ((attr   (otd--j-span-attr node))
                   (id     (otd--j-attr-kv attr "id"))
                   (author (otd--j-attr-kv attr "author"))
                   (body   (append (otd--j-span-inlines node) nil))
                   (rest   (cdr inlines))
                   (range  '())
                   (found  nil))
              ;; collect siblings up to the matching comment-end (by id);
              ;; inner comments in RANGE are handled by the recursive call.
              (while (and rest (not found))
                (let ((rn (car rest)))
                  (if (and (eq (otd--j-change-kind rn) 'comment-end)
                           (equal (otd--j-attr-kv (otd--j-span-attr rn) "id") id))
                      (setq found t rest (cdr rest))
                    (push rn range)
                    (setq rest (cdr rest)))))
              (setq range (nreverse range))
              (otd--j-bump :hi) (otd--j-bump :cmt)
              (when (and otd--j-done-set (gethash id otd--j-done-set))
                (otd--j-bump :done))
              (emit (otd--j-str "{=="))
              (splice (otd--j-map-inlines range))
              (emit (otd--j-str "==}"))
              (emit (otd--j-str
                     (format "{>>%s<<}" (otd--j-comment-text id author body))))
              (setq inlines rest)))
           ;; stray comment-end (its start was consumed as a range) -> drop
           ((eq kind 'comment-end)
            (setq inlines (cdr inlines)))
           ;; bibliography/bookmark anchor: drop the `.anchor' class so
           ;; pandoc emits an org `<<id>>' target (keeps in-text links live).
           ((otd--j-anchor-span-p node)
            (let* ((attr    (otd--j-span-attr node))
                   (classes (vconcat (remove "anchor" (otd--j-span-classes node))))
                   (new-attr (vector (aref attr 0) classes (aref attr 2))))
              (emit (list (cons 't "Span")
                          (cons 'c (vector new-attr
                                           (vconcat (otd--j-map-inlines
                                                     (append (otd--j-span-inlines node)
                                                             nil))))))))
            (setq inlines (cdr inlines)))
           ;; ordinary inline node: recurse into its children generically
           (t
            (emit (otd--j-walk node))
            (setq inlines (cdr inlines)))))))
    (nreverse out)))

;;;; --- top-level fix-ups --------------------------------------------------

(defun otd--j-downgrade-long-headers (blocks)
  "Convert any top-level `Header' whose visible text exceeds 80 chars
into a `Para'.  Word sometimes styles a lead paragraph (e.g. the
Abstract opener) as a heading; left as a Header it would become an
org section that swallows the following body.  Mirrors the markdown
backend's `# <long prose>' downgrade."
  (mapcar
   (lambda (b)
     (if (and (equal (otd--j-tag b) "Header")
              (let ((inls (append (aref (otd--j-c b) 2) nil)))
                (> (length (string-trim (otd--j-inlines->text inls))) 80)))
         (list (cons 't "Para") (cons 'c (aref (otd--j-c b) 2)))
       b))
   blocks))

(defun otd--j-transform (ast)
  "Return AST with tracked spans converted and top-level fix-ups applied.
Reads `otd--j-done-set'/`otd--j-counts' from the dynamic context."
  (let* ((blocks (otd--j-downgrade-long-headers
                  (append (alist-get 'blocks ast) nil)))
         (walked (vconcat (mapcar #'otd--j-walk blocks))))
    (cons (cons 'blocks walked)
          (assq-delete-all 'blocks (copy-alist ast)))))

;;;; --- driver -------------------------------------------------------------

;;;###autoload
(defun otd-import-json (docx &optional output no-select)
  "Import DOCX to CriticMarkup org via the JSON-AST backend.
With prefix arg, prompt for OUTPUT.  When NO-SELECT, don't visit the
result (used by tests).  Returns the output path."
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
         (json-tmp   (make-temp-file "otd-json-" nil ".json"))
         (json-tmp2  (make-temp-file "otd-json2-" nil ".json"))
         (docx-fixed (otd--preserve-leading-spaces (expand-file-name docx)))
         (otd--j-counts (make-hash-table :test 'eq)))
    (unwind-protect
        (progn
          ;; Stage 1: docx -> pandoc JSON AST (revisions preserved as Spans).
          (with-temp-buffer
            (let ((exit (call-process otd-pandoc-program nil t nil
                                      "-f" "docx" "-t" "json"
                                      "--track-changes=all"
                                      (format "--extract-media=%s" media)
                                      docx-fixed "-o" json-tmp)))
              (unless (zerop exit)
                (error "pandoc docx->json exit %s:\n%s" exit (buffer-string)))))
          ;; Stage 2: structured walk -> CriticMarkup delimiter Str nodes.
          (let* ((otd--j-done-set (otd--extract-done-comment-ids
                                   (expand-file-name docx)))
                 (ast (with-temp-buffer
                        (insert-file-contents json-tmp)
                        (goto-char (point-min))
                        (json-parse-buffer :object-type 'alist :array-type 'array)))
                 (ast2 (otd--j-transform ast)))
            ;; `json-serialize' returns a *unibyte* UTF-8 byte string.
            ;; Inserting it into a multibyte buffer and letting Emacs pick a
            ;; coding system on write raises an interactive coding-system
            ;; prompt on any non-ASCII byte -- which, under `--batch', reads
            ;; empty stdin and hangs forever.  Pin the coding system so the
            ;; UTF-8 bytes are written verbatim, no prompt.
            (let ((coding-system-for-write 'utf-8-unix))
              (with-temp-file json-tmp2 (insert (json-serialize ast2))))
            ;; Stage 3: AST' -> org.  pandoc emits the CriticMarkup delimiters
            ;; verbatim and renders the real content (tables, lists, cites).
            (with-temp-buffer
              (let ((exit (call-process otd-pandoc-program nil t nil
                                        "-f" "json" "-t" "org"
                                        "--wrap=none" json-tmp2 "-o" out)))
                (unless (zerop exit)
                  (error "pandoc json->org exit %s:\n%s" exit (buffer-string))))))
          ;; Stage 4: same org sanitiser as the markdown backend.
          (otd--postfix-org out)
          ;; Stage 5: merge against embedded canonical source, if present.
          (let ((embedded (otd--extract-org-source (expand-file-name docx))))
            (when (and embedded otd-import-auto-merge)
              (let* ((tracked (with-temp-buffer
                                (insert-file-contents out) (buffer-string)))
                     (result (otd--merge-content embedded tracked)))
                (with-temp-file out (insert (car result))))))
          (unless no-select
            (find-file out)
            (otd-criticmarkup-mode 1))
          (message
           "Imported (json) %s -> %s  (++%d --%d ~~%d ==%d >>%d ✓%d)"
           (file-name-nondirectory docx) (file-name-nondirectory out)
           (gethash :ins otd--j-counts 0) (gethash :del otd--j-counts 0)
           (gethash :sub otd--j-counts 0) (gethash :hi otd--j-counts 0)
           (gethash :cmt otd--j-counts 0) (gethash :done otd--j-counts 0))
          out)
      (dolist (f (list json-tmp json-tmp2 docx-fixed))
        (when (and f (file-exists-p f)) (delete-file f))))))

;;;###autoload
(defun otd-import-compare (docx)
  "Import DOCX with BOTH backends and `ediff' the two org outputs."
  (interactive
   (list (read-file-name "Docx file: " nil nil t nil
                         (lambda (f) (or (file-directory-p f)
                                         (string-match-p "\\.docx\\'" f))))))
  (let* ((base (file-name-sans-extension (expand-file-name docx)))
         (md   (concat base "--md.org"))
         (js   (concat base "--json.org"))
         (otd-import-auto-merge nil))
    (otd--import-markdown docx md)
    (otd-import-json docx js 'no-select)
    (ediff-files md js)))

(provide 'org-tracked-docx-json)
;;; org-tracked-docx-json.el ends here
