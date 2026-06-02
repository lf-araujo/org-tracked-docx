;;; org-tracked-docx-zotero.el --- Zotero-shielded docx roundtrip  -*- lexical-binding: t; -*-

;; Author: drafted with Claude
;; Keywords: org, pandoc, docx, track changes, zotero
;; Package-Requires: ((emacs "27.1") (org-tracked-docx "0.1"))

;;; Commentary:
;;
;; Companion to org-tracked-docx.el for documents authored in Google Docs
;; with the Zotero connector.  Such documents store citations as Word
;; HYPERLINK fields pointing at https://www.zotero.org/google-docs/?CODE
;; URLs, and a naive pandoc roundtrip silently strips every URL.
;;
;; This file shields those fields by:
;;   1. Pre-extracting the field-character XML from the source docx and
;;      replacing each group with an opaque placeholder token like
;;      `❰ZOT0001❱'.  Saved chunks live in a sidecar JSON.
;;   2. Running pandoc docx -> markdown with --track-changes=all on the
;;      shielded docx.  Tokens pass through pandoc as plain text.
;;   3. After editing in markdown (CriticMarkup highlighting via
;;      `otd-criticmarkup-mode'), pandoc markdown -> docx, then post-
;;      processing the result to substitute each placeholder with its
;;      saved Zotero field-group XML.
;;
;; Why markdown rather than org?  Pandoc's org reader hangs on the very
;; wide multi-column tables typical of NIH proposals authored in Google
;; Docs; markdown is unaffected.
;;
;; Workflow:
;;   M-x otd-zot-import   ; pick a docx; opens <basename>.shielded.md
;;   ... edit, accept/reject CriticMarkup ...
;;   M-x otd-zot-export   ; from the markdown buffer; writes
;;                          <basename>-edited.docx with all Zotero
;;                          hyperlinks restored.

;;; Code:

(require 'org-tracked-docx)
(require 'subr-x)

(defgroup org-tracked-docx-zotero nil
  "Zotero-shielded pandoc roundtrip for docx."
  :group 'org-tracked-docx)

(defcustom otd-zot-python "python3"
  "Python interpreter used to invoke the shield helper."
  :type 'string :group 'org-tracked-docx-zotero)

(defcustom otd-zot-helper
  (expand-file-name "otd_zotero_shield.py"
                    (file-name-directory
                     (or load-file-name buffer-file-name default-directory)))
  "Path to the otd_zotero_shield.py helper script."
  :type 'file :group 'org-tracked-docx-zotero)

(defcustom otd-zot-shielded-suffix ".shielded.md"
  "Suffix appended to the source docx basename for the editable markdown."
  :type 'string :group 'org-tracked-docx-zotero)

(defcustom otd-zot-sidecar-suffix ".zotrefs.json"
  "Suffix appended to the source docx basename for the sidecar JSON."
  :type 'string :group 'org-tracked-docx-zotero)

(defcustom otd-zot-output-suffix "-edited.docx"
  "Suffix appended to the source docx basename for the exported docx."
  :type 'string :group 'org-tracked-docx-zotero)

(defun otd-zot--call (subcmd &rest args)
  "Invoke the python helper SUBCMD with ARGS, return its parsed JSON result."
  (with-temp-buffer
    (let ((exit (apply #'call-process otd-zot-python nil t nil
                       otd-zot-helper subcmd args)))
      (unless (zerop exit)
        (error "otd_zotero_shield.py %s exit %s:\n%s"
               subcmd exit (buffer-string)))
      (goto-char (point-min))
      (let ((json-object-type 'alist) (json-array-type 'list))
        (json-read)))))

(defun otd-zot--header (source-docx sidecar)
  "Return the pinned-state header line for the editable markdown."
  (format "<!-- otd-zot: source=%S sidecar=%S -->\n\n"
          (expand-file-name source-docx)
          (expand-file-name sidecar)))

(defun otd-zot--header-read (md-file field)
  "Read FIELD (\"source\" or \"sidecar\") out of the otd-zot header in MD-FILE."
  (with-temp-buffer
    (insert-file-contents md-file nil 0 4096)
    (goto-char (point-min))
    (when (re-search-forward
           (format "<!-- otd-zot: .*?%s=\"\\([^\"]+\\)\"" field) nil t)
      (match-string 1))))

;;;###autoload
(defun otd-zot-import (docx)
  "Convert DOCX to a Zotero-shielded markdown buffer with track-changes.
The Zotero HYPERLINK fields are extracted into a sidecar JSON and replaced
with opaque placeholders so pandoc cannot mangle them.  The output buffer
runs `otd-criticmarkup-mode' so that {++ins++}/{--del--}/{>>cmt<<} markers
are font-locked, and `otd-accept-all'/`otd-reject-all' work as usual."
  (interactive
   (list (read-file-name "Docx file: " nil nil t nil
                         (lambda (f) (or (file-directory-p f)
                                         (string-match-p "\\.docx\\'" f))))))
  (unless (file-exists-p docx) (user-error "File not found: %s" docx))
  (unless (file-exists-p otd-zot-helper)
    (user-error "Helper script missing: %s" otd-zot-helper))
  (let* ((dir          (file-name-directory (expand-file-name docx)))
         (base         (file-name-sans-extension (file-name-nondirectory docx)))
         (sidecar      (expand-file-name (concat base otd-zot-sidecar-suffix) dir))
         (md-out       (expand-file-name (concat base otd-zot-shielded-suffix) dir))
         (shielded-dx  (make-temp-file "otd-zot-" nil ".docx"))
         (md-tmp       (make-temp-file "otd-zot-" nil ".md"))
         (media        (expand-file-name
                        (concat base "-" otd-extract-media-suffix) dir)))
    (unwind-protect
        (let* ((shield-result
                (otd-zot--call "shield"
                               (expand-file-name docx)
                               shielded-dx
                               sidecar))
               (n-shielded (alist-get 'shielded shield-result)))
          ;; pandoc docx -> markdown with track changes; placeholders pass through
          (with-temp-buffer
            (let ((exit (apply #'call-process otd-pandoc-program nil t nil
                               (list "-f" "docx" "-t" "markdown"
                                     "--wrap=none"
                                     "--track-changes=all"
                                     (format "--extract-media=%s" media)
                                     shielded-dx
                                     "-o" md-tmp))))
              (unless (zerop exit)
                (error "pandoc docx->md exit %s:\n%s" exit (buffer-string)))))
          ;; Reuse the existing CriticMarkup span rewriter
          (let ((counts (otd--rewrite-spans md-tmp)))
            (with-temp-file md-out
              (insert (otd-zot--header docx sidecar))
              (insert-file-contents md-tmp))
            (find-file md-out)
            (otd-criticmarkup-mode 1)
            (message
             "Imported %s -> %s  (zotero shielded:%d  ++%d --%d ~~%d ==%d >>%d)"
             (file-name-nondirectory docx)
             (file-name-nondirectory md-out)
             n-shielded
             (plist-get counts :ins) (plist-get counts :del)
             (plist-get counts :sub) (plist-get counts :hi)
             (plist-get counts :cmt))))
      (when (file-exists-p shielded-dx) (delete-file shielded-dx))
      (when (file-exists-p md-tmp) (delete-file md-tmp)))
    md-out))

;;;###autoload
(defun otd-zot-export (md-file &optional out-docx)
  "Export the shielded markdown MD-FILE back to a docx with Zotero links.
Reads the source docx and sidecar paths from the otd-zot header in
MD-FILE.  With prefix arg, prompts for OUT-DOCX; otherwise places the
output alongside the source docx with `otd-zot-output-suffix' appended.

CriticMarkup tokens left in the buffer are written through pandoc as
literal text.  Run `otd-accept-all' or `otd-reject-all' first if you
want to flatten them."
  (interactive
   (list (or (and (buffer-file-name)
                  (string-match-p "\\.shielded\\.md\\'" (buffer-file-name))
                  (buffer-file-name))
             (read-file-name "Shielded .md: " nil nil t nil
                             (lambda (f) (or (file-directory-p f)
                                             (string-match-p "\\.md\\'" f)))))
         (when current-prefix-arg (read-file-name "Output .docx: "))))
  (let* ((source (otd-zot--header-read md-file "source"))
         (sidecar (otd-zot--header-read md-file "sidecar")))
    (unless (and source sidecar)
      (user-error "No otd-zot header found in %s — was this file produced by `otd-zot-import'?"
                  md-file))
    (unless (file-exists-p sidecar)
      (user-error "Sidecar not found: %s" sidecar))
    (let* ((src-dir  (file-name-directory source))
           (src-base (file-name-sans-extension (file-name-nondirectory source)))
           (final    (or out-docx
                         (expand-file-name
                          (concat src-base otd-zot-output-suffix) src-dir)))
           (raw-dx   (make-temp-file "otd-zot-out-" nil ".docx")))
      (when (and (buffer-file-name)
                 (string= (expand-file-name md-file) (buffer-file-name))
                 (buffer-modified-p))
        (when (y-or-n-p (format "Save %s before exporting? "
                                (file-name-nondirectory md-file)))
          (save-buffer)))
      (unwind-protect
          (progn
            ;; pandoc md -> docx (still has placeholders inline as plain text)
            (with-temp-buffer
              (let ((exit (apply #'call-process otd-pandoc-program nil t nil
                                 (list "-f" "markdown" "-t" "docx"
                                       "--wrap=none"
                                       md-file
                                       "-o" raw-dx))))
                (unless (zerop exit)
                  (error "pandoc md->docx exit %s:\n%s" exit (buffer-string)))))
            ;; Substitute placeholders back with original Zotero XML
            (let* ((res (otd-zot--call "unshield" raw-dx sidecar final))
                   (rest (alist-get 'restored res))
                   (exp (alist-get 'expected res))
                   (leftover (alist-get 'leftover res))
                   (warns (alist-get 'warnings res)))
              (message "Exported %s  (zotero restored:%d/%d  leftover:%d  warnings:%d)"
                       (file-name-nondirectory final)
                       rest exp leftover (length warns))
              (when warns
                (with-current-buffer (get-buffer-create "*otd-zot-warnings*")
                  (let ((inhibit-read-only t))
                    (erase-buffer)
                    (dolist (w warns) (insert w "\n"))
                    (special-mode)
                    (display-buffer (current-buffer)))))))
        (when (file-exists-p raw-dx) (delete-file raw-dx)))
      final)))

;;;###autoload
(defun otd-zot-export-org (org-file &optional out-docx)
  "Export ORG-FILE to docx with Zotero links restored, via emacs ox-md + pandoc.
Bypasses pandoc's org reader (which hangs on wide tables in this corpus):
  1. emacs ox-md   converts ORG-FILE -> <basename>.md (in-process)
  2. pandoc        converts the .md -> .docx, with `otd-zotero-unshield.lua'
                   replacing every `❰ZOTNNNN❱' token in the AST with raw OOXML
                   from the sidecar JSON living next to ORG-FILE.

The sidecar file is expected at `<basename>.zotrefs.json' alongside ORG-FILE.
With prefix arg, prompts for OUT-DOCX; otherwise writes to
`<basename>-edited.docx'."
  (interactive
   (list (or (and (buffer-file-name)
                  (string-match-p "\\.org\\'" (buffer-file-name))
                  (buffer-file-name))
             (read-file-name "Org file: " nil nil t nil
                             (lambda (f) (or (file-directory-p f)
                                             (string-match-p "\\.org\\'" f)))))
         (when current-prefix-arg (read-file-name "Output .docx: "))))
  (require 'ox-md)
  (let* ((src-dir   (file-name-directory (expand-file-name org-file)))
         (src-base  (file-name-sans-extension (file-name-nondirectory org-file)))
         (sidecar   (expand-file-name (concat src-base otd-zot-sidecar-suffix) src-dir))
         (md-out    (expand-file-name (concat src-base ".md") src-dir))
         (final     (or out-docx
                        (expand-file-name (concat src-base otd-zot-output-suffix) src-dir)))
         (lua-filter (expand-file-name
                      "otd-zotero-unshield.lua"
                      (file-name-directory otd-zot-helper))))
    (unless (file-exists-p sidecar)
      (user-error "Sidecar not found: %s — run `otd-zot-import' on the source docx first" sidecar))
    (unless (file-exists-p lua-filter)
      (user-error "Lua filter not found: %s" lua-filter))
    ;; (1) Save current buffer if it's the source.
    (when (and (buffer-file-name)
               (string= (expand-file-name org-file) (buffer-file-name))
               (buffer-modified-p))
      (when (y-or-n-p (format "Save %s before exporting? "
                              (file-name-nondirectory org-file)))
        (save-buffer)))
    ;; (2) org -> markdown via ox-md (in-process).
    (let ((buf (find-file-noselect org-file)))
      (with-current-buffer buf
        (let ((org-export-with-toc nil)
              (org-export-with-section-numbers nil)
              (org-export-with-broken-links t))
          (org-md-export-to-markdown))))
    ;; (3) md -> docx with the Zotero-unshield Lua filter.
    (with-temp-buffer
      (let ((exit (apply #'call-process otd-pandoc-program nil t nil
                         (list "-f" "markdown" "-t" "docx"
                               "--wrap=none"
                               "--metadata" (concat "zotrefs-sidecar=" sidecar)
                               "--lua-filter" lua-filter
                               md-out
                               "-o" final))))
        (let ((stderr (buffer-string)))
          (unless (zerop exit)
            (error "pandoc md->docx exit %s:\n%s" exit stderr))
          (message "Exported %s -> %s\n%s"
                   (file-name-nondirectory org-file)
                   (file-name-nondirectory final)
                   (string-trim stderr)))))
    final))

(provide 'org-tracked-docx-zotero)
;;; org-tracked-docx-zotero.el ends here
