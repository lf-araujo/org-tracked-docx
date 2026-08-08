;;; otd-json-test.el --- ERT tests for the JSON import backend -*- lexical-binding: t; -*-

;; Run from the package root:
;;   emacs -Q --batch -L . -l tests/otd-json-test.el -f ert-run-tests-batch-and-exit
;;
;; These guard the tracked-change -> CriticMarkup conversion of the JSON
;; backend against the bug classes that motivated it: comment text containing
;; `]', nested comments, adjacent delete+insert substitutions, and tracked
;; changes inside table cells.  The fixtures are hand-crafted `.docx' files
;; (see tests/*.docx) exercising each construct.

;;; Code:

(require 'ert)
(require 'org-tracked-docx)
(require 'org-tracked-docx-json)

(defvar otd-json-test--dir
  (file-name-directory (or load-file-name buffer-file-name))
  "Directory holding the test fixtures.")

(defun otd-json-test--import (fixture)
  "Import FIXTURE (a basename under the tests dir) via the JSON backend,
returning the resulting org text.  Merge is disabled so the raw import
is tested."
  (let* ((docx (expand-file-name fixture otd-json-test--dir))
         (out  (make-temp-file "otd-json-test-" nil ".org"))
         (otd-import-auto-merge nil))
    (unwind-protect
        (progn (otd-import-json docx out 'no-select)
               (with-temp-buffer (insert-file-contents out) (buffer-string)))
      (when (file-exists-p out) (delete-file out)))))

(ert-deftest otd-json-insertion-deletion-substitution ()
  "Insertions, deletions and adjacent del+ins substitutions convert."
  (let ((org (otd-json-test--import "tracked-fixture.docx")))
    (should (string-match-p (regexp-quote "{++very++}") org))
    (should (string-match-p (regexp-quote "{--quickly--}") org))
    ;; `lazy' deleted then `sleepy' inserted -> one substitution token
    (should (string-match-p (regexp-quote "{~~lazy~>sleepy~~}") org))))

(ert-deftest otd-json-comment-with-bracket ()
  "A comment whose body contains `]' survives verbatim (no drop, no
escaping leak).  This is the regression the markdown backend fails."
  (let ((org (otd-json-test--import "tracked-fixture.docx")))
    (should (string-match-p
             (regexp-quote "inner note with a [bracket] and citation") org))
    ;; no markdown escaping artifact
    (should-not (string-match-p (regexp-quote "\\[bracket\\]") org))))

(ert-deftest otd-json-nested-comments ()
  "Nested comments each become their own {==..==}{>>..<<} run."
  (let ((org (otd-json-test--import "tracked-fixture.docx")))
    (should (string-match-p (regexp-quote "{>>[Reviewer B] outer comment<<}") org))
    (should (string-match-p (regexp-quote "{==inner range==}") org))))

(ert-deftest otd-json-resolved-comment-done ()
  "A comment resolved in Word (commentsExtended done=1) is tagged [DONE]."
  (let ((org (otd-json-test--import "tracked-fixture.docx")))
    (should (string-match-p
             (regexp-quote "{>>[Reviewer A] [DONE] check this phrase<<}") org))))

(ert-deftest otd-json-tracked-change-in-table-cell ()
  "A tracked insertion inside a table cell is converted (generic walk,
no table-specific code)."
  (let ((org (otd-json-test--import "table-cell-fixture.docx")))
    (should (string-match-p "|" org))                     ; rendered as an org table
    (should (string-match-p (regexp-quote "{++added++}") org))))

(provide 'otd-json-test)
;;; otd-json-test.el ends here
