;;; org-tracked-docx-pkg.el --- Package descriptor  -*- no-byte-compile: t -*-

;; Authoritative package metadata.  Hand-written so `package-vc' does NOT
;; auto-generate one by scraping `Package-Requires' across every .el in the
;; repo: that scrape merges the `(org-tracked-docx "0.1")' line from
;; org-tracked-docx-zotero.el into the main package, making it depend on
;; itself and sending `package-activate' into infinite recursion
;; ("excessive-lisp-nesting").

(define-package "org-tracked-docx" "0.1"
  "Round-trip Word docx with tracked changes via pandoc and CriticMarkup"
  '((emacs "27.1")))
