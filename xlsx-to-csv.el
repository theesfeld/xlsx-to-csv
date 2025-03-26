;;; xlsx-to-csv.el --- Convert .xlsx files to .csv using pure Elisp -*- lexical-binding: t; -*-

;; Copyright (C) 2025 Your Name

;; Author: William Theesfeld <william@theesfeld.net>
;; Version: 0.5
;; Keywords: files, data, spreadsheet, dired
;; Package-Requires: ((emacs "30.1"))

;; This program is free software; you can redistribute it and/or modify
;; it under the terms of the GNU General Public License as published by
;; the Free Software Foundation, either version 3 of the License, or
;; (at your option) any later version.

;;; Commentary:

;; This package converts .xlsx files to .csv files using only built-in
;; Emacs 30.1 functionality. It reads all sheets, supports Dired
;; integration, and provides programmatic access. Each sheet is exported
;; as a separate .csv file. Extensive error checking ensures graceful
;; failure with clear feedback.
;;
;; Usage:
;; - Interactive: `M-x xlsx-to-csv-convert-file RET /path/to/file.xlsx RET`
;; - In Dired: Mark .xlsx files, then `M-x dired-do-xlsx-to-csv RET`
;; - Programmatic: `(xlsx-to-csv-convert-file "/path/to/file.xlsx")`

;;; Code:

(require 'arc-mode) ;; For archive handling
(require 'xml) ;; For XML parsing
(require 'dired) ;; For Dired integration

(defun xlsx-to-csv--extract-to-buffer (archive file-name)
  "Extract FILE-NAME from ARCHIVE to a temporary buffer.
Return the buffer or nil if extraction fails."
  (condition-case err
      (let ((archive-buffer
             (and (file-exists-p archive)
                  (file-readable-p archive)
                  (find-file-noselect archive))))
        (unless archive-buffer
          (error "Cannot open archive %s" archive))
        (with-current-buffer archive-buffer
          (goto-char (point-min))
          (unless (re-search-forward (regexp-quote file-name) nil t)
            (error "File %s not found in archive" file-name))
          (archive-extract))
        (with-temp-buffer
          (insert-buffer-substring archive-buffer)
          (kill-buffer archive-buffer)
          (current-buffer)))
    (error
     (message "Failed to extract %s from %s: %s"
              file-name
              archive
              (error-message-string err))
     nil)))

(defun xlsx-to-csv--parse-shared-strings (xlsx-file)
  "Parse shared strings from XLSX-FILE's sharedStrings.xml.
Return a list of strings or nil on failure."
  (condition-case err
      (let* ((buffer
              (xlsx-to-csv--extract-to-buffer
               xlsx-file "xl/sharedStrings.xml"))
             (xml-tree
              (and buffer
                   (with-current-buffer buffer
                     (libxml-parse-xml-region
                      (point-min) (point-max)))))
             (strings '()))
        (unless xml-tree
          (error "Failed to parse sharedStrings.xml"))
        (dolist (si (xml-get-children xml-tree 'si))
          (let ((t-node (car (xml-get-children si 't))))
            (push (or (car (xml-node-children t-node)) "") strings)))
        (when buffer
          (kill-buffer buffer))
        (nreverse strings))
    (error
     (message "Error parsing shared strings in %s: %s"
              xlsx-file
              (error-message-string err))
     nil)))

(defun xlsx-to-csv--get-sheets (xlsx-file)
  "Return a list of (sheet-num . sheet-name) from XLSX-FILE's workbook.xml.
Return nil if parsing fails."
  (condition-case err
      (let* ((buffer
              (xlsx-to-csv--extract-to-buffer
               xlsx-file "xl/workbook.xml"))
             (xml-tree
              (and buffer
                   (with-current-buffer buffer
                     (libxml-parse-xml-region
                      (point-min) (point-max)))))
             (sheets '()))
        (unless xml-tree
          (error "Failed to parse workbook.xml"))
        (dolist (sheet
                 (xml-get-children
                  (car (xml-get-children xml-tree 'sheets)) 'sheet))
          (let ((sheet-id (xml-get-attribute sheet 'sheetId))
                (sheet-name (xml-get-attribute sheet 'name)))
            (push (cons
                   (string-to-number sheet-id)
                   (or sheet-name (format "Sheet%d" sheet-id)))
                  sheets)))
        (when buffer
          (kill-buffer buffer))
        (nreverse sheets))
    (error
     (message "Error parsing sheets in %s: %s"
              xlsx-file
              (error-message-string err))
     nil)))

(defun xlsx-to-csv--cell-to-coords (cell-ref)
  "Convert CELL-REF (e.g., \"A1\") to (row . col) coordinates.
Return nil if conversion fails."
  (condition-case err
      (let* ((col-str (replace-regexp-in-string "[0-9]+" "" cell-ref))
             (row-str (replace-regexp-in-string "[A-Z]+" "" cell-ref))
             (col-num
              (1- (string-to-number
                   (mapconcat
                    (lambda (c)
                      (number-to-string (- (upcase c) ?A -1)))
                    col-str
                    ""))))
             (row-num (1- (string-to-number row-str))))
        (if (and (>= row-num 0) (>= col-num 0))
            (cons row-num col-num)
          (error "Invalid cell reference: %s" cell-ref)))
    (error
     (message "Error converting cell reference %s: %s"
              cell-ref
              (error-message-string err))
     nil)))

(defun xlsx-to-csv--parse-sheet (xlsx-file sheet-num shared-strings)
  "Parse sheet SHEET-NUM from XLSX-FILE using SHARED-STRINGS.
Return a data structure or nil on failure."
  (condition-case err
      (let* ((file-name
              (format "xl/worksheets/sheet%d.xml" sheet-num))
             (buffer
              (xlsx-to-csv--extract-to-buffer xlsx-file file-name))
             (xml-tree
              (and buffer
                   (with-current-buffer buffer
                     (libxml-parse-xml-region
                      (point-min) (point-max)))))
             (rows '())
             (max-row 0)
             (max-col 0))
        (unless xml-tree
          (error "Failed to parse %s" file-name))
        (dolist (row
                 (xml-get-children
                  (car (xml-get-children xml-tree 'sheetData)) 'row))
          (let ((row-num
                 (string-to-number
                  (or (xml-get-attribute row 'r) "1")))
                (row-data '()))
            (dolist (c (xml-get-children row 'c))
              (let* ((cell-ref (xml-get-attribute c 'r))
                     (coords
                      (and cell-ref
                           (xlsx-to-csv--cell-to-coords cell-ref)))
                     (v-node (car (xml-get-children c 'v)))
                     (value (car (xml-node-children v-node)))
                     (t-attr (xml-get-attribute c 't)))
                (when (and coords value)
                  (when (string= t-attr "s")
                    (setq value
                          (nth
                           (string-to-number value) shared-strings)))
                  (push (cons coords value) row-data)
                  (setq max-row (max max-row (car coords)))
                  (setq max-col (max max-col (cdr coords))))))
            (when row-data
              (push (cons row-num row-data) rows))))
        (when buffer
          (kill-buffer buffer))
        (let ((matrix (make-vector (1+ max-row) nil)))
          (dolist (row rows)
            (let ((row-vec (make-vector (1+ max-col) nil)))
              (dolist (cell (cdr row))
                (aset row-vec (cdr (car cell)) (or (cdr cell) "")))
              (aset matrix (car row) row-vec)))
          (mapcar 'vector-to-list (append matrix nil))))
    (error
     (message "Error parsing sheet %d in %s: %s"
              sheet-num
              xlsx-file
              (error-message-string err))
     nil)))

(defun xlsx-to-csv--to-csv (data output-file)
  "Write DATA (list of lists) to OUTPUT-FILE in CSV format.
Return t on success, nil on failure."
  (condition-case err
      (progn
        (unless (file-writable-p (file-name-directory output-file))
          (error
           "Directory not writable: %s"
           (file-name-directory output-file)))
        (with-temp-file output-file
          (dolist (row data)
            (insert
             (mapconcat (lambda (cell)
                          (if (stringp cell)
                              (concat
                               "\""
                               (replace-regexp-in-string
                                "\"" "\"\"" cell)
                               "\"")
                            (or cell "")))
                        row
                        ","))
            (insert "\n")))
        t)
    (error
     (message "Failed to write CSV %s: %s"
              output-file
              (error-message-string err))
     nil)))

(defun xlsx-to-csv-convert-file (file)
  "Convert .xlsx FILE to .csv files.
Return the list of output file paths or nil on failure."
  (interactive "fXLSX file: ")
  (condition-case err
      (progn
        (unless (and (file-exists-p file) (file-readable-p file))
          (error "File does not exist or is not readable: %s" file))
        (unless (string-suffix-p ".xlsx" file)
          (error "File must be a .xlsx file: %s" file))
        (let* ((base-name (file-name-sans-extension file))
               (shared-strings
                (xlsx-to-csv--parse-shared-strings file))
               (sheets (xlsx-to-csv--get-sheets file))
               (output-files '()))
          (unless sheets
            (error "No sheets found in %s" file))
          (unless shared-strings
            (error "No shared strings parsed from %s" file))
          (dolist (sheet sheets)
            (let* ((sheet-num (car sheet))
                   (sheet-name (cdr sheet))
                   (data
                    (xlsx-to-csv--parse-sheet
                     file sheet-num shared-strings))
                   (output-file
                    (format "%s-sheet%d-%s.csv"
                            base-name sheet-num
                            (replace-regexp-in-string
                             "[^[:alnum:]]" "_" sheet-name))))
              (when data
                (if (xlsx-to-csv--to-csv data output-file)
                    (push output-file output-files)
                  (message
                   "Skipping sheet %d (%s) due to write failure"
                   sheet-num sheet-name)))))
          (when (called-interactively-p `interactive)
            (if output-files
                (message "Converted %s to %d CSV files: %s"
                         file
                         (length output-files)
                         (string-join output-files ", "))
              (message "Failed to convert %s: No CSV files generated"
                       file)))
          (if output-files
              (nreverse output-files)
            nil)))
    (error
     (message "Error converting %s: %s"
              file
              (error-message-string err))
     nil)))

(defun dired-do-xlsx-to-csv (&optional arg)
  "Convert marked .xlsx files in Dired to .csv files."
  (interactive "P")
  (let ((files (dired-get-marked-files nil arg))
        (success-count 0))
    (dolist (file files)
      (if (string-suffix-p ".xlsx" file)
          (let ((result (xlsx-to-csv-convert-file file)))
            (when result
              (setq success-count (1+ success-count))))
        (message "Skipping non-.xlsx file: %s" file)))
    (message "Processed %d .xlsx files successfully" success-count)))

(define-key dired-mode-map (kbd "C-c x") 'dired-do-xlsx-to-csv)

(provide 'xlsx-to-csv)
;;; xlsx-to-csv.el ends here
