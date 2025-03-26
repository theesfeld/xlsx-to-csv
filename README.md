# xlsx-to-csv - Convert .xlsx Files to .csv in GNU Emacs

`xlsx-to-csv` is a pure Elisp package for GNU Emacs 30.1+ that converts `.xlsx` files (Office Open XML format) to `.csv` files. It uses only built-in Emacs functionality—no external dependencies—to unzip `.xlsx` archives, parse OOXML, and export each sheet as a separate `.csv` file. The package includes robust error handling, Dired integration, and programmatic access.

Repository: [https://github.com/theesfeld/xlsx-to-csv](https://github.com/theesfeld/xlsx-to-csv)

## Features

- Reads all sheets from an `.xlsx` file and exports each as a separate `.csv` file (e.g., `file-sheet1.csv`, `file-sheet2.csv`).
- Interactive usage via `M-x xlsx-to-csv-convert-file` or Dired (`C-c x`).
- Programmatic usage with `(xlsx-to-csv-convert-file "/path/to/file.xlsx")`.
- Graceful error handling for invalid files, missing data, or write failures.
- Pure Elisp, leveraging Emacs 30.1's `arc-mode` and `libxml-parse-xml-region`.

## Requirements

- GNU Emacs 30.1 or later (released 2024, available at [gnu.org/software/emacs/](https://www.gnu.org/software/emacs/)).

## Installation

### Method 1: Manual Installation

1. Clone or download the repository:
```bash
git clone https://github.com/theesfeld/xlsx-to-csv.git ~/.emacs.d/lisp/xlsx-to-csv
```
2. Add to your `init.el`:
```elisp
(add-to-list 'load-path "~/.emacs.d/lisp/xlsx-to-csv")
(require 'xlsx-to-csv)
```
3. Restart Emacs or evaluate the above lines (`M-x eval-buffer`).

### Method 2: Using `straight.el`

1. Install `straight.el` if not already installed (see [github.com/raxod502/straight.el](https://github.com/raxod502/straight.el)).
2. Add to your `init.el`:
```elisp
(straight-use-package
 '(xlsx-to-csv :type git :host github :repo "theesfeld/xlsx-to-csv"))
(require 'xlsx-to-csv)
```
3. Restart Emacs or run `M-x straight-pull-all` and reload your config.

### Method 3: Using `use-package` with `straight.el`

1. Ensure `use-package` and `straight.el` are installed.
2. Add to your `init.el`:
```elisp
(use-package xlsx-to-csv
  :straight (xlsx-to-csv :type git :host github :repo "theesfeld/xlsx-to-csv")
  :demand t)
```
3. Restart Emacs or evaluate the block.

### Method 4: Using `quelpa` and `use-package`

1. Install `quelpa` (see [github.com/quelpa/quelpa](https://github.com/quelpa/quelpa)).
2. Add to your `init.el`:
```elisp
(use-package xlsx-to-csv
  :quelpa (xlsx-to-csv :fetcher github :repo "theesfeld/xlsx-to-csv")
  :demand t)
```
3. Run `M-x quelpa-upgrade` or restart Emacs.

### Notes on Installation

- As of March 26, 2025, this package is not available on MELPA or ELPA, so manual or Git-based methods are required.
- Ensure your Emacs is 30.1+ (`M-x emacs-version` to check).

## Usage

### Interactive Usage

1. **Convert a Single File**:
   - Run `M-x xlsx-to-csv-convert-file RET /path/to/file.xlsx RET`.
   - Output: `.csv` files (e.g., `file-sheet1-Name.csv`) in the same directory.
   - Example feedback: "Converted /path/to/file.xlsx to 2 CSV files: file-sheet1.csv, file-sheet2.csv".

2. **Using Dired**:
   - Open a Dired buffer (`C-x d` or `M-x dired`).
   - Mark `.xlsx` files with `m`.
   - Press `C-c x` (or `M-x dired-do-xlsx-to-csv RET`).
   - Feedback: "Processed 3 .xlsx files successfully".

### Programmatic Usage

- Call the function in Lisp code:
```elisp
(let ((csv-files (xlsx-to-csv-convert-file "/path/to/data.xlsx")))
  (if csv-files
      (message "Generated: %s" (string-join csv-files ", "))
    (message "Conversion failed")))
```
- Returns: A list of `.csv` file paths (e.g., `("/path/data-sheet1.csv" "/path/data-sheet2.csv")`) or `nil` on failure.

### Detailed Examples

1. **Simple Spreadsheet**:
   - Input: `test.xlsx` with one sheet ("Data"):

| A | B |
|```+```|
| 1 | 2 |
| 3 | 4 |

- Command: `M-x xlsx-to-csv-convert-file RET ~/test.xlsx RET`
- Output: `~/test-sheet1-Data.csv`:
```bash
"1","2"
"3","4"
```

2. **Multiple Sheets**:
- Input: `multi.xlsx` with two sheets:
- Sheet1 ("Numbers"): `((1 2) (3 4))`
- Sheet2 ("Text"): `(("a" "b") ("c" "d"))`
- Command: `(xlsx-to-csv-convert-file "~/multi.xlsx")`
- Output:
- `~/multi-sheet1-Numbers.csv`: `"1","2"\n"3","4"`
- `~/multi-sheet2-Text.csv`: `"a","b"\n"c","d"`
- Returns: `("~/multi-sheet1-Numbers.csv" "~/multi-sheet2-Text.csv")`

3. **Error Case**:
- Input: `broken.xlsx` (corrupt file).
- Command: `M-x xlsx-to-csv-convert-file RET ~/broken.xlsx RET`
- Feedback: "Error converting ~/broken.xlsx: Failed to parse workbook.xml"
- Returns: `nil`

## Customization

As of version 0.2, there are no customizable `setq` variables. The package operates with default behavior:
- Outputs `.csv` files in the same directory as the `.xlsx` file.
- Names files as `<base>-sheet<Num>-<SheetName>.csv`.
- Uses comma-separated values with quoted strings.

Future versions may add variables for:
- Output directory (`xlsx-to-csv-output-dir`).
- CSV delimiter (`xlsx-to-csv-delimiter`).
- File naming scheme (`xlsx-to-csv-naming-function`).

## Error Handling

The package gracefully handles:
- Non-existent or unreadable `.xlsx` files.
- Corrupt ZIP archives or missing XML files.
- Invalid sheet data or write failures.
- Non-`.xlsx` files in Dired (skipped with a message).

Feedback is provided via the minibuffer, and programmatic calls return `nil` on total failure or a partial list if some sheets fail.

## Contributing

Issues and pull requests are welcome at [https://github.com/theesfeld/xlsx-to-csv](https://github.com/theesfeld/xlsx-to-csv). Please include:
- A description of the bug or feature.
- Steps to reproduce (for bugs).
- Emacs version and relevant config.

## License

Distributed under the GNU General Public License v3 or later. See the source file (`xlsx-to-csv.el`) for details.

```
Last updated: March 26, 2025
