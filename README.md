# Excel Comment Column Copier

A Python script that uses `openpyxl` to find the last `COMMENTS DD-MM-YYYY` column in an Excel worksheet, duplicate it, and hide the original column while preserving values, formatting, width, and merged cells where possible.

## What it does

- Opens an Excel workbook.
- Searches a specific worksheet for the last header matching `COMMENTS DD-MM-YYYY`.
- Copies the entire target column.
- Inserts a new column to the right of the original.
- Pastes the copied values and formatting into the new column.
- Hides the original column.
- Saves the result as a new file with `_updated` appended to the filename.

## Features

- Regex-based header detection.
- Preserves:
  - cell values
  - fonts
  - fills
  - borders
  - alignment
  - number formats
  - protection
  - hyperlinks
  - comments
- Copies the original column width.
- Handles merged cells gracefully without crashing.
- Easy to configure at the top of the script.

## Requirements

- Python 3.9+
- `openpyxl`

Install dependencies:

```bash
pip install openpyxl
```

## Usage

1. Place the Excel file in the folder expected by the script, or update the `input_file` path.
2. Set the correct worksheet name.
3. Set the correct header row.
4. Run the script:

```bash
python script-perplexity.py
```

## Configuration

At the top of the script, update:

```python
input_file = Path(r"C:\Users\Jason\Projects\EXCEL-COPY-COLUMN\BARTRAC - KCC TRACKING AS OF 31-12-2025.xlsx")
sheet_name = "ENROUTE SITE"
header_row = 6
```

## How it works

The script scans the header row for the last column whose cell value matches this pattern:

```python
COMMENTS DD-MM-YYYY
```

For example:

```text
COMMENTS 31-12-2025
```

Once found, it:

1. Caches the entire column.
2. Inserts a blank column to the right.
3. Copies the cached content and formatting into the new column.
4. Hides the original column on the left.
5. Saves the workbook with `_updated` appended to the filename.

## Output

If the input file is:

```text
BARTRAC - KCC TRACKING AS OF 31-12-2025.xlsx
```

the output will be:

```text
BARTRAC - KCC TRACKING AS OF 31-12-2025_updated.xlsx
```

## Notes

- The script assumes the headers are on row `6`.
- If your worksheet name changes, update `sheet_name`.
- If the workbook has multiple matching comment columns, the **last** one in the header row is used.
- Merged cells are handled as best as possible, but complex merge layouts may still need manual review.

## Example

Input:

- Sheet: `ENROUTE SITE`
- Header row: `6`
- Matching header: `COMMENTS 31-12-2025`

Result:

- The original comments column is hidden.
- A duplicated visible copy is placed immediately to the right.

## License

MIT