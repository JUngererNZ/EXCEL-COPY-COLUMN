

---------------------------------
21-04-2026 1235
# Vibe coded this out to save me some time in the mornings to start the morning chores
# Excel Comment Column Copier

Python utilities built with `openpyxl` to find the last `COMMENTS DD-MM-YYYY` column in a worksheet, duplicate it, preserve formatting, and hide the original column.

## Overview

These scripts are designed for Excel tracking workbooks where comment columns are added over time using headers such as `COMMENTS 31-12-2025`.

The current behavior is:

- Find the **last** header matching `COMMENTS DD-MM-YYYY`.
- Copy the entire column, including values and formatting.
- Insert a new duplicate column immediately to the **right** of the original column.
- Keep the **original column on the left and hide it**.
- Save the workbook **in place**, overwriting the original file.

## Features

- Regex-based header detection using `^COMMENTS \d{2}-\d{2}-\d{4}$`
- Copies:
  - cell values
  - fonts
  - fills
  - borders
  - alignment
  - number formats
  - protection
  - comments
  - hyperlinks
- Preserves the original column width
- Handles merged cells gracefully
- Uses configurable file path, sheet name, and header row values at the top of each script
- Updates the workbook directly instead of creating a separate output file

## Requirements

- Python 3.9+
- `openpyxl`

Install dependency:

```bash
pip install openpyxl
```

## Scripts

### `script-perplexity-copy-2.py`

Configured for a workbook such as:

```python
input_file = Path(r"C:\Users\Jason\Projects\EXCEL-COPY-COLUMN\BARTRAC - KAMOA TRACKING AS OF 31-12-2025.xlsx")
sheet_name = "ENROUTE SITE"
header_row = 6
```

Behavior:

- Searches row `6` for the last `COMMENTS DD-MM-YYYY` header.
- Inserts the duplicate column to the **right**.
- Hides the original column on the **left**.
- Saves changes back into the **same file**.

### `script-perplexity-copy-3.py`

This script currently follows the same logic and structure as `script-perplexity-copy-2.py`.

Use it as a second variant for another workbook by changing:

```python
input_file
sheet_name
header_row
```

This is useful when working with multiple Excel files that share the same column-copy workflow.

### Related workbook variant

A similar variant was also used for a Congo tracking workbook where the target sheet name required special handling because the actual sheet name included a trailing space.

Example issue:

```text
CURRENT SHIPMENTS 
```

In cases like that, sheet matching should be normalized with `.strip()` before selecting the worksheet.

## How it works

1. Load the workbook with `openpyxl`.
2. Open the configured worksheet.
3. Scan the configured header row for the last matching `COMMENTS DD-MM-YYYY` column.
4. Cache all cells from that column, including style metadata.
5. Record the original column width.
6. Record merged ranges touching the target column.
7. Insert a blank column to the **right** of the target column.
8. Paste the copied data into the new column.
9. Recreate relevant merged ranges where possible.
10. Hide the original column on the left.
11. Save the workbook back to the same filename.

## Usage

1. Open the script you want to run.
2. Update the configuration block:

```python
input_file = Path(r"C:\path\to\your\file.xlsx")
sheet_name = "ENROUTE SITE"
header_row = 6
```

3. Run the script:

```bash
python script-perplexity-copy-2.py
```

or

```bash
python script-perplexity-copy-3.py
```

## Header matching

The scripts search for headers matching this regex:

```python
^COMMENTS \d{2}-\d{2}-\d{4}$
```

Example matches:

```text
COMMENTS 31-12-2025
COMMENTS 30-12-2025
COMMENTS 03-01-2025
```

If multiple comment columns exist, the **last** matching one is used.

## Notes

- Make sure `header_row` matches the actual header row in the workbook.
- Make sure `sheet_name` matches the actual worksheet name exactly.
- Some workbooks may contain hidden trailing spaces in sheet names.
- Because the workbook is overwritten in place, keep a backup if needed.
- Merged cells are handled carefully, but complex layouts should still be visually checked after execution.

## Example result

If the original comments column is in column `Z`:

- `Z` becomes the original hidden column.
- `AA` becomes the visible duplicated column.

## License

MIT

## Why this changed
Your current README still says the script creates a new _updated file, but both attached scripts now save directly back into the same workbook with wb.save(input_file). It also needed to describe the updated column placement logic, where the hidden original remains on the left and the new visible copy is inserted on the right.

---------------------------------
21-04-2026 1000

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