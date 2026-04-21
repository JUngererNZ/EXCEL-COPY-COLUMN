You are an expert Python developer specializing in Excel automation using `openpyxl`.

I have an Excel file where **Sheet1** contains data. The last column has a header matching the pattern **"COMMENTS DD-MM-YYYY"** (i.e., the word "COMMENTS" followed by a date in DD-MM-YYYY format — the exact date may vary).

Write a complete Python script that does the following:

1. **Locate the target column** — scan the headers in Sheet1 to find the last column whose header matches the pattern `COMMENTS \d{2}-\d{2}-\d{4}` (use regex).

2. **Copy the entire column** — capture all cell values, formatting (fonts, fills, borders, alignment, number formats), and column width from the target column.

3. **Insert a new blank column** immediately to the left of the target column, shifting the target column one position to the right.

4. **Paste the copied column data** — write all copied values and formatting into the newly inserted column. The new column should be an exact replica of the original.

5. **Hide the original column** (now shifted one to the right) by setting its `hidden` property to `True`.

**Requirements:**
- Use `openpyxl` for all Excel manipulation
- Preserve merged cells if encountered (handle gracefully without crashing)
- Accept the Excel filename as a variable at the top of the script so it's easy to change
- Save the modified file with a `_updated` suffix appended to the original filename
- Include clear inline comments explaining each step