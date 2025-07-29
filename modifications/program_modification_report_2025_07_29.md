
# 📝 Program Modification Report

**Date:** 2025-07-29  
**Project:** Hojo
**Author:** Zwe Htet Naing  

---

## Client Requests & Applied Fixes

---

### 1. Change mismatch highlight color from yellow to pink  

**Fix:**  
Updated the cell fill color used for highlighting mismatches.

**Before:**
```python
YELLOW_FILL = PatternFill(patternType="solid", fgColor='FFFF00')
```

**After:**
```python
PINK_FILL = PatternFill(patternType="solid", fgColor='FFC0CB')
```

---

### 2. Ignore "歳" when comparing age values  

**Fix:**  
Modified the normalization function to remove the "歳" suffix before comparison.

**Code Updated in `normalize_value` Method:**
```python
.replace('歳', '')
```

---

### 3. Compare rows using column 5 (unordered rows)  

**Fix:**  
Implemented a new method `get_row_headers()` to match rows based on combined values from column 1 and column 5.

#### New Method Added:
```python
def get_row_headers(sheet1, sheet2):
    try:
        headers1 = {}
        headers2 = {}
        current_main_header_v1 = None
        current_main_header_v2 = None

        for row in range(8, min(sheet1.max_row, sheet2.max_row) + 1):
            main_header_v1 = sheet1.cell(row, 1).value
            main_header_v2 = sheet2.cell(row, 1).value

            sub_header_v1 = sheet1.cell(row, 5).value
            sub_header_v2 = sheet2.cell(row, 5).value

            current_main_header_v1 = main_header_v1 if main_header_v1 else current_main_header_v1
            current_main_header_v2 = main_header_v2 if main_header_v2 else current_main_header_v2

            combined1 = f"{current_main_header_v1} - {sub_header_v1}" if sub_header_v1 and current_main_header_v1 else current_main_header_v1
            combined2 = f"{current_main_header_v2} - {sub_header_v2}" if sub_header_v2 and current_main_header_v2 else current_main_header_v2

            if combined1 and combined1 not in headers1.values():
                headers1[row] = combined1
            if combined2 and combined2 not in headers2.values():
                headers2[row] = combined2

        header_pairs = [
            (col1, col2)
            for col1, h1 in headers1.items()
            for col2, h2 in headers2.items()
            if h1 == h2
        ]

        return header_pairs, headers1, headers2

    except Exception as e:
        logging.error(f"Error processing header pairs: {str(e)}")
        return [], {}, {}
```

#### Update in `compare_excel_file`:
```python
if sheets1.get(sheet_id) == "補助調書2":
    header_pairs, headers1, headers2 = get_row_headers(sheet1, sheet2)
    row_pairs = header_pairs
    col_start = 5
else:
    row_pairs = [(row, row) for row in range(1, row_max + 1)]
    col_start = 1

for row1, row2 in row_pairs:
    for col in range(col_start, col_max + 1):
        # comparison logic
```

---

## Summary

- Color changed to pink for mismatched cells  
- Age comparisons now ignore "歳" suffix  
- Rows matched by column 5 values using custom matching logic  

---

_This report is intended for internal documentation and client transparency._
