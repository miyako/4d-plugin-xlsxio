![version](https://img.shields.io/badge/version-17%2B-3E8B93)
![platform](https://img.shields.io/static/v1?label=platform&message=mac-intel%20|%20mac-arm%20|%20win-64&color=blue)
[![license](https://img.shields.io/github/license/miyako/4d-plugin-xlsxio)](LICENSE)
![downloads](https://img.shields.io/github/downloads/miyako/4d-plugin-xlsxio/total)

XLSXIO reads and writes Excel `.xlsx` workbooks from 4D, using the open-source [xlsxio](https://github.com/brechtsanders/xlsxio) library under the hood. It exposes three commands: one to read a workbook into JSON, one to write a JSON structure out to a new workbook, and one to list a workbook's sheet names. All spreadsheet data crosses the boundary as JSON text (a `Text` parameter you parse/build yourself), not as a native 4D object or collection.

| Command | Returns | Purpose |
|---|---|---|
| [`XLSX TO JSON`](#xlsx-to-json) | Text | Read a workbook (one sheet or all sheets) into a JSON string |
| [`JSON TO XLSX`](#json-to-xlsx) | *(none)* | Write a JSON structure out to a new `.xlsx` workbook |
| [`XLSX SHEET NAMES`](#xlsx-sheet-names) | *(none)* | List the sheet names in a workbook into a text array |

**Platforms:** Windows (x64) and macOS (universal: Apple Silicon + Intel), macOS 10.13 or later.

---

## Requirements & platform notes

- **Cell values always come back as text.** `XLSX TO JSON` returns every cell — numbers, dates, booleans, whatever the spreadsheet cell actually contains — as a JSON string. There's no numeric/date/boolean typing on the read side; if you need a number or a date, convert it yourself after parsing the JSON.
- **`JSON TO XLSX` only ever writes `sheets[0]`.** Even though the input JSON's `sheets` key is an array, everything beyond the first entry is silently ignored — this is not a multi-sheet writer, despite accepting an array-shaped input. See [`JSON TO XLSX`](#json-to-xlsx) below.
- **`XLSX TO JSON`'s third parameter does double duty.** It's both the optional sheet-name filter you pass *in* and the destination for the error JSON written back *out* — the same variable is overwritten by the call. See [`XLSX TO JSON`](#xlsx-to-json) below.
- **Not every command reports errors.** `XLSX TO JSON` and `JSON TO XLSX` both have an error-output parameter; `XLSX SHEET NAMES` has none at all — its only observable failure signal is an empty array. See [Error handling & troubleshooting](#error-handling--troubleshooting).
- **Row/cell filtering flags are a bitmask**, defined in the plugin's `constants.json` (theme "XLSXIO options") and combinable with `+` or bitwise OR:

  | Constant | Value | Effect |
  |---|---|---|
  | `XLSXIOREAD_SKIP_NONE` | 0 | Don't skip anything |
  | `XLSXIOREAD_SKIP_EMPTY_ROWS` | 1 | Skip empty rows (note: a row can *look* empty while a cell in it still holds data — see [Error handling](#error-handling--troubleshooting)) |
  | `XLSXIOREAD_SKIP_EMPTY_CELLS` | 2 | Skip empty cells within a row |
  | `XLSXIOREAD_SKIP_ALL_EMPTY` | 3 | Both of the above (`1 + 2`) |
  | `XLSXIOREAD_SKIP_EXTRA_CELLS` | 4 | Skip cells to the right of the rightmost header cell |

---

## XLSX TO JSON

### Syntax

```4d
XLSX TO JSON ( path ; options ; sheet ) → jsonOrErrors : Text
```

| Parameter | Type | Description |
|---|---|---|
| `path` | Text | Full path to the `.xlsx` file to read. |
| `options` | Longint | Bitmask of `XLSXIOREAD_SKIP_*` constants (see [Requirements](#requirements--platform-notes)). Pass `XLSXIOREAD_SKIP_NONE` (`0`) for no filtering. |
| `sheet` | Text | **In:** sheet name to read, or an empty/undefined variable to read *every* sheet. **Out:** after the call, this same variable holds the error JSON for this call (see [Description](#description)) — its input value is overwritten. |
| Result | Text | JSON string: `{"sheets":[...]}` on success. If the workbook couldn't even be opened, this is `{}` — check `sheet` (the error parameter) for the reason. |

### Description

Reading with no `sheet` filter (an empty or undefined text variable) walks every sheet in the workbook and returns them all under `"sheets"`, each shaped as:

```json
{"name": "Sheet1", "rows": [{"values": ["a", "1", "TRUE"]}]}
```

Every value in `"values"` is a JSON **string**, regardless of what the underlying cell actually contained — dates, numbers, and booleans all come back as text. Rows are not padded to a fixed width: a row's `"values"` array only has as many entries as cells the underlying reader actually returned for that row, so rows can be different lengths, and — if you pass `XLSXIOREAD_SKIP_EMPTY_CELLS`/`XLSXIOREAD_SKIP_ALL_EMPTY` — the array is no longer reliably column-aligned, since skipped cells simply aren't present rather than being represented as empty placeholders.

Passing a specific `sheet` name reads only that sheet (still wrapped in the same `{"sheets":[{...}]}` shape, just with one entry) and stops as soon as it's done — it doesn't scan the rest of the workbook.

If the workbook can't be opened at all, or its sheet list can't be read, the `sheet` parameter comes back with an error object and `sheets` is absent from the function result entirely (so the result is just `{}`). If a *specific* sheet within the workbook fails to open, reading stops at that point — any sheets already read successfully before the failure are still included in the result, but nothing after it is attempted.

### Example

From the plugin's own test method (`TEST.4dm`):
```4d
$path:=System folder:C487(Desktop:K41:16)+"SB_data.xlsx"

ARRAY TEXT:C222($names; 0)
XLSX SHEET NAMES($path; $names)

For ($i; 1; Size of array:C274($names))
	$json_e:=$names{$i}
	$json:=XLSX TO JSON($path; XLSXIOREAD_SKIP_EMPTY_ROWS; $json_e)
End for 

//all values are returned as string
$json:=XLSX TO JSON($path; XLSXIOREAD_SKIP_EMPTY_ROWS; $json_e)
```

Reading every sheet in one call, with no filtering:
```4d
$path:=System folder:C487(Desktop:K41:16)+"report.xlsx"
$errors:=""
$json:=XLSX TO JSON($path; XLSXIOREAD_SKIP_NONE; $errors)
If ($errors#"{}")
	ALERT:C41("Could not read workbook: "+$errors)
End if 
```

Reading just one named sheet:
```4d
$sheet:="Q3 Summary"
$json:=XLSX TO JSON($path; XLSXIOREAD_SKIP_ALL_EMPTY; $sheet)
// $sheet now holds the error JSON for this call, not "Q3 Summary" anymore
```

---

## JSON TO XLSX

### Syntax

```4d
JSON TO XLSX ( path ; json ; rowHeight ; detectionRows ; errors )
```

| Parameter | Type | Description |
|---|---|---|
| `path` | Text | Full path where the new `.xlsx` file will be written. |
| `json` | Text | JSON text shaped like `{"sheets":[{"name": "...", "rows": [{"values": [...]}]}]}`. Only `sheets[0]` is used — see [Description](#description). |
| `rowHeight` | Longint | Row height in text lines. Pass `0` (or omit) for unspecified/default height. |
| `detectionRows` | Longint | Number of leading rows to buffer in memory for column-width auto-detection. Pass `0` (or omit) to disable auto-detection. |
| `errors` | Text | **Out only.** Error JSON for this call (`{}` if nothing went wrong). Can be omitted if you don't need it. |
| Result | *(none)* | This command has no function result — call it as a statement, not an assignment. |

### Description

`JSON TO XLSX` writes **only the first entry** of the `sheets` array in your input JSON (`sheets[0]`) — even though the input shape mirrors what `XLSX TO JSON` produces for a whole workbook, this command is not the write-side counterpart for multiple sheets at once. If you need to write several sheets, call it once per sheet with a different `path` (or build your own multi-file/multi-call workflow) rather than passing all sheets in one JSON payload.

`rowHeight`, `detectionRows`, and `errors` are all optional trailing parameters — the plugin's own test method calls this command with just `path` and `json`, omitting the other three entirely.

Each entry in a row's `"values"` array must be a JSON string, number, or boolean (or `null`, written as an empty cell) — a nested array or object anywhere in `values` is rejected as an error for that specific cell, but the rest of the row/sheet still gets written.

**Silent no-op case, worth knowing about explicitly:** if `sheets[0]` is present but is `null` or not a JSON object, the command does nothing at all — no file is written, but the `errors` parameter also comes back `{}`, indistinguishable from success. Make sure `sheets[0]` is a genuine object before relying on this command's error output alone to confirm success. See [Error handling](#error-handling--troubleshooting) for the full list of failure cases that *do* produce an error.

### Example

From the plugin's own test method (`TEST.4dm`), continuing the read example above:
```4d
$path:=System folder:C487(Desktop:K41:16)+"SB_data_cp.xlsx"

JSON TO XLSX($path; $json)
```

Writing a sheet built by hand, with explicit row height and error checking:
```4d
$json:="{\"sheets\":[{\"name\":\"Sheet1\",\"rows\":["+\
"{\"values\":[\"Name\",\"Score\"]},"+\
"{\"values\":[\"Ada\",97]},"+\
"{\"values\":[\"Grace\",99]}"+\
"]}]}"

$errors:=""
JSON TO XLSX($path; $json; 20; 0; $errors)
If ($errors#"{}")
	ALERT:C41("Could not write workbook: "+$errors)
End if 
```

---

## XLSX SHEET NAMES

### Syntax

```4d
XLSX SHEET NAMES ( path ; sheetNames )
```

| Parameter | Type | Description |
|---|---|---|
| `path` | Text | Full path to the `.xlsx` file to inspect. |
| `sheetNames` | Array Text | **Out only.** Must be declared as a `Text` array (via `ARRAY TEXT`) — or left undefined — before the call. Populated with one entry per sheet in the workbook, in workbook order. |
| Result | *(none)* | This command has no function result — call it as a statement, not an assignment. |

### Description

This command has **no error-output parameter at all** — it's the one command in this plugin with no way to distinguish "path not found" from "workbook has zero sheets" other than checking whether `sheetNames` came back empty. See [Error handling](#error-handling--troubleshooting).

The array you pass must already be typed as a `Text` array (or be a genuinely undefined variable) — if you pass a variable of any other type, the command silently does not touch it at all; you won't get an error, you'll just find your variable unchanged after the call.

### Example

From the plugin's own test method (`TEST.4dm`):
```4d
$path:=System folder:C487(Desktop:K41:16)+"SB_data.xlsx"

ARRAY TEXT:C222($names; 0)
XLSX SHEET NAMES($path; $names)

For ($i; 1; Size of array:C274($names))
	$json_e:=$names{$i}
	$json:=XLSX TO JSON($path; XLSXIOREAD_SKIP_EMPTY_ROWS; $json_e)
End for 
```

Checking for a workbook that couldn't be read at all:
```4d
ARRAY TEXT:C222($names; 0)
XLSX SHEET NAMES($path; $names)
If (Size of array:C274($names)=0)
	ALERT:C41("No sheets found — check the path or that the file is a valid .xlsx")
End if 
```

---

## Error handling & troubleshooting

- **`XLSX SHEET NAMES` never reports an error — an empty array is your only signal.** There's no error-output parameter for this command; a bad path, a corrupt file, and "this workbook genuinely has zero sheets" all look identical (`Size of array` returns `0`). Always check the array size after calling it.
- **Passing the wrong variable type to `XLSX SHEET NAMES`'s second parameter fails silently.** It must already be a `Text` array (or undefined) going in; any other type and the command does nothing to it, without raising any error.
- **`XLSX TO JSON`'s third parameter is overwritten by the call.** If you're passing a specific sheet name in, capture it in a different variable first if you need it again afterward — the same variable comes back holding the error JSON, not your original sheet name.
- **`XLSXIOREAD_SKIP_EMPTY_ROWS` can hide non-empty cells.** Per the underlying library's own documentation, a row can be treated as "empty" and skipped even if a cell within it actually contains data — don't rely on this flag if you need every populated cell to be visible in the result.
- **`XLSXIOREAD_SKIP_EMPTY_CELLS`/`XLSXIOREAD_SKIP_ALL_EMPTY` break column alignment.** With these flags, a row's `"values"` array only contains the cells that weren't skipped — so `values[2]` is not reliably "column C" once any cells have been skipped. Don't use these flags if you need to map values back to spreadsheet columns by position.
- **`JSON TO XLSX` only writes `sheets[0]`, silently ignoring the rest of the array.** If your JSON has multiple sheets and only one file gets written, this is why — call the command once per sheet.
- **`JSON TO XLSX` can silently do nothing.** If `sheets[0]` is `null` or not a JSON object, no file is written and no error is reported either (`errors` still comes back `{}`). Validate that `sheets[0]` is an actual object before you build the JSON you pass in.
- **Reading stops at the first sheet that fails to open, in "read all sheets" mode.** If sheet 3 of 5 fails, you'll get sheets 1–2 in the result plus an error describing sheet 3's failure — sheets 4–5 are never attempted.
- **Common `JSON TO XLSX` error messages and what they mean:**

  | Error | Cause |
  |---|---|
  | `failed to parse json` | `json` parameter isn't valid JSON. |
  | `sheets[0] is missing` | The `sheets` key is absent from the JSON, or is present but an empty array. |
  | `sheets is not an array` | `sheets` exists but isn't a JSON array. |
  | `sheets[0].name is missing` | `sheets[0]` has no `name` key. |
  | `sheets[0].name must be a string` | `sheets[0].name` exists but isn't a string. |
  | `failed to create xlsx` | Couldn't open `path` for writing (permissions, invalid path, disk full, etc.). |
  | `sheets[0].rows is missing` / `sheets[0].rows is not an array` | `rows` is absent, or isn't an array. |
  | `sheets[0].rows[].values is missing` / `... is not an array` | A specific row (reported by index) has no `values`, or `values` isn't an array. |
  | `sheets[0].rows[].values[] must be null, string, number or bool` | A specific cell (reported by row and column index) is a nested array or object, which isn't supported. |

---

## Quick reference

```4d
// List sheets, then read each one
ARRAY TEXT:C222($names; 0)
XLSX SHEET NAMES($path; $names)
For ($i; 1; Size of array:C274($names))
	$sheet:=$names{$i}
	$json:=XLSX TO JSON($path; XLSXIOREAD_SKIP_EMPTY_ROWS; $sheet)
End for 

// Read the whole workbook in one call
$errors:=""
$json:=XLSX TO JSON($path; XLSXIOREAD_SKIP_NONE; $errors)

// Write sheets[0] of a JSON payload back out
$errors:=""
JSON TO XLSX($outPath; $json; 0; 0; $errors)
```
