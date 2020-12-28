![platform](https://img.shields.io/static/v1?label=platform&message=osx-64%20|%20win-32%20|%20win-64&color=blue)
![version](https://img.shields.io/badge/version-16%2B-8331AE)
[![license](https://img.shields.io/github/license/miyako/4d-plugin-xlsxio/)](LICENSE)
![downloads](https://img.shields.io/github/downloads/miyako/4d-plugin-xlsxio/total)

# 4d-plugin-xlsxio
very simple XLSX library based on [XLSXIO](https://github.com/brechtsanders/xlsxio)

To use on v16 or v17, move manifest.json to contents.

## Syntax

```4d
json:=XLSX TO JSON (path;options;error)
```

Parameter|Type|Description
------------|------------|----
path|TEXT|location of ``.XLSX`` document
options|LONGINT|options
error|TEXT|on outout, error in ``JSON``; on input the sheet name to read
json|TEXT|values in ``JSON``

```4d
XLSX SHEET NAMES (path;names)
```

Parameter|Type|Description
------------|------------|----
path|TEXT|location of ``.XLSX`` document
names|TEXT ARRAY|sheet names


```4d
JSON TO XLSX (path;json;row_height;detection_rows;error)
```

Parameter|Type|Description
------------|------------|----
path|TEXT|location of ``.XLSX`` document
json|TEXT|values in ``JSON``
json|LONGINT|row height
json|LONGINT|number of rows used to detect column widths
error|TEXT|error in ``JSON``

* Options for ``XLSX TO JSON``

```c
XLSXIOREAD_SKIP_NONE 0
XLSXIOREAD_SKIP_EMPTY_ROWS 1
XLSXIOREAD_SKIP_EMPTY_CELLS 2
XLSXIOREAD_SKIP_ALL_EMPTY 3
XLSXIOREAD_SKIP_EXTRA_CELLS 4
```

## Remarks

Only the first sheet is exported with ``JSON TO XLSX``

All values are imported as text with ``XLSX TO JSON``

Values can be text, number, or bool with ``JSON TO XLSX``, but you would need to use ``Collection`` for that.

For more information see https://github.com/brechtsanders/xlsxio

## Examples

```4d
$path:=System folder(Desktop)+"sample.xlsx"

  //all values are returned as string
$json:=XLSX TO JSON ($path;XLSXIOREAD_SKIP_EMPTY_ROWS;$json_e)

$xlsx:=JSON Parse($json)
$_err:=JSON Parse($json_e)

$row_height:=1  //set row height
$detection_rows:=10  //how many rows to buffer to detect column widths

$path:=System folder(Desktop)+"sample-copy.xlsx"
  //only 1 sheet is supported
JSON TO XLSX ($path;$json;$row_height;$detection_rows;$json_e)

OPEN URL($path;"Microsoft Excel")
```
