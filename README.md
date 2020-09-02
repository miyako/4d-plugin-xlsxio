# 4d-plugin-xlsxio
very simple XLSX library based on [XLSXIO](https://github.com/brechtsanders/xlsxio)

### Platform

| carbon | cocoa | win32 | win64 |
|:------:|:-----:|:---------:|:---------:|
||<img src="https://cloud.githubusercontent.com/assets/1725068/22371562/1b091f0a-e4db-11e6-8458-8653954a7cce.png" width="24" height="24" />|<img src="https://cloud.githubusercontent.com/assets/1725068/22371562/1b091f0a-e4db-11e6-8458-8653954a7cce.png" width="24" height="24" />|<img src="https://cloud.githubusercontent.com/assets/1725068/22371562/1b091f0a-e4db-11e6-8458-8653954a7cce.png" width="24" height="24" />|

### Version

<img width="32" height="32" src="https://user-images.githubusercontent.com/1725068/73986501-15964580-4981-11ea-9ac1-73c5cee50aae.png"> <img src="https://user-images.githubusercontent.com/1725068/73987971-db2ea780-4984-11ea-8ada-e25fb9c3cf4e.png" width="32" height="32" />

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
