# 4d-plugin-xlsxio
very simply XLSX library based on [XLSXIO](https://github.com/brechtsanders/xlsxio)

### Platform

| carbon | cocoa | win32 | win64 |
|:------:|:-----:|:---------:|:---------:|
|<img src="https://cloud.githubusercontent.com/assets/1725068/22371562/1b091f0a-e4db-11e6-8458-8653954a7cce.png" width="24" height="24" />|<img src="https://cloud.githubusercontent.com/assets/1725068/22371562/1b091f0a-e4db-11e6-8458-8653954a7cce.png" width="24" height="24" />|||

### Version

<img src="https://cloud.githubusercontent.com/assets/1725068/18940649/21945000-8645-11e6-86ed-4a0f800e5a73.png" width="32" height="32" /> <img src="https://cloud.githubusercontent.com/assets/1725068/18940648/2192ddba-8645-11e6-864d-6d5692d55717.png" width="32" height="32" />

## Syntax

```
json:=XLSX TO JSON (path;options;error)
```

Parameter|Type|Description
------------|------------|----
path|TEXT|location of ``.XLSX`` document
options|LONGINT|options
error|TEXT|error in ``JSON``
json|TEXT|values in ``JSON``

```
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

## Remarks

Only the first sheet is exported with ``JSON TO XLSX``

All values are imported as text with ``XLSX TO JSON``

Values can be text, number, or bool with ``JSON TO XLSX``, but you would need to use ``Collection`` for that.

For more information see https://github.com/brechtsanders/xlsxio
