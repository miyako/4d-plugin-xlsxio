//%attributes = {}
$path:=System folder:C487(Desktop:K41:16)+"SB_data.xlsx"

ARRAY TEXT:C222($names;0)
XLSX SHEET NAMES ($path;$names)

For ($i;1;Size of array:C274($names))
	$json_e:=$names{$i}
	$json:=XLSX TO JSON ($path;XLSXIOREAD_SKIP_EMPTY_ROWS;$json_e)
End for 

  //all values are returned as string
$json:=XLSX TO JSON ($path;XLSXIOREAD_SKIP_EMPTY_ROWS;$json_e)

$path:=System folder:C487(Desktop:K41:16)+"SB_data_cp.xlsx"

JSON TO XLSX ($path;$json)