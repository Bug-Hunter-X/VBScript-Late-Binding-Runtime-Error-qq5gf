Late Binding in VBScript can lead to runtime errors if the object or method doesn't exist.  Consider this example:

```vbscript
Dim objExcel
Set objExcel = CreateObject("Excel.Application")

'This line will fail if Excel is not installed
MsgBox objExcel.ActiveWorkbook.Name
```

If Excel isn't installed,  `objExcel.ActiveWorkbook.Name` will throw a runtime error because `objExcel` will be `Nothing`.