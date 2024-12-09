The solution involves using error handling to gracefully manage the situation where the object or method might not exist.  Here's the improved code:

```vbscript
On Error Resume Next

Dim objExcel
Set objExcel = CreateObject("Excel.Application")

If Err.Number = 0 Then
  MsgBox objExcel.ActiveWorkbook.Name
Else
  MsgBox "Error accessing Excel: " & Err.Description
  Err.Clear
End If

Set objExcel = Nothing
```

This version checks `Err.Number` after attempting to access the object. If an error occurred (Excel not installed, file not found, etc.), an appropriate error message is displayed instead of a program crash.  Remember to always release objects using `Set objExcel = Nothing`.