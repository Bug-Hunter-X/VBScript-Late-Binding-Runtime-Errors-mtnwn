Early Binding and Error Handling:
```vbscript
On Error GoTo ErrHandler

Dim objExcel As Object
Set objExcel = GetObject(, "Excel.Application")
If Err.Number <> 0 Then
  Set objExcel = CreateObject("Excel.Application")
End If

If Err.Number <> 0 Then
    MsgBox "Error creating Excel object: " & Err.Description
    Exit Sub
End If

' ... use objExcel safely with error checks ...

' ... further error handling ...

ErrHandler:
  MsgBox "An error occurred: " & Err.Number & " - " & Err.Description
  ' ... optional cleanup ...
End Sub
```

The solution utilizes early binding by declaring the object type (`Dim objExcel As Object`). Additionally, it implements error handling (`On Error GoTo ErrHandler`) to catch and manage exceptions gracefully.  The `GetObject` function is used first, attempting to attach to a running instance of Excel, before creating a new instance with `CreateObject`.  This prevents unnecessary Excel process launches.