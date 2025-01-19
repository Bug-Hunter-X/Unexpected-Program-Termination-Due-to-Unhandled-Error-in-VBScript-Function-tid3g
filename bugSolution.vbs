Function MyFunction(param)
  On Error Resume Next
  If IsEmpty(param) Then
    Err.Raise vbError, , "Parameter cannot be empty"
  End If
  On Error GoTo 0
  ' ... rest of the function
End Function

Sub TestMyFunction()
  On Error GoTo ErrorHandler
  Dim result
  result = MyFunction("")
  If Err.Number <> 0 Then
    MsgBox "Error: " & Err.Description
  Else
    MsgBox "Function executed successfully."
  End If
  Exit Sub

ErrorHandler:
  MsgBox "An error occurred: " & Err.Number & " - " & Err.Description
End Sub

TestMyFunction