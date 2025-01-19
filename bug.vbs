Function MyFunction(param)
  If IsEmpty(param) Then
    Err.Raise vbError, , "Parameter cannot be empty"
  End If
  ' ... rest of the function
End Function