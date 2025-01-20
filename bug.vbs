Function MyFunc(param1, param2)
  ' Missing explicit data type declaration for parameters
  If IsEmpty(param1) Or IsEmpty(param2) Then
    Err.Raise 13, , "Parameters cannot be empty"
  End If
  ' ... function logic ...
End Function