Function MyFunc(param1 As Variant, param2 As Variant)
  ' Explicit data type declaration for parameters
  If IsEmpty(param1) Or IsEmpty(param2) Then
    Err.Raise 13, , "Parameters cannot be empty"
  ElseIf VarType(param1) <> vbString Or VarType(param2) <> vbString Then
    Err.Raise 13, , "Parameters must be strings"
  End If
  ' ... function logic ...
End Function