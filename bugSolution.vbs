Function overloading is not directly supported.  Instead, use conditional statements within a single function to handle different parameters:

Function MyFunction(param1, param2)
  If IsMissing(param2) Then
    ' Handle case with only param1
    MsgBox "One parameter: " & param1
  Else
    ' Handle case with param1 and param2
    MsgBox "Two parameters: " & param1 & ", " & param2
  End If
End Function

MyFunction 10
MyFunction 10, 20