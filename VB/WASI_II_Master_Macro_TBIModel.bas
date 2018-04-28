Attribute VB_Name = "Master_Macro_TBIModel"
Sub WASI_II_Master_Macro_TBIModel()

Worksheets("WASI_II_Raw_Scores").Activate
Dim age As Integer
age = Range("B1")
Dim cell As Range

If age > 45 Then
 MsgBox "Invalid age for TBI Model study"
 Exit Sub
End If

For Each cell In Range("B7:B10")
    If IsEmpty(cell) Then
        MsgBox "The program requires all subtests to have a raw score entered!"
        Exit Sub
    End If
Next

If age <= 19 Then
  Call WASI_II_17_19
ElseIf age <=24 Then
  Call WASI_II_20_24
ElseIf age <= 29 Then
  Call WASI_II_25_29
ElseIf age <= 34 Then
  Call WASI_II_30_34
ElseIf age <= 44 Then
  Call WASI_II_35_44
ElseIf age <= 54 Then
  Call WASI_II_45_54
End If

End Sub