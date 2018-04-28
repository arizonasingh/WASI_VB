Attribute VB_Name = "Master_Macro_BL2andPTSD"
Sub WASI_II_Master_Macro_Bl2andPTSD()

Worksheets("WASI_II_Raw_Scores").Activate
Dim age As Integer
age = Range("B1")
Dim cell As Range

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