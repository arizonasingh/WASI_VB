Attribute VB_Name = "WASI_Age_35_44"
Sub WASI_II_35_44()

Dim age, BD, VC, MR, SI, T1, T2, T3, T4, VCI, PRI, FSIQ4, FSIQ2 As Integer
Dim PRank, CI, WISC_CI90, WISC_CI68, WAIS_CI90, WAIS_CI68 As Variant

age = Range("B1").Value
BD = Range("B7").Value
VC = Range("B8").Value
MR = Range("B9").Value
SI = Range("B10").Value

'Total Raw Score to T score Conversion
Worksheets("T_Scores").Activate
Dim rownum As Integer
Dim cell As Range

'BD T score
For Each cell In Range("Z3:Z63")
    If cell = "-" Then
        'Do nothing
    ElseIf BD <= cell.Value Then
        rownum = cell.Row
        Exit For
    End If
Next

BD = Cells(rownum, 1) '1st column because that's where T scores are

'VC T score
For Each cell In Range("AA3:AA63")
    If cell = "-" Then
        'Do nothing
    ElseIf VC <= cell.Value Then
        rownum = cell.Row
        Exit For
    End If
Next

VC = Cells(rownum, 1)

'MR T score
For Each cell In Range("AB3:AB63")
    If cell = "-" Then
        'Do nothing
    ElseIf MR <= cell.Value Then
        rownum = cell.Row
        Exit For
    End If
Next

MR = Cells(rownum, 1)

'SI T score
For Each cell In Range("AC3:AC63")
    If cell = "-" Then
        'Do nothing
    ElseIf SI <= cell.Value Then
        rownum = cell.Row
        Exit For
    End If
Next

SI = Cells(rownum, 1)

Worksheets("WASI_II_Raw_Scores").Activate

Range("D7:E7") = BD
Range("C8,E8:F8") = VC
Range("D9:F9") = MR
Range("C10,E10") = SI

'Sum of T scores
T1 = VC + SI
T2 = BD + MR
T3 = BD + VC + MR + SI
T4 = VC + MR

Range("C11") = T1
Range("B17") = T1
Range("D11") = T2
Range("B18") = T2
Range("E11") = T3
Range("B19") = T3
Range("F11") = T4
Range("B20") = T4

'Sum of T Scores to Composite Score Conversion
Worksheets("Composite_Scores").Activate

'VCI
For Each cell In Range("A2:A122")
    If T1 <= cell.Value Then
        rownum = cell.Row
        Exit For
    End If
Next

VCI = Cells(rownum, 2)
Prank = Cells(rownum, 3)
CI = Cells(rownum, 4)

Worksheets("WASI_II_Raw_Scores").Activate
Range("D17") = VCI
Range("E17") =Prank
Range("F17") = CI

Worksheets("Composite_Scores").Activate

'PRI
For Each cell In Range("F2:F122")
    If T2 <= cell.Value Then
        rownum = cell.Row
        Exit For
    End If
Next

PRI = Cells(rownum, 7)
Prank = Cells(rownum, 8)
CI = Cells(rownum, 9)

Worksheets("WASI_II_Raw_Scores").Activate
Range("D18") = PRI
Range("E18") =Prank
Range("F18") = CI

Worksheets("Composite_Scores").Activate

'FSIQ-4
For Each cell In Range("K2:K242")
    If T3 <= cell.Value Then
        rownum = cell.Row
        Exit For
    End If
Next

FSIQ4 = Cells(rownum, 12)
Prank = Cells(rownum, 13)
CI = Cells(rownum, 14)

Worksheets("WASI_II_Raw_Scores").Activate
Range("D19") = FSIQ4
Range("E19") =Prank
Range("F19") = CI

Worksheets("Composite_Scores").Activate

'FSIQ-2
For Each cell In Range("P2:P122")
    If T4 <= cell.Value Then
        rownum = cell.Row
        Exit For
    End If
Next

FSIQ2 = Cells(rownum, 17)
Prank = Cells(rownum, 18)
CI = Cells(rownum, 19)

Worksheets("WASI_II_Raw_Scores").Activate
Range("D20") = FSIQ2
Range("E20") =Prank
Range("F20") = CI

'Range of Expected Scores
Range("B26:C26") = FSIQ4

Worksheets("Range_of_Expected_Scores").Activate

For Each cell In Range("A4:A124")
    If FSIQ4 <= cell.Value Then
        rownum = cell.Row
        Exit For
    End If
Next

WISC_CI90 = Cells(rownum, 2)
WISC_CI68 = Cells(rownum, 3)

For Each cell In Range("E4:E124")
    If FSIQ4 <= cell.Value Then
        rownum = cell.Row
        Exit For
    End If
Next

WAIS_CI90 = Cells(rownum, 6)
WAIS_CI68 = Cells(rownum, 7)

Worksheets("WASI_II_Raw_Scores").Activate
Range("B27") = WISC_CI90
Range("C27") = WISC_CI68
Range("B28") = WAIS_CI90
Range("C28") = WAIS_CI68

End Sub
