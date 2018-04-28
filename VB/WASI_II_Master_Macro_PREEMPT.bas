Attribute VB_Name = "Master_Macro_PREEMPT"
Sub WASI_II_Master_Macro_PREEMPT()

Worksheets("WASI_II_Raw_Scores").Activate
Dim age As Integer
age = Range("B1")
Dim cell As Range

If age > 30 Then
 MsgBox "Invalid age for PREEMPT study"
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
End If

Worksheets("WASI_II_Raw_Scores").Activate
Dim ID as Variant
ID = Application.InputBox("What is the Participant ID", "Participant ID (number only)", 1)

'Subtests
Dim BD_raw, BD_T, VO_raw, VO_T, MR_raw, MR_T, SM_raw, SM_T As Integer

BD_raw = Range("B7").Value
BD_T = Range("D7").Value
VO_raw = Range("B8").Value
VO_T = Range("C8").Value
MR_raw = Range("B9").Value
MR_T = Range("D9").Value
SM_raw = Range("B10").Value
SM_T = Range("C10").Value

'Scale, Composite, and Percentile Rank
Dim VC_T, PR_T, FSIQ4_T, FSIQ2_T, VCI, PRI, FSIQ4, FSIQ2, VC_Perc, PR_Perc, FSIQ4_Perc, FSIQ2_Perc As Integer

VC_T = Range("B17").Value
PR_T = Range("B18").Value
FSIQ4_T = Range("B19").Value
FSIQ2_T = Range("B20").Value
VCI = Range("D17").Value
PRI = Range("D18").Value
FSIQ4 = Range("D19").Value
FSIQ2 = Range("D20").Value
VC_Perc = Range("E17").Value
PR_Perc = Range("E18").Value
FSIQ4_Perc = Range("E19").Value
FSIQ2_Perc = Range("E20").Value

'Confidence Intervals
Dim VC_CI1, VC_CI2, PR_CI1, PR_CI2, FSIQ4_CI1, FSIQ4_CI2, FSIQ2_CI1, FSIQ2_CI2, WISC90_CI1, WISC90_CI2, WISC68_CI1, WISC68_CI2, WAIS90_CI1, WAIS90_CI2, WAIS68_CI1, WAIS68_CI2 As String

VC_CI1 = Left(Range("F17"), InStr(Range("F17"), "-") - 1)
VC_CI2 = Right(Range("F17"), InStr(Range("F17"), "-"))
PR_CI1 = Left(Range("F18"), InStr(Range("F18"), "-") - 1)
PR_CI2 = Right(Range("F18"), InStr(Range("F18"), "-"))
FSIQ4_CI1 = Left(Range("F19"), InStr(Range("F19"), "-") - 1)
FSIQ4_CI2 = Right(Range("F19"), InStr(Range("F19"), "-"))
FSIQ2_CI1 = Left(Range("F20"), InStr(Range("F20"), "-") - 1)
FSIQ2_CI2 = Right(Range("F20"), InStr(Range("F20"), "-"))
WISC90_CI1 = Left(Range("B27"), InStr(Range("B27"), "-") - 1)
WISC90_CI2 = Right(Range("B27"), InStr(Range("B27"), "-"))
WISC68_CI1 = Left(Range("C27"), InStr(Range("C27"), "-") - 1)
WISC68_CI2 = Right(Range("C27"), InStr(Range("C27"), "-"))
WAIS90_CI1 = Left(Range("B28"), InStr(Range("B28"), "-") - 1)
WAIS90_CI2 = Right(Range("B28"), InStr(Range("B28"), "-"))
WAIS68_CI1 = Left(Range("C28"), InStr(Range("C28"), "-") - 1)
WAIS68_CI2 = Right(Range("C28"), InStr(Range("C28"), "-"))

'Transfering data between sheets to PREEEMPT project
Worksheets("PREEMPT").Activate
Dim lastrow As long
lastrow = Cells.Find("*",SearchOrder:=xlByRows,SearchDirection:=xlPrevious).Row + 1 'Retrieving last row with data in it so you can add data to next row

'Based solely on PREEMPT REDCap data dictionary architecture
Range("A" & lastrow) = ID
Range("B" & lastrow) = "day_1_arm_1"
Range("C" & lastrow) = DateValue(Application.InputBox("WASI-II Administration Date?","Administration Date (mm-dd-yyyy)"))
Range("D" & lastrow) = BD_raw
Range("E" & lastrow) = BD_T
Range("F" & lastrow) = VO_raw
Range("G" & lastrow) = VO_T
Range("H" & lastrow) = MR_raw
Range("I" & lastrow) = MR_T
Range("J" & lastrow) = SM_raw
Range("K" & lastrow) = SM_T
Range("L" & lastrow) = VC_T
Range("M" & lastrow) = VCI
Range("N" & lastrow) = VC_Perc
Range("O" & lastrow) = Abs(CInt(VC_CI1))
Range("P" & lastrow) = Abs(CInt(VC_CI2))
Range("Q" & lastrow) = PR_T
Range("R" & lastrow) = PRI
Range("S" & lastrow) = PR_Perc
Range("T" & lastrow) = Abs(CInt(PR_CI1))
Range("U" & lastrow) = Abs(CInt(PR_CI2))
Range("V" & lastrow) = FSIQ4_T
Range("W" & lastrow) = FSIQ4
Range("X" & lastrow) = FSIQ4_Perc
Range("Y" & lastrow) = Abs(CInt(FSIQ4_CI1))
Range("Z" & lastrow) = Abs(CInt(FSIQ4_CI2))
Range("AA" & lastrow) = FSIQ2_T
Range("AB" & lastrow) = FSIQ2
Range("AC" & lastrow) = FSIQ2_Perc
Range("AD" & lastrow) = Abs(CInt(FSIQ2_CI1))
Range("AE" & lastrow) = Abs(CInt(FSIQ2_CI2))
Range("AF" & lastrow) = Abs(CInt(WISC90_CI1))
Range("AG" & lastrow) = Abs(CInt(WISC90_CI2))
Range("AH" & lastrow) = Abs(CInt(WISC68_CI1))
Range("AI" & lastrow) = Abs(CInt(WISC68_CI2))
Range("AJ" & lastrow) = Abs(CInt(WAIS90_CI1))
Range("AK" & lastrow) = Abs(CInt(WAIS90_CI2))
Range("AL" & lastrow) = Abs(CInt(WAIS68_CI1))
Range("AM" & lastrow) = Abs(CInt(WAIS68_CI2))

Worksheets("WASI_II_Raw_Scores").Activate

End Sub