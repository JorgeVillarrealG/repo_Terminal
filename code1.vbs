Sub grade()
Dim grade As Double
grade = Cells(2, 2).Value
If grade >= 90 Then
    Cells(2, 3).Interior.ColorIndex = 4
    Cells(2, 3).Value = "Pass"
    Cells(2, 4).Value = "A"
    ElseIf grade < 89 And grade >= 80 Then
        Cells(2, 3).Interior.ColorIndex = 4
        Cells(2, 3).Value = "Pass"
        Cells(2, 4).Value = "B"
            ElseIf grade < 70 And grade >= 79 Then
                Cells(2, 3).Interior.ColorIndex = 6
                Cells(2, 3).Value = "Warning"
                Cells(2, 4).Value = "C"
Else
    Cells(2, 3).Interior.ColorIndex = 3
    Cells(2, 3).Value = "Fail"
    Cells(2, 4).Value = "F"
End If
  
End Sub
Sub Reset()
LastRow = Cells(Rows.Count, 2).End(xlUp).Row + 1
Cells(LastRow, 2).Value = Cells(2, 2).Value
Cells(LastRow, 3).Value = Cells(2, 3).Value
Cells(LastRow, 4).Value = Cells(2, 4).Value

Range("B2:D2").Clear


End Sub
