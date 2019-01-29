Sub sheetLoop():
    Dim sheets As Worksheet
    Application.ScreenUpdating = False
    For Each sheets In Worksheets
        sheets.Select
        Call dataProcess
    Next
    Application.ScreenUpdating = True
End Sub

Sub dataProcess():
    Dim stock As String
    Dim lRow As Long
    Dim volume As Double
    Dim outputLocation As Integer
    Dim copyRange As String
    
    outputLocation = 1
    volume = 0
    
    lRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    copyRange = "A2:A" & lRow
    
    For i = 2 To lRow
        If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
            stock = Cells(i, 1).Value
            volume = volume + CDbl(Cells(i, 7).Value)
        ElseIf (i > 2 And (Cells(i + 1, 1).Value <> Cells(i, 1).Value)) Then
            Cells(outputLocation, 9).Value = stock
            Cells(outputLocation, 10).Value = volume
            outputLocation = outputLocation + 1
            volume = 0
        End If
    Next i
End Sub
