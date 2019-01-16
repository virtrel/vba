Public Sub removeDuplicateLines()
    Dim lastLine As Long
    Dim line As Long

    lastLine = Range("A" & Rows.Count).End(xlUp).Row

    For line = lastLine To 1 Step -1
        If WorksheetFunction.CountIf(Range("A1:A" & line), Range("A" & line)) > 1 Then
            Rows(line).EntireRow.Delete
        End If
    Next
End Sub