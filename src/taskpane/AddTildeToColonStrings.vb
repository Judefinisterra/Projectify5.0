Sub AddTildeToColonStrings()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String
    Dim modifiedValue As String
    Dim changesMade As Boolean
    
    ' Set the worksheet to "Training Data"
    Set ws = ThisWorkbook.Worksheets("Training Data")
    
    ' Find the last row with data in column B
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    ' Loop through each row starting from row 2
    For i = 2 To lastRow
        cellValue = ws.Cells(i, "B").Value
        changesMade = False
        
        ' Skip empty cells
        If Len(Trim(cellValue)) > 0 Then
            ' Process the cell value
            modifiedValue = AddTildeToColonStringsInRow(cellValue, changesMade)
            
            ' If changes were made, write to column C and highlight
            If changesMade Then
                ws.Cells(i, "C").Value = modifiedValue
                ws.Cells(i, "C").Interior.Color = RGB(255, 192, 0) ' Orange highlighting
            End If
        End If
    Next i
    
    MsgBox "Processing complete! Modified rows are highlighted in orange in column C."
End Sub

Function AddTildeToColonStringsInRow(inputRow As String, ByRef changesMade As Boolean) As String
    Dim result As String
    Dim parts As Variant
    Dim i As Integer
    Dim part As String
    Dim pipePos As Integer
    Dim beforePipe As String
    Dim afterPipe As String
    
    result = inputRow
    changesMade = False
    
    ' Find the pipe section in the row (everything after the first "|")
    pipePos = InStr(result, "|")
    
    If pipePos > 0 Then
        beforePipe = Left(result, pipePos - 1)
        afterPipe = Mid(result, pipePos)
        
        ' Split the pipe section by "|"
        parts = Split(afterPipe, "|")
        
        ' Process each part
        For i = 0 To UBound(parts)
            part = Trim(parts(i))
            
            ' Check if the part ends with ":" and doesn't already start with "~"
            If Len(part) > 0 And Right(part, 1) = ":" And Left(part, 1) <> "~" Then
                parts(i) = "~" & parts(i)
                changesMade = True
            End If
        Next i
        
        ' Reconstruct the row if changes were made
        If changesMade Then
            result = beforePipe & Join(parts, "|")
        End If
    End If
    
    AddTildeToColonStringsInRow = result
End Function 


