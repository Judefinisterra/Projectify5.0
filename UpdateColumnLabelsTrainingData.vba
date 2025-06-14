Sub AddColumnLabelsToTrainingData()
    Dim ws As Worksheet
    Dim cell As Range
    Dim lastRow As Long
    Dim originalText As String
    Dim updatedText As String
    Dim processedCount As Long
    Dim totalChanges As Long
    
    ' Error handling
    On Error GoTo ErrorHandler
    
    ' Get the Training data worksheet
    Set ws = ThisWorkbook.Worksheets("Training data")
    
    ' Find the last row with data in column B
    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).row
    
    ' Initialize counters
    processedCount = 0
    totalChanges = 0
    
    ' Show progress
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Debug.Print "Starting to process Training data column B..."
    Debug.Print "Last row with data: " & lastRow
    
    ' Loop through each cell in column B
    For Each cell In ws.Range("B1:B" & lastRow)
        If Not IsEmpty(cell.value) And cell.value <> "" Then
            originalText = cell.value
            updatedText = AddColumnLabelsToRowParams(originalText)
            
            ' Only update if there were changes
            If originalText <> updatedText Then
                cell.value = updatedText
                totalChanges = totalChanges + 1
                Debug.Print "Updated cell B" & cell.row & ": Found and updated row parameters"
            End If
            
            processedCount = processedCount + 1
            
            ' Show progress every 100 rows
            If processedCount Mod 100 = 0 Then
                Debug.Print "Processed " & processedCount & " cells..."
            End If
        End If
    Next cell
    
    ' Restore Excel settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    ' Show completion message
    MsgBox "Processing complete!" & vbNewLine & _
           "Processed " & processedCount & " cells" & vbNewLine & _
           "Updated " & totalChanges & " cells with row parameters", _
           vbInformation, "Column Labels Added"
    
    Debug.Print "Processing complete! Updated " & totalChanges & " cells out of " & processedCount & " processed."
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Error occurred: " & Err.Description, vbCritical, "Error"
    Debug.Print "Error: " & Err.Description
End Sub

Function AddColumnLabelsToRowParams(inputText As String) As String
    Dim result As String
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    Dim i As Long
    
    result = inputText
    
    ' Create regex object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = False
    
    ' Pattern to match row parameters like row1="...", row2="...", etc.
    regex.pattern = "(row\d+\s*=\s*""[^""]*"")"
    
    Set matches = regex.Execute(result)
    
    ' Process matches in reverse order to avoid position shifts
    For i = matches.count - 1 To 0 Step -1
        Set match = matches(i)
        Dim originalRowParam As String
        Dim updatedRowParam As String
        
        originalRowParam = match.value
        updatedRowParam = ProcessSingleRowParam(originalRowParam)
        
        ' Replace the original with updated version
        result = Left(result, match.FirstIndex) & updatedRowParam & Mid(result, match.FirstIndex + match.Length + 1)
    Next i
    
    AddColumnLabelsToRowParams = result
End Function

Function ProcessSingleRowParam(rowParam As String) As String
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    Dim rowValue As String
    Dim updatedRowValue As String
    
    ' Create regex to extract the value between quotes
    Set regex = CreateObject("VBScript.RegExp")
    regex.pattern = "row(\d+)\s*=\s*""([^""]*)"""
    
    Set matches = regex.Execute(rowParam)
    
    If matches.count > 0 Then
        Set match = matches(0)
        Dim rowNumber As String
        rowNumber = match.SubMatches(0)
        rowValue = match.SubMatches(1)
        
        ' Process the row value to add column labels
        updatedRowValue = AddColumnLabelsToRowValue(rowValue)
        
        ' Reconstruct the row parameter
        ProcessSingleRowParam = "row" & rowNumber & "=""" & updatedRowValue & """"
    Else
        ' Return original if pattern doesn't match
        ProcessSingleRowParam = rowParam
    End If
End Function

Function AddColumnLabelsToRowValue(rowValue As String) As String
    Dim segments() As String
    Dim result As String
    Dim i As Long
    Dim columnLabels As Variant
    
    ' Define the column labels
    columnLabels = Array("(D)", "(L)", "(C1)", "(C2)", "(C3)", "(C4)", "(C5)", "(C6)", "(C7)", "(Y1)", "(Y2)", "(Y3)", "(Y4)", "(Y5)", "(Y6)")
    
    ' Split the row value by "|"
    segments = Split(rowValue, "|")
    
    ' Process each segment and add appropriate labels
    result = ""
    For i = 0 To UBound(segments)
        If i < UBound(columnLabels) + 1 Then
            ' Add the column label to the segment
            If segments(i) <> "" Then
                result = result & segments(i) & columnLabels(i)
            Else
                result = result & columnLabels(i)
            End If
        Else
            ' If there are more segments than labels, just add the segment as-is
            result = result & segments(i)
        End If
        
        ' Add the separator "|" except for the last segment
        If i < UBound(segments) Then
            result = result & "|"
        End If
    Next i
    
    AddColumnLabelsToRowValue = result
End Function

' Helper subroutine to test the function on a specific cell
Sub TestOnSpecificCell()
    Dim ws As Worksheet
    Dim testCell As Range
    Dim originalText As String
    Dim updatedText As String
    
    Set ws = ThisWorkbook.Worksheets("Training data")
    
    ' Test on cell B10 (you can change this)
    Set testCell = ws.Range("B10")
    
    originalText = testCell.value
    updatedText = AddColumnLabelsToRowParams(originalText)
    
    Debug.Print "Original: " & originalText
    Debug.Print "Updated:  " & updatedText
    
    ' Uncomment the line below to actually update the cell
    ' testCell.Value = updatedText
End Sub

