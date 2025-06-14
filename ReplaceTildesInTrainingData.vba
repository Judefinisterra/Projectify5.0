Sub ReplaceTildesInTrainingData()
    '
    ' ReplaceTildesInTrainingData Macro
    ' Loops through every populated cell in column B of "Training data" worksheet
    ' and replaces "~~" with "~"
    '
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim currentRow As Long
    Dim cellValue As String
    Dim replacementCount As Long
    Dim processedCells As Long
    
    ' Initialize counters
    replacementCount = 0
    processedCells = 0
    
    ' Error handling
    On Error GoTo ErrorHandler
    
    ' Check if "Training data" worksheet exists
    Dim wsExists As Boolean
    wsExists = False
    Dim tempWS As Worksheet
    
    For Each tempWS In ThisWorkbook.Worksheets
        If tempWS.Name = "Training data" Then
            wsExists = True
            Exit For
        End If
    Next tempWS
    
    If Not wsExists Then
        MsgBox "Worksheet 'Training data' not found in this workbook.", vbCritical, "Worksheet Not Found"
        Exit Sub
    End If
    
    ' Get reference to "Training data" worksheet
    Set ws = ThisWorkbook.Worksheets("Training data")
    
    ' Find the last used row in column B
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    ' Check if there's any data
    If lastRow < 1 Then
        MsgBox "No data found in column B of Training data worksheet.", vbInformation
        Exit Sub
    End If
    
    ' Disable screen updating for better performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Display progress message
    Application.StatusBar = "Processing Training data column B..."
    
    ' Loop through each row in column B
    For currentRow = 1 To lastRow
        ' Check if cell is not empty
        If Not IsEmpty(ws.Cells(currentRow, "B").Value) Then
            cellValue = ws.Cells(currentRow, "B").Value
            
            ' Check if cell contains "~~"
            If InStr(cellValue, "~~") > 0 Then
                ' Replace "~~" with "~"
                cellValue = Replace(cellValue, "~~", "~")
                
                ' Update the cell with the new value
                ws.Cells(currentRow, "B").Value = cellValue
                
                ' Increment replacement counter
                replacementCount = replacementCount + 1
                
                ' Log the change (optional - can be commented out for performance)
                Debug.Print "Row " & currentRow & ": Replaced ~~ with ~ in cell B" & currentRow
            End If
            
            ' Increment processed cells counter
            processedCells = processedCells + 1
        End If
        
        ' Update progress every 100 rows
        If currentRow Mod 100 = 0 Then
            Application.StatusBar = "Processing row " & currentRow & " of " & lastRow & "..."
        End If
    Next currentRow
    
    ' Restore Excel settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    
    ' Display completion message
    MsgBox "Process completed!" & vbCrLf & vbCrLf & _
           "Processed cells: " & processedCells & vbCrLf & _
           "Replacements made: " & replacementCount & vbCrLf & _
           "Last row processed: " & lastRow, _
           vbInformation, "Replace Tildes Complete"
    
    Exit Sub
    
ErrorHandler:
    ' Restore Excel settings in case of error
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    
    MsgBox "An error occurred: " & Err.Description & vbCrLf & _
           "Error Number: " & Err.Number, vbCritical, "Error"
End Sub

' Alternative version that works on the active workbook
Sub ReplaceTildesInTrainingDataActiveWorkbook()
    '
    ' Alternative version that works on the currently active workbook
    '
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim currentRow As Long
    Dim cellValue As String
    Dim replacementCount As Long
    Dim processedCells As Long
    
    ' Initialize counters
    replacementCount = 0
    processedCells = 0
    
    ' Error handling
    On Error GoTo ErrorHandler
    
    ' Try to get reference to "Training data" worksheet in active workbook
    Set ws = ActiveWorkbook.Worksheets("Training data")
    
    ' Find the last used row in column B
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    ' Check if there's any data
    If lastRow < 1 Then
        MsgBox "No data found in column B of Training data worksheet.", vbInformation
        Exit Sub
    End If
    
    ' Disable screen updating for better performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Display progress message
    Application.StatusBar = "Processing Training data column B..."
    
    ' Loop through each row in column B
    For currentRow = 1 To lastRow
        ' Check if cell is not empty
        If Not IsEmpty(ws.Cells(currentRow, "B").Value) Then
            cellValue = ws.Cells(currentRow, "B").Value
            
            ' Check if cell contains "~~"
            If InStr(cellValue, "~~") > 0 Then
                ' Replace "~~" with "~"
                cellValue = Replace(cellValue, "~~", "~")
                
                ' Update the cell with the new value
                ws.Cells(currentRow, "B").Value = cellValue
                
                ' Increment replacement counter
                replacementCount = replacementCount + 1
            End If
            
            ' Increment processed cells counter
            processedCells = processedCells + 1
        End If
        
        ' Update progress every 100 rows
        If currentRow Mod 100 = 0 Then
            Application.StatusBar = "Processing row " & currentRow & " of " & lastRow & "..."
        End If
    Next currentRow
    
    ' Restore Excel settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    
    ' Display completion message
    MsgBox "Process completed!" & vbCrLf & vbCrLf & _
           "Processed cells: " & processedCells & vbCrLf & _
           "Replacements made: " & replacementCount & vbCrLf & _
           "Last row processed: " & lastRow, _
           vbInformation, "Replace Tildes Complete"
    
    Exit Sub
    
ErrorHandler:
    ' Restore Excel settings in case of error
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    
    MsgBox "An error occurred: " & Err.Description & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & vbCrLf & _
           "Make sure the 'Training data' worksheet exists.", _
           vbCritical, "Error"
End Sub

' Quick test subroutine to test on a specific cell
Sub TestTildeReplacement()
    '
    ' Test subroutine to verify the replacement logic works
    '
    
    Dim testString As String
    
    ' Test cases
    testString = "This has ~~ double tildes ~~"
    Debug.Print "Before: " & testString
    testString = Replace(testString, "~~", "~")
    Debug.Print "After: " & testString
    
    testString = "~~START~~ middle ~~END~~"
    Debug.Print "Before: " & testString
    testString = Replace(testString, "~~", "~")
    Debug.Print "After: " & testString
    
    MsgBox "Test completed. Check Immediate Window (Ctrl+G) for results.", vbInformation
End Sub 