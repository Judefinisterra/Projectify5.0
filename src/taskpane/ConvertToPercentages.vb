Sub ConvertToPercentageFormat()
    'Automatically converts numbers to percentage format in rows where column 2 contains % symbol
    'Processes codestrings in Training Data sheet column B and outputs results to column C
    'All non-empty columns after column 3 get % formatting
    'Decimals like 0.04 become 4%, whole numbers get % added
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String
    Dim updatedValue As String
    Dim changesMade As Long
    
    ' Set reference to Training Data worksheet
    Set ws = Worksheets("Training Data")
    
    ' Find last populated row in column B
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    Debug.Print "Starting percentage conversion process..."
    Debug.Print "Processing rows 2 to " & lastRow & " in Training Data sheet"
    Debug.Print "Looking for % symbols in column 2, then converting numbers to percentages"
    Debug.Print "Results will be written to column C"
    
    changesMade = 0
    
    ' Clear column C header and add new header
    ws.Cells(1, "C").Value = "Converted Percentages"
    
    ' Loop through each row from 2 to last populated row
    For i = 2 To lastRow
        cellValue = ws.Cells(i, "B").Value
        
        ' Skip empty cells
        If Trim(cellValue) <> "" Then
            ' Check if this row should be processed (contains % in column 2)
            If ShouldProcessRowForPercentages(cellValue) Then
                Debug.Print "Row " & i & ": Found % symbol, processing for percentage conversion"
                ' Process the cell value to convert numbers to percentages
                updatedValue = ConvertNumbersToPercentages(cellValue)
                changesMade = changesMade + 1
                ' Add visual indicator for changed rows
                ws.Cells(i, "C").Interior.Color = RGB(255, 255, 200) ' Light yellow background
            Else
                ' No % symbol found, copy as-is
                updatedValue = cellValue
                Debug.Print "Row " & i & ": No % symbol found"
                ' Clear any existing background color for unchanged rows
                ws.Cells(i, "C").Interior.ColorIndex = xlNone
            End If
            
            ' Always write the result to column C
            ws.Cells(i, "C").Value = updatedValue
        Else
            ' Clear column C if column B is empty
            ws.Cells(i, "C").Value = ""
            ws.Cells(i, "C").Interior.ColorIndex = xlNone
        End If
    Next i
    
    Debug.Print "Process complete. Converted " & changesMade & " codestrings to percentage format."
    MsgBox "Percentage conversion complete!" & vbCrLf & _
           "Processed " & (lastRow - 1) & " rows." & vbCrLf & _
           "Converted " & changesMade & " codestrings with % symbols." & vbCrLf & _
           "Results written to column C with yellow highlighting for changed rows.", vbInformation
End Sub

Function ShouldProcessRowForPercentages(codestring As String) As Boolean
    'Determines if a row should be processed based on presence of % symbol in column 2
    
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    Dim i As Long
    
    ' Create regex to find row parameters
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.pattern = "(row\d+\s*=\s*[""'])([^""']*?)([""'])"
    
    Set matches = regex.Execute(codestring)
    
    ' Check each row parameter
    For i = 0 To matches.Count - 1
        Set match = matches(i)
        Dim rowContent As String
        rowContent = match.SubMatches(1) ' The content between quotes
        
        ' Check if column 2 (position 1, 0-based) contains % symbol
        If HasPercentInColumn2(rowContent) Then
            ShouldProcessRowForPercentages = True
            Exit Function
        End If
    Next i
    
    ShouldProcessRowForPercentages = False
End Function

Function HasPercentInColumn2(rowContent As String) As Boolean
    'Checks if column 2 (position 1) contains a % symbol
    
    Dim parts() As String
    
    ' Split row content by pipes
    parts = Split(rowContent, "|")
    
    ' Check if we have at least 2 columns and column 2 contains %
    If UBound(parts) >= 1 Then
        If InStr(parts(1), "%") > 0 Then
            HasPercentInColumn2 = True
            Debug.Print "    Found % symbol in column 2: '" & parts(1) & "'"
        Else
            HasPercentInColumn2 = False
        End If
    Else
        HasPercentInColumn2 = False
    End If
End Function

Function ConvertNumbersToPercentages(codestring As String) As String
    'Converts numbers to percentage format in all row parameters
    
    Dim result As String
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    Dim i As Long
    
    result = codestring
    
    ' Create regex to find row parameters
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.pattern = "(row\d+\s*=\s*[""'])([^""']*?)([""'])"
    
    Set matches = regex.Execute(result)
    
    ' Process matches in reverse order to avoid position shifts
    For i = matches.Count - 1 To 0 Step -1
        Set match = matches(i)
        Dim originalRowContent As String
        Dim updatedRowContent As String
        
        originalRowContent = match.SubMatches(1) ' The content between quotes
        
        ' Convert numbers to percentages in the row content
        updatedRowContent = ConvertRowToPercentages(originalRowContent)
        
        ' Replace in result string
        Dim updatedMatch As String
        updatedMatch = match.SubMatches(0) & updatedRowContent & match.SubMatches(2)
        result = Left(result, match.FirstIndex) & updatedMatch & Mid(result, match.FirstIndex + match.Length + 1)
    Next i
    
    ConvertNumbersToPercentages = result
End Function

Function ConvertRowToPercentages(rowContent As String) As String
    'Converts numbers to percentage format starting from column 4 (position 3)
    
    Dim parts() As String
    Dim result As String
    Dim i As Long
    Dim value As String
    Dim convertedValue As String
    
    ' Split row content by pipes
    parts = Split(rowContent, "|")
    
    result = ""
    
    For i = 0 To UBound(parts)
        value = parts(i)
        
        ' Apply percentage conversion starting from position 3 (column D)
        If i >= 3 And value <> "" Then
            convertedValue = ConvertValueToPercentage(value)
            Debug.Print "    Position " & i & ": '" & value & "' -> '" & convertedValue & "'"
        Else
            ' Keep original value for positions 0-2 or empty values
            convertedValue = value
        End If
        
        ' Build result
        If i = 0 Then
            result = convertedValue
        Else
            result = result & "|" & convertedValue
        End If
    Next i
    
    ConvertRowToPercentages = result
End Function

Function ConvertValueToPercentage(value As String) As String
    'Converts a single value to percentage format
    
    Dim result As String
    Dim numValue As Double
    
    result = value
    
    ' Handle special case for "F"
    If UCase(Trim(value)) = "F" Then
        result = "%F"
        ConvertValueToPercentage = result
        Exit Function
    End If
    
    ' Check if value is numeric
    If IsNumeric(value) Then
        numValue = CDbl(value)
        
        ' Convert decimal to percentage (e.g., 0.04 -> 4%)
        If numValue > 0 And numValue < 1 Then
            ' Decimal value - multiply by 100 and add %
            result = Format(numValue * 100, "0") & "%"
        ElseIf numValue >= 1 Or numValue <= -1 Then
            ' Whole number - just add %
            result = Format(numValue, "0") & "%"
        ElseIf numValue = 0 Then
            ' Zero - add %
            result = "0%"
        Else
            ' Negative decimal - multiply by 100 and add %
            result = Format(numValue * 100, "0") & "%"
        End If
    Else
        ' Non-numeric value - check if it already has % or add it
        If Right(Trim(value), 1) <> "%" Then
            result = value & "%"
        Else
            result = value ' Already has %
        End If
    End If
    
    ConvertValueToPercentage = result
End Function

Sub TestPercentageConversion()
    'Test function to verify the percentage conversion logic
    
    Dim testCases() As String
    Dim i As Long
    
    ' Define test cases
    ReDim testCases(4)
    testCases(0) = "<CODE; row1=""V1|Growth %||||||||||0.04|0.15|0.25|0.35|0.45|0.55|"";>"
    testCases(1) = "<CODE; row1=""V2|Margin %||||||||||4|15|25|F|F|F|"";>"
    testCases(2) = "<CODE; row1=""V3|Tax Rate %||||||||||0.21|0.21|0.21|0.21|0.21|0.21|"";>"
    testCases(3) = "<CODE; row1=""V4|No Percent||||||||||100|200|300|400|500|600|"";>"
    testCases(4) = "<CODE; row1=""V5|Mixed % Values||||||||||0.05|10|0.125|F|20|0.03|"";>"
    
    Debug.Print "=== Testing Percentage Conversion ==="
    
    For i = 0 To UBound(testCases)
        Debug.Print "Test Case " & (i + 1) & ":"
        Debug.Print "Original: " & testCases(i)
        
        If ShouldProcessRowForPercentages(testCases(i)) Then
            Debug.Print "Should Process: YES"
            Debug.Print "Result: " & ConvertNumbersToPercentages(testCases(i))
        Else
            Debug.Print "Should Process: NO"
            Debug.Print "Result: " & testCases(i)
        End If
        Debug.Print ""
    Next i
End Sub 