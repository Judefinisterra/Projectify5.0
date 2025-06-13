Sub ConvertColumnFormatToInline()
    'Converts legacy columnformat parameters to new inline symbol-based approach
    'Processes codestrings in Training Data sheet column B and outputs results to column C
    'Example: columnformat="dollaritalic/date" applies dollaritalic to column 1 and date to column 2
    
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
    
    Debug.Print "Starting columnformat conversion process..."
    Debug.Print "Processing rows 2 to " & lastRow & " in Training Data sheet"
    Debug.Print "Results will be written to column C"
    
    changesMade = 0
    
    ' Clear column C header and add new header
    ws.Cells(1, "C").Value = "Converted ColumnFormat"
    
    ' Loop through each row from 2 to last populated row
    For i = 2 To lastRow
        cellValue = ws.Cells(i, "B").Value
        
        ' Skip empty cells
        If Trim(cellValue) <> "" Then
            ' Process the cell value to convert columnformat
            updatedValue = ConvertColumnFormatParameters(cellValue)
            
            ' Always write the result to column C (whether changed or not)
            ws.Cells(i, "C").Value = updatedValue
            
            ' Track if changes were made
            If updatedValue <> cellValue Then
                changesMade = changesMade + 1
                Debug.Print "Row " & i & ": Converted columnformat parameters"
                ' Add visual indicator for changed rows
                ws.Cells(i, "C").Interior.Color = RGB(255, 200, 255) ' Light purple background
            Else
                Debug.Print "Row " & i & ": No columnformat found"
                ' Clear any existing background color for unchanged rows
                ws.Cells(i, "C").Interior.ColorIndex = xlNone
            End If
        Else
            ' Clear column C if column B is empty
            ws.Cells(i, "C").Value = ""
            ws.Cells(i, "C").Interior.ColorIndex = xlNone
        End If
    Next i
    
    Debug.Print "Process complete. Converted " & changesMade & " codestrings."
    MsgBox "ColumnFormat conversion complete!" & vbCrLf & _
           "Processed " & (lastRow - 1) & " rows." & vbCrLf & _
           "Converted " & changesMade & " codestrings with columnformat." & vbCrLf & _
           "Results written to column C with purple highlighting for changed rows.", vbInformation
End Sub

Function ConvertColumnFormatParameters(codestring As String) As String
    'Converts legacy columnformat parameters to inline symbols in specific columns
    
    Dim result As String
    Dim columnFormatValue As String
    
    result = codestring
    
    ' Extract columnformat parameter
    columnFormatValue = ExtractColumnFormatParameter(result, "columnformat")
    
    ' If no columnformat parameter found, return as is
    If columnFormatValue = "" Then
        ConvertColumnFormatParameters = result
        Exit Function
    End If
    
    Debug.Print "  Found columnformat parameter: '" & columnFormatValue & "'"
    
    ' Remove the columnformat parameter from the codestring
    result = RemoveColumnFormatParameter(result, "columnformat")
    
    ' Apply formatting symbols to row parameters based on columnformat
    result = ApplyColumnFormattingToRowParameters(result, columnFormatValue)
    
    ConvertColumnFormatParameters = result
End Function

Function ExtractColumnFormatParameter(codestring As String, paramName As String) As String
    'Extracts the value of a parameter from the codestring (exact match only)
    
    Dim regex As Object
    Dim matches As Object
    Dim pattern As String
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    regex.IgnoreCase = True
    
    ' Pattern to match exact parameter name with word boundaries
    pattern = "\b" & paramName & "\b\s*=\s*[""']([^""']*)[""']"
    regex.pattern = pattern
    
    Set matches = regex.Execute(codestring)
    
    If matches.Count > 0 Then
        ExtractColumnFormatParameter = matches(0).SubMatches(0)
    Else
        ExtractColumnFormatParameter = ""
    End If
End Function

Function RemoveColumnFormatParameter(codestring As String, paramName As String) As String
    'Removes a parameter from the codestring (exact match only)
    
    Dim regex As Object
    Dim pattern As String
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    
    ' Pattern to match exact parameter name with word boundaries
    pattern = "\s*\b" & paramName & "\b\s*=\s*[""'][^""']*[""']\s*;"
    regex.pattern = pattern
    
    RemoveColumnFormatParameter = regex.Replace(codestring, "")
End Function

Function ApplyColumnFormattingToRowParameters(codestring As String, columnFormatValue As String) As String
    'Applies column-specific formatting symbols to row parameters
    
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
        
        ' Apply column formatting to the row content
        updatedRowContent = ApplyColumnFormatsToRow(originalRowContent, columnFormatValue)
        
        ' Replace in result string
        Dim updatedMatch As String
        updatedMatch = match.SubMatches(0) & updatedRowContent & match.SubMatches(2)
        result = Left(result, match.FirstIndex) & updatedMatch & Mid(result, match.FirstIndex + match.Length + 1)
    Next i
    
    ApplyColumnFormattingToRowParameters = result
End Function

Function ApplyColumnFormatsToRow(rowContent As String, columnFormatValue As String) As String
    'Applies column-specific formats to a single row
    'Column mapping: 1=I(pos 8), 2=H(pos 7), 3=G(pos 6), 4=F(pos 5), 5=E(pos 4), 6=D(pos 3), 7=C(pos 2), 8=B(pos 1), 9=A(pos 0)
    
    Dim parts() As String
    Dim formats() As String
    Dim result As String
    Dim i As Long
    Dim value As String
    Dim formattedValue As String
    Dim formatIndex As Long
    Dim currentFormat As String
    
    ' Split row content by pipes
    parts = Split(rowContent, "|")
    
    ' Split formats by / or \
    If InStr(columnFormatValue, "/") > 0 Then
        formats = Split(columnFormatValue, "/")
    Else
        formats = Split(columnFormatValue, "\")
    End If
    
    result = ""
    
    ' Define column position mapping (1-based format index to 0-based position)
    Dim columnPositions() As Long
    ReDim columnPositions(1 To 9)
    columnPositions(1) = 8  ' Column 1 -> Position I (8)
    columnPositions(2) = 7  ' Column 2 -> Position H (7)
    columnPositions(3) = 6  ' Column 3 -> Position G (6)
    columnPositions(4) = 5  ' Column 4 -> Position F (5)
    columnPositions(5) = 4  ' Column 5 -> Position E (4)
    columnPositions(6) = 3  ' Column 6 -> Position D (3)
    columnPositions(7) = 2  ' Column 7 -> Position C (2)
    columnPositions(8) = 1  ' Column 8 -> Position B (1)
    columnPositions(9) = 0  ' Column 9 -> Position A (0)
    
    For i = 0 To UBound(parts)
        value = parts(i)
        formattedValue = value
        
        ' Check if this position should have formatting applied
        For formatIndex = 1 To UBound(formats) + 1
            If formatIndex <= UBound(columnPositions) And columnPositions(formatIndex) = i Then
                ' This position matches a format column
                If formatIndex - 1 <= UBound(formats) Then
                    currentFormat = Trim(formats(formatIndex - 1))
                    formattedValue = ApplyColumnFormatToValue(value, currentFormat)
                    Debug.Print "    Applied format '" & currentFormat & "' to position " & i & ": '" & value & "' -> '" & formattedValue & "'"
                End If
                Exit For
            End If
        Next formatIndex
        
        ' Build result
        If i = 0 Then
            result = formattedValue
        Else
            result = result & "|" & formattedValue
        End If
    Next i
    
    ApplyColumnFormatsToRow = result
End Function

Function ApplyColumnFormatToValue(value As String, formatType As String) As String
    'Applies a specific format to a single value
    
    Dim result As String
    result = value
    
    ' Only apply formatting to non-empty values
    If value <> "" Then
        Select Case LCase(Trim(formatType))
            Case "dollaritalic"
                If UCase(value) = "F" Then
                    result = "~$F"
                Else
                    result = "~$" & value
                End If
                
            Case "dollar"
                If UCase(value) = "F" Then
                    result = "$F"
                Else
                    result = "$" & value
                End If
                
            Case "volume"
                If UCase(value) = "F" Then
                    result = "~F"
                Else
                    result = "~" & value
                End If
                
            Case "percent"
                If UCase(value) = "F" Then
                    result = "~F"
                Else
                    result = "~" & value
                End If
                
            Case "factor"
                If UCase(value) = "F" Then
                    result = "~F"
                ElseIf IsNumeric(value) Then
                    result = "~" & value & "x"
                Else
                    result = "~" & value
                End If
                
            Case "date"
                ' Date formatting typically doesn't add symbols
                ' Keep the value as is
                result = value
                
            Case ""
                ' Empty format, no change
                result = value
                
        End Select
    End If
    
    ApplyColumnFormatToValue = result
End Function

Sub TestColumnFormatConversion()
    'Test function to verify the column format conversion logic
    
    Dim testCases() As String
    Dim i As Long
    
    ' Define test cases
    ReDim testCases(3)
    testCases(0) = "<CODE; columnformat=""dollaritalic/date""; row1=""V1|Employee 1||||||1/1/2025|10|1000|2000|4000|8000|16000|32000|"";>"
    testCases(1) = "<CODE; columnformat=""dollar/volume/percent""; row1=""AS|Description|Count||||||||500|1000|2000|F|F|F|"";>"
    testCases(2) = "<CODE; columnformat=""volume\factor\date""; row1=""V|Items|Growth||||||1/1/2025|2.5|10|20|30|40|50|60|"";>"
    testCases(3) = "<CODE; columnformat=""dollaritalic""; row1=""V1|Single Format||||||||||1000|2000|4000|8000|16000|32000|"";>"
    
    Debug.Print "=== Testing Column Format Conversion ==="
    
    For i = 0 To UBound(testCases)
        Debug.Print "Test Case " & (i + 1) & ":"
        Debug.Print "Original: " & testCases(i)
        Debug.Print "Result: " & ConvertColumnFormatParameters(testCases(i))
        Debug.Print ""
    Next i
End Sub 