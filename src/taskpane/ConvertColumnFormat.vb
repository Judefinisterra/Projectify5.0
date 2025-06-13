Sub ConvertColumnFormatToInline()
    'Converts legacy columnformat parameters to new inline symbol-based approach
    'Processes codestrings in Training Data sheet column B and outputs results to column C
    'Each columnformat parameter applies only to the next row parameter that follows it
    'Example: columnformat="dollaritalic/date" applies dollaritalic to column 1 and date to column 2
    'Multiple columnformat parameters: columnformat="dollar"; row1="..."; columnformat="volume"; row1="...";
    
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
    'Handles multiple columnformat parameters - each applies only to the next row1 parameter
    
    Dim result As String
    result = codestring
    
    ' Process columnformat parameters and their corresponding row parameters sequentially
    result = ProcessSequentialColumnFormatAndRowParameters(result)
    
    ConvertColumnFormatParameters = result
End Function

Function ProcessSequentialColumnFormatAndRowParameters(codestring As String) As String
    'Processes columnformat parameters and row parameters in sequence
    'Each columnformat parameter applies only to the next row parameter that follows it
    
    Dim result As String
    Dim formatRegex As Object
    Dim rowRegex As Object
    Dim formatMatches As Object
    Dim rowMatches As Object
    Dim i As Long, j As Long
    
    result = codestring
    
    ' Create regex objects
    Set formatRegex = CreateObject("VBScript.RegExp")
    Set rowRegex = CreateObject("VBScript.RegExp")
    
    ' Setup columnformat parameter regex
    formatRegex.Global = True
    formatRegex.IgnoreCase = True
    formatRegex.pattern = "\bcolumnformat\s*=\s*""[^""]*""[;,]?|\bcolumnformat\s*=\s*[^;,\s]*[;,]?"
    
    ' Setup row parameter regex  
    rowRegex.Global = True
    rowRegex.IgnoreCase = True
    rowRegex.pattern = "row\d+\s*=\s*""[^""]*"""
    
    ' Get all matches
    Set formatMatches = formatRegex.Execute(codestring)
    Set rowMatches = rowRegex.Execute(codestring)
    
    Debug.Print "  Found " & formatMatches.Count & " columnformat parameters and " & rowMatches.Count & " row parameters"
    
    ' Process columnformat-row pairs
    For i = 0 To formatMatches.Count - 1
        Dim formatMatch As Object
        Dim formatValue As String
        Dim formatPosition As Long
        
        Set formatMatch = formatMatches(i)
        formatValue = ExtractColumnFormatValueFromMatch(formatMatch.Value)
        formatPosition = formatMatch.FirstIndex
        
        Debug.Print "    Processing columnformat parameter: '" & formatValue & "' at position " & formatPosition
        
        ' Find the next row parameter that comes after this columnformat parameter
        For j = 0 To rowMatches.Count - 1
            Dim rowMatch As Object
            Set rowMatch = rowMatches(j)
            
            If rowMatch.FirstIndex > formatPosition Then
                ' This row parameter comes after the current columnformat parameter
                Debug.Print "      Applying to row at position " & rowMatch.FirstIndex
                
                ' Apply the columnformat to this specific row
                Dim originalRowContent As String
                Dim formattedRowContent As String
                Dim originalRowParam As String
                Dim formattedRowParam As String
                
                originalRowContent = ExtractColumnFormatRowContentFromMatch(rowMatch.Value)
                formattedRowContent = ApplyColumnFormatsToRow(originalRowContent, formatValue)
                
                ' Replace the entire row parameter (not just the content) to be more precise
                originalRowParam = rowMatch.Value
                formattedRowParam = Replace(originalRowParam, originalRowContent, formattedRowContent)
                
                ' Replace the specific row parameter in the result
                result = Replace(result, originalRowParam, formattedRowParam, 1, 1)
                
                Exit For ' Move to next columnformat parameter
            End If
        Next j
    Next i
    
    ' Remove all columnformat parameters from the result
    result = formatRegex.Replace(result, "")
    
    ProcessSequentialColumnFormatAndRowParameters = result
End Function

Function ExtractColumnFormatValueFromMatch(formatMatch As String) As String
    'Extracts the columnformat value from a columnformat parameter match
    
    Dim regex As Object
    Dim matches As Object
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    regex.IgnoreCase = True
    regex.pattern = "\bcolumnformat\s*=\s*""([^""]*)""|columnformat\s*=\s*([^;,\s]*)"
    
    Set matches = regex.Execute(formatMatch)
    
    If matches.Count > 0 Then
        If matches(0).SubMatches(0) <> "" Then
            ExtractColumnFormatValueFromMatch = matches(0).SubMatches(0)
        Else
            ExtractColumnFormatValueFromMatch = matches(0).SubMatches(1)
        End If
    Else
        ExtractColumnFormatValueFromMatch = ""
    End If
End Function

Function ExtractColumnFormatRowContentFromMatch(rowMatch As String) As String
    'Extracts the content between quotes from a row parameter match
    
    Dim startQuote As Long
    Dim endQuote As Long
    
    startQuote = InStr(rowMatch, """")
    endQuote = InStrRev(rowMatch, """")
    
    If startQuote > 0 And endQuote > startQuote Then
        ExtractColumnFormatRowContentFromMatch = Mid(rowMatch, startQuote + 1, endQuote - startQuote - 1)
    Else
        ExtractColumnFormatRowContentFromMatch = ""
    End If
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
    'Test function to verify the column format conversion logic including multiple columnformat parameters
    
    Dim testCases() As String
    Dim i As Long
    
    ' Define test cases including multiple columnformat scenarios
    ReDim testCases(5)
    testCases(0) = "<CODE; columnformat=""dollaritalic/date""; row1=""V1|Employee 1||||||1/1/2025|10|1000|2000|4000|8000|16000|32000|"";>"
    testCases(1) = "<CODE; columnformat=""dollar/volume/percent""; row1=""AS|Description|Count||||||||500|1000|2000|F|F|F|"";>"
    testCases(2) = "<CODE; columnformat=""volume\factor\date""; row1=""V|Items|Growth||||||1/1/2025|2.5|10|20|30|40|50|60|"";>"
    testCases(3) = "<CODE; columnformat=""dollaritalic""; row1=""V1|Single Format||||||||||1000|2000|4000|8000|16000|32000|"";>"
    testCases(4) = "columnformat=""dollar/date""; row1=""V1|Revenue||||||1/1/2025|10|100|200|300|400|500|600|""; columnformat=""volume""; row1=""V2|Count||||||||||5|10|15|20|25|30|"";"
    testCases(5) = "indent=""2""; columnformat=""dollaritalic/volume""; bold=""true""; row1=""V3|Price||||||||||50|100|150|200|250|300|"";"
    
    Debug.Print "=== Testing Column Format Conversion ==="
    
    For i = 0 To UBound(testCases)
        Debug.Print "Test Case " & (i + 1) & ":"
        Debug.Print "Original: " & testCases(i)
        Debug.Print "Result: " & ConvertColumnFormatParameters(testCases(i))
        Debug.Print ""
    Next i
End Sub

Sub QuickColumnFormatTest()
    'Quick test to verify multiple columnformat parameters work correctly
    
    Dim testString As String
    Dim result As String
    
    Debug.Print "=== Quick Column Format Test ==="
    
    ' Test multiple columnformat parameters in one codestring
    testString = "columnformat=""dollaritalic""; row1=""V1|Test||||||||||||100|200|300|400|500|600|""; columnformat=""volume/date""; row1=""V2|Items||||||1/1/2025|50|10|20|30|40|50|60|"";"
    result = ConvertColumnFormatParameters(testString)
    Debug.Print "Multiple ColumnFormats Test:"
    Debug.Print "Input:  " & testString
    Debug.Print "Output: " & result
    Debug.Print "Expected: First row gets dollaritalic in column 1 (pos 8), second row gets volume in column 1 (pos 8) and date formatting in column 2 (pos 7)"
    Debug.Print ""
    
End Sub 