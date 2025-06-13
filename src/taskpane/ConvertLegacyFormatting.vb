Sub ConvertLegacyFormattingToInline()
    'Converts legacy format parameters to new inline symbol-based approach
    'Processes codestrings in Training Data sheet column B and outputs results to column C
    'Example: format="dollar" with row data "1000" becomes "$1000" in the row
    
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
    
    Debug.Print "Starting legacy formatting conversion process..."
    Debug.Print "Processing rows 2 to " & lastRow & " in Training Data sheet"
    Debug.Print "Results will be written to column C"
    
    changesMade = 0
    
    ' Clear column C header and add new header
    ws.Cells(1, "C").Value = "Converted Formatting"
    
    ' Loop through each row from 2 to last populated row
    For i = 2 To lastRow
        cellValue = ws.Cells(i, "B").Value
        
        ' Skip empty cells
        If Trim(cellValue) <> "" Then
            ' Process the cell value to convert legacy formatting
            updatedValue = ConvertLegacyFormatParameters(cellValue)
            
            ' Always write the result to column C (whether changed or not)
            ws.Cells(i, "C").Value = updatedValue
            
            ' Track if changes were made
            If updatedValue <> cellValue Then
                changesMade = changesMade + 1
                Debug.Print "Row " & i & ": Converted legacy formatting parameters"
                ' Add visual indicator for changed rows
                ws.Cells(i, "C").Interior.Color = RGB(200, 255, 200) ' Light green background
            Else
                Debug.Print "Row " & i & ": No legacy formatting found"
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
    MsgBox "Legacy formatting conversion complete!" & vbCrLf & _
           "Processed " & (lastRow - 1) & " rows." & vbCrLf & _
           "Converted " & changesMade & " codestrings with legacy formatting." & vbCrLf & _
           "Results written to column C with green highlighting for changed rows.", vbInformation
End Sub

Function ConvertLegacyFormatParameters(codestring As String) As String
    'Converts legacy format parameters to inline symbols in specific columns
    
    Dim result As String
    Dim formatType As String
    
    result = codestring
    
    ' Extract format parameter
    formatType = ExtractParameterValue(result, "format")
    
    ' If no format parameter found, return as is
    If formatType = "" Then
        ConvertLegacyFormatParameters = result
        Exit Function
    End If
    
    Debug.Print "  Found format parameter: '" & formatType & "'"
    
    ' Remove the format parameter from the codestring
    result = RemoveParameter(result, "format")
    
    ' Apply formatting symbols to row parameters based on format type
    result = ApplyFormattingToRowParameters(result, formatType)
    
    ConvertLegacyFormatParameters = result
End Function

Function ExtractParameterValue(codestring As String, paramName As String) As String
    'Extracts the value of a parameter from the codestring (exact match only)
    
    Dim regex As Object
    Dim matches As Object
    Dim pattern As String
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    regex.IgnoreCase = True
    
    ' Pattern to match exact parameter name with word boundaries to avoid matching "columnformat" when looking for "format"
    pattern = "\b" & paramName & "\b\s*=\s*[""']([^""']*)[""']"
    regex.pattern = pattern
    
    Set matches = regex.Execute(codestring)
    
    If matches.Count > 0 Then
        ExtractParameterValue = matches(0).SubMatches(0)
    Else
        ExtractParameterValue = ""
    End If
End Function

Function RemoveParameter(codestring As String, paramName As String) As String
    'Removes a parameter from the codestring (exact match only)
    
    Dim regex As Object
    Dim pattern As String
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    
    ' Pattern to match exact parameter name with word boundaries to avoid matching "columnformat" when looking for "format"
    pattern = "\s*\b" & paramName & "\b\s*=\s*[""'][^""']*[""']\s*;"
    regex.pattern = pattern
    
    RemoveParameter = regex.Replace(codestring, "")
End Function

Function ApplyFormattingToRowParameters(codestring As String, formatType As String) As String
    'Applies formatting symbols to specific columns in row parameters based on format type
    
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
        
        ' Apply formatting symbols to the row content
        updatedRowContent = ApplySymbolsToSpecificColumns(originalRowContent, formatType)
        
        ' Replace in result string
        Dim updatedMatch As String
        updatedMatch = match.SubMatches(0) & updatedRowContent & match.SubMatches(2)
        result = Left(result, match.FirstIndex) & updatedMatch & Mid(result, match.FirstIndex + match.Length + 1)
    Next i
    
    ApplyFormattingToRowParameters = result
End Function

Function ApplySymbolsToSpecificColumns(rowContent As String, formatType As String) As String
    'Applies symbols to specific columns based on format type:
    'Dollar: $ to columns 10-15 (positions 9-14)
    'Volume: ~ to column 2 and columns 10-15 (positions 1 and 9-14)
    'Percent: ~ to column 2 (position 1)
    'DollarItalic: ~$ to columns 10-15 and ~ to column 2 (positions 1 and 9-14)
    
    Dim parts() As String
    Dim result As String
    Dim i As Long
    Dim value As String
    Dim formattedValue As String
    
    ' Split row content by pipes
    parts = Split(rowContent, "|")
    result = ""
    
    For i = 0 To UBound(parts)
        value = parts(i) ' Don't trim to preserve spaces
        formattedValue = value
        
        ' Apply formatting based on format type and column position
        Select Case LCase(formatType)
            Case "dollar"
                ' Add $ to columns 10-15 (positions 9-14)
                If i >= 9 And i <= 14 And value <> "" Then
                    If UCase(value) = "F" Then
                        formattedValue = "$F"
                    Else
                        formattedValue = "$" & value
                    End If
                End If
                
            Case "volume"
                ' Add ~ to column 2 (position 1) and columns 10-15 (positions 9-14)
                If (i = 1 Or (i >= 9 And i <= 14)) And value <> "" Then
                    If UCase(value) = "F" Then
                        formattedValue = "~F"
                    Else
                        formattedValue = "~" & value
                    End If
                End If
                
            Case "percent"
                ' Add ~ to column 2 (position 1)
                If i = 1 And value <> "" Then
                    If UCase(value) = "F" Then
                        formattedValue = "~F"
                    Else
                        formattedValue = "~" & value
                    End If
                End If
                
            Case "dollaritalic"
                ' Add ~$ to columns 10-15 (positions 9-14) and ~ to column 2 (position 1)
                If i >= 9 And i <= 14 And value <> "" Then
                    If UCase(value) = "F" Then
                        formattedValue = "~$F"
                    Else
                        formattedValue = "~$" & value
                    End If
                ElseIf i = 1 And value <> "" Then
                    If UCase(value) = "F" Then
                        formattedValue = "~F"
                    Else
                        formattedValue = "~" & value
                    End If
                End If
                
        End Select
        
        ' Build result
        If i = 0 Then
            result = formattedValue
        Else
            result = result & "|" & formattedValue
        End If
    Next i
    
    ApplySymbolsToSpecificColumns = result
End Function

Sub TestLegacyConversion()
    'Test function to verify the conversion logic works correctly
    
    Dim testCases() As String
    Dim i As Long
    
    ' Define test cases
    ReDim testCases(4)
    testCases(0) = "<CODE; format=""dollar""; row1=""V1|Revenue||||||||1000|2000|4000|8000|16000|32000|"";>"
    testCases(1) = "<CODE; format=""dollaritalic""; row1=""AS|Description||||||||500|1000|2000|F|F|F|"";>"
    testCases(2) = "<CODE; format=""volume""; row1=""V|Count||||||||10|20|30|40|50|60|"";>"
    testCases(3) = "<CODE; format=""percent""; row1=""GR|Growth Rate||||||||2.5|3.0|3.5|4.0|4.5|5.0|"";>"
    testCases(4) = "<CODE; format=""dollar""; columnformat=""euro\date\volume""; row1=""V1|Revenue||||||||1000|2000|4000|8000|16000|32000|"";>"
    
    Debug.Print "=== Testing Legacy Format Conversion ==="
    
    For i = 0 To UBound(testCases)
        Debug.Print "Test Case " & (i + 1) & ":"
        Debug.Print "Original: " & testCases(i)
        Debug.Print "Result: " & ConvertLegacyFormatParameters(testCases(i))
        Debug.Print ""
    Next i
End Sub 