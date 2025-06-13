Sub ConvertLegacyFormattingToInline()
    'Converts legacy format parameters to new inline symbol-based approach
    'Processes codestrings in Training Data sheet column B and outputs results to column C
    'Example: format="dollar" with row data "1000" becomes "$1000" in the row
    '
    'IMPORTANT: This script ONLY processes format="..." parameters
    'It does NOT add ~ to labels ending with : - that's handled by a different script
    'It ONLY applies symbols to specific columns as defined in the format rules
    
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

Sub QuickFormatTest()
    'Quick test to verify the format conversion is working correctly
    
    Dim testString As String
    Dim result As String
    
    Debug.Print "=== Quick Format Test ==="
    
    ' Test Volume format (should add ~ to column 2 and columns 10-15 ONLY)
    testString = "format=""Volume""; row1 = ""V2|Beginning||||||||F|F|F|F|F|F|"";"
    result = ConvertLegacyFormatParameters(testString)
    Debug.Print "Volume Test:"
    Debug.Print "Input:  " & testString
    Debug.Print "Output: " & result
    Debug.Print ""
    
    ' Test Dollar format (should add $ to columns 10-15 ONLY)
    testString = "format=""Dollar""; row1 = ""R3|Revenue|is: revenue|||||||F|F|F|F|F|F|"";"
    result = ConvertLegacyFormatParameters(testString)
    Debug.Print "Dollar Test:"
    Debug.Print "Input:  " & testString
    Debug.Print "Output: " & result
    Debug.Print ""
    
    ' Test Dollaritalic format (should add ~$ to columns 10-15 and ~ to column 2 ONLY)
    testString = "format=""Dollaritalic""; row1 = ""V4|Price||||||||20|20|20|20|20|20|"";"
    result = ConvertLegacyFormatParameters(testString)
    Debug.Print "Dollaritalic Test:"
    Debug.Print "Input:  " & testString
    Debug.Print "Output: " & result
    Debug.Print ""
    
    ' Test the user's specific scenario - multiple formats in one codestring
    testString = "format=""volume""; row1 = ""V2|Beginning||||||||F|F|F|F|F|F|""; format=""Dollar""; row1 = ""R3|Revenue|is: revenue|||||||F|F|F|F|F|F|"";"
    result = ConvertLegacyFormatParameters(testString)
    Debug.Print "Multiple Formats Test:"
    Debug.Print "Input:  " & testString
    Debug.Print "Output: " & result
    Debug.Print ""
    
    ' Test user's exact problematic case
    testString = "topborder=""True""; bold=""true""; indent=""1""; format=""Dollar""; driver1=""V2""; driver2=""R2""; driver3=""V4""; row1 = ""R3|Revenue - Subscription|is: revenue|||||||F|F|F|F|F|F|"";"
    result = ConvertLegacyFormatParameters(testString)
    Debug.Print "User's Exact Case Test:"
    Debug.Print "Input:  " & testString
    Debug.Print "Output: " & result
    Debug.Print "Expected: Dollar format should add $ to columns 10-15 (positions 9-14) ONLY"
    Debug.Print ""
    
End Sub

Function ConvertLegacyFormatParameters(codestring As String) As String
    'Converts legacy format parameters to inline symbols in specific columns
    'Handles multiple format parameters - each applies only to the next row1 parameter
    
    Dim result As String
    result = codestring
    
    ' Process format parameters and their corresponding row parameters sequentially
    result = ProcessSequentialFormatAndRowParameters(result)
    
    ConvertLegacyFormatParameters = result
End Function

Function ProcessSequentialFormatAndRowParameters(codestring As String) As String
    'Processes format parameters and row parameters in sequence
    'Each format parameter applies only to the next row parameter that follows it
    
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
    
    ' Setup format parameter regex
    formatRegex.Global = True
    formatRegex.IgnoreCase = True
    formatRegex.pattern = "\bformat\s*=\s*""[^""]*""[;,]?|\bformat\s*=\s*[^;,\s]*[;,]?"
    
    ' Setup row parameter regex  
    rowRegex.Global = True
    rowRegex.IgnoreCase = True
    rowRegex.pattern = "row\d+\s*=\s*""[^""]*"""
    
    ' Get all matches
    Set formatMatches = formatRegex.Execute(codestring)
    Set rowMatches = rowRegex.Execute(codestring)
    
    Debug.Print "  Found " & formatMatches.Count & " format parameters and " & rowMatches.Count & " row parameters"
    
    ' Process format-row pairs
    For i = 0 To formatMatches.Count - 1
        Dim formatMatch As Object
        Dim formatValue As String
        Dim formatPosition As Long
        
        Set formatMatch = formatMatches(i)
        formatValue = ExtractFormatValueFromMatch(formatMatch.Value)
        formatPosition = formatMatch.FirstIndex
        
        Debug.Print "    Processing format parameter: '" & formatValue & "' at position " & formatPosition
        
        ' Find the next row parameter that comes after this format parameter
        For j = 0 To rowMatches.Count - 1
            Dim rowMatch As Object
            Set rowMatch = rowMatches(j)
            
            If rowMatch.FirstIndex > formatPosition Then
                ' This row parameter comes after the current format parameter
                Debug.Print "      Applying to row at position " & rowMatch.FirstIndex
                
                ' Apply the format to this specific row
                Dim originalRowContent As String
                Dim formattedRowContent As String
                Dim originalRowParam As String
                Dim formattedRowParam As String
                
                originalRowContent = ExtractRowContentFromMatch(rowMatch.Value)
                formattedRowContent = ApplySymbolsToSpecificColumns(originalRowContent, formatValue)
                
                ' Replace the entire row parameter (not just the content) to be more precise
                originalRowParam = rowMatch.Value
                formattedRowParam = Replace(originalRowParam, originalRowContent, formattedRowContent)
                
                ' Replace the specific row parameter in the result
                result = Replace(result, originalRowParam, formattedRowParam, 1, 1)
                
                Exit For ' Move to next format parameter
            End If
        Next j
    Next i
    
    ' Remove all format parameters from the result
    result = formatRegex.Replace(result, "")
    
    ProcessSequentialFormatAndRowParameters = result
End Function

Function ExtractFormatValueFromMatch(formatMatch As String) As String
    'Extracts the format value from a format parameter match
    
    Dim regex As Object
    Dim matches As Object
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    regex.IgnoreCase = True
    regex.pattern = "\bformat\s*=\s*""([^""]*)""|format\s*=\s*([^;,\s]*)"
    
    Set matches = regex.Execute(formatMatch)
    
    If matches.Count > 0 Then
        If matches(0).SubMatches(0) <> "" Then
            ExtractFormatValueFromMatch = matches(0).SubMatches(0)
        Else
            ExtractFormatValueFromMatch = matches(0).SubMatches(1)
        End If
    Else
        ExtractFormatValueFromMatch = ""
    End If
End Function

Function ExtractRowContentFromMatch(rowMatch As String) As String
    'Extracts the content between quotes from a row parameter match
    
    Dim startQuote As Long
    Dim endQuote As Long
    
    startQuote = InStr(rowMatch, """")
    endQuote = InStrRev(rowMatch, """")
    
    If startQuote > 0 And endQuote > startQuote Then
        ExtractRowContentFromMatch = Mid(rowMatch, startQuote + 1, endQuote - startQuote - 1)
    Else
        ExtractRowContentFromMatch = ""
    End If
End Function

Function ExtractParameterValue(codestring As String, paramName As String) As String
    'Extracts the value of a parameter from the codestring (exact match only)
    'IMPORTANT: This function will ONLY match exact parameter names and will NOT match:
    '- "columnformat" when looking for "format"
    '- "columnformats" when looking for "format"  
    
    Dim regex As Object
    Dim matches As Object
    Dim pattern As String
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    regex.IgnoreCase = True
    
    ' Pattern uses word boundaries (\b) to ensure exact match only
    ' Word boundaries prevent matching "columnformat" or "columnformats" when searching for "format"
    ' because there's no word boundary between "format" and following characters in those words
    pattern = "\b" & paramName & "\b\s*=\s*""([^""]*)""|" & "\b" & paramName & "\b\s*=\s*([^;,\s]*)"
    regex.pattern = pattern
    
    Set matches = regex.Execute(codestring)
    
    If matches.Count > 0 Then
        If matches(0).SubMatches(0) <> "" Then
            ExtractParameterValue = matches(0).SubMatches(0)
        Else
            ExtractParameterValue = matches(0).SubMatches(1)
        End If
    Else
        ExtractParameterValue = ""
    End If
End Function

Function RemoveParameter(codestring As String, paramName As String) As String
    'Removes a parameter from the codestring (exact match only)
    'IMPORTANT: This function will ONLY remove exact parameter names and will NOT remove:
    '- "columnformat" when looking for "format"
    '- "columnformats" when looking for "format"
    
    Dim regex As Object
    Dim pattern As String
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    
    ' Word boundaries ensure exact match only - prevents removing "columnformat" or "columnformats"
    ' when searching for "format" because there's no word boundary between "format" and following characters
    pattern = "\s*\b" & paramName & "\b\s*=\s*""[^""]*""\s*[;,]?|\s*\b" & paramName & "\b\s*=\s*[^;,\s]*\s*[;,]?"
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
    'Percent: ~ to column 2 and convert values in columns 10-15 to percentages with ~ prefix
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
        value = Trim(parts(i))
        formattedValue = parts(i) ' Preserve original spacing
        
        ' Apply formatting based on format type and column position
        ' Reset formattedValue to original to ensure we only change targeted columns
        formattedValue = parts(i)
        
        Select Case LCase(Trim(formatType))
            Case "dollar"
                ' Add $ to columns 10-15 ONLY (positions 9-14)
                If i >= 9 And i <= 14 And value <> "" Then
                    If Left(value, 1) <> "$" Then ' Don't add $ if already present
                        formattedValue = "$" & value
                    End If
                End If
                
            Case "volume"
                ' Add ~ to column 2 (position 1) AND columns 10-15 (positions 9-14) ONLY
                If i = 1 And value <> "" Then
                    If Left(value, 1) <> "~" Then
                        formattedValue = "~" & value
                    End If
                ElseIf i >= 9 And i <= 14 And value <> "" Then
                    If Left(value, 1) <> "~" Then
                        formattedValue = "~" & value
                    End If
                End If
                
            Case "percent"
                ' Add ~ to column 2 ONLY (position 1)
                If i = 1 And value <> "" Then
                    If Left(value, 1) <> "~" Then
                        formattedValue = "~" & value
                    End If
                ' Convert values in columns 10-15 to percentages with ~ prefix
                ElseIf i >= 9 And i <= 14 And value <> "" Then
                    formattedValue = ConvertToPercentageWithTilde(value)
                End If
                
            Case "dollaritalic"
                ' Add ~$ to columns 10-15 (positions 9-14) AND ~ to column 2 (position 1) ONLY
                If i >= 9 And i <= 14 And value <> "" Then
                    If Left(value, 2) <> "~$" Then ' Don't add ~$ if already present
                        formattedValue = "~$" & value
                    End If
                ElseIf i = 1 And value <> "" Then
                    If Left(value, 1) <> "~" Then
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

Function ConvertToPercentageWithTilde(value As String) As String
    'Converts a value to percentage format with ~ prefix
    'Examples: 0.01 -> ~1%, 0.5 -> ~50%, 1 -> ~100%, F -> ~F%
    
    Dim numValue As Double
    Dim result As String
    
    ' If value already ends with %, just add ~ if not present
    If Right(value, 1) = "%" Then
        If Left(value, 1) = "~" Then
            result = value ' Already has ~ and %
        Else
            result = "~" & value ' Add ~ prefix
        End If
    ' If value is numeric, convert to percentage
    ElseIf IsNumeric(value) Then
        numValue = CDbl(value)
        ' Convert decimal to percentage (e.g., 0.01 = 1%)
        If numValue < 1 And numValue > 0 Then
            result = "~" & CStr(Round(numValue * 100, 2)) & "%"
        Else
            result = "~" & value & "%"
        End If
    ' For non-numeric values like "F", just add ~ and %
    Else
        If Left(value, 1) = "~" Then
            result = value & "%"
        Else
            result = "~" & value & "%"
        End If
    End If
    
    ConvertToPercentageWithTilde = result
End Function

Sub TestLegacyConversion()
    'Test function to verify the conversion logic works correctly
    
    Dim testCases() As String
    Dim i As Long
    
    ' Define test cases based on user requirements including capitalized formats and multiple formats
    ReDim testCases(10)
    testCases(0) = "format=""dollar""; row1 = ""V1|Revenue||||||||1000|2000|4000|8000|16000|32000|"";"
    testCases(1) = "format=""dollaritalic""; row1 = ""AS|Description||||||||500|1000|2000|F|F|F|"";"
    testCases(2) = "format=""volume""; row1 = ""V|Count||||||||10|20|30|40|50|60|"";"
    testCases(3) = "format=""percent""; row1 = ""GR|Growth Rate||||||||0.025|0.03|0.035|0.04|0.045|0.05|"";"
    testCases(4) = "format=""percent""; row1 = ""PCT|Percentage||||||||2.5|3.0|F|5%|10|0.15|"";"
    testCases(5) = "columnformats=""euro\date""; format=""dollar""; row1 = ""V1|Revenue||||||||1000|2000|4000|8000|16000|32000|"";"
    testCases(6) = "format=""Volume""; row1 = ""V2|Beginning||||||||F|F|F|F|F|F|"";"
    testCases(7) = "format=""Dollaritalic""; row1 = ""V4|Price/Month/Subscriber||||||||20|20|20|20|20|20|"";"
    testCases(8) = "format=""Dollar""; row1 = ""R3|Revenue - Subscription|is: revenue|||||||F|F|F|F|F|F|"";"
    testCases(9) = "format=""volume""; row1 = ""V1|Test||||||||10|20|30|40|50|60|""; format=""dollar""; row1 = ""V2|Revenue||||||||100|200|300|400|500|600|"";"
    testCases(10) = "indent=""2""; format=""Dollar""; bold=""true""; row1 = ""R3|Revenue - Subscription|is: revenue|||||||F|F|F|F|F|F|"";"
    
    Debug.Print "=== Testing Legacy Format Conversion ==="
    
    For i = 0 To UBound(testCases)
        Debug.Print "Test Case " & (i + 1) & ":"
        Debug.Print "Original: " & testCases(i)
        Debug.Print "Result: " & ConvertLegacyFormatParameters(testCases(i))
        Debug.Print ""
    Next i
    
    ' Additional test for the percentage conversion function
    Debug.Print "=== Testing Percentage Conversion Function ==="
    Debug.Print "0.01 -> " & ConvertToPercentageWithTilde("0.01")
    Debug.Print "0.5 -> " & ConvertToPercentageWithTilde("0.5")
    Debug.Print "1 -> " & ConvertToPercentageWithTilde("1")
    Debug.Print "F -> " & ConvertToPercentageWithTilde("F")
    Debug.Print "5% -> " & ConvertToPercentageWithTilde("5%")
    Debug.Print "~3% -> " & ConvertToPercentageWithTilde("~3%")
End Sub 