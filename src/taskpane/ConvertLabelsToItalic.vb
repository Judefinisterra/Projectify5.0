Option Explicit

' ConvertLabelsToItalic Module - Adds ~ prefix to row data following specific label types
' Processes LABELH1, LABELH2, and COLUMNHEADER-E code types

Sub ConvertLabelsToItalic()
    'Finds LABELH1, LABELH2, and COLUMNHEADER-E code types and adds ~ prefix 
    'to all non-blank values in the next row1 parameter that follows each one
    'Processes codestrings in Training Data sheet column B and outputs results to column C
    
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
    
    Debug.Print "Starting label-to-italic conversion process..."
    Debug.Print "Processing rows 2 to " & lastRow & " in Training Data sheet"
    Debug.Print "Looking for LABELH1, LABELH2, and COLUMNHEADER-E code types"
    Debug.Print "Results will be written to column C"
    
    changesMade = 0
    
    ' Clear column C header and add new header
    ws.Cells(1, "C").Value = "Converted Labels to Italic"
    
    ' Loop through each row from 2 to last populated row
    For i = 2 To lastRow
        cellValue = ws.Cells(i, "B").Value
        
        ' Skip empty cells
        If Trim(cellValue) <> "" Then
            ' Process the cell value to convert labels to italic
            updatedValue = ConvertLabelsToItalicInString(cellValue)
            
            ' Always write the result to column C (whether changed or not)
            ws.Cells(i, "C").Value = updatedValue
            
            ' Track if changes were made
            If updatedValue <> cellValue Then
                changesMade = changesMade + 1
                Debug.Print "Row " & i & ": Converted labels to italic"
                ' Add visual indicator for changed rows
                ws.Cells(i, "C").Interior.Color = RGB(200, 255, 255) ' Light cyan background
            Else
                Debug.Print "Row " & i & ": No target labels found"
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
    MsgBox "Label-to-italic conversion complete!" & vbCrLf & _
           "Processed " & (lastRow - 1) & " rows." & vbCrLf & _
           "Converted " & changesMade & " codestrings with target labels." & vbCrLf & _
           "Results written to column C with cyan highlighting for changed rows.", vbInformation
End Sub

Function ConvertLabelsToItalicInString(codestring As String) As String
    'Processes label types and their corresponding row parameters sequentially
    'Each target label type applies italic formatting to the next row parameter that follows it
    
    Dim result As String
    result = codestring
    
    ' Process target labels and their corresponding row parameters sequentially
    result = ProcessSequentialLabelsAndRowParameters(result)
    
    ConvertLabelsToItalicInString = result
End Function

Function ProcessSequentialLabelsAndRowParameters(codestring As String) As String
    'Processes target label types and row parameters in sequence
    'Each LABELH1, LABELH2, or COLUMNHEADER-E applies italic formatting to the next row parameter
    
    Dim result As String
    Dim labelRegex As Object
    Dim rowRegex As Object
    Dim labelMatches As Object
    Dim rowMatches As Object
    Dim i As Long, j As Long
    
    result = codestring
    
    ' Create regex objects
    Set labelRegex = CreateObject("VBScript.RegExp")
    Set rowRegex = CreateObject("VBScript.RegExp")
    
    ' Setup target label types regex
    labelRegex.Global = True
    labelRegex.IgnoreCase = True
    labelRegex.pattern = "<(LABELH1|LABELH2|COLUMNHEADER-E)[\s;]"
    
    ' Setup row parameter regex  
    rowRegex.Global = True
    rowRegex.IgnoreCase = True
    rowRegex.pattern = "row\d+\s*=\s*""[^""]*"""
    
    ' Get all matches
    Set labelMatches = labelRegex.Execute(codestring)
    Set rowMatches = rowRegex.Execute(codestring)
    
    Debug.Print "  Found " & labelMatches.Count & " target labels and " & rowMatches.Count & " row parameters"
    
    ' Process label-row pairs
    For i = 0 To labelMatches.Count - 1
        Dim labelMatch As Object
        Dim labelType As String
        Dim labelPosition As Long
        
        Set labelMatch = labelMatches(i)
        labelType = labelMatch.SubMatches(0) ' Extract the label type (LABELH1, LABELH2, etc.)
        labelPosition = labelMatch.FirstIndex
        
        Debug.Print "    Processing " & labelType & " at position " & labelPosition
        
        ' Find the next row parameter that comes after this label
        For j = 0 To rowMatches.Count - 1
            Dim rowMatch As Object
            Set rowMatch = rowMatches(j)
            
            If rowMatch.FirstIndex > labelPosition Then
                ' This row parameter comes after the current label
                Debug.Print "      Applying italic formatting to row at position " & rowMatch.FirstIndex
                
                ' Apply italic formatting to this specific row
                Dim originalRowContent As String
                Dim formattedRowContent As String
                Dim originalRowParam As String
                Dim formattedRowParam As String
                
                originalRowContent = ExtractRowContentFromMatch(rowMatch.Value)
                formattedRowContent = AddTildePrefixToAllValues(originalRowContent)
                
                ' Replace the entire row parameter (not just the content) to be more precise
                originalRowParam = rowMatch.Value
                formattedRowParam = Replace(originalRowParam, originalRowContent, formattedRowContent)
                
                ' Replace the specific row parameter in the result
                result = Replace(result, originalRowParam, formattedRowParam, 1, 1)
                
                Exit For ' Move to next label
            End If
        Next j
    Next i
    
    ProcessSequentialLabelsAndRowParameters = result
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

Function AddTildePrefixToAllValues(rowContent As String) As String
    'Adds ~ prefix to all non-blank values in the row content
    
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
        
        ' Add ~ prefix to non-blank values that don't already have it
        If value <> "" And Left(value, 1) <> "~" Then
            formattedValue = "~" & value
            Debug.Print "      Added ~ to column " & (i + 1) & ": '" & value & "' -> '" & formattedValue & "'"
        End If
        
        ' Build result
        If i = 0 Then
            result = formattedValue
        Else
            result = result & "|" & formattedValue
        End If
    Next i
    
    AddTildePrefixToAllValues = result
End Function

Sub TestLabelsToItalicConversion()
    'Test function to verify the labels-to-italic conversion logic
    
    Dim testCases() As String
    Dim i As Long
    
    ' Define test cases for different label types
    ReDim testCases(5)
    testCases(0) = "<LABELH1; row1=""V1|Revenue Summary||||||||||||||"";>"
    testCases(1) = "<LABELH2; indent=""1""; row1=""V2|Quarterly Results||||||||||||||"";>"
    testCases(2) = "<COLUMNHEADER-E; bold=""true""; row1=""H1|Q1|Q2|Q3|Q4||||||||||||"";>"
    testCases(3) = "<LABELH1; row1=""V3|Financial Overview||||||||||||||""; <OTHER; row1=""X|Test||||||||||||||"";>"
    testCases(4) = "indent=""2""; <LABELH2; bold=""true""; row1=""V4|Department Analysis|Sales||||||||||||"";>"
    testCases(5) = "<LABELH1; row1=""V5|Existing ~Format|Mixed||||||||||||||""; <LABELH2; row1=""V6|Second Label|Data||||||||||||||"";>"
    
    Debug.Print "=== Testing Labels to Italic Conversion ==="
    
    For i = 0 To UBound(testCases)
        Debug.Print "Test Case " & (i + 1) & ":"
        Debug.Print "Original: " & testCases(i)
        Debug.Print "Result: " & ConvertLabelsToItalicInString(testCases(i))
        Debug.Print ""
    Next i
End Sub

Sub QuickLabelsTest()
    'Quick test to verify multiple labels work correctly
    
    Dim testString As String
    Dim result As String
    
    Debug.Print "=== Quick Labels Test ==="
    
    ' Test multiple target labels in one codestring
    testString = "<LABELH1; row1=""V1|Main Title|Data||||||||||||||""; <LABELH2; row1=""V2|Subtitle|Info||||||||||||||""; <COLUMNHEADER-E; row1=""H1|Col1|Col2|Col3||||||||||||"";"
    result = ConvertLabelsToItalicInString(testString)
    Debug.Print "Multiple Labels Test:"
    Debug.Print "Input:  " & testString
    Debug.Print "Output: " & result
    Debug.Print "Expected: All non-blank values in all three row parameters should have ~ prefix added"
    Debug.Print ""
    
    ' Test with mixed existing formatting
    testString = "<LABELH1; row1=""V3|~Existing|New Value|||||||||||||;""; <OTHER; row1=""X|Should Not Change||||||||||||||"";"
    result = ConvertLabelsToItalicInString(testString)
    Debug.Print "Mixed Formatting Test:"
    Debug.Print "Input:  " & testString
    Debug.Print "Output: " & result
    Debug.Print "Expected: Only the row after LABELH1 should be affected, existing ~ should not be duplicated"
    Debug.Print ""
    
End Sub 