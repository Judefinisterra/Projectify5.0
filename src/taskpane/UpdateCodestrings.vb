Sub UpdateCodestringsWithExtraPipes()
    'Updates codestrings in Training Data sheet to include extra pipes for columns D, E, F
    'Processes row1, row2, etc. parameters to ensure they have exactly 15 pipes
    'Results are printed in column C next to the original codestrings in column B
    
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
    
    Debug.Print "Starting codestring update process..."
    Debug.Print "Processing rows 2 to " & lastRow & " in Training Data sheet"
    Debug.Print "Results will be written to column C"
    
    changesMade = 0
    
    ' Clear column C header and add new header
    ws.Cells(1, "C").Value = "Updated Codestrings"
    
    ' Loop through each row from 2 to last populated row
    For i = 2 To lastRow
        cellValue = ws.Cells(i, "B").Value
        
        ' Skip empty cells
        If Trim(cellValue) <> "" Then
            ' Process the cell value to update row parameters
            updatedValue = UpdateRowParametersInCodestring(cellValue)
            
            ' Always write the result to column C (whether changed or not)
            ws.Cells(i, "C").Value = updatedValue
            
            ' Track if changes were made
            If updatedValue <> cellValue Then
                changesMade = changesMade + 1
                Debug.Print "Row " & i & ": Added pipes to row parameters"
                ' Optional: Add visual indicator for changed rows
                ws.Cells(i, "C").Interior.Color = RGB(255, 255, 200) ' Light yellow background
            Else
                Debug.Print "Row " & i & ": No changes needed"
                ' Clear any existing background color for unchanged rows
                ws.Cells(i, "C").Interior.ColorIndex = xlNone
            End If
        Else
            ' Clear column C if column B is empty
            ws.Cells(i, "C").Value = ""
            ws.Cells(i, "C").Interior.ColorIndex = xlNone
        End If
    Next i
    
    Debug.Print "Process complete. Updated " & changesMade & " codestrings."
    MsgBox "Codestring update complete!" & vbCrLf & _
           "Processed " & (lastRow - 1) & " rows." & vbCrLf & _
           "Updated " & changesMade & " codestrings with extra pipes." & vbCrLf & _
           "Results written to column C with yellow highlighting for changed rows.", vbInformation
End Sub

Function UpdateRowParametersInCodestring(codestring As String) As String
    'Updates all row parameters (row1, row2, etc.) in a codestring to have exactly 15 pipes
    
    Dim result As String
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    Dim i As Long
    Dim originalParam As String
    Dim updatedParam As String
    
    result = codestring
    
    ' Create regex object to find row parameters
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.Pattern = "row\d+\s*=\s*""([^""]*)""|row\d+\s*=\s*'([^']*)'"
    
    ' Find all row parameter matches
    Set matches = regex.Execute(codestring)
    
    ' Process matches in reverse order to avoid position shifts
    For i = matches.Count - 1 To 0 Step -1
        Set match = matches(i)
        originalParam = match.Value
        
        ' Extract the content between quotes
        Dim paramContent As String
        If InStr(originalParam, """") > 0 Then
            ' Double quotes
            paramContent = ExtractQuotedContent(originalParam, """")
        Else
            ' Single quotes
            paramContent = ExtractQuotedContent(originalParam, "'")
        End If
        
        ' Update the parameter content to have 15 pipes
        Dim updatedContent As String
        updatedContent = EnsureFifteenPipes(paramContent)
        
        ' Rebuild the parameter with updated content
        If InStr(originalParam, """") > 0 Then
            updatedParam = Replace(originalParam, """" & paramContent & """", """" & updatedContent & """")
        Else
            updatedParam = Replace(originalParam, "'" & paramContent & "'", "'" & updatedContent & "'")
        End If
        
        ' Replace in result string
        result = Left(result, match.FirstIndex) & updatedParam & Mid(result, match.FirstIndex + match.Length + 1)
    Next i
    
    UpdateRowParametersInCodestring = result
End Function

Function ExtractQuotedContent(parameter As String, quoteChar As String) As String
    'Extracts content between quote characters
    
    Dim startPos As Long
    Dim endPos As Long
    
    startPos = InStr(parameter, quoteChar)
    If startPos > 0 Then
        endPos = InStr(startPos + 1, parameter, quoteChar)
        If endPos > startPos Then
            ExtractQuotedContent = Mid(parameter, startPos + 1, endPos - startPos - 1)
        End If
    End If
End Function

Function EnsureFifteenPipes(content As String) As String
    'Ensures the content has exactly 15 pipes by adding them after the 3rd pipe (after column C)
    
    Dim pipeCount As Long
    Dim parts() As String
    Dim result As String
    Dim i As Long
    Dim pipesToAdd As Long
    Dim addedPipes As String
    
    ' Count existing pipes
    pipeCount = Len(content) - Len(Replace(content, "|", ""))
    
    Debug.Print "  Original content: " & content
    Debug.Print "  Current pipe count: " & pipeCount
    
    ' If already 15 pipes, return as is
    If pipeCount = 15 Then
        EnsureFifteenPipes = content
        Exit Function
    End If
    
    ' If more than 15 pipes, return as is (don't remove pipes)
    If pipeCount > 15 Then
        Debug.Print "  Warning: Content has more than 15 pipes (" & pipeCount & "). Leaving unchanged."
        EnsureFifteenPipes = content
        Exit Function
    End If
    
    ' Calculate how many pipes to add
    pipesToAdd = 15 - pipeCount
    
    ' Split by pipes to work with individual columns
    parts = Split(content, "|")
    
    ' Build result with extra pipes after column C (index 2)
    result = ""
    For i = 0 To UBound(parts)
        result = result & parts(i)
        
        ' Add original pipe
        If i < UBound(parts) Then
            result = result & "|"
        End If
        
        ' After column C (index 2), add the extra pipes
        If i = 2 Then
            addedPipes = String(pipesToAdd, "|")
            result = result & addedPipes
            Debug.Print "  Added " & pipesToAdd & " pipes after column C"
        End If
    Next i
    
    ' Verify final pipe count
    Dim finalPipeCount As Long
    finalPipeCount = Len(result) - Len(Replace(result, "|", ""))
    Debug.Print "  Updated content: " & result
    Debug.Print "  Final pipe count: " & finalPipeCount
    
    EnsureFifteenPipes = result
End Function

Sub TestUpdateCodestrings()
    'Test function to verify the update logic works correctly
    
    Dim testCases() As String
    Dim i As Long
    
    ' Define test cases
    ReDim testCases(3)
    testCases(0) = "AS|Description|Category|Value1|Value2|Value3|Value4|Value5|Value6|Value7|Value8|Value9"
    testCases(1) = "V|Test|Cat|D|E|F|G|H|I|K|L|M|N|O|P|R"
    testCases(2) = "A|B|C|G|H|I|K|L|M|N|O|P|R"
    testCases(3) = "|Label||||||||||||"
    
    Debug.Print "=== Testing EnsureFifteenPipes Function ==="
    
    For i = 0 To UBound(testCases)
        Debug.Print "Test Case " & (i + 1) & ":"
        Debug.Print "Result: " & EnsureFifteenPipes(testCases(i))
        Debug.Print ""
    Next i
End Sub 