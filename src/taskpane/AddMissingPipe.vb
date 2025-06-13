Sub AddMissingPipeToRows()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    Dim changesMade As Long
    changesMade = 0
    
    ' Clear column C first
    ws.Range("C:C").Clear
    
    Dim i As Long
    For i = 1 To lastRow
        Dim cellValue As String
        cellValue = ws.Cells(i, 2).Value ' Column B
        
        If ContainsRowParameters(cellValue) Then
            Dim result As String
            result = FixAllRowParameterPipes(cellValue)
            
            If result <> cellValue Then
                ' Changes made - highlight in green
                ws.Cells(i, 3).Value = result
                ws.Cells(i, 3).Interior.Color = RGB(144, 238, 144) ' Light green
                changesMade = changesMade + 1
            Else
                ' No changes - write original
                ws.Cells(i, 3).Value = result
                ws.Cells(i, 3).Interior.Color = RGB(255, 255, 255) ' White
            End If
        End If
    Next i
    
    MsgBox "Process completed. " & changesMade & " rows were modified."
End Sub

Function ContainsRowParameters(inputString As String) As Boolean
    ' Check if the string contains any row parameter patterns (row1, row2, etc.)
    ContainsRowParameters = False
    
    If Len(inputString) = 0 Then Exit Function
    
    ' Look for pattern like "row1 = " or "row2=" etc.
    If InStr(inputString, "row") > 0 And InStr(inputString, "=") > 0 And InStr(inputString, Chr(34)) > 0 Then
        ' Additional check to ensure it's actually a row parameter
        Dim i As Long
        For i = 1 To 999
            If InStr(inputString, "row" & i) > 0 Then
                ContainsRowParameters = True
                Exit Function
            End If
        Next i
    End If
End Function

Function FixAllRowParameterPipes(inputString As String) As String
    ' Process all row parameters in the input string
    Dim result As String
    result = inputString
    
    ' Find and fix each row parameter (row1, row2, row3, etc.)
    Dim i As Long
    For i = 1 To 999
        Dim rowPattern As String
        rowPattern = "row" & i
        
        ' Keep looking for ALL instances of this row parameter number
        Dim searchPos As Long
        searchPos = 1
        
        Do While True
            ' Find the next instance of this row parameter
            Dim startPos As Long
            startPos = InStr(searchPos, result, rowPattern)
            
            ' If no more instances found, move to next row number
            If startPos = 0 Then Exit Do
            
            ' Make sure this is actually a row parameter (has = and quotes after it)
            Dim checkStart As Long
            checkStart = startPos + Len(rowPattern)
            Dim hasEquals As Boolean
            Dim hasQuotes As Boolean
            hasEquals = InStr(checkStart, result, "=") > 0
            hasQuotes = InStr(checkStart, result, Chr(34)) > 0
            
            If Not (hasEquals And hasQuotes) Then
                ' Not a valid row parameter, continue searching after this position
                searchPos = startPos + Len(rowPattern)
                GoTo ContinueSearch
            End If
            
            ' Find the end of this row parameter instance
            Dim endPos As Long
            Dim searchStart As Long
            searchStart = startPos + Len(rowPattern)
            
            ' Look for the next row parameter of ANY number
            Dim nextRowPos As Long
            nextRowPos = 0
            Dim j As Long
            For j = 1 To 999 ' Check all possible row numbers
                Dim nextPattern As String
                nextPattern = "row" & j
                Dim tempPos As Long
                tempPos = InStr(searchStart, result, nextPattern)
                If tempPos > 0 And (nextRowPos = 0 Or tempPos < nextRowPos) Then
                    ' Make sure this next row parameter is valid too
                    Dim nextCheckStart As Long
                    nextCheckStart = tempPos + Len(nextPattern)
                    If InStr(nextCheckStart, result, "=") > 0 And InStr(nextCheckStart, result, Chr(34)) > 0 Then
                        nextRowPos = tempPos
                    End If
                End If
            Next j
            
            If nextRowPos > 0 Then
                endPos = nextRowPos - 1
            Else
                endPos = Len(result)
            End If
            
            ' Extract this single row parameter instance
            Dim singleRowParam As String
            singleRowParam = Mid(result, startPos, endPos - startPos + 1)
            
            ' Fix the pipes in this single row parameter
            Dim fixedRowParam As String
            fixedRowParam = FixPipeCountTo15(singleRowParam)
            
            ' Replace the original with the fixed version
            If fixedRowParam <> singleRowParam Then
                Dim lengthDifference As Long
                lengthDifference = Len(fixedRowParam) - Len(singleRowParam)
                result = Left(result, startPos - 1) & fixedRowParam & Mid(result, endPos + 1)
                ' Adjust end position for the length change
                endPos = endPos + lengthDifference
            End If
            
            ' Continue searching after this instance
            searchPos = endPos + 1
            
ContinueSearch:
        Loop
    Next i
    
    FixAllRowParameterPipes = result
End Function

Function FixPipeCountTo15(inputString As String) As String
    ' Find the quoted section
    Dim startQuote As Long
    Dim endQuote As Long
    
    startQuote = InStr(inputString, Chr(34))
    If startQuote = 0 Then
        FixPipeCountTo15 = inputString
        Exit Function
    End If
    
    endQuote = InStr(startQuote + 1, inputString, Chr(34))
    If endQuote = 0 Then
        FixPipeCountTo15 = inputString
        Exit Function
    End If
    
    ' Extract the content between quotes
    Dim quotedContent As String
    quotedContent = Mid(inputString, startQuote + 1, endQuote - startQuote - 1)
    
    ' Count pipes in the quoted content
    Dim pipeCount As Long
    Dim j As Long
    pipeCount = 0
    
    For j = 1 To Len(quotedContent)
        If Mid(quotedContent, j, 1) = "|" Then
            pipeCount = pipeCount + 1
        End If
    Next j
    
    ' If already exactly 15 pipes, no change needed
    If pipeCount = 15 Then
        FixPipeCountTo15 = inputString
        Exit Function
    End If
    
    ' Find the position after the 3rd pipe
    Dim pipeFoundCount As Long
    Dim insertPosition As Long
    pipeFoundCount = 0
    insertPosition = 0
    
    For j = 1 To Len(quotedContent)
        If Mid(quotedContent, j, 1) = "|" Then
            pipeFoundCount = pipeFoundCount + 1
            If pipeFoundCount = 3 Then
                insertPosition = j
                Exit For
            End If
        End If
    Next j
    
    ' If we couldn't find the 3rd pipe, return original
    If insertPosition = 0 Then
        FixPipeCountTo15 = inputString
        Exit Function
    End If
    
    Dim newQuotedContent As String
    
    If pipeCount < 15 Then
        ' Add pipes after the 3rd pipe
        Dim pipesToAdd As Long
        pipesToAdd = 15 - pipeCount
        
        Dim additionalPipes As String
        additionalPipes = String(pipesToAdd, "|")
        
        newQuotedContent = Left(quotedContent, insertPosition) & additionalPipes & Mid(quotedContent, insertPosition + 1)
        
    ElseIf pipeCount > 15 Then
        ' Remove pipes after the 3rd pipe
        Dim pipesToRemove As Long
        pipesToRemove = pipeCount - 15
        
        ' Find all pipes after the 3rd one and remove the excess
        Dim afterThirdPipe As String
        afterThirdPipe = Mid(quotedContent, insertPosition + 1)
        
        ' Remove pipes from the beginning of afterThirdPipe
        Dim pipesRemoved As Long
        pipesRemoved = 0
        Dim k As Long
        For k = 1 To Len(afterThirdPipe)
            If Mid(afterThirdPipe, k, 1) = "|" And pipesRemoved < pipesToRemove Then
                pipesRemoved = pipesRemoved + 1
            Else
                afterThirdPipe = Mid(afterThirdPipe, k)
                Exit For
            End If
        Next k
        
        ' If we removed all characters, set to empty
        If pipesRemoved = Len(Mid(quotedContent, insertPosition + 1)) Then
            afterThirdPipe = ""
        End If
        
        newQuotedContent = Left(quotedContent, insertPosition) & afterThirdPipe
    End If
    
    ' Reconstruct the full string
    Dim beforeQuote As String
    Dim afterQuote As String
    beforeQuote = Left(inputString, startQuote)
    afterQuote = Mid(inputString, endQuote)
    
    FixPipeCountTo15 = beforeQuote & newQuotedContent & afterQuote
End Function

Sub TestAddMissingPipe()
    ' Test the function with various scenarios
    Debug.Print "=== Testing FixAllRowParameterPipes Function ==="
    
    ' Test case 1: Single row parameter with 14 pipes (needs one added)
    Dim test1 As String
    test1 = "row1 = ""A1|CEO||||||~$100000|$F|$F|$F|$F|$F|$F|"""
    Debug.Print "Test 1 Input:  " & test1
    Debug.Print "Test 1 Output: " & FixAllRowParameterPipes(test1)
    Debug.Print "Test 1 Contains: " & ContainsRowParameters(test1)
    Debug.Print ""
    
    ' Test case 2: Multiple row parameters with different pipe counts
    Dim test2 As String
    test2 = "row1 = ""A1|CEO||||||~$100000|$F|$F|$F|$F|$F|$F|""" & vbCrLf & _
            "row2 = ""B2|Manager||||||$50000|$G|$G|""" & vbCrLf & _
            "row3 = ""C3|Director|||||||||$75000|$H|$H|$H|$H|$H|$H|"""
    Debug.Print "Test 2 Input:  " & Replace(test2, vbCrLf, " | ")
    Debug.Print "Test 2 Output: " & Replace(FixAllRowParameterPipes(test2), vbCrLf, " | ")
    Debug.Print "Test 2 Contains: " & ContainsRowParameters(test2)
    Debug.Print ""
    
    ' Test case 3: Already correct pipe counts
    Dim test3 As String
    test3 = "row1 = ""A1|CEO|||||||~$100000|$F|$F|$F|$F|$F|$F|""" & vbCrLf & _
            "row2 = ""B2|Manager|||||||$50000|$G|$G|$G|$G|$G|$G|"""
    Debug.Print "Test 3 Input:  " & Replace(test3, vbCrLf, " | ")
    Debug.Print "Test 3 Output: " & Replace(FixAllRowParameterPipes(test3), vbCrLf, " | ")
    Debug.Print "Test 3 Contains: " & ContainsRowParameters(test3)
    Debug.Print ""
    
    ' Test case 4: Non-sequential row numbers
    Dim test4 As String
    test4 = "row1 = ""A1|CEO||||||~$100000|$F|$F|$F|$F|$F|$F|""" & vbCrLf & _
            "row5 = ""E5|VP||||||$80000|$I|$I|""" & vbCrLf & _
            "row10 = ""J10|Analyst||||||||$40000|$K|$K|$K|$K|$K|"""
    Debug.Print "Test 4 Input:  " & Replace(test4, vbCrLf, " | ")
    Debug.Print "Test 4 Output: " & Replace(FixAllRowParameterPipes(test4), vbCrLf, " | ")
    Debug.Print "Test 4 Contains: " & ContainsRowParameters(test4)
    Debug.Print ""
    
    ' Test case 5: No row parameters
    Dim test5 As String
    test5 = "Some text without row parameters"
    Debug.Print "Test 5 Input:  " & test5
    Debug.Print "Test 5 Output: " & FixAllRowParameterPipes(test5)
    Debug.Print "Test 5 Contains: " & ContainsRowParameters(test5)
    Debug.Print ""
    
    ' Test case 6: Mixed content with row parameters
    Dim test6 As String
    test6 = "Some header text" & vbCrLf & _
            "row1 = ""A1|CEO||||||~$100000|$F|$F|$F|$F|$F|$F|""" & vbCrLf & _
            "Some middle text" & vbCrLf & _
            "row2 = ""B2|Manager||||||$50000|$G|$G|""" & vbCrLf & _
            "Some footer text"
    Debug.Print "Test 6 Input:  " & Replace(test6, vbCrLf, " | ")
    Debug.Print "Test 6 Output: " & Replace(FixAllRowParameterPipes(test6), vbCrLf, " | ")
    Debug.Print "Test 6 Contains: " & ContainsRowParameters(test6)
    Debug.Print ""
    
    ' Test case 7: Multiple instances of same row parameter number (user's reported issue)
    Dim test7 As String
    test7 = "<BR; row1 = ""|||"";><BR; row1 = ""|||"";>"
    Debug.Print "Test 7 Input:  " & test7
    Debug.Print "Test 7 Output: " & FixAllRowParameterPipes(test7)
    Debug.Print "Test 7 Contains: " & ContainsRowParameters(test7)
    Debug.Print ""
    
    ' Test case 8: Multiple instances with different row numbers and pipe counts
    Dim test8 As String
    test8 = "row1 = ""A|B|C||||||$100|$F|$F|$F|$F|$F|$F|""; row2 = ""D|E|F|G|H|I|J|K|L|M|N|O|P|Q|R|S|T|U|V|W|X|Y|Z|""; row1 = ""X|Y|Z|"""
    Debug.Print "Test 8 Input:  " & test8
    Debug.Print "Test 8 Output: " & FixAllRowParameterPipes(test8)
    Debug.Print "Test 8 Contains: " & ContainsRowParameters(test8)
    Debug.Print ""
    
    Debug.Print "=== Testing Complete ==="
End Sub 