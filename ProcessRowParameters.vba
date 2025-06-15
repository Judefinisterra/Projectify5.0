Sub ProcessRowParametersWithDollarSign()
    '
    ' This script iterates through every populated cell in column B
    ' and finds every row parameter (row1, row2, etc.)
    ' If column F has a value, it adds a $ in front of Y1-Y6 columns
    '
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String
    Dim modifiedValue As String
    Dim rowPattern As String
    Dim hasFinCode As Boolean
    
    ' Set the active worksheet (change this to your specific worksheet name if needed)
    Set ws = ActiveSheet
    
    ' Find the last row with data in column B
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    ' Loop through each populated cell in column B
    For i = 1 To lastRow
        cellValue = ws.Cells(i, "B").Value
        
        ' Check if the cell contains row parameter patterns (row1, row2, etc.)
        If InStr(cellValue, "row") > 0 And InStr(cellValue, "=") > 0 Then
            
            ' Check if the row parameter contains a value in the (F) section
            hasFinCode = CheckIfHasFinCode(cellValue)
            
            If hasFinCode Then
                ' Modify the row parameter to add $ signs to Y1-Y6 columns
                modifiedValue = AddDollarSignsToYColumns(cellValue)
                
                ' Update the cell with the modified value
                ws.Cells(i, "B").Value = modifiedValue
                
                ' Optional: Add some feedback
                Debug.Print "Modified row " & i & ": " & Left(modifiedValue, 100) & "..."
            End If
        End If
    Next i
    
    MsgBox "Processing complete! Modified rows with FinCode values in column F."
End Sub

Function CheckIfHasFinCode(rowString As String) As Boolean
    '
    ' This function checks if the row parameter has a value in the (F) section
    ' It looks for patterns like "CF: Non-cash(F)" or any text before "(F)"
    '
    
    Dim parts() As String
    Dim i As Integer
    
    CheckIfHasFinCode = False
    
    ' Split the string by pipe delimiter
    If InStr(rowString, "|") > 0 Then
        parts = Split(rowString, "|")
        
        ' Look for the (F) section (typically the 3rd element after splitting)
        For i = 0 To UBound(parts)
            If InStr(parts(i), "(F)") > 0 Then
                ' Check if there's content before "(F)" - meaning it has a FinCode
                If Len(Trim(Replace(parts(i), "(F)", ""))) > 0 Then
                    CheckIfHasFinCode = True
                    Exit Function
                End If
            End If
        Next i
    End If
End Function

Function AddDollarSignsToYColumns(rowString As String) As String
    '
    ' This function adds $ signs in front of Y1-Y6 columns
    ' It replaces F(Y1) with $F(Y1), F(Y2) with $F(Y2), etc.
    '
    
    Dim modifiedString As String
    Dim i As Integer
    
    modifiedString = rowString
    
    ' Replace F(Y1) through F(Y6) with $F(Y1) through $F(Y6)
    For i = 1 To 6
        ' Replace F(Y#) with $F(Y#) - but only if it's not already prefixed with $
        If InStr(modifiedString, "$F(Y" & i & ")") = 0 Then
            modifiedString = Replace(modifiedString, "F(Y" & i & ")", "$F(Y" & i & ")")
        End If
    Next i
    
    AddDollarSignsToYColumns = modifiedString
End Function

Sub TestProcessRowParameters()
    '
    ' Test function to demonstrate the functionality
    '
    
    Dim testString As String
    Dim result As String
    
    ' Test with the example provided
    testString = "row13=""(D)|Amortization - Capitalized Software Development(L)|CF: Non-cash(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|F(Y1)|F(Y2)|F(Y3)|F(Y4)|F(Y5)|F(Y6)|"""
    
    Debug.Print "Original: " & testString
    
    If CheckIfHasFinCode(testString) Then
        result = AddDollarSignsToYColumns(testString)
        Debug.Print "Modified: " & result
    Else
        Debug.Print "No FinCode found in (F) section"
    End If
End Sub

Sub ProcessSpecificRange()
    '
    ' Alternative version that allows you to specify a specific range
    ' Modify the range as needed
    '
    
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim cellValue As String
    Dim modifiedValue As String
    Dim hasFinCode As Boolean
    
    ' Set the worksheet and range
    Set ws = ActiveSheet
    Set rng = ws.Range("B:B") ' Process entire column B
    
    ' Or specify a specific range like:
    ' Set rng = ws.Range("B1:B1000")
    
    For Each cell In rng
        If cell.Value <> "" Then
            cellValue = cell.Value
            
            ' Check if the cell contains row parameter patterns
            If InStr(cellValue, "row") > 0 And InStr(cellValue, "=") > 0 Then
                
                ' Check if the row parameter contains a value in the (F) section
                hasFinCode = CheckIfHasFinCode(cellValue)
                
                If hasFinCode Then
                    ' Modify the row parameter to add $ signs to Y1-Y6 columns
                    modifiedValue = AddDollarSignsToYColumns(cellValue)
                    
                    ' Update the cell with the modified value
                    cell.Value = modifiedValue
                    
                    ' Optional: Add some feedback
                    Debug.Print "Modified cell " & cell.Address & ": " & Left(modifiedValue, 100) & "..."
                End If
            End If
        End If
    Next cell
    
    MsgBox "Processing complete! Modified cells with FinCode values in column F."
End Sub 