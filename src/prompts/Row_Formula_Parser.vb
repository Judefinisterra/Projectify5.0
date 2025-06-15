Sub Row_Formula_Parser()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim driverNumber As Long
    Dim codeString As String
    Dim rowString As String
    Dim customFormula As String
    Dim aeFormula As String
    Dim driverValue As String
    
    ' Dictionary to store driver mappings (row -> driver name)
    Dim driverMap As Object
    Set driverMap = CreateObject("Scripting.Dictionary")
    
    Set ws = ActiveSheet
    
    ' Find last used row in column B
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    ' Initialize driver counter
    driverNumber = 1
    
    ' First pass: Create driver mappings for column D values
    For i = 10 To lastRow
        If ws.Cells(i, "D").Value <> "" Then
            driverValue = "A" & driverNumber
            driverMap.Add CStr(i), driverValue
            driverNumber = driverNumber + 1
        End If
    Next i
    
    ' Create or clear the Custom Codes sheet
    Dim customCodesSheet As Worksheet
    Dim sheetExists As Boolean
    Dim tempWs As Worksheet
    
    ' Check if Custom Codes sheet already exists
    sheetExists = False
    For Each tempWs In ActiveWorkbook.Worksheets
        If tempWs.Name = "Custom Codes" Then
            sheetExists = True
            Set customCodesSheet = tempWs
            Exit For
        End If
    Next tempWs
    
    ' Create new sheet if it doesn't exist, otherwise clear existing content
    If Not sheetExists Then
        Set customCodesSheet = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
        customCodesSheet.Name = "Custom Codes"
    Else
        customCodesSheet.Cells.Clear
    End If
    
    ' Reset ws to original active sheet
    Set ws = ActiveSheet
    
    Dim outputRow As Long
    outputRow = 1
    
    ' Second pass: Process each row
    For i = 10 To lastRow
        ' Check if column B is empty to determine BR vs FORMULA-S
        If ws.Cells(i, "B").Value = "" Or IsEmpty(ws.Cells(i, "B").Value) Then
            ' BR case
            codeString = "<BR;"
        Else
            ' FORMULA-S case
            codeString = "<FORMULA-S;"
            
            ' Get the formula from column AE
            If ws.Cells(i, "AE").Formula <> "" Then
                aeFormula = ws.Cells(i, "AE").Formula
                ' Remove the leading "=" if present
                If Left(aeFormula, 1) = "=" Then
                    aeFormula = Mid(aeFormula, 2)
                End If
            Else
                aeFormula = CStr(ws.Cells(i, "AE").Value)
            End If
            
            ' Transform the formula
            customFormula = TransformFormula(aeFormula, i, driverMap)
            codeString = codeString & " customformula=""" & customFormula & """;"
        End If
        
        ' Build the row string
        rowString = BuildRowString(ws, i, driverMap)
        codeString = codeString & " row1 = """ & rowString & """;"
        
        codeString = codeString & ">"
        
        ' Output the code string to Custom Codes sheet
        customCodesSheet.Cells(outputRow, 1).Value = codeString
        outputRow = outputRow + 1
        
    Next i
    
End Sub

Function TransformFormula(formula As String, currentRow As Long, driverMap As Object) As String
    Dim transformedFormula As String
    Dim i As Long
    Dim char As String
    Dim refStart As Long
    Dim refEnd As Long
    Dim cellRef As String
    Dim refRow As Long
    Dim refCol As String
    
    transformedFormula = formula
    
    ' Handle AE column references (rd{driver} transformation)
    ' Look for patterns like AE## where ## is a row number
    Dim aePattern As String
    Dim matches As Object
    Set matches = CreateObject("VBScript.RegExp")
    matches.Global = True
    matches.Pattern = "AE(\d+)"
    
    Dim match As Object
    For Each match In matches.Execute(transformedFormula)
        refRow = CLng(match.SubMatches(0))
        If driverMap.Exists(CStr(refRow)) Then
            transformedFormula = Replace(transformedFormula, match.Value, "rd{" & driverMap(CStr(refRow)) & "}")
        End If
    Next match
    
    ' Handle same-row column references (cd{column} transformation)
    ' Look for patterns like D##, E##, F##, G##, H##, I## where ## is the current row
    Dim colPattern As String
    Dim colMatches As Object
    Set colMatches = CreateObject("VBScript.RegExp")
    colMatches.Global = True
    colMatches.Pattern = "([DEFGHI])" & currentRow & "(?!\d)"
    
    For Each match In colMatches.Execute(transformedFormula)
        refCol = match.SubMatches(0)
        Select Case refCol
            Case "D": transformedFormula = Replace(transformedFormula, match.Value, "cd{1}")
            Case "E": transformedFormula = Replace(transformedFormula, match.Value, "cd{2}")
            Case "F": transformedFormula = Replace(transformedFormula, match.Value, "cd{3}")
            Case "G": transformedFormula = Replace(transformedFormula, match.Value, "cd{4}")
            Case "H": transformedFormula = Replace(transformedFormula, match.Value, "cd{5}")
            Case "I": transformedFormula = Replace(transformedFormula, match.Value, "cd{6}")
        End Select
    Next match
    
    TransformFormula = transformedFormula
End Function

Function BuildRowString(ws As Worksheet, rowNum As Long, driverMap As Object) As String
    Dim rowArray As String
    Dim driverValue As String
    
    ' Get driver value or empty if not in map
    If driverMap.Exists(CStr(rowNum)) Then
        driverValue = driverMap(CStr(rowNum))
    Else
        driverValue = ""
    End If
    
    ' Column A = (D) - Output Driver
    rowArray = driverValue & "(D)"
    
    ' Column B = (L) - Label
    rowArray = rowArray & "|" & GetCellValueWithFormatSymbol(ws.Cells(rowNum, "B")) & "(L)"
    
    ' Column C = (F) - FinCode
    rowArray = rowArray & "|" & GetCellValueWithFormatSymbol(ws.Cells(rowNum, "C")) & "(F)"
    
    ' Column D = (C1) - Fixed Assumption 1
    rowArray = rowArray & "|" & GetCellValueIgnoreAllCaps(ws.Cells(rowNum, "D")) & "(C1)"
    
    ' Column E = (C2) - Fixed Assumption 2
    rowArray = rowArray & "|" & GetCellValueIgnoreAllCaps(ws.Cells(rowNum, "E")) & "(C2)"
    
    ' Column F = (C3) - Fixed Assumption 3
    rowArray = rowArray & "|" & GetCellValueIgnoreAllCaps(ws.Cells(rowNum, "F")) & "(C3)"
    
    ' Column G = (C4) - Fixed Assumption 4
    On Error Resume Next
    If ws.Cells(rowNum, "G").Formula <> "" Then
        If Left(ws.Cells(rowNum, "G").Formula, 1) = "=" Then
            rowArray = rowArray & "|" & GetCellFormatSymbol(ws.Cells(rowNum, "G")) & "F(C4)"
        Else
            rowArray = rowArray & "|" & GetCellValueWithFormatSymbol(ws.Cells(rowNum, "G")) & "(C4)"
        End If
    Else
        rowArray = rowArray & "|(C4)"
    End If
    On Error GoTo 0
    
    ' Column H = (C5) - Fixed Assumption 5
    On Error Resume Next
    If ws.Cells(rowNum, "H").Formula <> "" Then
        If Left(ws.Cells(rowNum, "H").Formula, 1) = "=" Then
            rowArray = rowArray & "|" & GetCellFormatSymbol(ws.Cells(rowNum, "H")) & "F(C5)"
        Else
            rowArray = rowArray & "|" & GetCellValueWithFormatSymbol(ws.Cells(rowNum, "H")) & "(C5)"
        End If
    Else
        rowArray = rowArray & "|(C5)"
    End If
    On Error GoTo 0
    
    ' Column I = (C6) - Fixed Assumption 6
    On Error Resume Next
    If ws.Cells(rowNum, "I").Formula <> "" Then
        If Left(ws.Cells(rowNum, "I").Formula, 1) = "=" Then
            rowArray = rowArray & "|" & GetCellFormatSymbol(ws.Cells(rowNum, "I")) & "F(C6)"
        Else
            rowArray = rowArray & "|" & GetCellValueWithFormatSymbol(ws.Cells(rowNum, "I")) & "(C6)"
        End If
    Else
        rowArray = rowArray & "|(C6)"
    End If
    On Error GoTo 0
    
    ' For FORMULA-S, years should have $F format, for BR they should be plain format
    Dim isFormula As Boolean
    isFormula = Not (ws.Cells(rowNum, "B").Value = "" Or IsEmpty(ws.Cells(rowNum, "B").Value))
    
    ' Column K = (Y1) - Year 1
    On Error Resume Next
    If ws.Cells(rowNum, "K").Formula <> "" Then
        If Left(ws.Cells(rowNum, "K").Formula, 1) = "=" Then
            If isFormula Then
                rowArray = rowArray & "|" & GetCellFormatSymbol(ws.Cells(rowNum, "K")) & "F(Y1)"
            Else
                rowArray = rowArray & "|" & GetCellFormatSymbol(ws.Cells(rowNum, "K")) & "F(Y1)"
            End If
        Else
            If isFormula Then
                rowArray = rowArray & "|$" & GetCellValueWithFormatSymbol(ws.Cells(rowNum, "K")) & "F(Y1)"
            Else
                rowArray = rowArray & "|" & GetCellValueWithFormatSymbol(ws.Cells(rowNum, "K")) & "(Y1)"
            End If
        End If
    Else
        If isFormula Then
            rowArray = rowArray & "|$F(Y1)"
        Else
            rowArray = rowArray & "|(Y1)"
        End If
    End If
    On Error GoTo 0
    
    ' Column L = (Y2) - Year 2
    On Error Resume Next
    If ws.Cells(rowNum, "L").Formula <> "" Then
        If Left(ws.Cells(rowNum, "L").Formula, 1) = "=" Then
            If isFormula Then
                rowArray = rowArray & "|" & GetCellFormatSymbol(ws.Cells(rowNum, "L")) & "F(Y2)"
            Else
                rowArray = rowArray & "|" & GetCellFormatSymbol(ws.Cells(rowNum, "L")) & "F(Y2)"
            End If
        Else
            If isFormula Then
                rowArray = rowArray & "|$" & GetCellValueWithFormatSymbol(ws.Cells(rowNum, "L")) & "F(Y2)"
            Else
                rowArray = rowArray & "|" & GetCellValueWithFormatSymbol(ws.Cells(rowNum, "L")) & "(Y2)"
            End If
        End If
    Else
        If isFormula Then
            rowArray = rowArray & "|$F(Y2)"
        Else
            rowArray = rowArray & "|(Y2)"
        End If
    End If
    On Error GoTo 0
    
    ' Column M = (Y3) - Year 3
    On Error Resume Next
    If ws.Cells(rowNum, "M").Formula <> "" Then
        If Left(ws.Cells(rowNum, "M").Formula, 1) = "=" Then
            If isFormula Then
                rowArray = rowArray & "|" & GetCellFormatSymbol(ws.Cells(rowNum, "M")) & "F(Y3)"
            Else
                rowArray = rowArray & "|" & GetCellFormatSymbol(ws.Cells(rowNum, "M")) & "F(Y3)"
            End If
        Else
            If isFormula Then
                rowArray = rowArray & "|$" & GetCellValueWithFormatSymbol(ws.Cells(rowNum, "M")) & "F(Y3)"
            Else
                rowArray = rowArray & "|" & GetCellValueWithFormatSymbol(ws.Cells(rowNum, "M")) & "(Y3)"
            End If
        End If
    Else
        If isFormula Then
            rowArray = rowArray & "|$F(Y3)"
        Else
            rowArray = rowArray & "|(Y3)"
        End If
    End If
    On Error GoTo 0
    
    ' Column N = (Y4) - Year 4
    On Error Resume Next
    If ws.Cells(rowNum, "N").Formula <> "" Then
        If Left(ws.Cells(rowNum, "N").Formula, 1) = "=" Then
            If isFormula Then
                rowArray = rowArray & "|" & GetCellFormatSymbol(ws.Cells(rowNum, "N")) & "F(Y4)"
            Else
                rowArray = rowArray & "|" & GetCellFormatSymbol(ws.Cells(rowNum, "N")) & "F(Y4)"
            End If
        Else
            If isFormula Then
                rowArray = rowArray & "|$" & GetCellValueWithFormatSymbol(ws.Cells(rowNum, "N")) & "F(Y4)"
            Else
                rowArray = rowArray & "|" & GetCellValueWithFormatSymbol(ws.Cells(rowNum, "N")) & "(Y4)"
            End If
        End If
    Else
        If isFormula Then
            rowArray = rowArray & "|$F(Y4)"
        Else
            rowArray = rowArray & "|(Y4)"
        End If
    End If
    On Error GoTo 0
    
    ' Column O = (Y5) - Year 5
    On Error Resume Next
    If ws.Cells(rowNum, "O").Formula <> "" Then
        If Left(ws.Cells(rowNum, "O").Formula, 1) = "=" Then
            If isFormula Then
                rowArray = rowArray & "|" & GetCellFormatSymbol(ws.Cells(rowNum, "O")) & "F(Y5)"
            Else
                rowArray = rowArray & "|" & GetCellFormatSymbol(ws.Cells(rowNum, "O")) & "F(Y5)"
            End If
        Else
            If isFormula Then
                rowArray = rowArray & "|$" & GetCellValueWithFormatSymbol(ws.Cells(rowNum, "O")) & "F(Y5)"
            Else
                rowArray = rowArray & "|" & GetCellValueWithFormatSymbol(ws.Cells(rowNum, "O")) & "(Y5)"
            End If
        End If
    Else
        If isFormula Then
            rowArray = rowArray & "|$F(Y5)"
        Else
            rowArray = rowArray & "|(Y5)"
        End If
    End If
    On Error GoTo 0
    
    ' Column P = (Y6) - Year 6
    On Error Resume Next
    If ws.Cells(rowNum, "P").Formula <> "" Then
        If Left(ws.Cells(rowNum, "P").Formula, 1) = "=" Then
            If isFormula Then
                rowArray = rowArray & "|" & GetCellFormatSymbol(ws.Cells(rowNum, "P")) & "F(Y6)"
            Else
                rowArray = rowArray & "|" & GetCellFormatSymbol(ws.Cells(rowNum, "P")) & "F(Y6)"
            End If
        Else
            If isFormula Then
                rowArray = rowArray & "|$" & GetCellValueWithFormatSymbol(ws.Cells(rowNum, "P")) & "F(Y6)"
            Else
                rowArray = rowArray & "|" & GetCellValueWithFormatSymbol(ws.Cells(rowNum, "P")) & "(Y6)"
            End If
        End If
    Else
        If isFormula Then
            rowArray = rowArray & "|$F(Y6)"
        Else
            rowArray = rowArray & "|(Y6)"
        End If
    End If
    On Error GoTo 0
    
    ' Final column - ends with "|"
    rowArray = rowArray & "|"
    
    BuildRowString = rowArray
End Function

' Helper functions (you may need to include these from your original script)
Function GetCellValueWithFormatSymbol(cell As Range) As String
    ' This function should return the cell value with appropriate format symbols
    ' You'll need to implement this based on your existing helper functions
    GetCellValueWithFormatSymbol = CStr(cell.Value)
End Function

Function GetCellFormatSymbol(cell As Range) As String
    ' This function should return the format symbol for formulas
    ' You'll need to implement this based on your existing helper functions
    GetCellFormatSymbol = "$"
End Function

Function GetCellValueIgnoreAllCaps(cell As Range) As String
    ' This function should return the cell value ignoring all caps formatting
    ' You'll need to implement this based on your existing helper functions
    GetCellValueIgnoreAllCaps = CStr(cell.Value)
End Function 