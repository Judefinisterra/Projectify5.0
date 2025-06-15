Sub Simple_Code_Generator()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim codeString As String
    Dim customFormula As String
    Dim aeFormula As String
    Dim outputText As String
    Dim filePath As String
    Dim fileNum As Integer
    Dim codestrings As Collection
    Dim driverCounter As Long
    Dim rowString As String
    Dim driverMap As Object
    Dim rowToDriverIndex As Object
    
    Set ws = ActiveSheet
    Set codestrings = New Collection
    Set driverMap = CreateObject("Scripting.Dictionary")
    Set rowToDriverIndex = CreateObject("Scripting.Dictionary")
    
    ' Find last used row in column B
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    ' Set file path
    filePath = "C:\Users\joeor\Dropbox\B - Freelance\C_Projectify\VanPC\Codestring outputs\codestrings.txt"
    
    ' First pass: Build driver map for AE references
    For i = 10 To lastRow
        If ws.Cells(i, "D").Value <> "" Then
            driverMap.Add CStr(i), CStr(ws.Cells(i, "D").Value)
        End If
    Next i
    
    ' Second pass: Generate initial codestrings and track driver rows
    Dim driverIndex As Long
    driverIndex = 1
    
    For i = 10 To lastRow
        
        ' Check if column D has all caps value for FORMULA-S rows
        Dim useColumnDAsCodeName As Boolean
        Dim columnDValue As String
        useColumnDAsCodeName = False
        
        If Not (ws.Cells(i, "B").Value = "" Or IsEmpty(ws.Cells(i, "B").Value)) And _
           Not (Right(Trim(CStr(ws.Cells(i, "B").Value)), 1) = ":") Then
            ' This will be a FORMULA-S row, check column D
            columnDValue = CStr(ws.Cells(i, "D").Value)
            If Trim(columnDValue) <> "" And UCase(columnDValue) = columnDValue And LCase(columnDValue) <> columnDValue Then
                useColumnDAsCodeName = True
            End If
        End If
        
        ' Build the row string with actual cell values and formatting
        rowString = BuildRowString(ws, i, useColumnDAsCodeName)
        
        ' Check if column B is empty
        If ws.Cells(i, "B").Value = "" Then
            ' BR case
            codeString = "<BR; row1 = """ & rowString & """;>"
        ElseIf Right(Trim(CStr(ws.Cells(i, "B").Value)), 1) = ":" Then
            ' LABELH3 case - label ends with colon
            codeString = "<LABELH3; row1 = """ & rowString & """;>"
        Else
            ' FORMULA-S case - track this as a driver row
            rowToDriverIndex.Add CStr(i), driverIndex
            driverIndex = driverIndex + 1
            
            ' Get formula from column AE
            If ws.Cells(i, "AE").Formula <> "" Then
                aeFormula = ws.Cells(i, "AE").Formula
                ' Remove the leading "=" if present
                If Left(aeFormula, 1) = "=" Then
                    customFormula = Mid(aeFormula, 2)
                Else
                    customFormula = aeFormula
                End If
            Else
                ' If no formula, use the value
                customFormula = CStr(ws.Cells(i, "AE").Value)
            End If
            
            ' Transform the formula references (only cd{} for now)
            customFormula = TransformFormula(customFormula, i, driverMap)
            
            ' Check if column B is bold
            Dim boldParam As String
            If ws.Cells(i, "B").Font.Bold Then
                boldParam = "bold=""true"""
            Else
                boldParam = "bold=""false"""
            End If
            
            ' Check if column K has a top border
            Dim topBorderParam As String
            If ws.Cells(i, "K").Borders(xlEdgeTop).LineStyle <> xlNone Then
                topBorderParam = "topborder=""true"""
            Else
                topBorderParam = "topborder=""false"""
            End If
            
            ' Check indent level of column B
            Dim indentParam As String
            Dim indentLevel As Long
            indentLevel = ws.Cells(i, "B").IndentLevel
            indentParam = "indent=""" & indentLevel & """"
            
            ' Check for specific sumif formulas in column K
            Dim sumifParam As String
            Dim kFormula As String
            sumifParam = ""
            
            If ws.Cells(i, "K").Formula <> "" Then
                kFormula = ws.Cells(i, "K").Formula
                
                If kFormula = "=SUMIF($4:$4, K$2, INDIRECT(ROW() & "":"" & ROW()))" Then
                    sumifParam = "; sumif=""yearend"""
                ElseIf kFormula = "=SUMIF($3:$3, K$2, INDIRECT(ROW() & "":"" & ROW()))" Then
                    sumifParam = "; sumif=""year"""
                ElseIf kFormula = "=INDEX(INDIRECT(ROW() & "":"" & ROW()),1,MATCH(INDIRECT(ADDRESS(2,COLUMN()-1,2)),$4:$4,0)+1)" Then
                    sumifParam = "; sumif=""offsetyear"""
                End If
            End If
            
            ' Determine code name and whether to include customformula
            Dim codeName As String
            Dim customFormulaParam As String
            
            If useColumnDAsCodeName Then
                codeName = columnDValue
                customFormulaParam = ""  ' No customformula for all-caps column D codes
            Else
                codeName = "FORMULA-S"
                customFormulaParam = "; customformula=""" & customFormula & """"
            End If
            
            ' Use placeholder for driver number - will be replaced in next pass
            codeString = "<" & codeName & "; " & boldParam & "; " & topBorderParam & "; " & indentParam & sumifParam & customFormulaParam & "; row1 = """ & rowString & """;>"
        End If
        
        ' Add codestring to collection
        codestrings.Add codeString
        
    Next i
    
    ' Third pass: Assign sequential driver numbers to FORMULA-S rows
    Dim updatedCodestrings As Collection
    Set updatedCodestrings = New Collection
    
    driverCounter = 1
    For i = 1 To codestrings.Count
        codeString = codestrings(i)
        
        ' Check if this is a FORMULA-S type codestring (could be FORMULA-S or custom name)
        If InStr(codeString, "PLACEHOLDER(D)") > 0 Then
            ' Replace placeholder with sequential driver number
            Dim finalDriverName As String
            finalDriverName = "A" & driverCounter
            codeString = Replace(codeString, "PLACEHOLDER(D)", finalDriverName & "(D)")
            
            driverCounter = driverCounter + 1
        End If
        
        updatedCodestrings.Add codeString
    Next i
    
    ' Build final driver map using rowToDriverIndex
    Dim finalDriverMap As Object
    Set finalDriverMap = CreateObject("Scripting.Dictionary")
    
    Dim rowNum As Variant
    For Each rowNum In rowToDriverIndex.Keys
        finalDriverMap.Add CStr(rowNum), "A" & rowToDriverIndex(rowNum)
    Next rowNum
    
    ' Fourth pass: Transform AE references to rd{} format using final driver numbers
    Dim finalCodestrings As Collection
    Set finalCodestrings = New Collection
    
    For i = 1 To updatedCodestrings.Count
        codeString = updatedCodestrings(i)
        
        ' Only transform FORMULA-S type codestrings (those with customformula parameter)
        If InStr(codeString, "customformula=") > 0 Then
            codeString = TransformAEReferences(codeString, finalDriverMap)
        End If
        
        finalCodestrings.Add codeString
    Next i
    
    ' Build final output text
    For i = 1 To finalCodestrings.Count
        outputText = outputText & finalCodestrings(i) & vbCrLf
    Next i
    
    ' Write to text file
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, outputText
    Close #fileNum
    
    ' Open the text file
    Shell "notepad.exe " & filePath, vbNormalFocus
    
End Sub

Function TransformFormula(formula As String, currentRow As Long, driverMap As Object) As String
    Dim transformedFormula As String
    Dim matches As Object
    Dim match As Object
    Dim refCol As String
    
    transformedFormula = formula
    
    ' Handle same-row column references (cd{column} transformation)
    ' Look for patterns like D##, E##, F##, G##, H##, I## where ## is the current row
    Set matches = CreateObject("VBScript.RegExp")
    matches.Global = True
    matches.Pattern = "([DEFGHI])" & currentRow & "(?!\d)"
    
    For Each match In matches.Execute(transformedFormula)
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

Function TransformAEReferences(codeString As String, finalDriverMap As Object) As String
    Dim transformedString As String
    Dim matches As Object
    Dim match As Object
    Dim refRow As Long
    
    transformedString = codeString
    
    ' Handle AE column references (rd{driver} transformation)
    ' Look for patterns like AE$## or AE## where ## is a row number
    Set matches = CreateObject("VBScript.RegExp")
    matches.Global = True
    matches.Pattern = "AE\$?(\d+)"
    
    For Each match In matches.Execute(transformedString)
        refRow = CLng(match.SubMatches(0))
        If finalDriverMap.Exists(CStr(refRow)) Then
            transformedString = Replace(transformedString, match.Value, "rd{" & finalDriverMap(CStr(refRow)) & "}")
        End If
    Next match
    
    TransformAEReferences = transformedString
End Function

Function BuildRowString(ws As Worksheet, rowNum As Long, Optional excludeColumnDFromC1 As Boolean = False) As String
    Dim rowArray As String
    Dim isFormula As Boolean
    Dim isLabelH3 As Boolean
    
    ' Determine row type
    If ws.Cells(rowNum, "B").Value = "" Or IsEmpty(ws.Cells(rowNum, "B").Value) Then
        ' BR case
        isFormula = False
        isLabelH3 = False
    ElseIf Right(Trim(CStr(ws.Cells(rowNum, "B").Value)), 1) = ":" Then
        ' LABELH3 case
        isFormula = False
        isLabelH3 = True
    Else
        ' FORMULA-S case
        isFormula = True
        isLabelH3 = False
    End If
    
    ' Column A = (D) - Output Driver
    If isFormula Then
        rowArray = "PLACEHOLDER(D)"
    ElseIf isLabelH3 Then
        rowArray = ""  ' No (D) for LABELH3
    Else
        rowArray = "(D)"  ' Empty (D) for BR
    End If
    
    ' Column B = (L) - Label
    If Not isFormula And Not isLabelH3 Then
        ' BR case - add "BR" to (L)
        rowArray = rowArray & "|BR(L)"
    Else
        rowArray = rowArray & "|" & GetCellValueWithFormatSymbol(ws.Cells(rowNum, "B")) & "(L)"
    End If
    
    ' Column C = (F) - FinCode
    rowArray = rowArray & "|" & GetCellValueWithFormatSymbol(ws.Cells(rowNum, "C")) & "(F)"
    
    ' Column D = (C1) - Fixed Assumption 1
    If excludeColumnDFromC1 Then
        rowArray = rowArray & "|(C1)"
    Else
        rowArray = rowArray & "|" & GetCellValueIgnoreAllCaps(ws.Cells(rowNum, "D")) & "(C1)"
    End If
    
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
    
    ' Column K = (Y1) - Year 1
    On Error Resume Next
    If ws.Cells(rowNum, "K").Formula <> "" Then
        If Left(ws.Cells(rowNum, "K").Formula, 1) = "=" Then
            rowArray = rowArray & "|" & GetCellFormatSymbol(ws.Cells(rowNum, "K")) & "F(Y1)"
        Else
            rowArray = rowArray & "|" & GetCellValueWithFormatSymbol(ws.Cells(rowNum, "K")) & "(Y1)"
        End If
    Else
        rowArray = rowArray & "|(Y1)"
    End If
    On Error GoTo 0
    
    ' Column L = (Y2) - Year 2
    On Error Resume Next
    If ws.Cells(rowNum, "L").Formula <> "" Then
        If Left(ws.Cells(rowNum, "L").Formula, 1) = "=" Then
            rowArray = rowArray & "|" & GetCellFormatSymbol(ws.Cells(rowNum, "L")) & "F(Y2)"
        Else
            rowArray = rowArray & "|" & GetCellValueWithFormatSymbol(ws.Cells(rowNum, "L")) & "(Y2)"
        End If
    Else
        rowArray = rowArray & "|(Y2)"
    End If
    On Error GoTo 0
    
    ' Column M = (Y3) - Year 3
    On Error Resume Next
    If ws.Cells(rowNum, "M").Formula <> "" Then
        If Left(ws.Cells(rowNum, "M").Formula, 1) = "=" Then
            rowArray = rowArray & "|" & GetCellFormatSymbol(ws.Cells(rowNum, "M")) & "F(Y3)"
        Else
            rowArray = rowArray & "|" & GetCellValueWithFormatSymbol(ws.Cells(rowNum, "M")) & "(Y3)"
        End If
    Else
        rowArray = rowArray & "|(Y3)"
    End If
    On Error GoTo 0
    
    ' Column N = (Y4) - Year 4
    On Error Resume Next
    If ws.Cells(rowNum, "N").Formula <> "" Then
        If Left(ws.Cells(rowNum, "N").Formula, 1) = "=" Then
            rowArray = rowArray & "|" & GetCellFormatSymbol(ws.Cells(rowNum, "N")) & "F(Y4)"
        Else
            rowArray = rowArray & "|" & GetCellValueWithFormatSymbol(ws.Cells(rowNum, "N")) & "(Y4)"
        End If
    Else
        rowArray = rowArray & "|(Y4)"
    End If
    On Error GoTo 0
    
    ' Column O = (Y5) - Year 5
    On Error Resume Next
    If ws.Cells(rowNum, "O").Formula <> "" Then
        If Left(ws.Cells(rowNum, "O").Formula, 1) = "=" Then
            rowArray = rowArray & "|" & GetCellFormatSymbol(ws.Cells(rowNum, "O")) & "F(Y5)"
        Else
            rowArray = rowArray & "|" & GetCellValueWithFormatSymbol(ws.Cells(rowNum, "O")) & "(Y5)"
        End If
    Else
        rowArray = rowArray & "|(Y5)"
    End If
    On Error GoTo 0
    
    ' Column P = (Y6) - Year 6
    On Error Resume Next
    If ws.Cells(rowNum, "P").Formula <> "" Then
        If Left(ws.Cells(rowNum, "P").Formula, 1) = "=" Then
            rowArray = rowArray & "|" & GetCellFormatSymbol(ws.Cells(rowNum, "P")) & "F(Y6)"
        Else
            rowArray = rowArray & "|" & GetCellValueWithFormatSymbol(ws.Cells(rowNum, "P")) & "(Y6)"
        End If
    Else
        rowArray = rowArray & "|(Y6)"
    End If
    On Error GoTo 0
    
    ' Final column - ends with "|"
    rowArray = rowArray & "|"
    
    BuildRowString = rowArray
End Function

Function GetCellFormatSymbol(cell As Range) As String
    'Returns the appropriate symbol prefix based on cell formatting
    Dim formatSymbol As String
    Dim numberFormat As String
    Dim isItalic As Boolean
    
    formatSymbol = ""
    numberFormat = cell.NumberFormat
    isItalic = cell.Font.Italic
    
    ' Check for currency formats
    If InStr(numberFormat, "$") > 0 Then
        formatSymbol = "$"
    ElseIf InStr(numberFormat, "£") > 0 Then
        formatSymbol = "£"
    ElseIf InStr(numberFormat, "€") > 0 Then
        formatSymbol = "€"
    ElseIf InStr(numberFormat, "¥") > 0 Then
        formatSymbol = "¥"
    ElseIf InStr(numberFormat, "%") > 0 Then
        ' Percentage formatting - no prefix needed as "%" suffix will be detected
        formatSymbol = ""
    ElseIf InStr(numberFormat, "#,##0.0x") > 0 Or InStr(numberFormat, "x") > 0 Then
        ' Factor formatting - no prefix needed as "x" suffix will be detected
        formatSymbol = ""
    ElseIf InStr(numberFormat, "#,##0") > 0 And InStr(numberFormat, "$") = 0 Then
        ' Volume formatting (numbers without currency) - no prefix needed
        formatSymbol = ""
    ElseIf InStr(numberFormat, "mmm") > 0 Or InStr(numberFormat, "yyyy") > 0 Then
        ' Date formatting - no prefix needed
        formatSymbol = ""
    End If
    
    ' Add tilde for italic formatting
    If isItalic And formatSymbol <> "" Then
        formatSymbol = "~" & formatSymbol
    ElseIf isItalic And formatSymbol = "" Then
        formatSymbol = "~"
    End If
    
    GetCellFormatSymbol = formatSymbol
End Function

Function GetCellValueWithFormatSymbol(cell As Range) As String
    'Returns the cell value with appropriate formatting symbols
    Dim cellValue As String
    Dim formatSymbol As String
    Dim numberFormat As String
    
    cellValue = CStr(cell.Value)
    
    ' Check if cell is empty or contains only spaces
    If Trim(cellValue) = "" Then
        GetCellValueWithFormatSymbol = ""
        Exit Function
    End If
    
    formatSymbol = GetCellFormatSymbol(cell)
    numberFormat = cell.NumberFormat
    
    ' Check if it's percentage formatting and add "%" suffix to numeric values
    If InStr(numberFormat, "%") > 0 And IsNumeric(cell.Value) Then
        cellValue = cellValue & "%"
    ' Check if it's factor formatting and add "x" suffix to numeric values
    ElseIf (InStr(numberFormat, "#,##0.0x") > 0 Or InStr(numberFormat, "x") > 0) And IsNumeric(cell.Value) Then
        cellValue = cellValue & "x"
    End If
    
    GetCellValueWithFormatSymbol = formatSymbol & cellValue
End Function

Function GetCellValueIgnoreAllCaps(cell As Range) As String
    'Returns the cell value with formatting symbols, but ignores all-caps values
    Dim cellValue As String
    
    cellValue = CStr(cell.Value)
    
    ' Check if cell is empty or contains only spaces
    If Trim(cellValue) = "" Then
        GetCellValueIgnoreAllCaps = ""
        Exit Function
    End If
    
    ' Check if value is all caps (and not just numbers/symbols)
    If UCase(cellValue) = cellValue And LCase(cellValue) <> cellValue Then
        ' Value is all caps, return empty string
        GetCellValueIgnoreAllCaps = ""
        Exit Function
    End If
    
    ' Not all caps, return formatted value
    GetCellValueIgnoreAllCaps = GetCellValueWithFormatSymbol(cell)
End Function 