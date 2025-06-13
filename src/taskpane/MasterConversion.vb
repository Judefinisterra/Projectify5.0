Option Explicit

' Master Conversion Module - Contains 5 independent subs for codestring processing
' All subs operate on "Training Data" sheet, process column B, and output to column C

' ==============================================================================
' SUB 1: EnsurePipes
' Ensures exactly 15 pipes in each row parameter
' ==============================================================================
Sub EnsurePipes()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String
    Dim modifiedValue As String
    
    Set ws = ThisWorkbook.Worksheets("Training Data")
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    For i = 2 To lastRow
        cellValue = ws.Cells(i, "B").Value
        modifiedValue = EnsureFifteenPipes(cellValue)
        ws.Cells(i, "C").Value = modifiedValue
        
        ' Highlight if changed
        If cellValue <> modifiedValue Then
            ws.Cells(i, "C").Interior.Color = RGB(255, 255, 0) ' Yellow
        End If
    Next i
    
    MsgBox "EnsurePipes completed. Modified rows highlighted in yellow."
End Sub

Private Function EnsureFifteenPipes(inputText As String) As String
    Dim result As String
    Dim rowPattern As String
    Dim matches As Object
    Dim match As Object
    Dim i As Long
    
    Set matches = CreateObject("VBScript.RegExp")
    result = inputText
    
    ' Pattern to match row parameters like row1 = "content"
    rowPattern = "(row\d+\s*=\s*""[^""]*"")"
    
    With matches
        .Global = True
        .Pattern = rowPattern
    End With
    
    Set matches = matches.Execute(inputText)
    
    For i = matches.Count - 1 To 0 Step -1 ' Process in reverse to maintain positions
        Set match = matches(i)
        Dim rowContent As String
        Dim quotedContent As String
        Dim pipeCount As Long
        Dim modifiedRowContent As String
        
        rowContent = match.Value
        
        ' Extract content between quotes
        Dim startQuote As Long, endQuote As Long
        startQuote = InStr(rowContent, """")
        endQuote = InStrRev(rowContent, """")
        
        If startQuote > 0 And endQuote > startQuote Then
            quotedContent = Mid(rowContent, startQuote + 1, endQuote - startQuote - 1)
            
            ' Count pipes
            pipeCount = Len(quotedContent) - Len(Replace(quotedContent, "|", ""))
            
            If pipeCount <> 15 Then
                ' Find position after column C (3rd pipe) to insert additional pipes
                Dim pipes As Variant
                pipes = Split(quotedContent, "|")
                
                If UBound(pipes) >= 2 Then ' At least 3 columns exist
                    ' Reconstruct with additional pipes after column C
                    Dim newContent As String
                    Dim j As Long
                    
                    ' Add first 3 columns (A, B, C)
                    newContent = pipes(0) & "|" & pipes(1) & "|" & pipes(2)
                    
                    ' Add required number of pipes to reach 15
                    Dim pipesToAdd As Long
                    pipesToAdd = 15 - pipeCount
                    
                    For j = 1 To pipesToAdd
                        newContent = newContent & "|"
                    Next j
                    
                    ' Add remaining columns
                    For j = 3 To UBound(pipes)
                        newContent = newContent & "|" & pipes(j)
                    Next j
                    
                    ' Replace in the row content
                    modifiedRowContent = Left(rowContent, startQuote) & newContent & Right(rowContent, Len(rowContent) - endQuote + 1)
                    
                    ' Replace in the result
                    result = Left(result, match.FirstIndex) & modifiedRowContent & Mid(result, match.FirstIndex + match.Length + 1)
                End If
            End If
        End If
    Next i
    
    EnsureFifteenPipes = result
End Function

' ==============================================================================
' SUB 2: ConvertLegacyFormat
' Converts format="..." parameters to inline symbols
' ==============================================================================
Sub ConvertLegacyFormat()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String
    Dim modifiedValue As String
    
    Set ws = ThisWorkbook.Worksheets("Training Data")
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    For i = 2 To lastRow
        cellValue = ws.Cells(i, "B").Value
        modifiedValue = ConvertLegacyFormatInString(cellValue)
        ws.Cells(i, "C").Value = modifiedValue
        
        ' Highlight if changed
        If cellValue <> modifiedValue Then
            ws.Cells(i, "C").Interior.Color = RGB(0, 255, 0) ' Green
        End If
    Next i
    
    MsgBox "ConvertLegacyFormat completed. Modified rows highlighted in green."
End Sub

Private Function ConvertLegacyFormatInString(inputText As String) As String
    Dim result As String
    Dim formatValue As String
    
    result = inputText
    formatValue = ExtractLegacyFormatParameter(result)
    
    If formatValue <> "" Then
        result = ApplyLegacyFormatToRow(result, formatValue)
        result = RemoveLegacyFormatParameter(result)
    End If
    
    ConvertLegacyFormatInString = result
End Function

Private Function ExtractLegacyFormatParameter(inputText As String) As String
    Dim regex As Object
    Dim matches As Object
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    regex.IgnoreCase = True
    regex.Pattern = "\bformat\s*=\s*""([^""]*)""|format\s*=\s*([^;,\s]*)"
    
    Set matches = regex.Execute(inputText)
    
    If matches.Count > 0 Then
        If matches(0).SubMatches(0) <> "" Then
            ExtractLegacyFormatParameter = matches(0).SubMatches(0)
        Else
            ExtractLegacyFormatParameter = matches(0).SubMatches(1)
        End If
    Else
        ExtractLegacyFormatParameter = ""
    End If
End Function

Private Function ApplyLegacyFormatToRow(inputText As String, formatType As String) As String
    Dim result As String
    result = inputText
    
    Select Case LCase(Trim(formatType))
        Case "dollar"
            result = ApplyLegacyDollarFormat(result)
        Case "volume"
            result = ApplyLegacyVolumeFormat(result)
        Case "percent"
            result = ApplyLegacyPercentFormat(result)
        Case "dollaritalic", "dollar italic"
            result = ApplyLegacyDollarItalicFormat(result)
    End Select
    
    ApplyLegacyFormatToRow = result
End Function

Private Function ApplyLegacyDollarFormat(inputText As String) As String
    ApplyLegacyDollarFormat = ApplyLegacySymbolToColumns(inputText, "$", "10,11,12,13,14,15")
End Function

Private Function ApplyLegacyVolumeFormat(inputText As String) As String
    Dim result As String
    result = ApplyLegacySymbolToColumns(inputText, "~", "2,10,11,12,13,14,15")
    ApplyLegacyVolumeFormat = result
End Function

Private Function ApplyLegacyPercentFormat(inputText As String) As String
    ApplyLegacyPercentFormat = ApplyLegacySymbolToColumns(inputText, "~", "2")
End Function

Private Function ApplyLegacyDollarItalicFormat(inputText As String) As String
    Dim result As String
    result = ApplyLegacySymbolToColumns(inputText, "~$", "10,11,12,13,14,15")
    result = ApplyLegacySymbolToColumns(result, "~", "2")
    ApplyLegacyDollarItalicFormat = result
End Function

Private Function ApplyLegacySymbolToColumns(inputText As String, symbol As String, columnList As String) As String
    Dim result As String
    Dim rowPattern As String
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    Dim i As Long
    
    Set regex = CreateObject("VBScript.RegExp")
    result = inputText
    
    rowPattern = "(row\d+\s*=\s*""[^""]*"")"
    
    With regex
        .Global = True
        .Pattern = rowPattern
    End With
    
    Set matches = regex.Execute(inputText)
    
    For i = matches.Count - 1 To 0 Step -1
        Set match = matches(i)
        Dim modifiedRow As String
        modifiedRow = ApplyLegacySymbolToSingleRow(match.Value, symbol, columnList)
        result = Left(result, match.FirstIndex) & modifiedRow & Mid(result, match.FirstIndex + match.Length + 1)
    Next i
    
    ApplyLegacySymbolToColumns = result
End Function

Private Function ApplyLegacySymbolToSingleRow(rowText As String, symbol As String, columnList As String) As String
    Dim quotedContent As String
    Dim startQuote As Long, endQuote As Long
    Dim columns As Variant
    Dim columnNumbers As Variant
    Dim i As Long, j As Long
    
    startQuote = InStr(rowText, """")
    endQuote = InStrRev(rowText, """")
    
    If startQuote > 0 And endQuote > startQuote Then
        quotedContent = Mid(rowText, startQuote + 1, endQuote - startQuote - 1)
        columns = Split(quotedContent, "|")
        columnNumbers = Split(columnList, ",")
        
        For i = LBound(columnNumbers) To UBound(columnNumbers)
            Dim colNum As Long
            colNum = CLng(Trim(columnNumbers(i))) - 1 ' Convert to 0-based
            
            If colNum >= 0 And colNum <= UBound(columns) Then
                Dim cellValue As String
                cellValue = Trim(columns(colNum))
                
                If cellValue <> "" And Not (symbol = "$" And Left(cellValue, 1) = "$") And _
                   Not (symbol = "~" And Left(cellValue, 1) = "~") And _
                   Not (symbol = "~$" And Left(cellValue, 2) = "~$") Then
                    columns(colNum) = symbol & cellValue
                End If
            End If
        Next i
        
        quotedContent = Join(columns, "|")
        ApplyLegacySymbolToSingleRow = Left(rowText, startQuote) & quotedContent & Right(rowText, Len(rowText) - endQuote + 1)
    Else
        ApplyLegacySymbolToSingleRow = rowText
    End If
End Function

Private Function RemoveLegacyFormatParameter(inputText As String) As String
    Dim regex As Object
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.Pattern = "\bformat\s*=\s*""[^""]*""\s*[;,]?|\bformat\s*=\s*[^;,\s]*\s*[;,]?"
    
    RemoveLegacyFormatParameter = regex.Replace(inputText, "")
End Function

' ==============================================================================
' SUB 3: ConvertLegacyColumnformat
' Converts columnformat="..." parameters to inline symbols
' ==============================================================================
Sub ConvertLegacyColumnformat()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String
    Dim modifiedValue As String
    
    Set ws = ThisWorkbook.Worksheets("Training Data")
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    For i = 2 To lastRow
        cellValue = ws.Cells(i, "B").Value
        modifiedValue = ConvertColumnFormatInString(cellValue)
        ws.Cells(i, "C").Value = modifiedValue
        
        ' Highlight if changed
        If cellValue <> modifiedValue Then
            ws.Cells(i, "C").Interior.Color = RGB(255, 0, 255) ' Purple
        End If
    Next i
    
    MsgBox "ConvertLegacyColumnformat completed. Modified rows highlighted in purple."
End Sub

Private Function ConvertColumnFormatInString(inputText As String) As String
    Dim result As String
    Dim columnFormatValue As String
    
    result = inputText
    columnFormatValue = ExtractColumnFormatParameter(result)
    
    If columnFormatValue <> "" Then
        result = ApplyColumnFormatToRow(result, columnFormatValue)
    End If
    
    ConvertColumnFormatInString = result
End Function

Private Function ExtractColumnFormatParameter(inputText As String) As String
    Dim regex As Object
    Dim matches As Object
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    regex.IgnoreCase = True
    regex.Pattern = "columnformat\s*=\s*""([^""]*)""|columnformat\s*=\s*([^;,\s]*)"
    
    Set matches = regex.Execute(inputText)
    
    If matches.Count > 0 Then
        If matches(0).SubMatches(0) <> "" Then
            ExtractColumnFormatParameter = matches(0).SubMatches(0)
        Else
            ExtractColumnFormatParameter = matches(0).SubMatches(1)
        End If
    Else
        ExtractColumnFormatParameter = ""
    End If
End Function

Private Function ApplyColumnFormatToRow(inputText As String, columnFormatValue As String) As String
    Dim result As String
    Dim formats As Variant
    Dim i As Long
    
    result = inputText
    
    ' Handle both / and \ as separators
    columnFormatValue = Replace(columnFormatValue, "\", "/")
    formats = Split(columnFormatValue, "/")
    
    For i = LBound(formats) To UBound(formats)
        Dim formatType As String
        Dim columnPosition As Long
        
        formatType = Trim(formats(i))
        columnPosition = i + 1 ' 1-based position
        
        If LCase(formatType) = "dollaritalic" Then
            result = ApplyColumnFormatSymbol(result, "~$", GetExcelColumnFromPosition(columnPosition))
        End If
        ' Add other format types as needed
    Next i
    
    ApplyColumnFormatToRow = result
End Function

Private Function GetExcelColumnFromPosition(position As Long) As Long
    ' Map format positions to Excel columns (1→I=9, 2→H=8, 3→G=7, etc.)
    Select Case position
        Case 1: GetExcelColumnFromPosition = 9  ' Column I
        Case 2: GetExcelColumnFromPosition = 8  ' Column H
        Case 3: GetExcelColumnFromPosition = 7  ' Column G
        Case 4: GetExcelColumnFromPosition = 6  ' Column F
        Case 5: GetExcelColumnFromPosition = 5  ' Column E
        Case 6: GetExcelColumnFromPosition = 4  ' Column D
        Case Else: GetExcelColumnFromPosition = 0 ' Invalid
    End Select
End Function

Private Function ApplyColumnFormatSymbol(inputText As String, symbol As String, excelColumn As Long) As String
    Dim result As String
    Dim rowPattern As String
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    Dim i As Long
    
    Set regex = CreateObject("VBScript.RegExp")
    result = inputText
    
    rowPattern = "(row\d+\s*=\s*""[^""]*"")"
    
    With regex
        .Global = True
        .Pattern = rowPattern
    End With
    
    Set matches = regex.Execute(inputText)
    
    For i = matches.Count - 1 To 0 Step -1
        Set match = matches(i)
        Dim modifiedRow As String
        modifiedRow = ApplyColumnFormatSymbolToSingleRow(match.Value, symbol, excelColumn)
        result = Left(result, match.FirstIndex) & modifiedRow & Mid(result, match.FirstIndex + match.Length + 1)
    Next i
    
    ApplyColumnFormatSymbol = result
End Function

Private Function ApplyColumnFormatSymbolToSingleRow(rowText As String, symbol As String, excelColumn As Long) As String
    Dim quotedContent As String
    Dim startQuote As Long, endQuote As Long
    Dim columns As Variant
    Dim colIndex As Long
    
    startQuote = InStr(rowText, """")
    endQuote = InStrRev(rowText, """")
    
    If startQuote > 0 And endQuote > startQuote Then
        quotedContent = Mid(rowText, startQuote + 1, endQuote - startQuote - 1)
        columns = Split(quotedContent, "|")
        colIndex = excelColumn - 1 ' Convert to 0-based
        
        If colIndex >= 0 And colIndex <= UBound(columns) Then
            Dim cellValue As String
            cellValue = Trim(columns(colIndex))
            
            If cellValue <> "" And Left(cellValue, Len(symbol)) <> symbol Then
                columns(colIndex) = symbol & cellValue
            End If
        End If
        
        quotedContent = Join(columns, "|")
        ApplyColumnFormatSymbolToSingleRow = Left(rowText, startQuote) & quotedContent & Right(rowText, Len(rowText) - endQuote + 1)
    Else
        ApplyColumnFormatSymbolToSingleRow = rowText
    End If
End Function

' ==============================================================================
' SUB 4: Percentify
' Adds % to values when codestring contains %
' ==============================================================================
Sub Percentify()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String
    Dim modifiedValue As String
    
    Set ws = ThisWorkbook.Worksheets("Training Data")
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    For i = 2 To lastRow
        cellValue = ws.Cells(i, "B").Value
        
        ' Only process if the codestring contains %
        If InStr(cellValue, "%") > 0 Then
            modifiedValue = ConvertToPercentages(cellValue)
            ws.Cells(i, "C").Value = modifiedValue
            
            ' Highlight if changed
            If cellValue <> modifiedValue Then
                ws.Cells(i, "C").Interior.Color = RGB(255, 255, 0) ' Yellow
            End If
        Else
            ws.Cells(i, "C").Value = cellValue
        End If
    Next i
    
    MsgBox "Percentify completed. Modified rows highlighted in yellow."
End Sub

Private Function ConvertToPercentages(inputText As String) As String
    Dim result As String
    Dim rowPattern As String
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    Dim i As Long
    
    Set regex = CreateObject("VBScript.RegExp")
    result = inputText
    
    rowPattern = "(row\d+\s*=\s*""[^""]*"")"
    
    With regex
        .Global = True
        .Pattern = rowPattern
    End With
    
    Set matches = regex.Execute(inputText)
    
    For i = matches.Count - 1 To 0 Step -1
        Set match = matches(i)
        Dim modifiedRow As String
        modifiedRow = ConvertRowToPercentages(match.Value)
        result = Left(result, match.FirstIndex) & modifiedRow & Mid(result, match.FirstIndex + match.Length + 1)
    Next i
    
    ConvertToPercentages = result
End Function

Private Function ConvertRowToPercentages(rowText As String) As String
    Dim quotedContent As String
    Dim startQuote As Long, endQuote As Long
    Dim columns As Variant
    Dim i As Long
    
    startQuote = InStr(rowText, """")
    endQuote = InStrRev(rowText, """")
    
    If startQuote > 0 And endQuote > startQuote Then
        quotedContent = Mid(rowText, startQuote + 1, endQuote - startQuote - 1)
        columns = Split(quotedContent, "|")
        
        ' Process columns after column C (index 3 and beyond)
        For i = 3 To UBound(columns)
            Dim cellValue As String
            cellValue = Trim(columns(i))
            
            If cellValue <> "" Then
                columns(i) = ConvertValueToPercentage(cellValue)
            End If
        Next i
        
        quotedContent = Join(columns, "|")
        ConvertRowToPercentages = Left(rowText, startQuote) & quotedContent & Right(rowText, Len(rowText) - endQuote + 1)
    Else
        ConvertRowToPercentages = rowText
    End If
End Function

Private Function ConvertValueToPercentage(value As String) As String
    Dim result As String
    Dim numValue As Double
    
    result = value
    
    ' Handle decimal conversion (e.g., 0.04 -> 4%)
    If IsNumeric(value) Then
        numValue = CDbl(value)
        If numValue < 1 And numValue > 0 Then
            result = CStr(Round(numValue * 100, 2)) & "%"
        ElseIf numValue >= 1 Then
            result = value & "%"
        End If
    ElseIf Right(value, 1) <> "%" Then
        ' For non-numeric values like "F", just add %
        result = value & "%"
    End If
    
    ConvertValueToPercentage = result
End Function

' ==============================================================================
' SUB 5: PrefixTildeForLabels
' Adds ~ prefix to strings ending with : in row parameters
' ==============================================================================
Sub PrefixTildeForLabels()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String
    Dim modifiedValue As String
    
    Set ws = ThisWorkbook.Worksheets("Training Data")
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    For i = 2 To lastRow
        cellValue = ws.Cells(i, "B").Value
        modifiedValue = AddTildePrefixToLabels(cellValue)
        ws.Cells(i, "C").Value = modifiedValue
        
        ' Highlight if changed
        If cellValue <> modifiedValue Then
            ws.Cells(i, "C").Interior.Color = RGB(0, 255, 255) ' Cyan
        End If
    Next i
    
    MsgBox "PrefixTildeForLabels completed. Modified rows highlighted in cyan."
End Sub

Private Function AddTildePrefixToLabels(inputText As String) As String
    Dim result As String
    Dim rowPattern As String
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    Dim i As Long
    
    Set regex = CreateObject("VBScript.RegExp")
    result = inputText
    
    rowPattern = "(row\d+\s*=\s*""[^""]*"")"
    
    With regex
        .Global = True
        .Pattern = rowPattern
    End With
    
    Set matches = regex.Execute(inputText)
    
    For i = matches.Count - 1 To 0 Step -1
        Set match = matches(i)
        Dim modifiedRow As String
        modifiedRow = AddTildePrefixToLabelsInRow(match.Value)
        result = Left(result, match.FirstIndex) & modifiedRow & Mid(result, match.FirstIndex + match.Length + 1)
    Next i
    
    AddTildePrefixToLabels = result
End Function

Private Function AddTildePrefixToLabelsInRow(rowText As String) As String
    Dim quotedContent As String
    Dim startQuote As Long, endQuote As Long
    Dim columns As Variant
    Dim i As Long
    
    startQuote = InStr(rowText, """")
    endQuote = InStrRev(rowText, """")
    
    If startQuote > 0 And endQuote > startQuote Then
        quotedContent = Mid(rowText, startQuote + 1, endQuote - startQuote - 1)
        columns = Split(quotedContent, "|")
        
        For i = LBound(columns) To UBound(columns)
            Dim cellValue As String
            cellValue = Trim(columns(i))
            
            ' Check if the cell value ends with : and doesn't already start with ~
            If Len(cellValue) > 0 And Right(cellValue, 1) = ":" And Left(cellValue, 1) <> "~" Then
                columns(i) = "~" & cellValue
            End If
        Next i
        
        quotedContent = Join(columns, "|")
        AddTildePrefixToLabelsInRow = Left(rowText, startQuote) & quotedContent & Right(rowText, Len(rowText) - endQuote + 1)
    Else
        AddTildePrefixToLabelsInRow = rowText
    End If
End Function
