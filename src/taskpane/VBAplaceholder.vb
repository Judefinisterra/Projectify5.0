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
    
    ' Check if it's factor formatting and add "x" suffix to numeric values
    If (InStr(numberFormat, "#,##0.0x") > 0 Or InStr(numberFormat, "x") > 0) And IsNumeric(cell.Value) Then
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

Sub Populate_Code_DB_V3()
    'Add rowLabels and columnLabels
    
    Dim CodeItem As CodeDBClass
    Dim temparray() As Variant
    Dim tempCollection As Collection
    Dim linkrange As Range
    Dim nm As Name
    Dim lastRow As Long
    Dim nameExists As Boolean
    Dim codeString As String
    Dim fincodecount As Long
    Dim outputdrivercount As Long
    
    
    Set CodeCollection = New Collection

    With Codes
        .Outline.ShowLevels RowLevels:=2, ColumnLevels:=2
    
        lastRow = .Cells(.Rows.count, "B").End(xlUp).row
        
'        codei = 1
        
        For i = 9 To lastRow
            If .Cells(i, "D").value <> .Cells(i - 1, "D").value And .Cells(i, "D") <> "" Then
                
                Set CodeItem = New CodeDBClass
                
                
            
                CodeItem.codeName = .Cells(i, "D").value
                CodeItem.CodeDescription = .Cells(i, "G").value
                CodeItem.category = .Cells(i, "B").value
                CodeItem.Examples = .Cells(i, "H").value
                CodeNameString = .Cells(i, "D").value
                CodeCollection.Add CodeItem
              
                codeString = "<" & CodeNameString & ";"
                
                ' Add format string for specific codes
                ' Special formatting for CONST-E
                If CodeItem.codeName = "CONST-E" Then
                    codeString = codeString & " bold=""false""; format=""DollarItalic""; topborder=""false""; negative=""false"";"
                ElseIf CodeItem.codeName = "FORMULA-S" Then
                    codeString = codeString & " customformula="""";"
                Else
                    ' Check other codes for standard formatting
                    Dim formatCodes As Variant
                    formatCodes = Array("SPREAD-E", "ENDPOINT-E", "GROWTH-E", "FINANCIALS-S", _
                        "MULT3-S", "DIVIDE2-S", "SUBTRACT2-S", "SUBTOTAL2-S", "SUBTOTAL3-S", _
                        "AVGMULT3-S", "ANNUALIZE-S", "DEANNUALIZE-S", "AVGDEANNUALIZE2-S", _
                        "DIRECT-S", "CHANGE-S", "INCREASE-S", "DECREASE-S", "GROWTH-S", _
                        "OFFSETCOLUMN-S", "OFFSET2-S", "SUM2-S", "DISCOUNT2-S")
                    
                    Dim codeIndex As Long
                    For codeIndex = LBound(formatCodes) To UBound(formatCodes)
                        If CodeItem.codeName = formatCodes(codeIndex) Then
                            codeString = codeString & " format=""Dollar""; topborder=""false""; bold=""false""; indent=""1""; negative=""false"";"
                            Exit For
                        End If
                    Next codeIndex
                End If
                
                'Define Codetype
                If InStr(1, CodeItem.codeName, "-") Then
                    codeparts = Split(CodeItem.codeName, "-")
                    CodeItem.codeType = codeparts(1)
                End If

                
                'Define code location
                CodeItem.RowStart = .Columns("D").Find(CodeItem.codeName, LookAt:=xlWhole, SearchDirection:=xlNext).row
                CodeItem.RowEnd = .Columns("D").Find(CodeItem.codeName, LookAt:=xlWhole, SearchDirection:=xlPrevious).row
                
               ' Define non-green rows
                With .Cells(CodeItem.RowStart, "D")
                    Dim currentRow As Range
                    Set currentRow = .Cells
                    Do While currentRow.Interior.Color = RGB(204, 255, 204)
                        Set currentRow = currentRow.Offset(1, 0) ' Move to the next row
                    Loop
                    CodeItem.NonGreenStart = currentRow.row ' Set the first non-green row
                End With
                
                Dim labelRange As Range
                
                count = 1
                Dim rowarray As String
                
                rowString = ""
                columnLabels = ""
                
                If CodeItem.NonGreenStart <= CodeItem.RowEnd Then
                    For L = CodeItem.NonGreenStart To CodeItem.RowEnd
                        
                        columnLabels = columnLabels & GetCellValueWithFormatSymbol(.Cells(L, "B")) & "|"
                        
                        
                        rowarray = GetCellValueWithFormatSymbol(.Cells(L, "A"))
                        rowarray = rowarray & "|" & GetCellValueWithFormatSymbol(.Cells(L, "B"))
                        rowarray = rowarray & "|" & GetCellValueWithFormatSymbol(.Cells(L, "C"))
                        
                        ' Columns D, E, F (ignore all-caps values)
                        rowarray = rowarray & "|" & GetCellValueIgnoreAllCaps(.Cells(L, "D"))
                        rowarray = rowarray & "|" & GetCellValueIgnoreAllCaps(.Cells(L, "E"))
                        rowarray = rowarray & "|" & GetCellValueIgnoreAllCaps(.Cells(L, "F"))
                        
                        ' Column G
                        On Error Resume Next
                        If .Cells(L, "G").formula <> "" Then
                            If Left(.Cells(L, "G").formula, 1) = "=" Then
                                rowarray = rowarray & "|" & GetCellFormatSymbol(.Cells(L, "G")) & "F"
                            Else
                                rowarray = rowarray & "|" & GetCellValueWithFormatSymbol(.Cells(L, "G"))
                            End If
                        Else
                            rowarray = rowarray & "|"
                        End If
                        On Error GoTo 0
                        
                        ' Column H
                        On Error Resume Next
                        If .Cells(L, "H").formula <> "" Then
                            If Left(.Cells(L, "H").formula, 1) = "=" Then
                                rowarray = rowarray & "|" & GetCellFormatSymbol(.Cells(L, "H")) & "F"
                            Else
                                rowarray = rowarray & "|" & GetCellValueWithFormatSymbol(.Cells(L, "H"))
                            End If
                        Else
                            rowarray = rowarray & "|"
                        End If
                        On Error GoTo 0
                        
                        ' Column I
                        On Error Resume Next
                        If .Cells(L, "I").formula <> "" Then
                            If Left(.Cells(L, "I").formula, 1) = "=" Then
                                rowarray = rowarray & "|" & GetCellFormatSymbol(.Cells(L, "I")) & "F"
                            Else
                                rowarray = rowarray & "|" & GetCellValueWithFormatSymbol(.Cells(L, "I"))
                            End If
                        Else
                            rowarray = rowarray & "|"
                        End If
                        On Error GoTo 0
                        
                        ' Column K
                        On Error Resume Next
                        If .Cells(L, "K").formula <> "" Then
                            If Left(.Cells(L, "K").formula, 1) = "=" Then
                                rowarray = rowarray & "|" & GetCellFormatSymbol(.Cells(L, "K")) & "F"
                            Else
                                rowarray = rowarray & "|" & GetCellValueWithFormatSymbol(.Cells(L, "K"))
                            End If
                        Else
                            rowarray = rowarray & "|"
                        End If
                        On Error GoTo 0
                        
                        ' Column L
                        On Error Resume Next
                        If .Cells(L, "L").formula <> "" Then
                            If Left(.Cells(L, "L").formula, 1) = "=" Then
                                rowarray = rowarray & "|" & GetCellFormatSymbol(.Cells(L, "L")) & "F"
                            Else
                                rowarray = rowarray & "|" & GetCellValueWithFormatSymbol(.Cells(L, "L"))
                            End If
                        Else
                            rowarray = rowarray & "|"
                        End If
                        On Error GoTo 0
                        
                        ' Column M
                        On Error Resume Next
                        If .Cells(L, "M").formula <> "" Then
                            If Left(.Cells(L, "M").formula, 1) = "=" Then
                                rowarray = rowarray & "|" & GetCellFormatSymbol(.Cells(L, "M")) & "F"
                            Else
                                rowarray = rowarray & "|" & GetCellValueWithFormatSymbol(.Cells(L, "M"))
                            End If
                        Else
                            rowarray = rowarray & "|"
                        End If
                        On Error GoTo 0
                        
                        ' Column N
                        On Error Resume Next
                        If .Cells(L, "N").formula <> "" Then
                            If Left(.Cells(L, "N").formula, 1) = "=" Then
                                rowarray = rowarray & "|" & GetCellFormatSymbol(.Cells(L, "N")) & "F"
                            Else
                                rowarray = rowarray & "|" & GetCellValueWithFormatSymbol(.Cells(L, "N"))
                            End If
                        Else
                            rowarray = rowarray & "|"
                        End If
                        On Error GoTo 0
                        
                        ' Column O
                        On Error Resume Next
                        If .Cells(L, "O").formula <> "" Then
                            If Left(.Cells(L, "O").formula, 1) = "=" Then
                                rowarray = rowarray & "|" & GetCellFormatSymbol(.Cells(L, "O")) & "F"
                            Else
                                rowarray = rowarray & "|" & GetCellValueWithFormatSymbol(.Cells(L, "O"))
                            End If
                        Else
                            rowarray = rowarray & "|"
                        End If
                        On Error GoTo 0
                        
                        ' Column P
                        On Error Resume Next
                        If .Cells(L, "P").formula <> "" Then
                            If Left(.Cells(L, "P").formula, 1) = "=" Then
                                rowarray = rowarray & "|" & GetCellFormatSymbol(.Cells(L, "P")) & "F"
                            Else
                                rowarray = rowarray & "|" & GetCellValueWithFormatSymbol(.Cells(L, "P"))
                            End If
                        Else
                            rowarray = rowarray & "|"
                        End If
                        On Error GoTo 0
                        
                        ' Column R
                        On Error Resume Next
                        If .Cells(L, "R").formula <> "" Then
                            If Left(.Cells(L, "R").formula, 1) = "=" Then
                                rowarray = rowarray & "|" & GetCellFormatSymbol(.Cells(L, "R")) & "F"
                            Else
                                rowarray = rowarray & "|" & GetCellValueWithFormatSymbol(.Cells(L, "R"))
                            End If
                        Else
                            rowarray = rowarray & "|"
                        End If
                        On Error GoTo 0
                        
                        rowString = rowString & " " & "row" & count & " = """ & rowarray & """" & ";"
'                        CallByName CodeItem, "row" & Format(count, "000"), VbLet, rowarray
                        
'
                        count = count + 1
                    Next L
                End If
                columnLabels = columnLabels & """; "
                CallByName CodeItem, "columnLabels", VbLet, columnLabels
                
                
''
                    Set tempCollection = New Collection
                    
                    For j = CodeItem.RowStart To CodeItem.RowEnd
                    
                        With .Cells(j, "AE")
                            If .Interior.Color = RGB(204, 255, 204) And .formula <> "" Then
                                    tempCollection.Add .row
                            End If
                        End With
                    Next j
                    
                        CodeItem.InputDriverCount = tempCollection.count
                
'                codeString = codeString & " labelRow="""";"
                
                For x = 1 To CodeItem.InputDriverCount
                    codeString = codeString & " driver" & x & "="""";"
                Next x
                
                
                'FinDriver
                If .Cells(i, "I").value <> 0 Then
                    codeString = codeString & " financialsdriver=""" & .Cells(i, "I").value & """;"
                End If
                    
                codeString = codeString & rowString
                
'                'outputdrivers
                If outputdrivercount > 0 Then
                    codeString = codeString & " outputdrivers = """ & Join(CodeItem.Outputdrivers, "|") & """;"


                End If
                
                
                codeString = codeString & ">"
                CodeItem.codeString = codeString
                
            
            End If
            

MODEL_TAB_Codes:
        
        Next i
    End With

    ' Add a new CodeItem for "TAB"
    Set CodeItem = New CodeDBClass
    CodeItem.codeName = "TAB"
    CodeItem.codeString = "<TAB; label1="""";>"
    CodeCollection.Add CodeItem

'    Call Print_CodeDB
End Sub