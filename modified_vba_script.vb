Sub Populate_Code_DB_V3()
    'Add rowLabels and columnLabels
    
    Dim CodeItem As CodeDBClass
    Dim temparray() As Variant
    Dim tempCollection As Collection
    Dim linkrange As Range
    Dim nm As Name
    Dim lastRow As Long
    Dim nameExists As Boolean
    Dim codestring As String
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
              
                codestring = "<" & CodeNameString & ";"
                
                ' Add format string for specific codes
                ' Special formatting for CONST-E
                If CodeItem.codeName = "CONST-E" Then
                    codestring = codestring & " bold=""false""; format=""DollarItalic""; topborder=""false""; negative=""false"";"
                ElseIf CodeItem.codeName = "FORMULA-S" Then
                    codestring = codestring & " customformula="""";"
                ElseIf CodeItem.codeName = "COLUMNFORMULA-S" Then
                    codestring = codestring & " customformula="""";"
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
                            codestring = codestring & " format=""Dollar""; topborder=""false""; bold=""false""; indent=""1""; negative=""false"";"
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
                CodeItem.RowStart = .columns("D").Find(CodeItem.codeName, LookAt:=xlWhole, SearchDirection:=xlNext).row
                CodeItem.RowEnd = .columns("D").Find(CodeItem.codeName, LookAt:=xlWhole, SearchDirection:=xlPrevious).row
                
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
                        
                        ' Build row array with column labels
                        ' Column A = (D) - Output Driver
                        rowarray = GetCellValueWithFormatSymbol(.Cells(L, "A")) & "(D)"
                        
                        ' Column B = (L) - Label
                        rowarray = rowarray & "|" & GetCellValueWithFormatSymbol(.Cells(L, "B")) & "(L)"
                        
                        ' Column C = (F) - FinCode
                        rowarray = rowarray & "|" & GetCellValueWithFormatSymbol(.Cells(L, "C")) & "(F)"
                        
                        ' Column D = (C1) - Fixed Assumption 1
                        rowarray = rowarray & "|" & GetCellValueIgnoreAllCaps(.Cells(L, "D")) & "(C1)"
                        
                        ' Column E = (C2) - Fixed Assumption 2
                        rowarray = rowarray & "|" & GetCellValueIgnoreAllCaps(.Cells(L, "E")) & "(C2)"
                        
                        ' Column F = (C3) - Fixed Assumption 3
                        rowarray = rowarray & "|" & GetCellValueIgnoreAllCaps(.Cells(L, "F")) & "(C3)"
                        
                        ' Column G = (C4) - Fixed Assumption 4
                        On Error Resume Next
                        If .Cells(L, "G").formula <> "" Then
                            If Left(.Cells(L, "G").formula, 1) = "=" Then
                                rowarray = rowarray & "|" & GetCellFormatSymbol(.Cells(L, "G")) & "F(C4)"
                            Else
                                rowarray = rowarray & "|" & GetCellValueWithFormatSymbol(.Cells(L, "G")) & "(C4)"
                            End If
                        Else
                            rowarray = rowarray & "|(C4)"
                        End If
                        On Error GoTo 0
                        
                        ' Column H = (C5) - Fixed Assumption 5
                        On Error Resume Next
                        If .Cells(L, "H").formula <> "" Then
                            If Left(.Cells(L, "H").formula, 1) = "=" Then
                                rowarray = rowarray & "|" & GetCellFormatSymbol(.Cells(L, "H")) & "F(C5)"
                            Else
                                rowarray = rowarray & "|" & GetCellValueWithFormatSymbol(.Cells(L, "H")) & "(C5)"
                            End If
                        Else
                            rowarray = rowarray & "|(C5)"
                        End If
                        On Error GoTo 0
                        
                        ' Column I = (C6) - Fixed Assumption 6
                        On Error Resume Next
                        If .Cells(L, "I").formula <> "" Then
                            If Left(.Cells(L, "I").formula, 1) = "=" Then
                                rowarray = rowarray & "|" & GetCellFormatSymbol(.Cells(L, "I")) & "F(C6)"
                            Else
                                rowarray = rowarray & "|" & GetCellValueWithFormatSymbol(.Cells(L, "I")) & "(C6)"
                            End If
                        Else
                            rowarray = rowarray & "|(C6)"
                        End If
                        On Error GoTo 0
                        
                        ' Column K = (Y1) - Year 1
                        On Error Resume Next
                        If .Cells(L, "K").formula <> "" Then
                            If Left(.Cells(L, "K").formula, 1) = "=" Then
                                rowarray = rowarray & "|" & GetCellFormatSymbol(.Cells(L, "K")) & "F(Y1)"
                            Else
                                rowarray = rowarray & "|" & GetCellValueWithFormatSymbol(.Cells(L, "K")) & "(Y1)"
                            End If
                        Else
                            rowarray = rowarray & "|(Y1)"
                        End If
                        On Error GoTo 0
                        
                        ' Column L = (Y2) - Year 2
                        On Error Resume Next
                        If .Cells(L, "L").formula <> "" Then
                            If Left(.Cells(L, "L").formula, 1) = "=" Then
                                rowarray = rowarray & "|" & GetCellFormatSymbol(.Cells(L, "L")) & "F(Y2)"
                            Else
                                rowarray = rowarray & "|" & GetCellValueWithFormatSymbol(.Cells(L, "L")) & "(Y2)"
                            End If
                        Else
                            rowarray = rowarray & "|(Y2)"
                        End If
                        On Error GoTo 0
                        
                        ' Column M = (Y3) - Year 3
                        On Error Resume Next
                        If .Cells(L, "M").formula <> "" Then
                            If Left(.Cells(L, "M").formula, 1) = "=" Then
                                rowarray = rowarray & "|" & GetCellFormatSymbol(.Cells(L, "M")) & "F(Y3)"
                            Else
                                rowarray = rowarray & "|" & GetCellValueWithFormatSymbol(.Cells(L, "M")) & "(Y3)"
                            End If
                        Else
                            rowarray = rowarray & "|(Y3)"
                        End If
                        On Error GoTo 0
                        
                        ' Column N = (Y4) - Year 4
                        On Error Resume Next
                        If .Cells(L, "N").formula <> "" Then
                            If Left(.Cells(L, "N").formula, 1) = "=" Then
                                rowarray = rowarray & "|" & GetCellFormatSymbol(.Cells(L, "N")) & "F(Y4)"
                            Else
                                rowarray = rowarray & "|" & GetCellValueWithFormatSymbol(.Cells(L, "N")) & "(Y4)"
                            End If
                        Else
                            rowarray = rowarray & "|(Y4)"
                        End If
                        On Error GoTo 0
                        
                        ' Column O = (Y5) - Year 5
                        On Error Resume Next
                        If .Cells(L, "O").formula <> "" Then
                            If Left(.Cells(L, "O").formula, 1) = "=" Then
                                rowarray = rowarray & "|" & GetCellFormatSymbol(.Cells(L, "O")) & "F(Y5)"
                            Else
                                rowarray = rowarray & "|" & GetCellValueWithFormatSymbol(.Cells(L, "O")) & "(Y5)"
                            End If
                        Else
                            rowarray = rowarray & "|(Y5)"
                        End If
                        On Error GoTo 0
                        
                        ' Column P = (Y6) - Year 6
                        On Error Resume Next
                        If .Cells(L, "P").formula <> "" Then
                            If Left(.Cells(L, "P").formula, 1) = "=" Then
                                rowarray = rowarray & "|" & GetCellFormatSymbol(.Cells(L, "P")) & "F(Y6)"
                            Else
                                rowarray = rowarray & "|" & GetCellValueWithFormatSymbol(.Cells(L, "P")) & "(Y6)"
                            End If
                        Else
                            rowarray = rowarray & "|(Y6)"
                        End If
                        On Error GoTo 0
                        
                        ' Column R - Final column (should end with "|" and remain empty)
                        rowarray = rowarray & "|"
                        
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
                    codestring = codestring & " driver" & x & "="""";"
                Next x
                
                
                'FinDriver
                If .Cells(i, "I").value <> 0 Then
                    codestring = codestring & " financialsdriver=""" & .Cells(i, "I").value & """;"
                End If
                    
                codestring = codestring & rowString
                
'                'outputdrivers
                If outputdrivercount > 0 Then
                    codestring = codestring & " outputdrivers = """ & Join(CodeItem.Outputdrivers, "|") & """;"


                End If
                
                
                codestring = codestring & ">"
                CodeItem.codestring = codestring
                
            
            End If
            

MODEL_TAB_Codes:
        
        Next i
    End With

    ' Add a new CodeItem for "TAB"
    Set CodeItem = New CodeDBClass
    CodeItem.codeName = "TAB"
    CodeItem.codestring = "<TAB; label1="""";>"
    CodeCollection.Add CodeItem

'    Call Print_CodeDB
End Sub 