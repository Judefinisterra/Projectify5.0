

Public CodeCollection As Collection

Sub updated_all_DBs()
    Call Toggle_Settings(False)
    Call Update_DB_And_Print
'    Call CreateDBFiles("Call 1 Training Data.txt", "Training Data", "A")
    Call UpdateColumnC
    Call CreateDBFiles("Call 2 Training Data.txt", "Training Data", "C")
    Call CreateDBFiles("Call 1 Context.txt", "Call 1 Context", "A")
    Call CreateDBFiles("Call 2 Context.txt", "Call 2 Context", "B")
'    Call CreateDBFiles("Code Descriptions.txt", "Code DB", "X")
    Call CreateCodestringDBtxtJSassetfile
    Call ExportCodes
    Call CopyAndSaveSheets
    
    Dim PCUploadPath As String
    Dim OAIUploadPath As String
    OAIUploadPath = "C:\Users\joeor\Dropbox\B - Freelance\C_Projectify\VanPC\Codes\OAIupdate11.js"
    
    PCUploadPath = "C:\Users\joeor\Dropbox\B - Freelance\C_Projectify\VanPC\Codes\PCUpload39.js"
    'Path to second PC database
'    PCUploadPath = "C:\Users\joeor\Dropbox\B - Freelance\C_Projectify\VanPC\Codes\PCUpload38_old app_6.1.25.js"
    
    
    Call UpdatePinecone(PCUploadPath)
'    Call UpdatePinecone(OAIUploadPath)

    Call Toggle_Settings(True)
End Sub

Sub CreateCodestringDBtxtJSassetfile()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim filePath As String
    Dim fileNum As Integer
    Dim textLine As String
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Worksheets("code DB")
    
    ' Get the last row with data in column A
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    
    ' Set the file path
    filePath = "C:\Users\joeor\Desktop\Projectify6.0\Projectify4.0\assets\codestringDB.txt"
    
    ' Get the next available file number
    fileNum = FreeFile
    
    ' Open the file for output
    Open filePath For Output As #fileNum
    
    ' Loop through each row and write values to the text file
    For i = 1 To lastRow
        ' Combine values from column A and K
        textLine = ws.Cells(i, "A").value & vbTab & ws.Cells(i, "K").value
        
        ' Write the line to the text file
        Print #fileNum, textLine
    Next i
    
    ' Close the file
    Close #fileNum
    
      Debug.Print "Text file created successfully at: " & filePath
End Sub




Sub UpdateColumnC()
    Dim lastRow As Long
    Dim i As Long
    Dim dataArray() As String    'Array to hold the concatenated strings
    
    With Worksheets("Training Data")
        ' Find the last used row in column A
        lastRow = .Cells(Rows.count, "A").End(xlUp).row
        
        'Resize array to match the number of rows we need
        ReDim dataArray(1 To lastRow - 1)
        
        'Populate array with concatenated strings
        For i = 2 To lastRow
            dataArray(i - 1) = "Client input: " & .Cells(i, "A").value & ". Output: " & .Cells(i, "B").value
        Next i
        
        'Write the entire array to column C in one operation
        .Range("C2:C" & lastRow).value = Application.Transpose(dataArray)
        
        'Set column C to not wrap text
        .Range("C2:C" & lastRow).WrapText = False
    End With
End Sub

Sub CreateDBFiles(fileName As String, worksheetName As String, columnLetter As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim fileNum As Integer
    Dim i As Long
    Dim cellValue As String
    Dim directoryPath As String
    Dim fullPath As String
    
    ' Define the directory path
    directoryPath = "C:\Users\joeor\Dropbox\B - Freelance\C_Projectify\VanPC\Training Data\Main Script Training and Context Data\"
    
    ' Combine directory path and file name to get the full path
    fullPath = directoryPath & fileName
    
    ' Check if worksheet exists
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(worksheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Worksheet '" & worksheetName & "' does not exist.", vbExclamation, "Error"
        Exit Sub
    End If
    
    ' Find the last row with data in the specified column
    lastRow = ws.Cells(ws.Rows.count, columnLetter).End(xlUp).row
    
    ' Check if there's data to export
    If lastRow < 2 Then
        MsgBox "No data found in column " & columnLetter & " starting from row 2.", vbInformation, "No Data"
        Exit Sub
    End If
    
    ' Get a free file number
    fileNum = FreeFile
    
    ' Create and open the text file
    On Error Resume Next
    Open fullPath For Output As #fileNum
    
    If Err.Number <> 0 Then
        MsgBox "Error creating file: " & Err.Description, vbCritical, "File Error"
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Handle Training Data worksheet differently to separate columns A and B
    If worksheetName = "Training Data" Then
        ' For Training Data, create structured format with separate input and output
        For i = 2 To lastRow
            Dim inputValue As String
            Dim outputValue As String
            inputValue = ws.Cells(i, "A").value
            outputValue = ws.Cells(i, "B").value
            
            If Not IsEmpty(inputValue) And Not IsEmpty(outputValue) Then
                ' Create JSON-like structure for easier parsing
                Print #fileNum, "INPUT:" & inputValue & vbCrLf & "OUTPUT:" & outputValue & vbCrLf & vbCrLf & "***" & vbCrLf & vbCrLf
            End If
        Next i
    Else
        ' For other worksheets, use original logic
        For i = 2 To lastRow
            cellValue = ws.Cells(i, columnLetter).value
            If Not IsEmpty(cellValue) Then
                Print #fileNum, cellValue & vbCrLf & vbCrLf & "***" & vbCrLf & vbCrLf
            End If
        Next i
    End If
    
    ' Close the file
    Close #fileNum
    
       Debug.Print fileName & "Export complete"
End Sub


Function UpdatePinecone(scriptPath As String)
    Dim objShell As Object
    Dim strCommand As String
    Dim strOutput As String
    
    ' Create Shell object
    Set objShell = CreateObject("WScript.Shell")
    
    ' Command to run Node.js script
    strCommand = "node """ & scriptPath & """"
    
    ' Debug print start
    Debug.Print "Starting JavaScript execution at " & Now()
    Debug.Print "Using script: " & scriptPath
    
    ' Execute command and capture output
    On Error Resume Next
    strOutput = objShell.Run(strCommand, 1, True)
    
    ' Check for errors
    If Err.Number <> 0 Then
        Debug.Print "Error executing JavaScript: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    
    ' Debug print completion
    Debug.Print "JavaScript execution completed at " & Now()
    Debug.Print "Return value: " & strOutput
    
    ' Clean up
    Set objShell = Nothing
End Function


Sub Update_DB_And_Print()
    Toggle_Settings (False)
    Populate_Code_DB_V3
    
    Call Print_CodeDB
    Call ExportCodes
    
    
    Toggle_Settings (True)
    
'    Call ExportCodestringsCSV
'    Call ExportCodeDescriptionsCSV
'    Call UpdatePinecone
    
End Sub

Function Toggle_Settings(settings As Boolean)
    Dim RunTime As Variant
    
    With Application
        .DisplayAlerts = settings
        .Iteration = settings
        .ScreenUpdating = settings
'        .CutCopyMode = False
        
        If settings Then
            .Calculation = xlAutomatic
        Else
            .Calculation = xlManual
        End If
        
        .EnableEvents = settings
        
        .CalculateBeforeSave = settings
    End With
End Function




Sub Print_CodeDB()
    
    Dim CodeItem As CodeDBClass
    Dim outputArray() As Variant
    Dim rowCount As Long
    Dim i As Long
    Dim j As Long

    ' Set target worksheet
   With CodeDB
   
        .Range("A1:N1").value = Array("CodeName", "CodeType", "Category", "RowEnd", _
                                    "Assumptions", "InputDriverCount", "OutputDrivers", _
                                    "FinCodes", "Code Description", "Call 1 Description", "CodeString", "Examples", "Call 2 Description", "Call 3 Description")
        
        ' Determine the number of rows in the collection
        rowCount = CodeCollection.count
        ReDim outputArray(1 To rowCount, 1 To 12)
    
        ' Populate the output array
        For i = 1 To rowCount
            Set CodeItem = CodeCollection(i)
            outputArray(i, 1) = CodeItem.codeName
            outputArray(i, 2) = CodeItem.codeType
            outputArray(i, 3) = CodeItem.category
            outputArray(i, 4) = CodeItem.RowEnd
    
            ' Flatten arrays
            If Not IsEmpty(CodeItem.Assumptions) Then
                outputArray(i, 5) = Join(CodeItem.Assumptions, ",") ' Join Assumptions array
            Else
                outputArray(i, 5) = ""
            End If
    
            outputArray(i, 6) = CodeItem.InputDriverCount
    
            If Not IsEmpty(CodeItem.Outputdrivers) Then
                outputArray(i, 7) = Join(CodeItem.Outputdrivers, ",") ' Join OutputDrivers array
            Else
                outputArray(i, 7) = ""
            End If
    
            If Not IsEmpty(CodeItem.FinCodes) Then
                outputArray(i, 8) = Join(CodeItem.FinCodes, ",") ' Join FinCodes array
            Else
                outputArray(i, 8) = ""
            End If
    
            outputArray(i, 9) = CodeItem.CodeDescription
            outputArray(i, 10) = 0
            outputArray(i, 11) = CodeItem.codestring  ' Still populate array but we'll handle this column separately
            outputArray(i, 12) = CodeItem.Examples
        Next i
    
        ' Write the array to the worksheet
        ' Debug to see what we're actually assigning
'        Debug.Print "Array dimensions: " & LBound(outputArray, 1) & " to " & UBound(outputArray, 1) & " x " & _
'                    LBound(outputArray, 2) & " to " & UBound(outputArray, 2)
                    
        ' Write all columns except column 11 using array approach
        Dim colArray() As Variant
        
        For j = 1 To 12
            ' Skip column 11 - we'll handle it separately
            If j <> 11 Then
                ReDim colArray(1 To rowCount, 1 To 1)
                For i = 1 To rowCount
                    colArray(i, 1) = outputArray(i, j)
                Next i
                .Range(.Cells(2, j), .Cells(rowCount + 1, j)).value = colArray
            End If
        Next j
        
        ' Handle column 11 (CodeString) cell by cell
        For i = 1 To rowCount
            Set CodeItem = CodeCollection(i)
            ' Write directly to the cell
            .Cells(i + 1, 11).value = CodeItem.codestring
        Next i
        
        ' Rest of the code remains the same
        lastRow = .Range("A1").End(xlDown).row
        lastcolumn = .Range("A1").End(xlToRight).Column
        
        'Call 1 Description
    .Range("J2").formula = "=CONCATENATE(""Code: "",IF(A2<>"""",A2,""None""),""; Category: "",IF(C2<>"""",C2,""None""),""; Description: "",IF(I2<>"""",I2,""None""),""; Codestring: "",IF(K2<>"""",K2,""None""))"
       .Range("J2").Copy
        .Range("J2:J" & lastRow).PasteSpecial Paste:=xlPasteFormulas
        
        Application.Calculate
        
        With .Range("J:J")
            .value = .value
        End With
        
        'Call 2 and 3 Description
        .Range("M2").formula = "=CONCATENATE(""Description: "",IF(I2<>"""",I2,""None""),"" & Codestring: "",IF(K2<>"""",K2,""None""))"
        .Range("M2").Copy
        .Range("M2:M" & lastRow).PasteSpecial Paste:=xlPasteFormulas
        
        Application.Calculate
        
        With .Range("M:M")
            .value = .value
        End With
'

    End With
    
'        Worksheets("Code DB").Sort.SortFields.Clear
'        Worksheets("Code DB").Sort.SortFields.Add2 key:=Range("A1:A" & Cells(Rows.count, 1).End(xlUp).row), _
'        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        
     
    With Worksheets("Code DB").Sort
        
        .SortFields.Clear
        .SortFields.Add key:=Worksheets("Code DB").Range("A1:A" & lastRow), Order:=xlAscending
        .SetRange Worksheets("Code DB").Range("A1").Resize(lastRow, lastcolumn)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With


End Sub


Function GetCellFormatSymbol(cell As Range) As String
    'Returns the appropriate symbol prefix based on cell formatting
    Dim formatSymbol As String
    Dim numberFormat As String
    Dim isItalic As Boolean
    
    formatSymbol = ""
    numberFormat = cell.numberFormat
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
    
    cellValue = CStr(cell.value)
    
    ' Check if cell is empty or contains only spaces
    If Trim(cellValue) = "" Then
        GetCellValueWithFormatSymbol = ""
        Exit Function
    End If
    
    formatSymbol = GetCellFormatSymbol(cell)
    numberFormat = cell.numberFormat
    
    ' Check if it's factor formatting and add "x" suffix to numeric values
    If (InStr(numberFormat, "#,##0.0x") > 0 Or InStr(numberFormat, "x") > 0) And IsNumeric(cell.value) Then
        cellValue = cellValue & "x"
    End If
    
    GetCellValueWithFormatSymbol = formatSymbol & cellValue
End Function

Function GetCellValueIgnoreAllCaps(cell As Range) As String
    'Returns the cell value with formatting symbols, but ignores all-caps values
    Dim cellValue As String
    
    cellValue = CStr(cell.value)
    
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


Sub ExportCodestringsCSV()
    Dim ws As Worksheet
    Dim cell As Range
    Dim csvFile As String
    Dim fileNum As Integer
    Dim rowData As String
    Dim lastRow As Long
    
'    Call Update_DB_And_Print
    
    ' Set your worksheet (modify "Sheet1" if needed)
    Set ws = ThisWorkbook.Sheets("Code DB")
    
    ' Find last row with data in Column A
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    ' Define CSV file path
    csvFile = "C:\Users\joeor\Dropbox\B - Freelance\C_Projectify\VanPC\Training Data\Main Script Training and Context Data\Codestrings.csv"
    
    ' Open file for writing
    fileNum = FreeFile
    Open csvFile For Output As #fileNum
    
    ' Loop through each row and write "Column A, Column K" to CSV
    For Each cell In ws.Range("A1:A" & lastRow)
        rowData = cell.value & "," & cell.Offset(0, 10).value ' Offset(0,10) moves from A to K
        Print #fileNum, rowData
    Next cell
    
    ' Close file
    Close #fileNum
    
    Debug.Print "Codestrings data updated to text file successfully"
End Sub

Sub ExportCodeDescriptionsCSV()
    Dim ws As Worksheet
    Dim cell As Range
    Dim csvFile As String
    Dim fileNum As Integer
    Dim rowData As String
    Dim lastRow As Long
    
'    Call Update_DB_And_Print
    
    ' Set your worksheet (modify "Sheet1" if needed)
    Set ws = ThisWorkbook.Sheets("Code DB")
    
    ' Find last row with data in Column A
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    ' Define CSV file path
    csvFile = "C:\Users\joeor\Dropbox\B - Freelance\C_Projectify\VanPC\Training Data\Main Script Training and Context Data\Codes Descriptions.TXT"
    
    ' Open file for writing
    fileNum = FreeFile
    Open csvFile For Output As #fileNum
    
    ' Loop through each row and write "Column A, Column K" to CSV
    For Each cell In ws.Range("A1:A" & lastRow)
        rowData = cell.value & "," & cell.Offset(0, 9).value ' Offset(0,10) moves from A to K
        Print #fileNum, rowData
    Next cell
    
    ' Close file
    Close #fileNum
    
    Debug.Print "Code descriptions data updated to text file successfully"
End Sub

Sub ExportCall2Context()
    Dim lastRow As Long
    Dim i As Long
    Dim filePath As String
    Dim fileNum As Integer
    Dim cellValueA As String
    Dim cellValueB As String
    
    ' Set the file path
    filePath = "C:\Users\joeor\Dropbox\B - Freelance\C_Projectify\VanPC\Training Data\Main Script Training and Context Data\"
    
    ' Find the last row with data in column A
    lastRow = Worksheets("Call 2 Context").Cells(Worksheets("Call 2 Context").Rows.count, "A").End(xlUp).row
    
    ' Get free file number
    fileNum = FreeFile
    
    ' Open the file for writing
    Open filePath & "Call2Context.TXT" For Output As #fileNum
    
    ' Loop through each row starting from row 1
    For i = 1 To lastRow
        With Worksheets("Call 2 Context")
            ' Get values from columns A and B
            cellValueA = Trim(.Cells(i, "A").value)
'            cellValueB = Trim(.Cells(i, "B").value)
            
            ' Clean up the values
            cellValueA = Replace(cellValueA, ",", ";")
            cellValueA = Replace(cellValueA, vbCrLf, "br/")
            cellValueA = Replace(cellValueA, vbLf, " ")
            cellValueA = Replace(cellValueA, "   ", " ")
            
            cellValueB = Replace(cellValueB, ",", ";")
            cellValueB = Replace(cellValueB, vbCrLf, "br/")
            cellValueB = Replace(cellValueB, vbLf, " ")
            cellValueB = Replace(cellValueB, "   ", " ")
            
            ' Write to file if there's data
            If Len(cellValueA) > 0 Then
                Print #fileNum, cellValueA & "," & cellValueB
            End If
        End With
    Next i
    
    ' Close the file
    Close #fileNum
    
    Debug.Print "Call 2 Context data exported to text file successfully"
End Sub

Sub ExportCall1Context()
    Dim lastRow As Long
    Dim i As Long
    Dim filePath As String
    Dim fileNum As Integer
    Dim cellValueA As String
    Dim cellValueB As String
    
    ' Set the file path
    filePath = "C:\Users\joeor\Dropbox\B - Freelance\C_Projectify\VanPC\Training Data\Main Script Training and Context Data\"
    
    ' Find the last row with data in column A
    lastRow = Worksheets("Call 1 Context").Cells(Worksheets("Call 2 Context").Rows.count, "A").End(xlUp).row
    
    ' Get free file number
    fileNum = FreeFile
    
    ' Open the file for writing
    Open filePath & "Call1Context.TXT" For Output As #fileNum
    
    ' Loop through each row starting from row 1
    For i = 2 To lastRow
        With Worksheets("Call 1 Context")
            ' Get values from columns A and B
            cellValueA = Trim(.Cells(i, "A").value)
            ' cellValueB = Trim(.Cells(i, "B").value)
            
            ' Clean up the values
            cellValueA = Replace(cellValueA, ",", ";")
            cellValueA = Replace(cellValueA, vbCrLf, "br/")
            cellValueA = Replace(cellValueA, vbLf, " ")
            cellValueA = Replace(cellValueA, "   ", " ")
            
            cellValueB = Replace(cellValueB, ",", ";")
            cellValueB = Replace(cellValueB, vbCrLf, "br/")
            cellValueB = Replace(cellValueB, vbLf, " ")
            cellValueB = Replace(cellValueB, "   ", " ")
            
            ' Write to file if there's data
            If Len(cellValueA) > 0 Then
                Print #fileNum, cellValueA & "," & cellValueB
            End If
        End With
    Next i
    
    ' Close the file
    Close #fileNum
    
    Debug.Print "Call 1 Context data exported to text file successfully"
End Sub

 Sub ExportTrainingData()
    Dim lastRow As Long
    Dim i As Long
    Dim filePath As String
    Dim fileNum1 As Integer
    Dim fileNum2 As Integer
    Dim cellValue As String
    
    ' Set the file path
    filePath = "C:\Users\joeor\Dropbox\B - Freelance\C_Projectify\VanPC\Training Data\Main Script Training and Context Data\"
    
    ' Find the last row with data in column A
    lastRow = Worksheets("Training Data").Cells(Worksheets("Training Data").Rows.count, "B").End(xlUp).row
    
    ' Get free file numbers
    fileNum1 = FreeFile
    fileNum2 = FreeFile + 1
    
    ' Open the two files for writing
    Open filePath & "Call 1 Training Data.txt" For Output As #fileNum1
    Open filePath & "Call 2 Training Data.txt" For Output As #fileNum2
    
    ' Loop through each row starting from row 2
    For i = 2 To lastRow
        With Worksheets("Training Data")
            ' Write column A data to first file
            cellValue = Trim(.Cells(i, "A").value)
            If Len(cellValue) > 0 Then
                Print #fileNum1, cellValue
            End If
            
            ' Write column B data to second file
            cellValue = Trim(.Cells(i, "B").value)
            If Len(cellValue) > 0 Then
                Print #fileNum2, cellValue
            End If
        End With
    Next i
    
    ' Close all files
    Close #fileNum1
    Close #fileNum2
    
    Debug.Print "Training data exported to text files successfully"
End Sub

Sub ExportCodes()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim filePath As String
    Dim fileNum As Integer
    

    
    ' Set reference to the Code DB worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Code DB")
    On Error GoTo 0
    
    ' Check if the worksheet exists
    If ws Is Nothing Then
        MsgBox "The 'Code DB' worksheet does not exist.", vbExclamation
        Exit Sub
    End If
    
    ' Find the last populated row in column A
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    
    ' Exit if there's no data (only header row)
    If lastRow <= 1 Then
        MsgBox "No data found in column A of the Code DB worksheet.", vbInformation
        Exit Sub
    End If
    
    ' Create a file path for the CSV
    filePath = "C:\Users\joeor\Desktop\Projectify6.0\Projectify4.0\src\prompts\Codes.txt"
    
    ' Open the file for writing
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    
    ' Write each populated cell from column A (starting from row 2) to the CSV file
    For i = 2 To lastRow
        If Not IsEmpty(ws.Cells(i, 1).value) Then
            Print #fileNum, ws.Cells(i, 1).value
        End If
    Next i
    
    ' Close the file
    Close #fileNum
    
    ' Inform the user
    Debug.Print "CSV file created successfully at:" & vbCrLf & filePath
End Sub

Sub CopyAndSaveSheets()
    Dim wbSource As Workbook
    Dim wbNew As Workbook
    Dim ws As Worksheet
    Dim sheetNames As Variant
    Dim savePath As String
    Dim sheetExists As Boolean
    Dim i As Integer

    ' --- Configuration ---
    ' List of worksheet names to copy
    sheetNames = Array("Financials", "Calcs", "Misc.", "Actuals", "Codes")
    ' Path to save the new workbook
    savePath = "C:\Users\joeor\Desktop\Projectify6.0\Projectify4.0\assets\Worksheets_4.3.25 v1.xlsx"
    ' --- End Configuration ---

    Call DeleteRowsBasedOnColumnD
    
    On Error GoTo ErrorHandler

    ' Set the source workbook (the one containing this code)
    Set wbSource = ThisWorkbook

    ' Check if all specified sheets exist in the source workbook
    For i = LBound(sheetNames) To UBound(sheetNames)
        sheetExists = False
        For Each ws In wbSource.Worksheets
            If ws.Name = sheetNames(i) Then
                sheetExists = True
                Exit For
            End If
        Next ws
        If Not sheetExists Then
            MsgBox "Error: Worksheet '" & sheetNames(i) & "' not found in the source workbook.", vbCritical, "Sheet Not Found"
            GoTo CleanExit
        End If
    Next i

    ' Turn off screen updating and alerts to speed up and hide process
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Ensure "Calcs" sheet is visible before copying
    On Error Resume Next ' In case "Calcs" is already visible
    wbSource.Worksheets("Calcs").Visible = xlSheetVisible
    On Error GoTo ErrorHandler ' Re-enable default error handling

    ' Select the sheets in the specified order and copy them
    wbSource.Worksheets(sheetNames(LBound(sheetNames))).Select Replace:=True ' Select the first sheet, replacing any current selection
    For i = LBound(sheetNames) + 1 To UBound(sheetNames) ' Select remaining sheets, adding to the selection
        wbSource.Worksheets(sheetNames(i)).Select Replace:=False
    Next i
    ActiveWindow.SelectedSheets.Copy

    ' The new workbook created by the Copy method is now the active workbook
    Set wbNew = ActiveWorkbook

    ' Save the new workbook
    ' Use xlOpenXMLWorkbook for .xlsx format
    Debug.Print "Attempting to save multi-sheet workbook to: " & savePath
    wbNew.SaveAs fileName:=savePath, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Debug.Print "Successfully saved multi-sheet workbook."

    ' Close the new workbook without saving changes again
    wbNew.Close SaveChanges:=False

    ' --- Start: Add logic to copy and save "Codes" sheet separately ---
    
    Dim codesSavePath As String
    codesSavePath = "C:\Users\joeor\Desktop\Projectify6.0\Projectify4.0\assets\Codes.xlsx"
    
    ' Select and copy the "Codes" sheet
    wbSource.Worksheets("Codes").Copy
    
    ' The new workbook created by the Copy method is now the active workbook
    Set wbNew = ActiveWorkbook ' Reusing wbNew object variable
    
    ' Save the new workbook containing only the "Codes" sheet
    Debug.Print "Attempting to save Codes sheet workbook to: " & codesSavePath
    On Error Resume Next ' Temporarily disable global handler for specific check
    wbNew.SaveAs fileName:=codesSavePath, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    If Err.Number <> 0 Then
        Debug.Print "SaveAs Error Number: " & Err.Number
        Debug.Print "SaveAs Error Description: " & Err.Description
        MsgBox "Error saving " & codesSavePath & vbCrLf & _
               "Error Number: " & Err.Number & vbCrLf & _
               "Description: " & Err.Description, vbCritical, "SaveAs Error"
        Err.Clear ' Clear the specific error
        On Error GoTo ErrorHandler ' Optional: Re-enable global handler if needed, or just go to CleanExit
        Resume CleanExit ' Go to cleanup code
    End If
    On Error GoTo ErrorHandler ' Re-enable global handler if save was successful
    Debug.Print "Successfully saved Codes sheet workbook."
    
    ' Close the new workbook
    wbNew.Close SaveChanges:=False
    
    ' --- End: Add logic to copy and save "Codes" sheet separately ---

CleanExit:
    ' Restore screen updating and alerts
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    ' Clean up object variables
    Set wbNew = Nothing
    Set wbSource = Nothing
    Set ws = Nothing

    On Error GoTo 0 ' Turn error handling off
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & vbCrLf & Err.Description, vbCritical, "Error"
    Resume CleanExit ' Go to cleanup code on error

    

End Sub
Sub DeleteRowsBasedOnColumnD()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' --- Configuration ---
    Const TARGET_SHEET_NAME As String = "Financials"
    Const TARGET_COLUMN As Integer = 4 ' Column D
    ' --- End Configuration ---

    On Error GoTo ErrorHandler

    ' Set the workbook (the one containing this code)
    Set wb = ThisWorkbook

    ' Attempt to set the worksheet
    On Error Resume Next ' Temporarily disable error handling for sheet check
    Set ws = wb.Worksheets(TARGET_SHEET_NAME)
    On Error GoTo ErrorHandler ' Re-enable default error handling
    If ws Is Nothing Then
        MsgBox "Error: Worksheet '" & TARGET_SHEET_NAME & "' not found.", vbCritical, "Sheet Not Found"
        GoTo CleanExit
    End If

    ' Turn off screen updating to speed up the process
    Application.ScreenUpdating = False

    ' Find the last row with data in the target column
    lastRow = ws.Cells(ws.Rows.count, TARGET_COLUMN).End(xlUp).row

    ' Loop backwards from the last row to the first row (important when deleting rows)
    For i = lastRow To 1 Step -1
        ' Check if the cell in the target column is not empty
        If Not IsEmpty(ws.Cells(i, TARGET_COLUMN).value) And Trim(ws.Cells(i, TARGET_COLUMN).value) <> "" Then
            ' Delete the entire row
            ws.Rows(i).Delete
        End If
    Next i

CleanExit:
    ' Restore screen updating
    Application.ScreenUpdating = True

    ' Clean up object variables
    Set ws = Nothing
    Set wb = Nothing

    On Error GoTo 0 ' Turn error handling off
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & vbCrLf & Err.Description, vbCritical, "Error"
    Resume CleanExit ' Go to cleanup code on error

End Sub

Sub Update_Training_Data_Parameters()
    'Updates Training Data tab with formatting parameters for specific codes
    'Only adds parameters that don't already exist
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String
    Dim newCellValue As String
    Dim codeSegments() As String
    Dim j As Long
    Dim currentSegment As String
    Dim codeName As String
    Dim parametersToAdd As String
    Dim wasChanged As Boolean
    
    ' Define the codes to check
    Dim formatCodes As Variant
    formatCodes = Array("SPREAD-E", "ENDPOINT-E", "GROWTH-E", "FINANCIALS-S", _
        "MULT3-S", "DIVIDE2-S", "SUBTRACT2-S", "SUBTOTAL2-S", "SUBTOTAL3-S", _
        "AVGMULT3-S", "ANNUALIZE-S", "DEANNUALIZE-S", "AVGDEANNUALIZE2-S", _
        "DIRECT-S", "CHANGE-S", "INCREASE-S", "DECREASE-S", "GROWTH-S", _
        "OFFSETCOLUMN-S", "OFFSET2-S", "SUM2-S", "DISCOUNT2-S")
    
    ' Set worksheet
    Set ws = ThisWorkbook.Worksheets("Training Data")
    
    ' Find last row in column B
    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).row
    
    ' Iterate through column B
    For i = 1 To lastRow
        cellValue = ws.Cells(i, "B").value
        wasChanged = False
        
        ' Skip empty cells
        If cellValue <> "" Then
            newCellValue = ""
            
            ' Split by >< to get individual code segments
            cellValue = Replace(cellValue, "><", ">|<")
            codeSegments = Split(cellValue, "|")
            
            ' Process each code segment
            For j = 0 To UBound(codeSegments)
                currentSegment = codeSegments(j)
                
                ' Extract code name from current segment
                If InStr(currentSegment, "<") > 0 And InStr(currentSegment, ";") > 0 Then
                    codeName = Mid(currentSegment, InStr(currentSegment, "<") + 1, InStr(currentSegment, ";") - InStr(currentSegment, "<") - 1)
                    
                    parametersToAdd = ""
                    
                    ' Special case for CONST-E
                    If codeName = "CONST-E" Then
                        ' Check which parameters are missing in this segment
                        If InStr(currentSegment, "bold=") = 0 Then
                            parametersToAdd = parametersToAdd & " bold=""false"";"
                        End If
                        If InStr(currentSegment, "format=") = 0 Then
                            parametersToAdd = parametersToAdd & " format=""DollarItalic"";"
                        End If
                        If InStr(currentSegment, "topborder=") = 0 Then
                            parametersToAdd = parametersToAdd & " topborder=""false"";"
                        End If
                        If InStr(currentSegment, "negative=") = 0 Then
                            parametersToAdd = parametersToAdd & " negative=""false"";"
                        End If
                    Else
                        ' Check if it's one of the other codes
                        Dim codeIndex As Long
                        Dim isTargetCode As Boolean
                        isTargetCode = False
                        
                        For codeIndex = LBound(formatCodes) To UBound(formatCodes)
                            If codeName = formatCodes(codeIndex) Then
                                isTargetCode = True
                                Exit For
                            End If
                        Next codeIndex
                        
                        If isTargetCode Then
                            ' Check which parameters are missing
                            If InStr(currentSegment, "format=") = 0 Then
                                parametersToAdd = parametersToAdd & " format=""Dollar"";"
                            End If
                            If InStr(currentSegment, "topborder=") = 0 Then
                                parametersToAdd = parametersToAdd & " topborder=""false"";"
                            End If
                            If InStr(currentSegment, "bold=") = 0 Then
                                parametersToAdd = parametersToAdd & " bold=""false"";"
                            End If
                            If InStr(currentSegment, "indent=") = 0 Then
                                parametersToAdd = parametersToAdd & " indent=""1"";"
                            End If
                            If InStr(currentSegment, "negative=") = 0 Then
                                parametersToAdd = parametersToAdd & " negative=""false"";"
                            End If
                        End If
                    End If
                    
                    ' If there are parameters to add, insert them
                    If parametersToAdd <> "" Then
                        ' Find the position after the semicolon following the code name
                        Dim insertPos As Long
                        insertPos = InStr(currentSegment, codeName & ";") + Len(codeName & ";")
                        currentSegment = Left(currentSegment, insertPos - 1) & parametersToAdd & Mid(currentSegment, insertPos)
                        wasChanged = True
                    End If
                End If
                
                ' Add the processed segment to the new cell value
                If j > 0 Then
                    newCellValue = newCellValue & currentSegment
                Else
                    newCellValue = currentSegment
                End If
            Next j
            
            ' Update the cell if changes were made
            If wasChanged Then
                ws.Cells(i, "B").value = newCellValue
                ' Highlight the entire row in blue
                ws.Rows(i).Interior.Color = RGB(173, 216, 230) ' Light blue color
            End If
        End If
    Next i
    
    MsgBox "Training Data parameters update complete!", vbInformation
End Sub


Sub HighlightCellsWithLI()
    Dim ws As Worksheet
    Dim cell As Range
    Dim lastRow As Long
    
    ' Set reference to "training data" worksheet
    Set ws = ThisWorkbook.Worksheets("training data")
    
    ' Find last row in column B
    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).row
    
    ' Loop through each cell in column B
    For Each cell In ws.Range("B1:B" & lastRow)
        ' Check if cell contains "*LI" (case sensitive)
        If InStr(1, cell.value, "*LI", vbBinaryCompare) > 0 Then
            ' Highlight cell in red
            cell.Interior.Color = RGB(255, 0, 0)
        End If
    Next cell
End Sub


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
    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).row
    
    Debug.Print "Starting codestring update process..."
    Debug.Print "Processing rows 2 to " & lastRow & " in Training Data sheet"
    Debug.Print "Results will be written to column C"
    
    changesMade = 0
    
    ' Clear column C header and add new header
    ws.Cells(1, "C").value = "Updated Codestrings"
    
    ' Loop through each row from 2 to last populated row
    For i = 2 To lastRow
        cellValue = ws.Cells(i, "B").value
        
        ' Skip empty cells
        If Trim(cellValue) <> "" Then
            ' Process the cell value to update row parameters
            updatedValue = UpdateRowParametersInCodestring(cellValue)
            
            ' Always write the result to column C (whether changed or not)
            ws.Cells(i, "C").value = updatedValue
            
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
            ws.Cells(i, "C").value = ""
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
    regex.pattern = "row\d+\s*=\s*""([^""]*)""|row\d+\s*=\s*'([^']*)'"
    
    ' Find all row parameter matches
    Set matches = regex.Execute(codestring)
    
    ' Process matches in reverse order to avoid position shifts
    For i = matches.count - 1 To 0 Step -1
        Set match = matches(i)
        originalParam = match.value
        
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

Sub ConvertLegacyFormattingToInline()
    'Converts legacy format parameters to new inline symbol-based approach
    'Processes codestrings in Training Data sheet column B and outputs results to column C
    'Example: format="dollar" with row data "1000" becomes "$1000" in the row
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String
    Dim updatedValue As String
    Dim changesMade As Long
    
    ' Set reference to Training Data worksheet
    Set ws = Worksheets("Training Data")
    
    ' Find last populated row in column B
    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).row
    
    Debug.Print "Starting legacy formatting conversion process..."
    Debug.Print "Processing rows 2 to " & lastRow & " in Training Data sheet"
    Debug.Print "Results will be written to column C"
    
    changesMade = 0
    
    ' Clear column C header and add new header
    ws.Cells(1, "C").value = "Converted Formatting"
    
    ' Loop through each row from 2 to last populated row
    For i = 2 To lastRow
        cellValue = ws.Cells(i, "B").value
        
        ' Skip empty cells
        If Trim(cellValue) <> "" Then
            ' Process the cell value to convert legacy formatting
            updatedValue = ConvertLegacyFormatParameters(cellValue)
            
            ' Always write the result to column C (whether changed or not)
            ws.Cells(i, "C").value = updatedValue
            
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
            ws.Cells(i, "C").value = ""
            ws.Cells(i, "C").Interior.ColorIndex = xlNone
        End If
    Next i
    
    Debug.Print "Process complete. Converted " & changesMade & " codestrings."
    MsgBox "Legacy formatting conversion complete!" & vbCrLf & _
           "Processed " & (lastRow - 1) & " rows." & vbCrLf & _
           "Converted " & changesMade & " codestrings with legacy formatting." & vbCrLf & _
           "Results written to column C with green highlighting for changed rows.", vbInformation
End Sub

Function ConvertLegacyFormatParameters(codestring As String) As String
    'Converts legacy format parameters (format="", italic="") to inline symbols in row data
    
    Dim result As String
    Dim formatType As String
    Dim isItalic As Boolean
    Dim hasFormatParam As Boolean
    Dim hasItalicParam As Boolean
    
    result = codestring
    
    ' Extract format parameter
    formatType = ExtractParameterValue(result, "format")
    hasFormatParam = (formatType <> "")
    
    ' Extract italic parameter
    Dim italicValue As String
    italicValue = ExtractParameterValue(result, "italic")
    isItalic = (LCase(italicValue) = "true")
    hasItalicParam = (italicValue <> "")
    
    ' If no legacy formatting parameters found, return as is
    If Not hasFormatParam And Not hasItalicParam Then
        ConvertLegacyFormatParameters = result
        Exit Function
    End If
    
    Debug.Print "  Found legacy formatting - Format: '" & formatType & "', Italic: '" & italicValue & "'"
    
    ' Remove the legacy format and italic parameters from the codestring
    If hasFormatParam Then
        result = RemoveParameter(result, "format")
    End If
    If hasItalicParam Then
        result = RemoveParameter(result, "italic")
    End If
    
    ' Apply formatting symbols to row parameters
    result = ApplyFormattingToRowParameters(result, formatType, isItalic)
    
    ConvertLegacyFormatParameters = result
End Function

Function ExtractParameterValue(codestring As String, paramName As String) As String
    'Extracts the value of a parameter from the codestring
    
    Dim regex As Object
    Dim matches As Object
    Dim pattern As String
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    regex.IgnoreCase = True
    
    ' Pattern to match parameter="value" or parameter='value'
    pattern = paramName & "\s*=\s*[""']([^""']*)[""']"
    regex.pattern = pattern
    
    Set matches = regex.Execute(codestring)
    
    If matches.count > 0 Then
        ExtractParameterValue = matches(0).SubMatches(0)
    Else
        ExtractParameterValue = ""
    End If
End Function

Function RemoveParameter(codestring As String, paramName As String) As String
    'Removes a parameter from the codestring
    
    Dim regex As Object
    Dim pattern As String
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    
    ' Pattern to match parameter="value"; or parameter='value';
    pattern = "\s*" & paramName & "\s*=\s*[""'][^""']*[""']\s*;"
    regex.pattern = pattern
    
    RemoveParameter = regex.Replace(codestring, "")
End Function

Function ApplyFormattingToRowParameters(codestring As String, formatType As String, isItalic As Boolean) As String
    'Applies formatting symbols to values in row parameters based on format type and italic setting
    
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
    For i = matches.count - 1 To 0 Step -1
        Set match = matches(i)
        Dim originalRowContent As String
        Dim updatedRowContent As String
        Dim fullMatch As String
        
        originalRowContent = match.SubMatches(1) ' The content between quotes
        fullMatch = match.value ' The entire match including quotes
        
        ' Apply formatting symbols to the row content
        updatedRowContent = ApplySymbolsToRowContent(originalRowContent, formatType, isItalic)
        
        ' Replace in result string
        Dim updatedMatch As String
        updatedMatch = match.SubMatches(0) & updatedRowContent & match.SubMatches(2)
        result = Left(result, match.FirstIndex) & updatedMatch & Mid(result, match.FirstIndex + match.Length + 1)
    Next i
    
    ApplyFormattingToRowParameters = result
End Function

Function ApplySymbolsToRowContent(rowContent As String, formatType As String, isItalic As Boolean) As String
    'Applies appropriate symbols to values in a single row based on format type
    
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
        formattedValue = value
        
        ' Only apply formatting to non-empty values that aren't "F"
        If value <> "" And UCase(value) <> "F" Then
            ' Apply format-specific symbols
            Select Case LCase(formatType)
                Case "dollar"
                    If IsNumeric(value) Then
                        formattedValue = "$" & value
                    ElseIf value <> "" Then
                        formattedValue = "$" & value ' Apply to non-numeric values too
                    End If
                    
                Case "dollaritalic"
                    If IsNumeric(value) Then
                        formattedValue = "~$" & value
                    ElseIf value <> "" Then
                        formattedValue = "~$" & value
                    End If
                    
                Case "volume"
                    ' Volume formatting - apply italic if specified
                    If isItalic And value <> "" Then
                        formattedValue = "~" & value
                    End If
                    
                Case "percent"
                    ' Percent values - apply italic if specified
                    If isItalic And value <> "" Then
                        formattedValue = "~" & value
                    End If
                    
                Case "factor"
                    ' Factor formatting - add x suffix and italic
                    If IsNumeric(value) Then
                        If isItalic Then
                            formattedValue = "~" & value & "x"
                        Else
                            formattedValue = value & "x"
                        End If
                    ElseIf value <> "" Then
                        If isItalic Then
                            formattedValue = "~" & value
                        End If
                    End If
                    
                Case "date"
                    ' Date formatting - no symbol changes needed
                    If isItalic And value <> "" Then
                        formattedValue = "~" & value
                    End If
                    
                Case ""
                    ' No format specified, just apply italic if needed
                    If isItalic And value <> "" Then
                        formattedValue = "~" & value
                    End If
                    
            End Select
        ElseIf UCase(value) = "F" Then
            ' Handle "F" values (formulas)
            Select Case LCase(formatType)
                Case "dollar"
                    formattedValue = "$F"
                Case "dollaritalic"
                    formattedValue = "~$F"
                Case "volume", "percent", "factor"
                    If isItalic Then
                        formattedValue = "~F"
                    End If
                Case "date"
                    If isItalic Then
                        formattedValue = "~F"
                    End If
                Case ""
                    If isItalic Then
                        formattedValue = "~F"
                    End If
            End Select
        End If
        
        ' Build result
        If i = 0 Then
            result = formattedValue
        Else
            result = result & "|" & formattedValue
        End If
    Next i
    
    ApplySymbolsToRowContent = result
End Function

Sub TestLegacyConversion()
    'Test function to verify the conversion logic works correctly
    
    Dim testCases() As String
    Dim i As Long
    
    ' Define test cases
    ReDim testCases(4)
    testCases(0) = "<CODE; format=""dollar""; row1=""V1|# of subscribers||||||||1000|2000|4000|F|F|F|"";>"
    testCases(1) = "<CODE; format=""dollaritalic""; row1=""AS|Description||||||||500|1000|2000|F|F|F|"";>"
    testCases(2) = "<CODE; format=""volume""; italic=""true""; row1=""V|Count||||||||10|20|30|F|F|F|"";>"
    testCases(3) = "<CODE; format=""factor""; row1=""GR|Growth Rate||||||||2.5|3.0|3.5|F|F|F|"";>"
    testCases(4) = "<CODE; italic=""true""; row1=""LB|Label||||||||Text1|Text2|Text3|F|F|F|"";>"
    
    Debug.Print "=== Testing Legacy Format Conversion ==="
    
    For i = 0 To UBound(testCases)
        Debug.Print "Test Case " & (i + 1) & ":"
        Debug.Print "Original: " & testCases(i)
        Debug.Print "Result: " & ConvertLegacyFormatParameters(testCases(i))
        Debug.Print ""
    Next i
End Sub


