Sub main()

    Dim dataSheet As Worksheet
    Dim wb As Workbook
    Dim roughSheets() As String
    Dim pivotTables() As String
    Dim claimStatus() As String
    ' Add Microsoft Scripting Runtime in references
    Dim dict As New Scripting.Dictionary
    Dim map As New Scripting.Dictionary
    Dim pivot As New Scripting.Dictionary
    Dim hubMap As Worksheet
    Dim rowNum As Integer
    Dim claimWiseRowCount As Integer
    Dim RegEx As New regexp
    Dim tvsPattern As String
    
    ' Environment check
    Set wb = ActiveWorkbook
    If sheetExists("Sheet1") Then
        Set dataSheet = wb.sheets("Sheet1")
    Else
        MsgBox "Data not available. Exiting Code...!"
        GoTo errorHandle
    End If
    If dataSheet.Cells(1, Columns.Count).End(xlToLeft).Column < 40 Then
        MsgBox "Insufficient columns in data. Exiting Code...!"
        GoTo errorHandle
    End If
    
    ' Initialization
    tvsPattern = "^TVS\s*-?\s*(\w+)(\s*-?\s*\w+)?$"
    With RegEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .Pattern = tvsPattern
    End With
    roughSheets = Split("Sheet1,Sheet2,Sheet4,CLAIMWISE", ",")
    pivotTables = Split("SumOfClaimAmounts,FaceSheetPivot", ",")
    
    dict.Add "B01X", "Returned"
    dict.Add "B001", "Claim not uploaded"
    dict.Add "", "Claim to be generated"
    
    map.Add "S054", "TN 2"
    map.Add "S055", "Karnataka"
    map.Add "S056", "Kerala"
    map.Add "S057", "MP"
    map.Add "S058", "Madurai"
    map.Add "S059", "CBE"
    map.Add "S063", "Karnataka"
    
    ' create empty sheets if not available
    For i = LBound(roughSheets) To UBound(roughSheets)
        If Not sheetExists(roughSheets(i), wb) Then
            wb.sheets.Add(After:=dataSheet).Name = roughSheets(i)
        End If
    Next i
    
    ' Step 1
    ' Adding the values in column AO
    With wb.sheets(roughSheets(0))
        .Activate
        .Cells(1, 41) = "REMARKS"
        .Cells(1, 42) = "Status"
        .Cells(1, 43) = "Month"
        Columns("AO:AO").Select
        Selection.NumberFormat = "@"
        For Each cell In .Range("$G$2:$G$" & CStr(Cells(Rows.Count, 1).End(xlUp).Row)).Cells
            rowNum = cell.Row
            .Cells(rowNum, 5).Value = RegEx.Replace(.Cells(rowNum, 5).Value, "$1$2")
            .Cells(rowNum, 41).Value = "" & .Cells(rowNum, 7) & .Cells(rowNum, 35)
            If dict.Exists(.Cells(rowNum, 26).Value) Then
                .Cells(rowNum, 42).Value = dict(.Cells(rowNum, 26).Value)
            End If
            .Cells(rowNum, 43).Value = Month(.Cells(rowNum, 9))
        Next cell
    End With
    
    ' Step 2
        
    With wb.sheets(roughSheets(1))
        ' truncating old data from pages
        .Select
        Cells.Select
        Selection.ClearContents
        
        ' Making Sheet2 pivot table
        wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
            roughSheets(0) & "!R1C1:R" & Rows.Count & "C41").CreatePivotTable _
            TableDestination:=roughSheets(1) & "!R1C1", TableName:=pivotTables(0)
        With .pivotTables(pivotTables(0)).PivotFields("REMARKS")
            .Orientation = xlRowField
            .Position = 1
        End With
        .pivotTables(pivotTables(0)).AddDataField ActiveSheet.pivotTables( _
            pivotTables(0)).PivotFields("Claim Amount"), "Sum of Claim Amount", xlSum
            
        ' reading entire pivot into a dictionary
        For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row - 1
            pivot.Add Cells(i, 1).Value, Cells(i, 2).Value
        Next i
        
    End With
        
    ' Step 3
    With wb.sheets("Sheet1")
        .Activate
        .Cells.Select
        .Cells.EntireColumn.AutoFit
        With Selection
            .RowHeight = 15
            .Font.Name = "Liberation Sans"
            .Font.Size = 9
        End With
        .Range("$A$1:$AQ$" & CStr(Cells(Rows.Count, 1).End(xlUp).Row)).Select
    End With
    Selection.Copy
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    With wb.sheets(roughSheets(3))
        .Activate
        .Cells(1, 1).Select
        .Paste
        .Cells(1, 1).Select
    End With
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    
'    ActiveSheet.Range("$A$1:$AO$" & CStr(Cells(Rows.Count, 1).End(xlUp).Row)).removeduplicates Columns:=41, Header:=xlYes
    
    ' Second last step
    ' Removing unwanted columns from CLAIMWISE
     With wb.sheets(roughSheets(3))
        .Activate
        .Cells.Select
        .Cells.EntireColumn.AutoFit
        With Selection
            .RowHeight = 15
            .Font.Name = "Liberation Sans"
            .Font.Size = 9
        End With
        .Range("$A$1:$AO$" & CStr(Cells(Rows.Count, 1).End(xlUp).Row)).removeduplicates Columns:=41, Header:=xlYes
        .Columns("AC:AF").Select
        Selection.Delete Shift:=xlToLeft
        .Columns("AD:AD").Select
        Selection.Delete Shift:=xlToLeft
        .Columns("AE:AH").Select
        Selection.Delete Shift:=xlToLeft
        claimWiseRowCount = Cells(Rows.Count, 1).End(xlUp).Row
        ' adding VLOOKUP and autofill other rows
        For i = 2 To claimWiseRowCount
            If pivot.Exists(Cells(i, 32).Value) Then
                Cells(i, 29).Value = pivot(Cells(i, 32).Value)
            End If
            ' Cells(i, 29).FormulaR1C1 = "=VLOOKUP(R" & i & "C32," & roughSheets(1) & "!R" & i & "C1:R" & claimWiseRowCount & "C2, 2,)"
            ' Range("AC2").Select
            ' Selection.AutoFill Destination:=Range("AC2:AC" & claimWiseRowCount)
            ' Range("AC2:AC" & claimWiseRowCount).FillDown
        Next i
    End With
    
    ' Final step
    
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "CLAIMWISE!R1C1:R" & Cells(Rows.Count, 1).End(xlUp).Row & "C34").CreatePivotTable _
        TableDestination:="Sheet4!R3C1", TableName:=pivotTables(1)
    sheets("Sheet4").Select
    Cells(3, 1).Select
    With ActiveSheet.pivotTables(pivotTables(1)).PivotFields("Sales Organisasation")
        .Orientation = xlRowField
        .Position = 1
        .Caption = "Hub"
    End With
    With ActiveSheet.pivotTables(pivotTables(1)).PivotFields("Plant Name")
        .Orientation = xlRowField
        .Position = 2
        .Caption = "Outlet"
    End With
    With ActiveSheet.pivotTables(pivotTables(1)).PivotFields("Status")
        .Orientation = xlRowField
        .Position = 3
    End With
    With ActiveSheet.pivotTables(pivotTables(1)).PivotFields("Status")
        .Orientation = xlColumnField
        .Position = 1
    End With
    With ActiveSheet.pivotTables(pivotTables(1)).PivotFields("REMARKS")
        .Orientation = xlRowField
        .Position = 3
    End With
    ActiveSheet.pivotTables(pivotTables(1)).AddDataField ActiveSheet.pivotTables( _
        pivotTables(1)).PivotFields("REMARKS"), "Count of REMARKS", xlCount
    ActiveSheet.pivotTables(pivotTables(1)).AddDataField ActiveSheet.pivotTables( _
        pivotTables(1)).PivotFields("Claim Amount"), "Sum of Claim Amount", xlSum
    
    ' pivot table formatting
    With ActiveSheet.pivotTables(pivotTables(1))
        .MergeLabels = True
        .InGridDropZones = True
        .RowAxisLayout xlTabularRow
    End With
    ActiveSheet.pivotTables(pivotTables(1)).DataPivotField.PivotItems( _
        "Count of REMARKS").Caption = "Nos."
    ActiveSheet.pivotTables(pivotTables(1)).DataPivotField.PivotItems( _
        "Sum of Claim Amount").Caption = "Amount"
    ActiveSheet.pivotTables(pivotTables(1)).PivotSelect "'Claim not uploaded'", _
        xlDataAndLabel, True
    ActiveSheet.pivotTables("FaceSheetPivot").FieldListSortAscending = True
    
    Cells.Select
    With Selection
        .RowHeight = 15
        .Font.Name = "Liberation Sans"
        .Font.Size = 9
    End With
    
    Range("C:C,E:E,G:G,I:I").Select
    With Selection
        .ColumnWidth = 9
    End With
    Range("D:D,F:F,H:H,J:J").Select
    With Selection
        .NumberFormat = "0"
        .ColumnWidth = 9
    End With
    ' Range("A:A").Select
    ' Selection.ColumnWidth = 9
    Range("B:B").Select
    Selection.HorizontalAlignment = xlLeft
        
    ' Final touches
    wb.sheets(roughSheets(2)).Range("A1:J1").Select
    With Selection
        .MergeCells = True
        .Font.Size = 11
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    Selection.Merge
    ActiveCell.Value = "SAP DBM Claims status as on " & Format(Now, "dd.mm.yyyy")
    
    MsgBox "Script completed successfully"
    
errorHandle:
    Debug.Print Err.Description
    Err.Clear

End Sub

Function sheetExists(sheetToFind As String, Optional InWorkbook As Workbook) As Boolean
    If InWorkbook Is Nothing Then Set InWorkbook = ThisWorkbook
    On Error Resume Next
    sheetExists = Not InWorkbook.sheets(sheetToFind) Is Nothing
End Function

Sub cleanUp()
    Dim sheets() As String
    sheets = Split("Sheet2,Sheet4,CLAIMWISE", ",")
    For i = LBound(sheets) To UBound(sheets)
        If sheetExists(sheets(i)) Then
            ActiveWorkbook.sheets(sheets(i)).Select
            ActiveWindow.SelectedSheets.Delete
        End If
    Next i
    ActiveWorkbook.sheets("Sheet1").Activate
    Columns("AO:AQ").Select
    Selection.ClearContents
End Sub
