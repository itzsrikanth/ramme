Sub main()

    Dim dataSheet As Worksheet
    Dim wb As Workbook
    Dim roughSheets() As String
    Dim pivotTables() As String
    Dim claimStatus() As String
    ' Add Microsoft Scripting Runtime in references
    Dim dict As New Scripting.Dictionary
    Dim map As New Scripting.Dictionary
    Dim hubMap As Worksheet
    Dim rowNum As Integer
    
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
            roughSheets(0) & "!R1C1:R" & Rows.Count & "C41", Version:=xlPivotTableVersion15).CreatePivotTable _
            TableDestination:=roughSheets(1) & "!R1C1", TableName:=pivotTables(0), DefaultVersion _
            :=xlPivotTableVersion15
        With .pivotTables(pivotTables(0)).PivotFields("REMARKS")
            .Orientation = xlRowField
            .Position = 1
        End With
        .pivotTables(pivotTables(0)).AddDataField ActiveSheet.pivotTables( _
            pivotTables(0)).PivotFields("Claim Amount"), "Sum of Claim Amount", xlSum
        
    End With
        
    ' Step 3
    With wb.sheets("Sheet1")
        .Activate
        .Range("$A$1:$AQ$" & CStr(Cells(Rows.Count, 1).End(xlUp).Row)).Select
    End With
    Selection.Copy
    With wb.sheets(roughSheets(3))
        .Activate
        .Cells(1, 1).Select
        .Paste
        .Cells(1, 1).Select
    End With
    
'    ActiveSheet.Range("$A$1:$AO$" & CStr(Cells(Rows.Count, 1).End(xlUp).Row)).removeduplicates Columns:=41, Header:=xlYes
    
    ' Second last step
    ' Removing unwanted columns from CLAIMWISE
     With wb.sheets(roughSheets(3))
        .Activate
        .Range("$A$1:$AO$" & CStr(Cells(Rows.Count, 1).End(xlUp).Row)).removeduplicates Columns:=41, Header:=xlYes
        .Columns("AC:AF").Select
        Selection.Delete Shift:=xlToLeft
        .Columns("AD:AD").Select
        Selection.Delete Shift:=xlToLeft
        .Columns("AE:AH").Select
        Selection.Delete Shift:=xlToLeft
    End With
    
    ' Final step
    
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "CLAIMWISE!R1C1:R14856C34", Version:=xlPivotTableVersion15).CreatePivotTable _
        TableDestination:="Sheet4!R3C1", TableName:=pivotTables(1), DefaultVersion _
        :=xlPivotTableVersion15
    sheets("Sheet4").Select
    Cells(3, 1).Select
    With ActiveSheet.pivotTables(pivotTables(1)).PivotFields("Sales Organisasation")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.pivotTables(pivotTables(1)).PivotFields("Plant Name")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.pivotTables(pivotTables(1)).PivotFields("Status")
        .Orientation = xlColumnField
        .Position = 1
    End With
    ActiveSheet.pivotTables(pivotTables(1)).AddDataField ActiveSheet.pivotTables( _
        pivotTables(1)).PivotFields("Claim Amount"), "Sum of Claim Amount", xlSum
    With ActiveSheet.pivotTables(pivotTables(1)).PivotFields("REMARKS")
        .Orientation = xlRowField
        .Position = 3
    End With
    ActiveSheet.pivotTables(pivotTables(1)).AddDataField ActiveSheet.pivotTables( _
        pivotTables(1)).PivotFields("REMARKS"), "Count of REMARKS", xlCount
        
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
        ActiveWorkbook.sheets(sheets(i)).Select
        ActiveWindow.SelectedSheets.Delete
    Next i
    ActiveWorkbook.sheets("Sheet1").Activate
    Columns("AO:AQ").Select
    Selection.ClearContents
End Sub
