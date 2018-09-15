Sub main()

    ' Variable declarations
    Dim dataSheet As Worksheet
    Dim wb As Workbook
    Dim roughSheets() As String
    Dim pivotTables() As String
    Dim claimStatus() As String
    Dim dict As New Scripting.Dictionary
    Dim map As New Scripting.Dictionary
    Dim hubMap As Worksheet
    
    ' Environment check
    Set wb = ActiveWorkbook
    If sheetExists("Data") Then
        Set dataSheet = wb.Sheets("Data")
    Else
        MsgBox "Data not available. Exiting Code...!"
        GoTo errorHandle
    End If
    If sheetExists("Hub Map") Then
        Set hubMap = wb.Sheets("Hub Map")
    Else
        MsgBox "Hub Maps not available. Exiting Code...!"
        GoTo errorHandle
    End If
    If dataSheet.Cells(1, Columns.Count).End(xlToLeft).Column < 40 Then
        MsgBox "Insufficient columns in data. Exiting Code...!"
        GoTo errorHandle
    End If
    
    ' Initialization
    roughSheets = Split("Returned claims,Claim not uploaded,Claim to be generated,Face Sheet,All status,Sheet8", ",")
    pivotTables = Split("ReturnedClaimsPT,ClaimNotUploadedPT,ClaimToBeGeneratedPT", ",")
    claimStatus = Split("B01X,B001,", ",")
    
    ' create empty sheets if not available
    For i = LBound(roughSheets) To UBound(roughSheets)
        If Not sheetExists(roughSheets(i), wb) Then
            wb.Sheets.Add(After:=dataSheet).Name = roughSheets(i)
        End If
    Next i
    
    ' truncating old data from pages
    wb.Sheets(roughSheets(5)).Select
    Cells.Select
    Selection.ClearContents
    
    ' Generating other worksheets from master data
    For i = LBound(claimStatus) To UBound(claimStatus)
        With dataSheet
            .Activate
            If ActiveSheet.AutoFilterMode Then
                Selection.AutoFilter
            End If
            'On Error GoTo errorHandle
            .Range("$A:$AN").AutoFilter Field:=26, Criteria1:=claimStatus(i)
            .Range("$A$1:$AN$" & CStr(Cells(Rows.Count, 1).End(xlUp).Row)).Select
        End With
        Selection.Copy
        
        With wb.Sheets(roughSheets(i))
            .Activate
            .Cells(1, 1).Select
            .Paste
            .Cells(1, 1).Select
        End With
    Next i
    
    Set map = hubMapFn
    
    ' "Returned claims"
    With wb.Sheets(roughSheets(0))
        .Activate
        .Range("$A$1:$AN$" & CStr(Cells(Rows.Count, 1).End(xlUp).Row)).RemoveDuplicates Columns:=7, Header:=xlYes
        .Columns("AC:AG").Select
        Selection.Delete Shift:=xlToLeft
        .Columns("AE:AF").Select
        Selection.Delete Shift:=xlToLeft
        dataSheet.Activate
        ' grouping the claim amount based on "Job Card Number",
        ' and getting return in dictionary
        Set dict = groupAdd("X")
        ' looping cells to paste values in dictionary to column "AC" - #29
        .Activate
        For Each cell In .Range("X2:X" & CStr(Cells(Rows.Count, 1).End(xlUp).Row)).Cells
            If dict.Exists(cell.Value) Then
                .Cells(cell.Row, 29) = dict(cell.Value)
            End If
            .Cells(cell.Row, 41) = map(.Cells(cell.Row, 3).Value)
        Next cell
        .Cells(1, 29) = "Claim Amount"
        .Cells(1, 41) = "Hub"
        
        ' inserting pivot table
        ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
            roughSheets(0) & "!R1C1:R" & Rows.Count & "C34", Version:=xlPivotTableVersion15). _
            CreatePivotTable TableDestination:=roughSheets(5) & "!R3C1", TableName:=pivotTables(0) _
            , DefaultVersion:=xlPivotTableVersion15
        Sheets(roughSheets(5)).Select
        With ActiveSheet.pivotTables(pivotTables(0)).PivotFields("Hub")
            .Orientation = xlRowField
            .Position = 1
        End With
        With ActiveSheet.pivotTables(pivotTables(0)).PivotFields("Plant Name")
            .Orientation = xlRowField
            .Position = 2
        End With
        ActiveSheet.pivotTables(pivotTables(0)).AddDataField ActiveSheet.pivotTables( _
            pivotTables(0)).PivotFields("Active Claim Number"), _
            "No. of Claims", xlCount
        ActiveSheet.pivotTables(pivotTables(0)).AddDataField ActiveSheet.pivotTables( _
            pivotTables(0)).PivotFields("Claim Amount"), "Total Amount", xlSum
    End With
    
    ' "Claim not uploaded"
    With wb.Sheets(roughSheets(1))
        .Activate
        .Range("$A$1:$AN$" & CStr(Cells(Rows.Count, 1).End(xlUp).Row)).RemoveDuplicates Columns:=7, Header:=xlYes
        .Columns("AB:AF").Select
        Selection.Delete Shift:=xlToLeft
        .Columns("AC:AC").Select
        Selection.Delete Shift:=xlToLeft
        .Columns("AE:AH").Select
        Selection.Delete Shift:=xlToLeft
        Set dict = groupAdd("X")
        ' looping cells to paste values in dictionary to column "AB" - #28
        .Activate
        For Each cell In .Range("X2:X" & CStr(Cells(Rows.Count, 1).End(xlUp).Row)).Cells
            If dict.Exists(cell.Value) Then
                .Cells(cell.Row, 28) = dict(cell.Value)
            End If
            .Cells(cell.Row, 31) = map(Cells(cell.Row, 3).Value)
        Next cell
        .Cells(1, 31) = "Hub"
    End With
    
    ' "Claim to be generated"
    With wb.Sheets(roughSheets(2))
        .Activate
        .Range("$A$1:$AN$" & CStr(Cells(Rows.Count, 1).End(xlUp).Row)).RemoveDuplicates Columns:=7, Header:=xlYes
        .Columns("T:AF").Select
        Selection.Delete Shift:=xlToLeft
        .Columns("U:AB").Select
        Selection.Delete Shift:=xlToLeft
        Set dict = groupAdd("G")
        .Activate
        For Each cell In .Range("G2:G" & CStr(Cells(Rows.Count, 1).End(xlUp).Row)).Cells
            If dict.Exists(cell.Value) Then
                .Cells(cell.Row, 20) = dict(cell.Value)
            End If
            .Cells(cell.Row, 21) = map(.Cells(cell.Row, 3).Value)
        Next cell
        .Cells(1, 21) = "Hub"
    End With
    
errorHandle:
     Debug.Print Err.Description
    Err.Clear


    
End Sub

Function sheetExists(sheetToFind As String, Optional InWorkbook As Workbook) As Boolean
    If InWorkbook Is Nothing Then Set InWorkbook = ThisWorkbook
    On Error Resume Next
    sheetExists = Not InWorkbook.Sheets(sheetToFind) Is Nothing
End Function

Function groupAdd(colName As String, Optional ws As Worksheet) As Scripting.Dictionary

    Dim dict As New Scripting.Dictionary

    ' selecting worksheet
    If ws Is Nothing Then
        Set ws = ActiveSheet
    Else
        ws.Activate
    End If
    'removing any active filters
    If ActiveSheet.AutoFilterMode Then
        Selection.AutoFilter
    End If
    '
    With ws
        For Each cell In .Range(colName & "2:" & colName & CStr(Cells(Rows.Count, 1).End(xlUp).Row)).Cells
            If Not dict.Exists(cell.Value) Then
                dict.Add cell.Value, 0
            End If
            dict(cell.Value) = dict(cell.Value) + .Cells(cell.Row, 33)
        Next cell
    End With
    
    Set groupAdd = dict
    
End Function

Function hubMapFn() As Scripting.Dictionary
    Dim map As New Scripting.Dictionary
    With ActiveWorkbook.Sheets("Hub Map")
        .Activate
        For Each cell In .Range("$A$1:$A$" & .Cells(1, 1).End(xlDown).Row).Cells
                map.Add cell.Value, .Cells(cell.Row, 2)
        Next cell
    End With
    Set hubMapFn = map
End Function
