Sub main()

    ' Variable declarations
    Dim dataSheet As Worksheet
    Dim wb As Workbook
    Dim roughSheets() As String
    Dim claimStatus() As String
    Dim dict As New Scripting.Dictionary
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
    roughSheets = Split("Returned claims,Claim not uploaded,Claim to be generated,Face Sheet,All status", ",")
    claimStatus = Split("B01X,B001,", ",")
    
    ' create empty sheets if not available
    For i = LBound(roughSheets) To UBound(roughSheets)
        If Not sheetExists(roughSheets(i), wb) Then
            wb.Sheets.Add(After:=dataSheet).Name = roughSheets(i)
        End If
    Next i
    
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
            '.Cells(cell.Row, 34) =
            MsgBox hubMap("S055")  'y ' hubMap(.Cells(cell.Row, 3).Value)
        Next cell
        .Cells(1, 29) = "Claim Amount"
        .Cells(1, 41) = "Hub"
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
            .Cells(cell.Row, 34) = hubMap(Cells(cell.Row, 3).Value)
        Next cell
        .Cells(1, 34) = "Hub"
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
            .Cells(cell.Row, 21) = hubMap(Cells(cell.Row, 3).Value)
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

Function hubMap(Optional hubCode As String)
    With ActiveWorkbook.Sheets("Hub Map")
        .Activate
        For Each cell In .Range("$A$1:$A$" & .Cells(1, 1).End(xlDown).Row).Cells
            If hubCode = cell.Value Then
                hubMap = .Cells(cell.Row, 2)
                Exit For
            End If
        Next cell
    End With
End Function
