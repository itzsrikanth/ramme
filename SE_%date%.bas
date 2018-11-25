Sub main()

    Dim HubMap As New Scripting.dictionary
    Dim roughSheets() As String
    Dim splits() As String
    
    
    ' Environment check
    Set wb = ActiveWorkbook
    If sheetExists("Sheet1") Then
        Set dataSheet = wb.sheets("Sheet1")
    Else
        MsgBox "Data not available. Exiting Code...!"
        GoTo errorHandle
    End If
    If dataSheet.Cells(1, Columns.Count).End(xlToLeft).Column < 12 Then
        MsgBox "Insufficient columns in data. Exiting Code...!"
        GoTo errorHandle
    End If
    
    ' Initialization
    roughSheets = Split("Sheet1,Sheet2", ",")
    
    HubMap.Add 3644, "CBE;Nilambur"
    HubMap.Add 3647, "MDU;Madurai"
    HubMap.Add 3648, "MDU;Pudukottai"
    HubMap.Add 3649, "CBE;Salem"
    HubMap.Add 3650, "MDU;Tirunelveli"
    HubMap.Add 3651, "MDU;Trichy"
    HubMap.Add 7310, "MDU;Namakkal"
    HubMap.Add 7877, "MDU;Tuticorin"
    HubMap.Add 7997, "CBE;Nilambur"
    HubMap.Add 8160, "MDU;Theni"
    HubMap.Add 8236, "MDU;Perambalur"
    HubMap.Add 8245, "CBE;Sankagiri"
    HubMap.Add 8335, "MDU;Paramakudi"
    HubMap.Add 8338, "CBE;Krishnagiri"
    HubMap.Add 8482, "MDU;Nagercoil"
    HubMap.Add 8521, "MDU;Karur"
    HubMap.Add 25856, "CBE;Salem - II"
    HubMap.Add 25857, "MDU;Madurai _II"
    HubMap.Add 33032, "MDU;Oddanchatram"
    HubMap.Add 33033, "CBE;Tiruppur"
    HubMap.Add 34998, "CBE;Dharmapuri"
    HubMap.Add 36280, "CBE;Mettupalayam"
    HubMap.Add 36377, "MDU;Kumbakonam"
    HubMap.Add 41333, "CBE;Hosur"
    HubMap.Add 42290, "CBE;Pollachi"
    HubMap.Add 42527, "MDU;Ariyalur"
    HubMap.Add 71160, "MDU;Kumbakonam WOW"
    HubMap.Add 71161, "MDU;Madurai WOW"
    HubMap.Add 71178, "CBE;Mettur"
    HubMap.Add 71179, "CBE;WOW-Nilambur"
    HubMap.Add 71187, "MDU;Tuticorin - WOW"
    HubMap.Add 71220, "CBE;Atthur"
    HubMap.Add 71356, "MDU;Rajapalayam"
    HubMap.Add 72926, "CBE;Thiruchengodu"
    HubMap.Add 74402, "CBE;Bhavani"
    HubMap.Add 75562, "MDU;Namakkal II"
    HubMap.Add 76133, "MDU;Dindigul"
    
    ' create empty sheets if not available
    For i = LBound(roughSheets) To UBound(roughSheets)
        If Not sheetExists(roughSheets(i), ActiveWorkbook) Then
            wb.sheets.Add(After:=dataSheet).Name = roughSheets(i)
        End If
    Next i

    ' Process starts
    With wb.sheets(roughSheets(0))
        .Activate
        .Cells(1, 7) = "Outlet"
        .Cells(2, 8) = "Hub"
        For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
            splits = Split(HubMap(Cells(i, 6).Value), ";")
            Cells(i, 7).Value = splits(1)
            Cells(i, 8).Value = splits(0)
        Next i
    End With
    
    ' Final touches
    wb.sheets(roughSheets(1)).Activate
    Range("A1:H1").Select
    With Selection
        .MergeCells = True
        .Font.Size = 11
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    Selection.Merge
    ActiveCell.Value = "Awaiting SE approval claims as on " & Format(Now, "dd.mm.yyyy")
    
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
    sheets = Split("Sheet2", ",")
    For i = LBound(sheets) To UBound(sheets)
        If sheetExists(sheets(i)) Then
            ActiveWorkbook.sheets(sheets(i)).Select
            ActiveWindow.SelectedSheets.Delete
        End If
    Next i
    ActiveWorkbook.sheets("Sheet1").Activate
    Columns("G:H").Select
    Selection.ClearContents
End Sub
