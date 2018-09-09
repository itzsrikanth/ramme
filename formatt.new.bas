Sub rowIns(rowmax As Integer, colmax As Integer)
    
    Dim col As String
    col = Left(Cells(1, colmax + 1).Address(False, False), 1)
    
    For i = 2 To rowmax
        Select Case Range("A" & i).Value
            Case 3644
                Range(col & i).Value = "Nilambur"
            Case 3647
                Range(col & i).Value = "Madurai"
            Case 3648
                Range(col & i).Value = "Pudukottai"
            Case 3649
                Range(col & i).Value = "Salem"
            Case 3650
                Range(col & i).Value = "Tirunelveli"
            Case 3651
                Range(col & i).Value = "Trichy"
            Case 7310
                Range(col & i).Value = "Namakkal"
            Case 7877
                Range(col & i).Value = "Tuticorin"
            Case 7997
                Range(col & i).Value = "MTP Road"
            Case 8160
                Range(col & i).Value = "Theni"
            Case 8236
                Range(col & i).Value = "Perambalur"
            Case 8245
                Range(col & i).Value = "Sankagiri"
            Case 8335
                Range(col & i).Value = "Paramakudi"
            Case 8338
                Range(col & i).Value = "Krishnagiri"
            Case 8482
                Range(col & i).Value = "Nagercoil"
            Case 8521
                Range(col & i).Value = "Karur"
            Case 25856
                Range(col & i).Value = "Salem II"
            Case 25857
                Range(col & i).Value = "Madurai II"
            Case 33032
                Range(col & i).Value = "Oddenchatram"
            Case 33033
                Range(col & i).Value = "Tiruppur"
            Case 34998
                Range(col & i).Value = "Dharmapuri"
            Case 36280
                Range(col & i).Value = "Mettupalayam"
            Case 36377
                Range(col & i).Value = "Kumbakonam"
            Case 41333
                Range(col & i).Value = "Hosur"
            Case 42290
                Range(col & i).Value = "Pollachi"
            Case 42527
                Range(col & i).Value = "Ariyalur"
        End Select
    Next i
End Sub
Sub f()

    Dim rowmaxPack As Integer, _
        colmaxPack As Integer, _
        rowmaxShip As Integer, _
        colmaxShip As Integer
    Dim temp As String

    rowmaxPack = Worksheets(1).Cells(1, 1).End(xlDown).Row
    colmaxPack = Worksheets(1).Cells(1, 1).End(xlToRight).Column
    rowmaxShip = Worksheets(2).Cells(1, 1).End(xlDown).Row
    colmaxShip = Worksheets(2).Cells(1, 1).End(xlToRight).Column
    
    '''''''''''Sheet1 - Generating J and K columns'''''''''''''
    Worksheets(1).Select
    Call rowIns(rowmaxPack, colmaxPack)
    Range("K1").Value = "> 30 days"
    Range("J1").Value = "Outlet"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-2]>30,""YES"",""NO"")"
    Selection.AutoFill Destination:=Range("K2:K" & rowmaxPack), Type:=xlFillDefault
    colmaxPack = Cells(1, 1).End(xlToRight).Column
    
    '''''''''''Sheet2 - Generating E and F columns'''''''''''''
    Worksheets(2).Select
    Call rowIns(rowmaxShip, colmaxShip)
    Range("F1").Value = "> 7 days"
    Range("E1").Value = "Outlet"
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-2]>7,""YES"",""NO"")"
    Selection.AutoFill Destination:=Range("F2:F" & rowmaxShip), Type:=xlFillDefault
    colmaxShip = Cells(1, 1).End(xlToRight).Column

    ''''''''''Format Sheet1 Cells''''''''''''
    Worksheets(1).Select
    Columns("A:K").Select
    Columns("A:K").EntireColumn.AutoFit
    Range("A1:K1").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    Columns("E:G").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A2").Select
    Range("A:D,H:K").Select
    Range("H1").Activate
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A2").Select
    
    For i = 1 To 11
        temp = Left(Cells(1, i).Address(False, False), 1)
        border (temp & "1:" & temp & rowmaxPack)
    Next i
    Range("A2").Select
    
    ''''''''''Format Sheet2 Cells''''''''''''
    Worksheets(2).Select
    Columns("A:F").Select
    Columns("A:F").EntireColumn.AutoFit
    Range("A1:F1").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    
    Range("A2").Select
    Range("A:F").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A2").Select
    
    For i = 1 To 6
        temp = Left(Cells(1, i).Address(False, False), 1)
        border (temp & "1:" & temp & rowmaxShip)
    Next i
    Range("A2").Select
    
    ''''''''''''Position sheets in order'''''''''''
    If Sheets.Count < 3 Then
        For i = Sheets.Count + 1 To 3
            Sheets.Add after:=Worksheets(Sheets.Count)
        Next i
    End If
    Sheets.Add after:=Worksheets(Sheets.Count)
    Sheets(1).Name = "NOT PACKED MATERIALS"
    Sheets(2).Name = "Not shipped pkg slips"
    Sheets("NOT PACKED MATERIALS").Move after:=Worksheets(4)
    Sheets("Not shipped pkg slips").Move after:=Sheets("NOT PACKED MATERIALS")
    
    
    ''''''''''''' Editing Sheet3,4 ''''''''''''''''''''
    Sheets(2).Select
    Sheets(2).Range("A1:D1").Merge
    ActiveCell.Value = "Materials not packed as on - " & format(now, "dd.mm.yyyy")
    Sheets(2).Range("F1:I1").Merge
    Sheets(2).Range("F1").Select
    ActiveCell.Value = "Packing slips not dispatched - " & format(now, "dd.mm.yyyy")
    
    Sheets(1).Select
    Sheets(1).Range("A2:B2").Merge
    Sheets(1).Range("A2").Select
    ActiveCell.Value = "Materials not packed as on - " & format(now, "dd.mm.yyyy")
    Sheets(1).Range("G2:H2").Merge
    Sheets(1).Range("G2").Select
    ActiveCell.Value = "Packing slips not dispatched - " & format(now, "dd.mm.yyyy")
    
    
    '''''''''''''Pivot table - Materials not packed - Sheet4''''''''''''
    Sheets("NOT PACKED MATERIALS").Select
    temp = "NOT PACKED MATERIALS!R1C1:R" & rowmaxPack & "C" & colmaxPack
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=temp).CreatePivotTable tabledestination:=Sheets(2).Name & "!R3C1", _
                    tablename:="MaterialNotPacked3"
    With Worksheets(2)
        .Select
        .Cells(3, 1).Select
    End With
    With ActiveSheet.PivotTables("MaterialNotPacked3").PivotFields("Outlet")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("MaterialNotPacked3").AddDataField ActiveSheet.PivotTables( _
        "MaterialNotPacked3").PivotFields("> 30 days"), "Count of > 30 days", xlCount
    With ActiveSheet.PivotTables("MaterialNotPacked3").PivotFields("> 30 days")
        .Orientation = xlColumnField
        .Position = 1
    End With
    
    ''''''''''''''Pivot table - Materials not packed - Sheet3''''''''''''
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=temp).CreatePivotTable tabledestination:=Sheets(1).Name & "!R5C1", _
                    tablename:="MaterialNotPacked2"
    With Worksheets(1)
        .Select
        .Cells(5, 1).Select
    End With
    With ActiveSheet.PivotTables("MaterialNotPacked2").PivotFields("Outlet")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("MaterialNotPacked2").AddDataField ActiveSheet. _
        PivotTables("MaterialNotPacked2").PivotFields("Reference No"), _
        "Count of Reference No", xlCount
    
    '''''''''''''Pivot table - Materials not shipped - Sheet4''''''''''''
    Sheets("Not shipped pkg slips").Select
    temp = "Not shipped pkg slips!R1C1:R" & rowmaxShip & "C" & colmaxShip
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=temp).CreatePivotTable tabledestination:=Sheets(2).Name & "!R3C6", _
                    tablename:="MaterialNotShipped3"
    With Worksheets(2)
        .Select
        .Cells(3, 6).Select
    End With
    With ActiveSheet.PivotTables("MaterialNotShipped3").PivotFields("Outlet")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("MaterialNotShipped3").AddDataField ActiveSheet.PivotTables( _
        "MaterialNotShipped3").PivotFields("> 7 days"), "Count of > 7 days", xlCount
    With ActiveSheet.PivotTables("MaterialNotShipped3").PivotFields("> 7 days")
        .Orientation = xlColumnField
        .Position = 1
    End With
        
    ''''''''''''''Pivot table - Materials not shipped - Sheet3''''''''''''
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=temp).CreatePivotTable tabledestination:=Sheets(1).Name & "!R5C7", _
                    tablename:="MaterialNotShipped2"
    With Worksheets(1)
        .Select
        .Cells(5, 7).Select
    End With
    With ActiveSheet.PivotTables("MaterialNotShipped2").PivotFields("Outlet")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("MaterialNotShipped2").AddDataField ActiveSheet. _
        PivotTables("MaterialNotShipped2").PivotFields("Packing Slip No"), _
        "Count of Packing Slip No", xlCount
    pivotformat
End Sub
Sub pivotformat()
    '''''''Formatting pivot''''''''''
    
    With Worksheets(2)
        .Select
        .Cells(4, 1).Value = "Outlet"
        .Cells(4, 2).Value = "# claims < 30 days"
        .Cells(4, 3).Value = "# claims > 30 days"
        .Cells(4, 4).Value = "Total"
        .Cells(.Cells(3, 1).End(xlDown).Row, 1) = "Total"
        .Cells(4, 6) = "Outlet"
        .Cells(4, 7) = "# Pkg slips < 7 Days"
        .Cells(4, 8) = "# Pkg slips > 7 Days"
        .Cells(4, 9) = "Total"
        .Cells(.Cells(3, 6).End(xlDown).Row, 1) = "Total"
        .Name = "Sheet2"
    End With
    With Worksheets(1)
        .Select
        .Cells(4, 1).Value = "No. of claims"
        .Cells(5, 1).Value = "Outlet"
        .Cells(5, 2).Value = "Total"
        .Cells(4, 7) = "No. of packing slips"
        .Cells(5, 7) = "Outlet"
        .Cells(5, 8) = "Total"
        .Name = "Sheet1"
    End With
    '''''''''Sheet1'''''''''''''
    Worksheets("Sheet1").Select
    Range("B5").Select
    Selection.Copy
    Range("A4:B4").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("H5").Select
    Selection.Copy
    Range("G4:H4").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A4:B" & Cells(4, 1).End(xlDown).Row).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("G4:H" & Cells(4, 7).End(xlDown).Row).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    Range("B5:B" & Cells(4, 1).End(xlDown).Row).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
    End With
    Range("H5:H" & Cells(4, 7).End(xlDown).Row).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
    End With
    Range("A2:B2").Select
    Selection.Font.Bold = True
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("G2:H2").Select
    Selection.Font.Bold = True
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    Cells(1, 1).Select
    
    ''''''''''''Sheet2'''''''''''''
    Worksheets("Sheet2").Select
    Range("A1:D1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    Range("F1:I1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    Range("A4:D" & Cells(4, 1).End(xlDown).Row).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    Range("B4:D" & Cells(4, 1).End(xlDown).Row).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
    End With
    Range("F4:I" & Cells(4, 6).End(xlDown).Row).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    Range("G4:I" & Cells(4, 6).End(xlDown).Row).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
    End With
    Columns("C:C").ColumnWidth = 10
    Columns("D:D").ColumnWidth = 10
    Columns("H:H").ColumnWidth = 10
    Columns("I:I").ColumnWidth = 10
    Columns("G:G").ColumnWidth = 15
    Columns("B:B").ColumnWidth = 15
    Rows("4:4").Select
    With Selection
        .VerticalAlignment = xlBottom
        .WrapText = True
        .ReadingOrder = xlContext
    End With
    Cells(1, 1).Select

End Sub
    ''''borders''''''''''''
Sub border(str As String)
    Range(str).Select     ''''  "A1:A119"
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub
