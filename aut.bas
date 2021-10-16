' Dependencies:
' - Microsoft Outlook 14 Object Library
' - HTML Object Library


Sub GetMailSubjects()
    Dim myOlApp As New Outlook.Application
    Dim objNamespace As Outlook.Namespace
    Dim objFolder As Outlook.MAPIFolder
    Dim filteredItems As Outlook.Items
    Dim strFilter As String
    
    Set objNamespace = myOlApp.GetNamespace("MAPI")
    Set objFolder = objNamespace.GetDefaultFolder(olFolderInbox)
    strFilter = "@SQL=" & Chr(34) & "urn:schemas:httpmail:subject" & Chr(34) & " like '%Authorization%'" _
        & " AND %today(" & Chr(34) & "urn:schemas:httpmail:datereceived" & Chr(34) & ")%"
    Set filteredItems = objFolder.Items.Restrict(strFilter)
    
    Dim Html As New HTMLDocument
    Dim Td As IHTMLElementCollection
    Dim i As Long
    Dim ws As Worksheet
    
    Set ws = ActiveWorkbook.Sheets("Sheet1")
    ws.Activate
    irow = Cells(Rows.Count, 1).End(xlUp).Row + 1
    
    If filteredItems.Count = 0 Then
        Found = False
    Else
        Found = True
        For Each Item In filteredItems
            Html.Body.innerHTML = Item.HTMLBody
            Set Td = Html.getElementsByTagName("td")
            
            Dim ICol As Long
            Dim stringOverride As Boolean
            Dim mainValue As String
            Dim result As Boolean
            Dim searchParam As String

            ws.Range("A1:A" & irow).Select
            searchParam = Val(Td.Item(3).innerText)
            For Each c In Selection
                If c.Value = searchParam Then
                    result = True
                    Exit For
                End If
            Next
            If Not result Then
                For i = 1 To Td.Length
                    stringOverride = False
                    Select Case i
                        Case 4:
                            ICol = 1
                        Case 6:
                            ICol = 2
                        Case 8:
                            ICol = 3
                            stringOverride = True
                        Case 10:
                            ICol = 4
                        Case 16:
                            ICol = 7
                        Case 18:
                            ICol = 8
                        Case 20:
                            ICol = 9
                        Case 22:
                            ICol = 20
                        Case 24:
                            ICol = 11
                        Case Else:
                            ICol = 0
                    End Select
                    mainValue = Trim(Td.Item(i - 1).innerText)
                    If IsNumeric(ICol) And ICol > 0 Then
                        If stringOverride Then
                            ws.Cells(irow, ICol).Value = "'" & mainValue
                        Else
                            ws.Cells(irow, ICol).Value = mainValue
                        End If
                    End If
                Next
            End If
        Next
    End If
End Sub
