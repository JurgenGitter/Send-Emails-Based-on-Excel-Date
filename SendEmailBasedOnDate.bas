Sub SendEmailBasedOnDate()
 
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim BMemail As String
    Dim DSMemail As String
    Dim dueDate As Date
    Dim OutApp As Object
    Dim OutMail As Object
    Const COL_BMMAIL As String = "D"
    Const COL_DMSMAIL As String = "E"
    Const COL_DATE As String = "F"
 
    Set ws = ThisWorkbook.Sheets("Sheet1") ' <-- change sheet1 to name of sheet
    lastRow = ws.Cells(ws.Rows.Count, COL_DATE).End(xlUp).Row
    ' Create Outlook instance
    Set OutApp = CreateObject("Outlook.Application")
    For i = 2 To lastRow
        BMemail = ws.Cells(i, COL_BMMAIL).Value
        DSMemail = ws.Cells(i, COL_DSMMAIL).Value
        dueDate = ws.Cells(i, COL_DATE).Value
        If IsDate(dueDate) Then
            If dueDate = Date Then
                Set OutMail = OutApp.CreateItem(0)
                With OutMail
                    .To = BMemail & ";" & DSMemail
                    .Subject = "Automated Reminder"
                    .Body = "Hello," & vbCrLf & vbCrLf & _
                            "This is a reminder that your Agility appointment is occurring today (" & Date & ")." & vbCrLf & vbCrLf & _
                            "Thank you for your cooperation," & vbCrLf & "The Agility Team"
                    .Send
                End With
            End If
        End If
    Next i
 
    Set OutMail = Nothing
    Set OutApp = Nothing
    MsgBox "Emails sent for today's date!", vbInformation
End Sub
