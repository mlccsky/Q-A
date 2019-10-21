Private Sub Workbook_Activate()

'Call checkUserDomain("CBRE")

Call fillCompanyCombo ' fill client drop down box on worksheet Quote

Call populate_cboReportType(True)
TrackEvent "BDCC Quote & Acceptance"

End Sub

Private Sub Workbook_Open()
On Error GoTo Err_Workbook_Open
Dim X As String
Dim sDomain As String
Dim user As String
Dim currentUser As String
Dim ws As Worksheet

    Call settings(False)

    ' check file path for store is valid, otherwise default to user's local drive
    On Error Resume Next
    If Len(Dir(wsQuote.Range("zzListFilePath"), vbDirectory)) = 0 Then
        wsQuote.Range("zzListFilePath") = Application.DefaultFilePath & "\"
    End If
    
    X = "CBRE"
    sDomain = Environ("USERDNSDOMAIN")

    ' check if user is on the CBRE domain
    ' otherwise remove worksheets
    If InStr(1, sDomain, X) = 0 Then
        ' unhide all worksheets
        Call unhideWS
    
        For Each ws In Worksheets
            Select Case ws.CodeName
                'Include sheet names to keep on next line (with comma between)
                Case wsVersion.CodeName
                    ' do nothing
                Case Else
                    ws.Delete
            End Select
        Next ws
    
        GoTo Err_QuitApp
    
    End If

    ' get current username
    'currentUser = Environ$("USERNAME")

Exit_Workbook_Open:
    Call settings(True)
    Exit Sub

Err_Workbook_Open:
    MsgBox "Error Number: " & Err.Number & chr(12) & "Error Description: " & Err.Description
    Resume Exit_Workbook_Open:

Err_QuitApp:
    Call settings(True)
    Application.Quit
    ThisWorkbook.Close SaveChanges:=True
    Resume Exit_Workbook_Open

End Sub