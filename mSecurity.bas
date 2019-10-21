Attribute VB_Name = "mSecurity"
Option Explicit

Public Sub checkUserDomain(sDomainName As String, Optional sPasswd As String)
On Error GoTo Err_checkUserDomain
Dim X As String
Dim UserName As String
Dim ws As Worksheet

Dim wsNewBlank As String

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    X = sDomainName
    UserName = Environ("USERDNSDOMAIN")

    ' check if user is on the CBRE domain
    ' otherwise remove worksheets
    If InStr(1, UserName, X) = 0 Then
    
        ' add a new worksheet as you need to have at least one worksheet in a workbook
        ' delete the rest of the worksheets
        ActiveWorkbook.Worksheets.Add
        wsNewBlank = ActiveSheet.CodeName
        
        For Each ws In Worksheets
            
            If ws.CodeName <> wsNewBlank Then
                ws.Visible = True
                ws.Activate
                'ActiveSheet.Unprotect Password:=sPasswd
                ws.Delete
            End If
            
        Next ws
    
        Warning.Show
    
        GoTo Err_QuitApp
    
    End If

Exit_checkUserDomain:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

Err_checkUserDomain:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Error Description: " & Err.Description
    Resume Exit_checkUserDomain

Err_QuitApp:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    
    ThisWorkbook.Close SaveChanges:=True
    Resume Exit_checkUserDomain

End Sub

