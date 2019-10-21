Attribute VB_Name = "mUtil"
Option Explicit

Public Function lastRow(sSheet As String, sCellAdd As String) As Long
On Error GoTo Err_lastRow

Dim LastCellAddress As String
Dim cPos As Integer
Dim rPos As Integer
Dim iLen As Integer
Dim col As String
Dim row As String
Dim i As Integer
Dim aCellAdd As Variant

Sheets(sSheet).Select
Range(sCellAdd).Select

Selection.End(xlDown).Select

LastCellAddress = ActiveCell.Cells.Address

lastRow = Split(Selection.Address, "$")(2)

Exit_lastRow:

    Exit Function

Err_lastRow:
    Err.Raise Err.Number, "lastRow", Err.Description
    Resume Exit_lastRow

End Function

Public Sub settings(bEnabled As Boolean)

If bEnabled Then
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

Else
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

End If

End Sub
Public Sub unhideWS()

Dim ws As Worksheet

    For Each ws In ActiveWorkbook.Worksheets
        ws.Visible = xlSheetVisible
    Next ws

End Sub
Public Sub saveFile()
On Error GoTo Err_saveFile
Dim sTempFile As String
Dim sAddress As String
Dim i As Integer
    
'    Call settings(False)
    ' check the validity of the directory
    If checkFilePathExist(wsQuote.Range("zzListFilePath")) = True Then
    
        ' check that the last character is a back slash
        If Right(wsQuote.Range("zzListFilePath"), 1) <> "\" Then
            wsQuote.Range("zzListFilePath") = wsQuote.Range("zzListFilePath") + "\"
        End If
        
        ' determine whether it is a portfolio or not
        If wsLists.Range("zzPFStatus") = True Then
            sAddress = wsQuote.Range("PFName")
        Else
            sAddress = wsQuote.Range("PFAddress_01")
        End If
        
        ' check whether an address/portfolio exists
        If Len(wsQuote.Range("PFAddress_01")) > 0 Then
            sAddress = cleanAddress(sAddress)
            sTempFile = wsQuote.Range("zzListFilePath") & sAddress & "_" & Format(Now, "yyyymmdd") & ".xlsb"
        Else
            sTempFile = wsQuote.Range("zzListFilePath") & "Q&A_" & Format(Now, "yyyymmdd_hhmm") & ".xlsb"
        End If
        
        ' check if the file exists. if it does, add/increment the version
        Application.DisplayAlerts = False
        i = 1
        
        ActiveWorkbook.SaveAs sTempFile, AccessMode:=xlExclusive, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
        'ActiveWorkbook.SaveAs sTempFile
        Application.DisplayAlerts = True
    
    Else
        MsgBox "The file path location for the reports is invalid. Please select a valid path.", vbExclamation
    End If

Exit_saveFile:
'    Call settings(True)
    Exit Sub

Save_Default:
    ' save in user's local default location
    wsQuote.Range("zzListFilePath") = Application.DefaultFilePath & "\"
    sTempFile = wsQuote.Range("zzListFilePath") & "Q&A_" & wsQuote.Range("PropertyMID_01a") & "_" & Format(Now, "yyyymmdd") & ".xlsb"
    
    ' check if file exists
    If Len(Dir(sTempFile, vbNormal)) > 0 Then
        ActiveWorkbook.Save
    Else
        ActiveWorkbook.SaveAs sTempFile
    End If
    
    Resume Exit_saveFile

Err_saveFile:
    
    Select Case Err.Number
    Case 52
        Debug.Print Err.Number & ": " & Err.Description
        'MsgBox "You currently don't have access to drive " & wsQuote.Range("zzListFilePath") & ". The file will be save in your local Documents folder.", vbInformation
        GoTo Save_Default
    'Case 1004
    '    Debug.Print Err.Number & ": " & Err.Description
        'MsgBox "You currently don't have access to drive " & wsQuote.Range("zzListFilePath") & ". The file will be save in your local Documents folder.", vbInformation
    '    GoTo Save_Default
    Case -2147024891
        Debug.Print Err.Number & ": " & Err.Description
        'MsgBox "You currently don't have access to drive " & wsQuote.Range("zzListFilePath") & ". The file will be save in your local Documents folder.", vbInformation
        GoTo Save_Default
    Case 80070005
        Debug.Print Err.Number & ": " & Err.Description
        'MsgBox "You currently don't have access to drive " & wsQuote.Range("zzListFilePath") & ". The file will be save in your local Documents folder.", vbInformation
        GoTo Save_Default
    Case Else
        MsgBox "Error Number: " & Err.Number & chr(13) & "Procedure: saveFile" & chr(13) & "Description: " & Err.Description
    End Select
    
    Resume Exit_saveFile

End Sub
Public Sub selectUserDir()
Dim sFolderName As String


    With Application.FileDialog(msoFileDialogFolderPicker)
        
        If Len(wsQuote.Range("zzListFilePath")) > 0 Then .InitialFileName = wsQuote.Range("zzListFilePath")
            
        .AllowMultiSelect = False
        .Title = "File Path selected is invalid. Please select a valid path."
        .Show
        
        On Error Resume Next
        sFolderName = .SelectedItems(1)
        Err.Clear
        On Error GoTo 0
    End With
    
    If sFolderName <> "" And Not IsNull(sFolderName) Then
        wsQuote.Range("zzListFilePath").Value = sFolderName & "\"
    Else
        MsgBox "Please select a valid folder"
    End If

End Sub

Public Sub getFolder(sWorksheet As String, sFolderRangeName As String)
Dim ws As Worksheet
Dim sFolderName As String

    Set ws = Worksheets(sWorksheet)

    With Application.FileDialog(msoFileDialogFolderPicker)
        
        If Len(ws.Range(sFolderRangeName)) > 0 Then
            .InitialFileName = ws.Range(sFolderRangeName)
        End If
            
        .AllowMultiSelect = False
        .Show
        
        On Error Resume Next
        sFolderName = .SelectedItems(1) & "\"
        Err.Clear
        On Error GoTo 0
    End With
    
    If sFolderName <> "" And Not IsNull(sFolderName) Then
        ws.Range(sFolderRangeName).Value = sFolderName
    Else
        MsgBox "Please select a valid folder"
    End If
    
End Sub

Public Sub getReportsFolder()

Call getFolder(wsQuote.Name, "zzListFilePath")

End Sub

Public Function checkFilePathExist(sPath As String) As Boolean
    
    If Dir(sPath, vbDirectory) <> "" Then
        checkFilePathExist = True
    Else
        checkFilePathExist = False
    End If



End Function

Public Function removeInvalidChars(sTemp As String) As String
Dim tempString As String
Dim chr As Variant

tempString = sTemp

For Each chr In Split(SpecialCharacters, "|")
'For Each chr In SpecialCharacters
    tempString = Replace(tempString, chr, "_")
Next

removeInvalidChars = tempString
    

End Function

Public Function cleanAddress(sAddress As String) As String
Dim tempStr As String

'tempStr = removeInvalidChars("1/29 Pleasant Street, QLD, 4069(Australia)")

tempStr = removeInvalidChars(sAddress)

' remove double "_"
tempStr = Replace(tempStr, "__", "_")

'remove any "_" from front and end of string
If Right(tempStr, 1) = "_" Then
    tempStr = Left(tempStr, Len(tempStr) - 1)
End If

If Left(tempStr, 1) = "_" Then
    tempStr = Right(tempStr, Len(tempStr) - 1)
End If

cleanAddress = tempStr

'Debug.Print tempStr

End Function
Public Sub resetComboBoxPhysicalAttributes(sWrkSht As String, sComboName As String, dHeight As Double, dWidth As Double, dTop As Double)
Dim ws As Worksheet
Dim sCombobox As Shape
Dim oCombo As Object

Set ws = Worksheets(sWrkSht)
Set sCombobox = ws.Shapes(sComboName)

With oCombo
    
End With

With sCombobox

    .AutoSize = True
End With

With sCombobox
    .AutoSize = False
    .height = dHeight
    .width = dWidth
    .Top = dTop
End With


End Sub
Public Sub ResetButton(ByRef btn As Object)
' Purpose:      Reset button size and font size for form command button on worksheet
'               Addresses known Excel bug(s) which alters button size and/or apparent font size
' Parameters:   Reference to button object
' Remarks:      Getting/setting font size fails since font size remains the same; display (apparent) size changes
'               AutoSize maximizes the font size to fit the current button size in case it has changed
'               Button size is reset in case it has changed
'               Finally, font size is reset to adjust for font changes applied by AutoSize
'               This fix seems to handle shrinking button icon sizes as well
Dim h As Integer    'command button height
Dim w As Integer    '               width
Dim fs As Integer   '               font size
    On Error Resume Next
    With btn
        h = .height             'capture original values
        w = .width
        fs = .Font.Size
        .AutoSize = True        'apply maximum font size to fit button
        .AutoSize = False
        .height = h             'reset original button and font sizes
        .width = w
        .Font.Size = fs
    End With
End Sub


Sub test21()
Dim i As Integer
Dim aAssetClasses() As Variant
'
'aAssetClasses = Array("AssetClass_BC", "AssetClass_CC", "AssetClass_Tax")
'
'For i = LBound(aAssetClasses) To UBound(aAssetClasses)
'    Debug.Print aAssetClasses(i)
'Next i

'Call settings(True)

'If Dir("\\ausdcfnp02\Capital_Allowances\2. Templates\16. QA Form\QA Form Generated LOEs\", vbDirectory) > 0 Then
If checkFilePathExist("\\ausdfnp02\Capital_Allowances\2. Templates\16. QA Form\QA Form Generated LOEs\") Then
    MsgBox "Valid"
Else
    MsgBox "Not valid"
End If

'If Dir("C:\QA Form\QA Form Generated LOEs\", vbDirectory) <> "" Then
'    MsgBox "Valid"
'Else
'    MsgBox "Not valid"
'End If

End Sub
