Attribute VB_Name = "mEmail"
Option Explicit
Sub sendEmail_QuoteJob()
Dim aEmailMsgRng() As String
Dim rUnion As Range


    If validateForm Then
        
        If checkFilePathExist(wsQuote.Range("zzListFilePath")) = True Then
    
            ' check if portfolio. if true, need to include all properties in the portfolio
            ReDim aEmailMsgRng(1 To 9, 1 To 2) As String
            ' name ranges representing parts of the email
            aEmailMsgRng(1, 1) = wsQuote.Name
            aEmailMsgRng(1, 2) = "EmailContent_QuoteJob_L1"
            
            aEmailMsgRng(2, 1) = wsLists.Name
            aEmailMsgRng(2, 2) = "keyFields_forDataEntryTeam"
            
            aEmailMsgRng(3, 1) = wsQuote.Name
            aEmailMsgRng(3, 2) = "Client_Details"
            
            aEmailMsgRng(4, 1) = wsQuote.Name
            aEmailMsgRng(4, 2) = "PF_PropertyAddresses_Selected"
            
            aEmailMsgRng(5, 1) = wsQuote.Name
            aEmailMsgRng(5, 2) = "AutoQuote_Fees_PrintArea"
            
            aEmailMsgRng(6, 1) = wsQuote.Name
            aEmailMsgRng(6, 2) = "AutoQuote_Allocations_PrintArea"
            
            aEmailMsgRng(7, 1) = wsQuote.Name
            aEmailMsgRng(7, 2) = "Disbursements_List_PrintArea"
            
            aEmailMsgRng(8, 1) = wsQuote.Name
            aEmailMsgRng(8, 2) = "Subconsultants_List_PrintArea"
            
            aEmailMsgRng(9, 1) = wsQuote.Name
            aEmailMsgRng(9, 2) = "EmailContent_QuoteJob_L2"
                
            Call settings(False)
            
            Call SendEmail("Quote", wsQuote.Name, aEmailMsgRng, "EmailTo_NewQuote", "EmailSubjectLine_NewQuote", "EmailCC_NewQuote")
            
            Call settings(True)
    
        Else
            MsgBox "The file path location for the reports is invalid. Please select a valid path.", vbExclamation
            wsQuote.Range("C3:C4").Select
        End If
        
    End If

End Sub
Sub sendEmail_NewJob()
Dim aEmailMsgRng() As String
Dim rUnion As Range


    If checkFilePathExist(wsQuote.Range("zzListFilePath")) = True Then

        ' check if portfolio. if true, need to include all properties in the portfolio
        ReDim aEmailMsgRng(1 To 9, 1 To 2) As String
        ' name ranges representing parts of the email
        aEmailMsgRng(1, 1) = wsQuote.Name
        aEmailMsgRng(1, 2) = "EmailContent_NewJob_L1"
        
        aEmailMsgRng(2, 1) = wsLists.Name
        aEmailMsgRng(2, 2) = "keyFields_forDataEntryTeam"
        
        aEmailMsgRng(3, 1) = wsQuote.Name
        aEmailMsgRng(3, 2) = "Client_Details"
        
        aEmailMsgRng(4, 1) = wsQuote.Name
        aEmailMsgRng(4, 2) = "PF_PropertyAddresses_Selected"
        
        aEmailMsgRng(5, 1) = wsQuote.Name
        aEmailMsgRng(5, 2) = "AutoQuote_Fees_PrintArea"
        
        aEmailMsgRng(6, 1) = wsQuote.Name
        aEmailMsgRng(6, 2) = "AutoQuote_Allocations_PrintArea"
        
        aEmailMsgRng(7, 1) = wsQuote.Name
        aEmailMsgRng(7, 2) = "Disbursements_List_PrintArea"
        
        aEmailMsgRng(8, 1) = wsQuote.Name
        aEmailMsgRng(8, 2) = "Subconsultants_List_PrintArea"
        
        aEmailMsgRng(9, 1) = wsQuote.Name
        aEmailMsgRng(9, 2) = "EmailContent_NewJob_L2"
            
        Call settings(False)
        
        Call SendEmail("NewJob", wsQuote.Name, aEmailMsgRng, "EmailTo_NewJob", "EmailSubjectLine_NewJob", "EmailCC_NewJob")
        
        Call settings(True)

    Else
        MsgBox "The file path location for the reports is invalid. Please select a valid path.", vbExclamation
        wsQuote.Range("C3:C4").Select
    End If

End Sub
Sub sendEmail_FinalInvoice()
Dim aEmailMsgRng() As String

    If checkFilePathExist(wsQuote.Range("zzListFilePath")) = True Then
        ' check if portfolio. if true, need to include all properties in the portfolio
        ReDim aEmailMsgRng(1 To 6, 1 To 2) As String
        
        aEmailMsgRng(1, 1) = wsQuote.Name
        aEmailMsgRng(1, 2) = "EmailContent_FinalInvoice_L1"
        
        aEmailMsgRng(2, 1) = wsLists.Name
        aEmailMsgRng(2, 2) = "keyFields_forDataEntryTeam"
        
        aEmailMsgRng(3, 1) = wsQuote.Name
        aEmailMsgRng(3, 2) = "Client_Details"
        
        aEmailMsgRng(4, 1) = wsQuote.Name
        aEmailMsgRng(4, 2) = "PF_PropertyAddresses_Selected"
        
        aEmailMsgRng(5, 1) = wsQuote.Name
        aEmailMsgRng(5, 2) = "AutoQuote_Fees_PrintArea"
        
        aEmailMsgRng(6, 1) = wsQuote.Name
        aEmailMsgRng(6, 2) = "AutoQuote_Allocations_PrintArea"
        
        
        
        Call settings(False)
        
        Call SendEmail("FinalInvoice", wsQuote.Name, aEmailMsgRng, "EmailTo_FinalInvoice", "EmailSubjectLine_FinalInvoice", "EmailCC_FinalInvoice")
        
        Call settings(True)
    
    Else
        MsgBox "The file path location for the reports is invalid. Please select a valid path.", vbExclamation
        wsQuote.Range("C3:C4").Select
    End If


End Sub

Public Sub SendEmail(sJobType As String, sWrkShtName As String, sRngEmailMsg() As String, sEmailTo As String, sEmailSubject As String, Optional sEmailCC As String)
'Macro requires RangetoHTML Function (below)

Dim rng As Range
Dim sRange As String
Dim outApp As Object
Dim OutMail As Object
Dim strSig As String
Dim strPath As String
Dim rUnion As Range
Dim vEmailBody As Variant

Dim bNothingToAdd As Boolean

Dim iNoOfDisbursements As Long
Dim iNoOfSubconsultants As Long
Dim i As Integer

Dim sFileName As String

Dim wbMaster As Workbook
Dim wbTemp As Workbook
Dim sTempFile As String

bNothingToAdd = False


    Application.StatusBar = True
    Application.StatusBar = "Generating final invoice email..."
    
    Set wbMaster = ActiveWorkbook
    
    'Email-generation
    vEmailBody = ""
    'Create an Outlook session
    Set outApp = CreateObject("Outlook.Application")
    Set OutMail = outApp.CreateItem(0)
    
    On Error Resume Next
    
    
    'Email details
    With OutMail
        .To = wbMaster.Worksheets(sWrkShtName).Range(sEmailTo)
        
        If sEmailCC <> "" Then ' optional
            .CC = wbMaster.Worksheets(sWrkShtName).Range(sEmailCC)
        End If
        .Subject = wbMaster.Worksheets(sWrkShtName).Range(sEmailSubject)
        
        ' ************************************
        For i = LBound(sRngEmailMsg, 1) To UBound(sRngEmailMsg, 1)
        
            Set rng = Nothing
            On Error Resume Next
            
            Select Case sRngEmailMsg(i, 2)
            Case "Disbursements_List_PrintArea"
                iNoOfDisbursements = lastRow(wsDisbursements.Name, iStartCell)
                
                If iNoOfDisbursements > 500000 Then  ' indicates no disbursements
                    'Exit For
                    bNothingToAdd = True
                Else
                    ' amend the disbursement named range as the number of disbursements will vary with each job
                    ActiveWorkbook.Names("Disbursements_List_PrintArea").RefersToR1C1 = "='" & wsDisbursements.Name & "'!R7C1:R" & iNoOfDisbursements & "C6"
                    Set rng = ActiveWorkbook.Worksheets(wsDisbursements.Name).Range(sRngEmailMsg(i, 2)).SpecialCells(xlCellTypeVisible)
                    bNothingToAdd = False
                End If
            
            Case "Subconsultants_List_PrintArea"
                iNoOfSubconsultants = lastRow(wsSubConsultants.Name, iStartCell)
                
                If iNoOfSubconsultants > 500000 Then  ' indicates no disbursements
                    'Exit For
                    bNothingToAdd = True
                Else
                    ' amend the disbursement named range as the number of disbursements will vary with each job
                    ActiveWorkbook.Names("Subconsultants_List_PrintArea").RefersToR1C1 = "='" & wsSubConsultants.Name & "'!R7C1:R" & iNoOfSubconsultants & "C7"
                    Set rng = ActiveWorkbook.Worksheets(wsSubConsultants.Name).Range(sRngEmailMsg(i, 2)).SpecialCells(xlCellTypeVisible)
                    bNothingToAdd = False
                End If
            
            Case "PF_PropertyAddresses_Selected"
                sRange = getPortfolioPropertiesRange
                Set rng = ActiveWorkbook.Worksheets(wsLists.Name).Range(sRange).SpecialCells(xlCellTypeVisible)
                bNothingToAdd = False

            Case Else
                Set rng = ActiveWorkbook.Worksheets(sRngEmailMsg(i, 1)).Range(sRngEmailMsg(i, 2)).SpecialCells(xlCellTypeVisible)
                On Error GoTo 0
            End Select
        
            ' add table only if there are disbursements and/or subconsultant fees
            If bNothingToAdd = False Then
                vEmailBody = vEmailBody & chr(12) & RangetoHTML(rng)
            End If
            
        Next i
        ' save file before attaching
        Call saveFile
        
        Application.StatusBar = "Inserting tables and text into email..."
        
        .HTMLBody = vEmailBody
        
        Application.StatusBar = "Attaching Q&A workbook..."
        .Attachments.Add wsQuote.Range("zzListFilePath") & ActiveWorkbook.Name
        
        ' attach templates only for quotes
        If sJobType = "Quote" Then
            
            ' get Filename - remove ".xlsb"
            sFileName = Left(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 4)
            
            Application.StatusBar = "Generating quote template..."
            
            ' check for TOE attachments
            If AssetClassStatus_BC Then
                ' prepare the template to attach to the email
                sTempFile = prepareTemplate("TOE_BC", sFileName)
                .Attachments.Add sTempFile
            End If
            
            If AssetClassStatus_CC Then
                sTempFile = prepareTemplate("TOE_CC", sFileName)
                .Attachments.Add sTempFile
            End If
                
            If AssetClassStatus_Tax Then
                sTempFile = prepareTemplate("TOE_Tax", sFileName)
                .Attachments.Add sTempFile
            End If
        End If
        
        .Display
    
    End With
    On Error GoTo 0
    
    Application.StatusBar = False
    
    wsQuote.Activate
    
Set rng = Nothing
Set OutMail = Nothing
Set outApp = Nothing
    
End Sub

Function RangetoHTML(rng As Range)
'Function required for SendEmail macros

Dim fso As Object
Dim ts As Object
Dim TempFile As String
Dim TempWB As Workbook

TempFile = Environ$("temp") & "\" & VBA.Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy range and create new workbook to paste data
    Set TempWB = Workbooks.Add(1)
    rng.Copy
    
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        '.Cells(1).PasteSpecial xlPasteAllExceptBorders, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
'        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With
    
    'Publish sheet to .htm file
    With TempWB.PublishObjects.Add( _
        SourceType:=xlSourceRange, _
        Filename:=TempFile, _
        Sheet:=TempWB.Sheets(1).Name, _
        Source:=TempWB.Sheets(1).UsedRange.Address, _
        HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With
    
    'Read all data from .htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.ReadAll
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
    "align=left x:publishsource=")
    
    'Close TempWB
    TempWB.Close SaveChanges:=False
    
    'Delete .htm file used in this function
    Kill TempFile
    
Set ts = Nothing
Set fso = Nothing
Set TempWB = Nothing

End Function

Sub test2()

Call settings(True)

End Sub
