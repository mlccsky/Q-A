Attribute VB_Name = "mAutoQuote"
Option Explicit
Dim Cell As Range

' ---------- PORTFOLIO STATUS ----------

Sub PFStatus_Yes()

Call portfolioStatus(True)
Call updateInvoiceWording_Address

End Sub

Sub PFStatus_No()

    Call portfolioStatus(False)
    Call updateInvoiceWording_Address

End Sub

Public Sub portfolioStatus(bStatus As Boolean)

    ' clear all addresses in portfolio
    Call clearPortfolioData
    
    Call settings(False)
    
    If bStatus Then ' if true
        
        
        ' if PORTFOLIO, remove formula
        wsQuote.Range("PFAddress_01") = "Edit_Address"
        wsQuote.Range("PFAddress_01_Postcode") = ""
        wsQuote.Range("PFAddress_01_MID") = ""
        wsQuote.Range("PFName") = ""
        wsQuote.Range("PFAddress_01_VASFile") = ""
        wsLists.Range("zzPFStatus") = True
    
        ' unhide Portfolio Name row and hide single address row
        'wsQuote.Range("zzSglPropRowRange").Select
        'Selection.EntireRow.Hidden = True
        wsQuote.Range("zzPFRowRange").Select
        Selection.EntireRow.Hidden = False
    
        ' format button
        wsQuote.Shapes.Range("btn_PFStatus_True").Select
        Selection.Font.ColorIndex = 5
        Selection.Font.FontStyle = "Book Bold"
    
        wsQuote.Shapes.Range("btn_PFStatus_False").Select
        Selection.Font.ColorIndex = 1
        Selection.Font.FontStyle = "Book"
    
    
    Else
        wsLists.Range("zzPFStatus") = False
    
        ' hide Portfolio Name row and unhide single address row
        wsQuote.Activate
        wsQuote.Range("zzPFRowRange").Select
        Selection.EntireRow.Hidden = True
        
        ' format button
        wsQuote.Shapes.Range("btn_PFStatus_False").Select
        Selection.Font.ColorIndex = 5
        Selection.Font.FontStyle = "Book Bold"
        
        wsQuote.Shapes.Range("btn_PFStatus_True").Select
        Selection.Font.ColorIndex = 1
        Selection.Font.FontStyle = "Book"
        
    
    End If
    wsQuote.Range("PFAddress_01").Select
    
    Call settings(True)

End Sub

Private Sub clearPortfolioData()
Dim i As Integer
    
    Call settings(False)
    wsQuote.Range("PF_PropertyData").ClearContents
    With wsQuote
    
    'Portfolio
        '.Range("PFName") = "=PropertyAddress_01b"
        .Activate
        '.Range("zzSglPropRowRange").Select
        'Selection.EntireRow.Hidden = False
        wsQuote.Range("zzPFRowRange").Select
        Selection.EntireRow.Hidden = False
    
    'Property Address/es
        '.Range("PropertyAddress_01a") = "Edit_Address"
        '.Range("PropertyPostCode_01a") = ""
        '.Range("PropertyVASFile_01a") = ""
        '.Range("PropertyMID_01a") = ""
        
        '.Range("PropertyAddress_01b") = ""
        '.Range("PropertyPostCode_01b") = ""
        '.Range("PropertyVASFile_01b") = ""
        '.Range("PropertyMID_01b") = ""
        
        For i = 1 To 20
            If i < 10 Then
                .Range("PFAddress_0" & i) = "Edit_Address"
                .Range("PFAddress_0" & i & "_Postcode") = ""
                .Range("PFAddress_0" & i & "_MID") = ""
                .Range("PFAddress_0" & i & "_VASFile") = ""
            Else
                .Range("PFAddress_" & i) = "Edit_Address"
                .Range("PFAddress_" & i & "_Postcode") = ""
                .Range("PFAddress_" & i & "_MID") = ""
                .Range("PFAddress_" & i & "_VASFile") = ""
            End If
        Next i
    
    End With

    Call settings(True)



End Sub


' ---------- CLIENT ----------
Sub Client_ANZ()

Call settings(False)

    wsQuote.Range("Client") = "ANZ"

Call settings(True)

End Sub

Sub Client_Bendigo()

Call settings(False)

    wsQuote.Range("Client") = "Bendigo"

Call settings(True)


End Sub

Sub Client_CBA()

Call settings(False)

    wsQuote.Range("Client") = "CBA"

Call settings(True)


End Sub

Sub Client_NAB()

Call settings(False)

    wsQuote.Range("Client") = "NAB"

Call settings(True)


End Sub

Sub Client_Suncorp()

Call settings(False)
    
    wsQuote.Range("Client") = "Suncorp"

Call settings(True)


End Sub

Sub Client_Westpac()

Call settings(False)

    wsQuote.Range("Client") = "Westpac"

Call settings(True)


End Sub

Sub Client_Other()

Call settings(False)

    wsQuote.Range("Client") = "Other"

Call settings(True)


End Sub



Sub AssetClass_ScheduleOfCondition()

Call settings(False)

Call setAssetClassTypeStatus_Single("btn_BC_2", True)
Call BC_Purpose_SetUpDropdown

Call settings(True)

End Sub

Sub AssetClass_ScheduleOfMakeGood()

Call settings(False)

Call setAssetClassTypeStatus_Single("btn_BC_3", True)
Call setAssetClassTypeStatus_Single("btn_BC_4", False)

If wsLists.Range("btn_BC_status_3") = True Then
    ' enable selection Interim/Terminal/Final for Make Good Schedule
    wsQuote.cboMakeGood_ScopeOfService.Visible = True
    wsQuote.Range("cboMakeGood_ScopeOfService_Label") = "Scope of Service"
Else
    wsQuote.cboMakeGood_ScopeOfService = ""
    wsQuote.cboMakeGood_ScopeOfService.Visible = False
    wsQuote.Range("cboMakeGood_ScopeOfService_Label") = ""
End If

Call BC_Purpose_SetUpDropdown

Call settings(True)

End Sub

Sub AssetClass_TDD()

Call settings(False)

Call setAssetClassTypeStatus_Single("btn_BC_5", True)
Call BC_Purpose_SetUpDropdown

Call settings(True)

End Sub
Sub AssetClass_BCother()

Call settings(False)
'Call populateBuildingConsult_AssetClass("BC Other", iHigh)

Call setAssetClassTypeStatus_Single("btn_BC_6", True)
Call BC_Purpose_SetUpDropdown

Call settings(True)

End Sub


Sub AssetClassCC_Acq_ReinstCostAssess()

Call settings(False)

Call setAssetClassTypeStatus_Single("btn_CC_1", True)

Call settings(True)

End Sub

Sub AssetClassCC_CC_Other()

Call settings(False)

Call setAssetClassTypeStatus_Single("btn_CC_5", True)


Call settings(True)

End Sub

Sub AssetClassCC_CostPlanning()

Call settings(False)

Call setAssetClassTypeStatus_Single("btn_CC_3", True)

Call settings(True)

End Sub
Sub AssetClassCC_IRQSVerifyCC()

Call settings(False)

Call setAssetClassTypeStatus_Single("btn_CC_2", True)

Call settings(True)

End Sub
Sub AssetClassCC_InsuranceReinstateCostAssessment()

Call settings(False)

Call setAssetClassTypeStatus_Single("btn_CC_6", True)

Call settings(True)

End Sub

Sub AssetClassCC_BCLifeCycleCosting()

Call settings(False)

Call setAssetClassTypeStatus_Single("btn_BC_1", True)
Call BC_Purpose_SetUpDropdown

Call settings(True)

End Sub
Sub AssetClassCC_ProgressClaim()

Call settings(False)

Call setAssetClassTypeStatus_Single("btn_CC_4", True)

Call settings(True)

End Sub

Sub AssetClassTax_AquiAssessment()

Call settings(False)

'Call populateTax_AssetClass("Acquisition Assessment", iLowMedium)

Call setAssetClassTypeStatus_Single("btn_Tax_1", True)

' determine whether at least one service within Tax has been selected
If AssetClassStatus_Tax = True Then
    ' unhide the discount row
    wsQuote.Range("ClientFeeTotalDiscount").Select
    Selection.EntireRow.Hidden = False
Else
    wsQuote.Range("ClientFeeTotalDiscount").Select
    Selection.EntireRow.Hidden = True
End If


Call settings(True)

End Sub

Sub AssetClassTax_Other()

Call settings(False)

Call setAssetClassTypeStatus_Single("btn_Tax_10", True)

Call settings(True)

End Sub

Sub AssetClassTax_ComplementaryReview()

Call settings(False)

Call setAssetClassTypeStatus_Single("btn_Tax_2", True)

Call settings(True)

End Sub

Sub AssetClassTax_ConstructionAssessmentDepreciation()

Call settings(False)

Call setAssetClassTypeStatus_Single("btn_Tax_3", True)

Call settings(True)

End Sub

Sub AssetClassTax_DepreciatedReplacementCost()

Call settings(False)

Call setAssetClassTypeStatus_Single("btn_Tax_4", True)

Call settings(True)

End Sub

Sub AssetClassTax_FitOutAbdRefurb()

Call settings(False)

Call setAssetClassTypeStatus_Single("btn_Tax_5", True)

Call settings(True)

End Sub

Sub AssetClassTax_FitOutAbd()

Call settings(False)

Call setAssetClassTypeStatus_Single("btn_Tax_6", True)

Call settings(True)

End Sub

Sub AssetClassTax_FixedAssetRegister()

Call settings(False)

Call setAssetClassTypeStatus_Single("btn_Tax_7", True)

Call settings(True)

End Sub

Sub AssetClassTax_IndicativeDepreciationSched()

Call settings(False)

Call setAssetClassTypeStatus_Single("btn_Tax_8", True)

Call settings(True)

End Sub

Sub AssetClassTax_RefurbExtAssessmentDepreciation()

Call settings(False)

Call setAssetClassTypeStatus_Single("btn_Tax_9", True)

Call settings(True)

End Sub

Sub riskLevel(iLevel As Integer, sRange As String)

    ' determine level of risk
    Select Case iLevel
    Case iLow
        wsQuote.Range(sRange).Value = "Low Risk"
    Case iLowMedium
        wsQuote.Range(sRange).Value = "Low to Medium Risk"
    Case iMedium
        wsQuote.Range(sRange).Value = "Medium Risk"
    Case iMediumHigh
        wsQuote.Range(sRange).Value = "Medium to High Risk"
    Case iHigh
        wsQuote.Range(sRange).Value = "High Risk"
    End Select


End Sub


Sub VPStatus_No()

Call settings(False)

    wsQuote.Range("VPStatus") = "No"

Call settings(True)

End Sub

Sub VPStatus_Unknown()

Call settings(False)

    If wsQuote.Range("zzPFStatus") = True Then
        wsQuote.Range("VPStatus") = "Unknown/PF"
    Else
        wsQuote.Range("VPStatus") = "Unknown"
    End If

Call settings(True)

End Sub

' ---------- TENANTS QTY ----------

Sub TenantQty_1()

Call settings(False)

    wsQuote.Range("TenantQty") = 1

Call settings(True)

End Sub

Sub TenantQty_2()

Call settings(False)

    wsQuote.Range("TenantQty") = 2

Call settings(True)

End Sub

Sub TenantQty_3()

Call settings(False)

    wsQuote.Range("TenantQty") = 3

Call settings(True)

End Sub

Sub TenantQty_4()

Call settings(False)

    wsQuote.Range("TenantQty") = 4

Call settings(True)

End Sub

Sub TenantQty_5Plus()

Call settings(False)

    wsQuote.Range("TenantQty") = "5+"

Call settings(True)

End Sub

Sub TenantQty_Unknown()

Call settings(False)

    wsQuote.Range("TenantQty") = "Unknown"

Call settings(True)

End Sub

' ---------- RESET AUTO-QUOTE SHEET ----------

Sub AutoQuote_Reset_Worksheet()
Dim i As Integer
Dim rCell As Range

'Msgbox warning
Dim Msg, Style, Title, Help, Ctxt, Response, myString
Msg = "This will clear the entire worksheet." '& vbCrLf & "Do you want to proceed?"
Style = vbOKCancel
Title = "Reset Inputs"
Ctxt = 1000
Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    
    If Response = vbOK Then
    
        Call settings(False)
    
        'Reset input section
        With wsQuote
            Application.StatusBar = True
            Application.StatusBar = "Clearing Quote worksheet..."
        'Portfolio
            .Range("PFName") = ""
            .Activate
            'wsQuote.Range("zzSglPropRowRange").Select
            'Selection.EntireRow.Hidden = False
            'wsQuote.Range("zzPFRowRange").EntireRow.Hidden = False
            wsQuote.Range("zzPFRowRange").Select
            Selection.EntireRow.Hidden = False
            Application.EnableEvents = False
        
            .cboPrimaryOperator = ""
            .Range("PrimaryOperator") = ""
            
            .cboLOESignatory = ""
            .Range("LOESignatory") = ""
            
            .cboCompanyContact = ""
            
            .Range("SelectedCompany") = ""
            
            
        ' Client Details
            '.Range("ClientName") = "=IFNA(VLOOKUP(SelectedCompany,Lists!$A$72:$H$88,2,FALSE)," & chr(34) & chr(34) & ")"
            '.Range("ClientCompany") = "=IFNA(VLOOKUP(SelectedCompany,Lists!$A$72:$H$88,1,FALSE)," & chr(34) & chr(34) & ")"
            '.Range("ClientStreet") = "=IFNA(VLOOKUP(SelectedCompany,Lists!$A$72:$H$88,3,FALSE)," & chr(34) & chr(34) & ")"
            '.Range("ClientSuburb") = "=IFNA(VLOOKUP(SelectedCompany,Lists!$A$72:$H$88,4,FALSE)," & chr(34) & chr(34) & ")"
            '.Range("ClientState") = "=IFNA(VLOOKUP(SelectedCompany,Lists!$A$72:$H$88,5,FALSE)," & chr(34) & chr(34) & ")"
            '.Range("ClientPostcode") = "=IFNA(VLOOKUP(SelectedCompany,Lists!$A$72:$H$88,6,FALSE)," & chr(34) & chr(34) & ")"
            '.Range("ClientEmailAddress") = "=IFNA(VLOOKUP(SelectedCompany,Lists!$A$72:$H$88,8,FALSE)," & chr(34) & chr(34) & ")"
            '.Range("ClientPhone") = "=IFNA(VLOOKUP(SelectedCompany,Lists!$A$72:$H$88,7,FALSE)," & chr(34) & chr(34) & ")"
            '.Range("InvoiceWording") = ""
            
            .Range("ClientID") = ""
            .Range("ClientName") = ""
            .Range("ClientCompany") = ""
            .Range("ClientAddressLine1") = ""
            .Range("ClientAddressLine2") = ""
            .Range("ClientSuburb") = ""
            .Range("ClientState") = ""
            .Range("ClientPostcode") = ""
            .Range("ClientEmailAddress") = ""
            .Range("ClientPhone") = ""
            .Range("InvoiceWording_Address").Formula = "=PFAddress_01 & " + chr(34) + " " + chr(34) + " & PFAddress_01_Postcode"
            .Range("InvoiceWording") = ""
            
        ' Fees
            .Range("feeSchedOfCondition") = ""
            .Range("feeAdditional") = ""
            
            .Range("btn_BC_Fee_1").Formula = "=SUM(btn_BC_column_feesRange_1)"
            .Range("btn_BC_Fee_2").Formula = "=SUM(btn_BC_column_feesRange_2)"
            .Range("btn_BC_Fee_3").Formula = "=SUM(btn_BC_column_feesRange_3)"
            .Range("btn_BC_Fee_4").Formula = "=SUM(btn_BC_column_feesRange_4)"
            .Range("btn_BC_Fee_5").Formula = "=SUM(btn_BC_column_feesRange_5)"
            .Range("btn_BC_Fee_6").Formula = "=SUM(btn_BC_column_feesRange_6)"
            
            .Range("btn_CC_Fee_1").Formula = "=SUM(btn_CC_column_feesRange_1)"
            .Range("btn_CC_Fee_2").Formula = "=SUM(btn_CC_column_feesRange_2)"
            .Range("btn_CC_Fee_3").Formula = "=SUM(btn_CC_column_feesRange_3)"
            .Range("btn_CC_Fee_4").Formula = "=SUM(btn_CC_column_feesRange_4)"
            .Range("btn_CC_Fee_5").Formula = "=SUM(btn_CC_column_feesRange_5)"
            
            .Range("btn_Tax_Fee_1").Formula = "=SUM(btn_Tax_column_feesRange_1)"
            .Range("btn_Tax_Fee_2").Formula = "=SUM(btn_Tax_column_feesRange_2)"
            .Range("btn_Tax_Fee_3").Formula = "=SUM(btn_Tax_column_feesRange_3)"
            .Range("btn_Tax_Fee_4").Formula = "=SUM(btn_Tax_column_feesRange_4)"
            .Range("btn_Tax_Fee_5").Formula = "=SUM(btn_Tax_column_feesRange_5)"
            .Range("btn_Tax_Fee_6").Formula = "=SUM(btn_Tax_column_feesRange_6)"
            .Range("btn_Tax_Fee_7").Formula = "=SUM(btn_Tax_column_feesRange_7)"
            .Range("btn_Tax_Fee_8").Formula = "=SUM(btn_Tax_column_feesRange_8)"
            .Range("btn_Tax_Fee_9").Formula = "=SUM(btn_Tax_column_feesRange_9)"
            .Range("btn_Tax_Fee_10").Formula = "=SUM(btn_Tax_column_feesRange_10)"
            
            .Range("ClientFeeTotal").Formula = "=SUM(totalFeeRange,ClientFeeTotalDiscount)"
            .Range("SCDisbursementFeeTotal ").Formula = "=SUM(subConsultantDisbursementFeeRange)"
            .Range("ClientFeeTotalDiscountPerc") = 0
            
            
        ' Emails - New Invoice
            
            .Range("EmailSubjectLine_NewJob") = "=" & chr(34) & "New job to be created | " & chr(34) & "& IF(zzPFStatus=True" & ",PFName, PFAddress_01)"
            .Range("EmailSubjectLine_FinalInvoice") = "=" & chr(34) & "Final Invoice | " & chr(34) & "& IF(zzPFStatus=True" & ",PFName, PFAddress_01)"
        
        End With
        
        wsQuote.cboPrimaryOperator = ""

        Call setCountry("Australia")
        
        Application.StatusBar = "Resetting Asset Class Type Status..."
        ' reset all the buttons
        Call setResetAssetClassTypeStatus_All(True)
        
        ' hide BC combo boxes
        Call BC_Purpose_SetUpDropdown
        
        ' clear buttons selected texts
        wsQuote.Range("AssetClass_BC_Selected") = ""
        wsQuote.Range("AssetClass_CC_Selected") = ""
        wsQuote.Range("AssetClass_Tax_Selected") = ""
        
        wsQuote.cboMakeGood_ScopeOfService.Visible = False
        wsQuote.Range("cboMakeGood_ScopeOfService_Label") = ""
        
        wsQuote.Range("AllocationRange_Percentage").ClearContents
        wsQuote.Range("subConsultantDisbursementFeeRange").ClearContents
        wsQuote.Range("SCAllocationFeeRange").ClearContents
        
        Application.StatusBar = "Clearing disbursements..."
        Call clearDisbursements(wsDisbursements.Name)
        Call clearDisbursements(wsSubConsultants.Name)
        
        ' reset all the subconsultant roles to FALSE
        For Each rCell In wsLists.Range("zzList_SubcontractorRoles")
            ' reset all Roles to FALSE
            rCell.Offset(0, 1) = False
        Next rCell
        
        Application.StatusBar = "Resetting Portfolio Status..."
        Call portfolioStatus(False)
        
        Application.StatusBar = "Resetting Allocation formulas..."
        Call resetAllocationFormulas
        
        ' navigate back to the autoquote ws
        wsQuote.Activate
        wsQuote.Range("A1").Select
        
        ' hide Asset and Scope description ws
        wsAssetDesc.Visible = xlSheetHidden
        wsScopeDesc.Visible = xlSheetHidden
        wsAttachment2.Visible = xlSheetHidden
        wsMakeGood.Visible = xlSheetHidden
        wsSchedOfCondition.Visible = xlSheetHidden
        wsScopeOfService.Visible = xlSheetHidden
        
        
        Application.StatusBar = False
        
        MsgBox "Worksheet has been cleared and reset.", vbInformation
    
    End If

Call settings(True)

End Sub

Sub currentUserEmail()
'Macro requires currentUserEmailAddress Function (below)

Dim outApp As Object, outSession As Object
Dim currentUserEmailAddress As String
Dim sh As Worksheet

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False

    Set sh = wsQuote

    'Create an Outlook session
    Set outApp = CreateObject("Outlook.Application")

    'Check if session is created
    If outApp Is Nothing Then
        currentUserEmailAddress = "Cannot create Microsoft Outlook session."
        currentUserEmailAddress = "Not found"
        Exit Sub
    End If

'''    'Set object variable with .Session property to access existing Outlook items and get current username
'''    Set outSession = outApp.Session.CurrentUser
'''
'''    'Get current user email address
'''    'currentUserEmailAddress = outSession.AddressEntry.GetExchangeUser().PrimarySmtpAddress
'''    currentUserEmailAddress = "commercialvaluations@cbre.com.au"
'''
'''    sh.Range("AutoQuoteOperator") = currentUserEmailAddress

    Set outApp = Nothing

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True

End Sub

Function currentUserEmailAddress() As String
'Function required for currentUserEmail macro (above)

Dim outApp As Object, outSession As Object

    'Create an Outlook session
    Set outApp = CreateObject("Outlook.Application")

    'Check if session is created
    If outApp Is Nothing Then
        currentUserEmailAddress = "Cannot create Microsoft Outlook session."
        currentUserEmailAddress = "Not found"
        Exit Function
    End If

'''    'Set object variable with .Session property to access existing Outlook items and get current username
'''    Set outSession = outApp.Session.CurrentUser
'''
'''    'Get current user email address
'''    'currentUserEmailAddress = outSession.AddressEntry.GetExchangeUser().PrimarySmtpAddress
'''    currentUserEmailAddress = "commercialvaluations@cbre.com.au"
    
    Set outApp = Nothing

End Function

Public Function AssetClassStatus_BC() As Boolean
Dim rCell As Range

AssetClassStatus_BC = False

For Each rCell In wsLists.Range("AssetClass_BC")
    If rCell = True Then
        AssetClassStatus_BC = True
        Exit For
    End If

Next rCell

End Function

Public Function AssetClassStatus_CC() As Boolean
Dim rCell As Range

AssetClassStatus_CC = False

For Each rCell In wsLists.Range("AssetClass_CC")
    If rCell = True Then
        AssetClassStatus_CC = True
        Exit For
    End If

Next rCell


End Function
Public Function AssetClassStatus_Tax() As Boolean
Dim rCell As Range

AssetClassStatus_Tax = False

For Each rCell In wsLists.Range("AssetClass_Tax")
    If rCell = True Then
        AssetClassStatus_Tax = True
        Exit For
    End If

Next rCell

End Function

Public Sub resetSubconsultantRoles()

Dim iLastRow As Long
Dim rSortCell_K1 As Range
Dim rSortCell_K2 As Range
Dim i As Integer
Dim ws As Worksheet
Dim rCell As Range

    Call settings(False)
    
    Set ws = Worksheets(wsSubConsultants.Name)

    ' navigate to the start row
    ws.Range(iStartCell).Select
    
    iLastRow = lastRow(ws.Name, sortCell_K1)
    
    If iLastRow > 500000 Then
        iLastRow = 120
        ws.Range("A1").Select
        
        ' no subconsultant fees so set all flags subcontractor roles to false
        For Each rCell In wsLists.Range("zzList_SubcontractorRoles")
            rCell.Offset(0, 1) = False
        Next rCell
    Else
        
        ' determine all the subconsultant roles in the current job
        For Each rCell In wsLists.Range("zzList_SubcontractorRoles")
            For i = 9 To iLastRow
                If ws.Range(colSubconsultantRow & i) = rCell Then
                    rCell.Offset(0, 1) = True
                    Exit For
                Else
                    rCell.Offset(0, 1) = False
                End If
            Next i
        Next rCell
    End If
    
    ' amend formulas in the wsLists worksheet that keeps track of the fees
    For Each rCell In wsLists.Range("zzList_SubcontractorRoles")
        rCell.Offset(0, 2).Formula = "=SUMIF(SubConsultants!$C$9:$C$" & iLastRow & ",Lists!N" & rCell.row & ",SubConsultants!$D$9:$D$" & iLastRow & ")"
    Next rCell
    
    Call settings(True)
    
    Set ws = Nothing
        
End Sub

Public Sub fillCompanyCombo()
Dim rCell As Range


wsQuote.Activate
gsClientCompany = wsQuote.cboClients
wsQuote.cboClients.Clear

On Error Resume Next
For Each rCell In wsUniqueList.Range("zzList_CompanyUnique")
    wsQuote.cboClients.AddItem rCell
Next rCell

If gsClientCompany <> "" Then
    wsQuote.cboClients = gsClientCompany
End If


End Sub
