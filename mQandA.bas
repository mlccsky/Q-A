Attribute VB_Name = "mQandA"

Sub cmdAddDisbursement()
Dim iLastRow As Long
Dim rSortCell_K1 As Range
Dim rSortCell_K2 As Range
Dim i As Integer

    Call settings(False)

    ' navigate to the start row
    wsDisbursements.Range(iStartCell).Select
    
    If insertValues(wsDisbursements.Name, Selection.row + 1) Then
        ' sort list in order of user only if data was successfully inserted
        iLastRow = lastRow(wsDisbursements.Name, sortCell_K1)
        
        If iLastRow > 500000 Then
            iLastRow = 120
            wsDisbursements.Range("A1").Select
        End If
        
        Set rSortCell_K1 = Range(sortCell_K1)
        Set rSortCell_K2 = Range(sortCell_K2)
        
        wsDisbursements.Range(sortCell_K1 & ":F" & iLastRow).Sort rSortCell_K1, xlAscending, rSortCell_K2
        
        ' amend range for total disbursements
        wsDisbursements.Range("TotalDisbursements").Formula = "=SUM(F9:F" & iLastRow & ")"
        
        ' amend range for each surveyor for the Disbursements in Auto Quote worksheet
        For i = AutoQuote_UserStartRow To AutoQuote_UserLastRow
            wsQuote.Range("f" & i).Formula = "=SUMIF(Disbursements!$A$9:$A$" & iLastRow & ",'Quote'!C" & i & ",Disbursements!$C$9:$C$" & iLastRow & ")"
        Next i
        
        ' amend range for primary operator
        wsQuote.Range("Fee_Disbursements_PrimOp").Formula = "=SUMIF(Disbursements!$A$9:$A$" & iLastRow _
        & ",'Quote'!Allocation_PrimOperator" & ",Disbursements!$C$9:$C$" & iLastRow & ")"
        
    End If
    
    Set rSortCell_K1 = Nothing
    Set rSortCell_K2 = Nothing

    Call settings(True)
    
    wsDisbursements.Activate


End Sub

Sub cmdAddSubConsultants()
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
    
    If insertValues(ws.Name, Selection.row + 1) Then
        ' sort list in order of user only if data was successfully inserted
        iLastRow = lastRow(ws.Name, sortCell_K1)
        
        If iLastRow > 500000 Then
            iLastRow = 120
            ws.Range("A1").Select
        End If
        
        Set rSortCell_K1 = Range(sortCell_K1)
        Set rSortCell_K2 = Range(sortCell_K2)
        
        ws.Range(sortCell_K1 & ":G" & iLastRow).Sort rSortCell_K1, xlAscending, rSortCell_K2
        
        ' amend range for total disbursements
        ws.Range("TotalSubConsultants").Formula = "=SUM(G9:G" & iLastRow & ")"
        
        ' amend range for scope item (i.e. TDD, Sched Make Good, FARR, Progress Claim etc..) in Auto Quote worksheet
        For i = AutoQuote_ScopeStartRow To AutoQuote_ScopeLastRow
            wsQuote.Range("F" & i).Formula = "=SUMIF(SubConsultants!$B$9:$B$" & iLastRow & ",'Quote'!C" & i & ",SubConsultants!$D$9:$D$" & iLastRow & ")"
        Next i
        
    End If
    
    Set rSortCell_K1 = Nothing
    Set rSortCell_K2 = Nothing

    Call resetSubconsultantRoles

    Call settings(True)
    
    ws.Activate


End Sub
Function insertValues(sWorksheet As String, iRow As Long) As Boolean
Dim bAdd As Boolean
Dim ws As Worksheet
Dim rCell As Range

Set ws = Worksheets(sWorksheet)
    
    
    ' Report/Disbursements_Type
    If ws.Range(ws.Name & "_Type") = "" Then
        MsgBox "Please select a valid type and try again.", vbCritical, "Report/Disbursement Type selection error"
        insertValues = False
        Exit Function
    End If
    
    ' Cost Excluding GST
    If ws.Range(ws.Name & "_CostExGST") = 0 Or ws.Range(ws.Name & "_CostExGST") = "" Or Not IsNumeric(ws.Range(ws.Name & "_CostExGST")) Then
        MsgBox "Please enter a valid numerical value and try again.", vbCritical, "Cost input error"
        insertValues = False
        Exit Function
    End If
    
    ' Inclusive or Exclusive Of Fee
    If ws.Range(ws.Name & "_IncExcOfFee") = "" Then
        MsgBox "Please select whether the fee is inclusive or exclusive of the fee.", vbCritical, "Inclusive or Exclusive of Fee selection error"
        insertValues = False
        Exit Function
    End If
    
    ' Taxable
    If ws.Range(ws.Name & "_Tax") = "" Then
        MsgBox "Please select whether the fee is taxable.", vbCritical, "Fee Tax selection error"
        insertValues = False
        Exit Function
    End If
    
'    If bAdd = False Then
'        MsgBox "Please validate your data and try again.", vbCritical, "Error"
'    Else
        
        ' insert new row
    Selection.EntireRow.Offset(1, 0).Insert Shift:=xlDown ', CopyOrigin:=xlFormatFromLeftOrAbove
        
    Select Case ws.Name
    Case "Disbursements"
        ' insert default values in new row
        ws.Range("A" & iRow).Value = ws.Range(ws.Name & "_User")
        ws.Range("B" & iRow).Value = ws.Range(ws.Name & "_Type")
        ws.Range("C" & iRow).Value = ws.Range(ws.Name & "_CostExGST")
        ws.Range("D" & iRow).Value = ws.Range(ws.Name & "_IncExcOfFee")
        ws.Range("E" & iRow).Value = ws.Range(ws.Name & "_Tax")
        ws.Range("F" & iRow).Value = ws.Range(ws.Name & "_TotalCost")
        
        ' insert default values
        ws.Range(ws.Name & "_Type") = ""
        ws.Range(ws.Name & "_CostExGST") = 0
        ws.Range(ws.Name & "_IncExcOfFee") = "Inclusive"
        ws.Range(ws.Name & "_Tax") = "Y"
        
        insertValues = True
        
    Case "SubConsultants"
        ws.Range("A" & iRow).Value = ws.Range(ws.Name & "_User")
        ws.Range("B" & iRow).Value = ws.Range(ws.Name & "_Type")
        
        ws.Range("C" & iRow).Value = ws.Range(ws.Name & "_Role")
        
        ws.Range("D" & iRow).Value = ws.Range(ws.Name & "_CostExGST")
        ws.Range("E" & iRow).Value = ws.Range(ws.Name & "_IncExcOfFee")
        ws.Range("F" & iRow).Value = ws.Range(ws.Name & "_Tax")
        ws.Range("G" & iRow).Value = ws.Range(ws.Name & "_TotalCost")
        
        ' insert default values
        ws.Range(ws.Name & "_Type") = ""
        ws.Range(ws.Name & "_Role") = ""
        ws.Range(ws.Name & "_CostExGST") = 0
        ws.Range(ws.Name & "_IncExcOfFee") = "Inclusive"
        ws.Range(ws.Name & "_Tax") = "Y"
        
        insertValues = True
    
    End Select
        
    ws.Range(ws.Name & "_User").Select
        
        
Set ws = Nothing

End Function
Sub UpdateUserAllocationList()

Dim i As Integer
Dim iNoOfUsers As Integer
Dim sPrimaryUser As String
Dim iAllocationStartRow As Integer
Dim iAllocationRow As Integer
Dim iUserCounter As Integer
Dim rCell As Range

iNoOfUsers = wsLists.Range("zzList_Surveyors").Count
sPrimaryUser = wsQuote.Range("PrimaryOperator")
iAllocationStartRow = wsQuote.Range("AdditionalOperatorList_StartRow").row
iAllocationRow = iAllocationStartRow

iUserCounter = 0

' make sure the number of users in the list does not exceed the maximum
If iNoOfUsers > iMaxNoOfUsers Then
    MsgBox "There are too many users in the list. Contact support to have this issue resolved."
    Exit Sub
End If

' need to remove the primary operator from the list of all valuers/surveyors/operators
For Each rCell In wsLists.Range("zzList_Surveyors")
    ' go through the list of users. if the user does not match the primary user,
    ' paste the user into the Additional Operators list
    If rCell.Value <> sPrimaryUser Then
        wsQuote.Range("C" & iAllocationRow).Value = rCell
        iAllocationRow = iAllocationRow + 1
        iUserCounter = iUserCounter + 1
    End If

Next rCell

' populate any unused cell with the text "Other"
If iUserCounter < iMaxNoOfUsers Then
    For i = iAllocationRow To (iAllocationStartRow + iMaxNoOfUsers) - 1
        wsQuote.Range("C" & i).Value = "Other"
    Next i
End If


End Sub

Sub clearDisbursements(sWorksheet As String)

Dim iRow As Long
Dim iLastRow As Long
Dim rSortCell_K1 As Range
Dim iDeleteRangeStart As Long
Dim iDeleteRangeEnd As Long
Dim i As Integer
Dim ws As Worksheet

    Set ws = Worksheets(sWorksheet)

    ' navigate to the start row
    ws.Activate
    ws.Range(sortCell_K1).Select
    iDeleteRangeStart = Selection.row
    iDeleteRangeEnd = lastRow(ws.Name, sortCell_K1)
        
    If iDeleteRangeEnd < 500000 Then ' indicates that there's unlikely to be any disbursements if > 499,999
        ws.Range("$" & iDeleteRangeStart & ":$" & iDeleteRangeEnd).ClearContents
    Else
        ws.Range("$" & iDeleteRangeStart & ":$" & iDeleteRangeStart).ClearContents
        ws.Range("A1").Select
    End If
    
    ' clear combo boxes
    If ws.CodeName = "wsSubConsultants" Then
        If wsSubConsultants.cboReportType.ListCount > 0 Then
            For i = wsSubConsultants.cboReportType.ListCount - 1 To 0 Step -1
                wsSubConsultants.cboReportType.RemoveItem i
            Next i
        End If
        
        wsSubConsultants.cboReportType = ""
        wsSubConsultants.cboSurveyors = ""
        wsSubConsultants.Range("SubConsultants_Role") = ""
    
    End If
    
    If ws.CodeName = "wsDisbursements" Then
        wsDisbursements.cboSurveyors = ""
        wsDisbursements.Range("Disbursements_Type") = ""
    End If
    
    ws.Range(sortCell_K1).Select

End Sub

Public Sub resetButtons()
Dim rCell As Range

    Call settings(False)

    Call setResetAssetClassTypeStatus_All(True)
    
    Call settings(True)


End Sub

Public Sub resetAllocationFormulas()

'Dim rSortCell_K1 As Range
'Dim rSortCell_K2 As Range
Dim i As Integer
Dim iLastRow As Long

    'Call settings(False)
    ' need the last row of the disbursement list in order to amend the formulas for each of he operators
    ' navigate to the start row
    wsDisbursements.Activate
    'wsDisbursements.Range(iStartCell).Select
    
    ' sort list in order of user only if data was successfully inserted
    iLastRow = lastRow(wsDisbursements.Name, sortCell_K1)

    If iLastRow > 500000 Then  ' indicates no disbursements
       ' default to 100
       iLastRow = 120
       wsDisbursements.Range(sortCell_K1).Select
    End If

    ' navigate to the start row on auto quote ws
    wsQuote.Activate
    wsQuote.Range("C" & AutoQuote_UserStartRow).Select
    
    ' Primary Operator Fees
    ' disbursements
    wsQuote.Range("Fee_Disbursements_PrimOp").Formula = "=SUMIF(Disbursements!$A$9:$A$" & iLastRow & ",'Quote'!Allocation_PrimOperator,Disbursements!$C$9:$C$" & iLastRow & ")"
    
    ' SC Allocation
    wsQuote.Range("Fee_SCAllocation_PrimOp").Formula = "=SCDisbursementFeeTotal"

    ' net fee = Gross - (Disbursements + SC Allocation)
    wsQuote.Range("Fee_Net_PrimOp").Formula = "=Fee_GrossFee_PrimOp-Fee_Disbursements_PrimOp-Fee_SCAllocation_PrimOp"
    
    ' amend formulas for the additional operators in Quote worksheet
    Call Allocation_InputFormulas

End Sub

Public Sub AllocationSwitch()

Call settings(False)

If wsQuote.Range("Allocation_Status") = "Gross Fee" Then
    wsQuote.Range("Allocation_Status") = "Percentage"
    wsQuote.Range("AllocationRange_Percentage").Interior.Color = 10092543
    wsQuote.Range("AllocationRange_GrossFee").Interior.Color = 16777215
Else
    wsQuote.Range("Allocation_Status") = "Gross Fee"
    wsQuote.Range("AllocationRange_Percentage").Interior.Color = 16777215
    wsQuote.Range("AllocationRange_GrossFee").Interior.Color = 10092543
End If

' populate the formulas in the correct column
Call Allocation_InputFormulas

Call settings(True)


End Sub

Public Sub Allocation_InputFormulas()
Dim rSortCell_K1 As Range
Dim rSortCell_K2 As Range
Dim i As Integer
Dim iLastRow As Long

    Call settings(False)
    ' need the last row of the disbursement list in order to amend the formulas for each of he operators
    ' navigate to the start row
    wsDisbursements.Activate
    'wsDisbursements.Range(iStartCell).Select
    
    ' sort list in order of user only if data was successfully inserted
    iLastRow = lastRow(wsDisbursements.Name, sortCell_K1)

    If iLastRow > 500000 Then  ' indicates no disbursements
       ' default to 100
       iLastRow = 120
       wsDisbursements.Range(sortCell_K1).Select
    End If

    ' navigate to the start row on auto quote ws
    wsQuote.Activate
    wsQuote.Range("G" & AutoQuote_UserStartRow).Select
    
    For i = AutoQuote_UserStartRow To AutoQuote_UserLastRow
        
        If wsQuote.Range("Allocation_Status") = "Gross Fee" Then
            ' gross fee input, column I - insert formula into percentage column
            'wsQuote.Range("H" & i).Formula = "=I" & i & "/ClientFeeTotal"
            wsQuote.Range("D" & i).Formula = "=IFERROR(E" & i & "/ClientFeeTotal,0)"
            wsQuote.Range("E" & i) = 0
        Else
            ' percentage fee input, column H - insert formula into Gross Fee column
            'wsQuote.Range("I" & i).Formula = "=H" & i & "*ClientFeeTotal"
            wsQuote.Range("E" & i).Formula = "=IFERROR(D" & i & "*ClientFeeTotal,0)"
            wsQuote.Range("D" & i) = 0
        End If
        
        ' disbursements
        wsQuote.Range("F" & i).Formula = "=SUMIF(Disbursements!$A$9:$A$" & iLastRow & ",'Quote'!C" & i & ",Disbursements!$C$9:$C$" & iLastRow & ")"
        
        ' net fee = Gross - (Disbursements + SC Allocation)
        wsQuote.Range("H" & i).Formula = "=E" & i & "-F" & i & "-G" & i
    
    Next i

    ' for Primary Valuer
    If wsQuote.Range("Allocation_Status") = "Gross Fee" Then
        ' % of client fee
        wsQuote.Range("Fee_FeePercentage_PrimOp").Formula = "=IFERROR(Fee_GrossFee_PrimOp/ClientFeeTotal,0)"
        wsQuote.Range("Fee_GrossFee_PrimOp") = 0
    Else
        ' gross fee = % x Total Client Fee
        wsQuote.Range("Fee_GrossFee_PrimOp").Formula = "=IFERROR(Fee_FeePercentage_PrimOp*ClientFeeTotal,0)"
        wsQuote.Range("Fee_FeePercentage_PrimOp") = 0
    End If

    ' reset formula range for scope item (i.e. TDD, Sched Make Good, FARR, Progress Claim etc..) in Auto Quote worksheet
    For i = AutoQuote_ScopeStartRow To AutoQuote_ScopeLastRow
        wsQuote.Range("F" & i).Formula = "=SUMIF(SubConsultants!$B$9:$B$200" & ",'Quote'!C" & i & ",SubConsultants!$D$9:$D$200" & ")"
    Next i
    Call settings(True)


End Sub
Public Sub Allocation_ClearNoFormulasColumn()
Dim rSortCell_K1 As Range
Dim rSortCell_K2 As Range
Dim i As Integer
Dim iLastRow As Long

    wsQuote.Activate
    wsQuote.Range("G" & AutoQuote_UserStartRow).Select
    
    For i = AutoQuote_UserStartRow To AutoQuote_UserLastRow
        
        If wsQuote.Range("Allocation_Status") = "Gross Fee" Then
            wsQuote.Range("I" & i).ClearContents ' clear gross fee column
        Else
            wsQuote.Range("H" & i).ClearContents ' clear % column
        End If
    
    Next i

    If wsQuote.Range("Allocation_Status") = "Gross Fee" Then
        wsQuote.Range("I60").ClearContents ' clear gross fee column
    Else
        wsQuote.Range("H60").ClearContents ' clear % column
    End If


End Sub

Public Function getLastPropertyRow_InPortfolio() As Long
Dim rCell As Range
Dim iLastRow As Integer

iLastRow = 0

' find position of last property
For Each rCell In wsQuote.Range("PF_PropertyAddresses_All")
    
    If rCell.Column = 4 Then ' process only if column is looking at D
        If rCell <> "Edit_Address" And Len(rCell) > 0 Then
            iLastRow = rCell.row
        Else
            Exit For
        End If
    End If

Next rCell

'  return last row
If iLastRow > 0 Then
    getLastPropertyRow_InPortfolio = iLastRow
Else
    getLastPropertyRow_InPortfolio = 0
End If


End Function
Public Function getPortfolioPropertiesRange() As String
Dim rCell As Range
Dim iRow As Integer
Dim sPropertiesRange As String
Dim sTemp As String

' get the cell range for the list of properties in a portfolio
sTemp = wsLists.Range("zzPropertyList").Rows.AddressLocal
sPropertiesRange = Left(sTemp, InStrRev(sTemp, "$"))
iRow = 0

For Each rCell In wsLists.Range("zzPropertyList")
    ' locate the last row of the last property on the list
    If Trim(rCell) <> "Edit_Address" Then
        
        iRow = rCell.row
        
    Else
        Exit For
    End If

Next rCell

getPortfolioPropertiesRange = sPropertiesRange & iRow

End Function

Public Function amendPortfolioPropertiesRange() As Boolean
Dim iLastPropertyRow As Long

    ' retrieve the last row of the last property in the portfolio. if no properties found, set to flag to false
    iLastPropertyRow = getLastPropertyRow_InPortfolio
    
    
    If iLastPropertyRow > 0 Then
        ' amend the disbursement named range as the number of disbursements will vary with each job
        ActiveWorkbook.Names("PF_PropertyAddresses_Selected").RefersToR1C1 = "='" & wsQuote.Name & "'!R38C4:R" & iLastPropertyRow & "C12"
        amendPortfolioPropertiesRange = True
'                   wbMaster.Names("Disbursements_List_PrintArea").RefersToR1C1 = "='" & wsDisbursements.Name & "'!R7C1:R" & iNoOfDisbursements & "C6"

    Else ' no property listed in portfolio
        amendPortfolioPropertiesRange = False
    
    End If

End Function

Public Sub setAssetClassTypeStatus_Single(sButtonName As String, isButton As Boolean)

Dim iPos As String
Dim sIndex As String
Dim sButtonNamePrefix As String
Dim sRange As String
Dim sColRange As String
Dim sColFeesRange As String
Dim sSummaryFeeRow As String
Dim sBusLine As String
Dim sSelectedReports As String
Dim iSelectedReportCount As Integer
Dim bSelectedReportFound As Boolean
Dim aAssetClass() As Variant
Dim i As Integer
Dim bBC_Sched_MG As Boolean
Dim bBC_SchedOfCond As Boolean
Dim rCell As Range

' sample button name passed - btn_BC_3

bBC_Sched_MG = False
bBC_SchedOfCond = False
iPos = InStrRev(sButtonName, "_")
sButtonNamePrefix = Left(sButtonName, iPos)

' get the button index
sIndex = Right(sButtonName, Len(sButtonName) - iPos)

' holds status of button i.e. true means button pressed
sRange = sButtonNamePrefix & "status_" & sIndex

' entire column range name for fee
sColRange = sButtonNamePrefix & "column_" & sIndex

' column fees only range
sColFeesRange = sButtonNamePrefix & "column_feesRange_" & sIndex

' row range for fee summary
sSummaryFeeRow = sButtonNamePrefix & "Fee_" & sIndex

If wsLists.Range(sRange) = True Then
    
    If isButton Then
        wsQuote.Shapes.Range(Array(sButtonName)).Select
        ' format button
        Selection.Font.ColorIndex = 1
        Selection.Font.FontStyle = "Book"
    End If

    wsLists.Range(sRange) = False

    ' delete all values
    wsQuote.Range(sColFeesRange).ClearContents

    ' hide fee column
    wsQuote.Range(sColRange).Select
    Selection.EntireColumn.Hidden = True

    ' hide fee summary row
    wsQuote.Range(sSummaryFeeRow).Select
    Selection.EntireRow.Hidden = True

Else
    
    ' if Schedule of Make Good (btn_BC_3) is selected, disable all other buttons in BC
    If sButtonName = "btn_BC_3" Then
        bBC_Sched_MG = True ' flag to indicate MakeGood has been selected
    End If
    
    ' if Schedule of Condition (btn_BC_2) is selected, disable all other buttons in BC
    If sButtonName = "btn_BC_2" Then
        bBC_SchedOfCond = True ' flag to indicate MakeGood has been selected
    End If
    
    If isButton Then
        ' format button
        wsQuote.Shapes.Range(Array(sButtonName)).Select
        Selection.Font.ColorIndex = 5
        Selection.Font.FontStyle = "Book Bold"
    End If

    wsLists.Range(sRange) = True
    
    ' unhide fee column
    wsQuote.Range(sColRange).Select
    Selection.EntireColumn.Hidden = False

    ' unhide fee summary row
    wsQuote.Range(sSummaryFeeRow).Select
    Selection.EntireRow.Hidden = False
    
End If


' test to see if button is clicked
If wsLists.Range("btn_BC_status_3") = True Then
    bBC_Sched_MG = True
End If

' hide buttons for specific reports
If isButton Then
    If bBC_Sched_MG Then
        Call disableButton("btn_BC_1", True)
        Call disableButton("btn_BC_2", True)
        Call disableButton("btn_BC_5", True)
        Call disableButton("btn_BC_6", True)
        
        wsMakeGood.Visible = xlSheetVisible
    Else
        If bBC_SchedOfCond = False Then
            Call disableButton("btn_BC_1", False)
            Call disableButton("btn_BC_2", False)
            Call disableButton("btn_BC_5", False)
            Call disableButton("btn_BC_6", False)
        
            wsMakeGood.Visible = xlSheetHidden
        End If
    End If
    

    If bBC_SchedOfCond Then
        Call disableButton("btn_BC_1", True)
        Call disableButton("btn_BC_3", True)
        Call disableButton("btn_BC_5", True)
        Call disableButton("btn_BC_6", True)
        
        wsSchedOfCondition.Visible = xlSheetVisible
    Else
        If bBC_Sched_MG = False Then
            Call disableButton("btn_BC_1", False)
            Call disableButton("btn_BC_3", False)
            Call disableButton("btn_BC_5", False)
            Call disableButton("btn_BC_6", False)
        
            wsSchedOfCondition.Visible = xlSheetHidden
        End If
        
    End If

End If

' enable or disable worksheets that tied to specific scopes
Call hideUnhide_Worksheets

' populates the property type combo box
Call populate_cboReportType(isButton)

' populate the invoice wording
Call populate_InvoiceWording

wsQuote.Activate
wsQuote.Range("A1").Select

End Sub
Private Sub hideUnhide_Worksheets()

wsMakeGood.Visible = xlSheetHidden
wsSchedOfCondition.Visible = xlSheetHidden
wsAssetDesc.Visible = xlSheetHidden
wsScopeDesc.Visible = xlSheetHidden
wsScopeOfService.Visible = xlSheetHidden
wsAttachment2.Visible = xlSheetHidden

' determine service specific worksheets to display
aAssetClass = Array("AssetClass_BC", "AssetClass_CC", "AssetClass_Tax")
For i = LBound(aAssetClass) To UBound(aAssetClass)
    
    bSelectedReportFound = False
    
    ' get description of selected report/service types
    For Each rCell In wsLists.Range(aAssetClass(i))
        
        
        If rCell = True Then
            Select Case aAssetClass(i)
            Case "AssetClass_BC"
                Select Case rCell.Offset(0, -1)
                Case "BC_SchedMakeGoodStage1", "BC_SchedMakeGoodStage2"
                    wsMakeGood.Visible = xlSheetVisible
                Case "BC_SchedCond"
                    wsSchedOfCondition.Visible = xlSheetVisible
                Case "BC_LifeCycleCost", "BC_TDD_Purchaser", "BC_TDD_Vendor", "BC_Other"
                    wsAssetDesc.Visible = xlSheetVisible
                    wsScopeDesc.Visible = xlSheetVisible
                End Select
            Case "AssetClass_CC"
                wsScopeOfService.Visible = xlSheetVisible
            Case "AssetClass_Tax"
                wsAttachment2.Visible = xlSheetVisible
            End Select
        
        End If

    Next rCell

Next i


End Sub

Private Sub disableButton(sButtonName As String, bDisable As Boolean)
Dim iPos As String
Dim sIndex As String
Dim sButtonNamePrefix As String
Dim sRange As String
Dim sColRange As String
Dim sColFeesRange As String
Dim sSummaryFeeRow As String
Dim isAlreadyHidden As Boolean

'check if already hidden
If wsQuote.Shapes.Range(Array(sButtonName)).Visible = msoFalse Then
    isAlreadyHidden = True
Else
    isAlreadyHidden = False
End If

' sample button name passed - btn_BC_3
iPos = InStrRev(sButtonName, "_")
sButtonNamePrefix = Left(sButtonName, iPos)

' get the button index
sIndex = Right(sButtonName, Len(sButtonName) - iPos)

' holds status of button i.e. true means button pressed
sRange = sButtonNamePrefix & "status_" & sIndex

' entire column range name for fee
sColRange = sButtonNamePrefix & "column_" & sIndex

' column fees only range
sColFeesRange = sButtonNamePrefix & "column_feesRange_" & sIndex

' row range for fee summary
sSummaryFeeRow = sButtonNamePrefix & "Fee_" & sIndex

If isAlreadyHidden = False Or bDisable = False Then
    If bDisable Then

        ' format button
        wsQuote.Shapes.Range(Array(sButtonName)).Select
        Selection.Font.ColorIndex = 1
        Selection.Font.FontStyle = "Book"
    
        ' set flag to False
        wsLists.Range(sRange) = False
    
        ' delete all values
        wsQuote.Range(sColFeesRange).ClearContents
    
        ' hide fee column
        wsQuote.Activate
        wsQuote.Range(sColRange).Select
        Selection.EntireColumn.Hidden = True
    
        ' hide fee summary row
        wsQuote.Range(sSummaryFeeRow).Select
        Selection.EntireRow.Hidden = True
        
        ' hide button
        wsQuote.Shapes.Range(Array(sButtonName)).Visible = msoFalse
        
    Else
        ' unhide button
        wsQuote.Shapes.Range(Array(sButtonName)).Visible = msoTrue
    End If
End If

End Sub


Public Sub setResetAssetClassTypeStatus_All(bReset As Boolean)
Dim i As Integer

Call settings(False)
' run through all the buttons to evaluate button status
' BUILDING CONSULTANCY BUTTONS
For i = 1 To btn_BC_Count

    If i <> 4 Then ' 4 doesn't have an associated button. Sched of Make Good comes in 2 stages
        Call disableButton("btn_BC_" & i, False) ' make sure all buttons are made visible again
        wsQuote.Shapes.Range(Array("btn_BC_" & i)).Select
        
        If wsLists.Range("btn_BC_status_" & i) = False Or bReset = True Then
            ' set buttons to false i.e. not selected
            Selection.Font.ColorIndex = 1
            Selection.Font.FontStyle = "Book"
            
            ' set status to false
            wsLists.Range("btn_BC_status_" & i) = False
            
            If i = 3 Then ' if i = 3 (Sched of Make Good, also need to set Sched of Make Good Stage 2 to false)
                wsLists.Range("btn_BC_status_4") = False
            
                ' hide fee column
                wsQuote.Range("btn_BC_column_4").Select
                Selection.EntireColumn.Hidden = True
            
                ' hide fee summary row
                wsQuote.Range("btn_BC_Fee_4").Select
                Selection.EntireRow.Hidden = True
            End If
        
            ' hide fee column
            wsQuote.Range("btn_BC_column_" & i).Select
            Selection.EntireColumn.Hidden = True
        
            ' hide fee summary row
            wsQuote.Range("btn_BC_Fee_" & i).Select
            Selection.EntireRow.Hidden = True
        
        Else
            Selection.Font.ColorIndex = 5
            Selection.Font.FontStyle = "Book Bold"
        
            ' unhide fee column
            wsQuote.Range("btn_BC_column_" & i).Select
            Selection.EntireColumn.Hidden = False
            
            ' unhide fee summary row
            wsQuote.Range("btn_BC_Fee_" & i).Select
            Selection.EntireRow.Hidden = False
            
        End If
    End If
Next i

' COST CONSULTANCY BUTTONS
For i = 1 To btn_CC_Count

    wsQuote.Shapes.Range(Array("btn_CC_" & i)).Select
    If wsLists.Range("btn_CC_status_" & i) = False Or bReset = True Then
        ' set buttons to false
        Selection.Font.ColorIndex = 1
        Selection.Font.FontStyle = "Book"
    
        ' set status to false
        wsLists.Range("btn_CC_status_" & i) = False
    
        ' hide fee column
        wsQuote.Range("btn_CC_column_" & i).Select
        Selection.EntireColumn.Hidden = True
    
        ' hide fee summary row
        wsQuote.Range("btn_CC_Fee_" & i).Select
        Selection.EntireRow.Hidden = True
    
    Else
        Selection.Font.ColorIndex = 5
        Selection.Font.FontStyle = "Book Bold"
    
        ' unhide fee column
        wsQuote.Range("btn_CC_column_" & i).Select
        Selection.EntireColumn.Hidden = False
    
        ' unhide fee summary row
        wsQuote.Range("btn_CC_Fee_" & i).Select
        Selection.EntireRow.Hidden = False
    
    End If

Next i

' TAX CONSULTANCY BUTTONS
For i = 1 To btn_Tax_Count

    wsQuote.Shapes.Range(Array("btn_Tax_" & i)).Select
    If wsLists.Range("btn_Tax_status_" & i) = False Or bReset = True Then
        Selection.Font.ColorIndex = 1
        Selection.Font.FontStyle = "Book"
    
        ' set status to false
        wsLists.Range("btn_Tax_status_" & i) = False
    
        ' hide fee column
        wsQuote.Range("btn_Tax_column_" & i).Select
        Selection.EntireColumn.Hidden = True
    
        ' hide fee summary row
        wsQuote.Range("btn_Tax_Fee_" & i).Select
        Selection.EntireRow.Hidden = True
    
    Else
        Selection.Font.ColorIndex = 5
        Selection.Font.FontStyle = "Book Bold"
    
        ' unhide fee column
        wsQuote.Range("btn_Tax_column_" & i).Select
        Selection.EntireColumn.Hidden = False
        
        ' unhide fee summary row
        wsQuote.Range("btn_Tax_Fee_" & i).Select
        Selection.EntireRow.Hidden = False
    
    End If

Next i

Call settings(True)


End Sub

Public Sub setCountry(sCountry As String)

'wsQuote.Shapes.Range(Array(sButtonName)).Select
If sCountry = "Australia" Then
    
    ' format button
    wsQuote.Shapes.Range(Array("btn_Country_NZ")).Select
    Selection.Font.ColorIndex = 1
    Selection.Font.FontStyle = "Book"

    wsQuote.Shapes.Range(Array("btn_Country_Aust")).Select
    Selection.Font.ColorIndex = 5
    Selection.Font.FontStyle = "Book Bold"
    
    ' display selected country
    wsQuote.Range("Country_Selected") = "Australia"
    
'    wsQuote.cboStates_NZ.Text = ""
'    wsQuote.cboStates_NZ.Visible = False
'    wsQuote.cboStates_Aust.Visible = True

Else

    ' format button
    wsQuote.Shapes.Range(Array("btn_Country_Aust")).Select
    Selection.Font.ColorIndex = 1
    Selection.Font.FontStyle = "Book"

    wsQuote.Shapes.Range(Array("btn_Country_NZ")).Select
    Selection.Font.ColorIndex = 5
    Selection.Font.FontStyle = "Book Bold"

    ' display selected country
    wsQuote.Range("Country_Selected") = "New Zealand"

'    wsQuote.cboStates_Aust.Text = ""
'    wsQuote.cboStates_Aust.Visible = False
'    wsQuote.cboStates_NZ.Visible = True

End If

wsQuote.Range("Country_Selected").Select

End Sub

Public Sub selectAustralia()

Call settings(False)

Call setCountry("Australia")

Call settings(True)

End Sub

Public Sub selectNZ()

Call settings(False)

Call setCountry("NZ")

Call settings(True)

End Sub
Public Function validateScopeSelection() As Boolean
Dim ws As Worksheet
Dim rCell As Range

Set ws = Worksheets(wsLists.Name)

validateScopeSelection = False

' search for any scope that has been selected
For Each rCell In ws.Range("AssetClass_BC")
    If rCell = True Then
        validateScopeSelection = True
        GoTo Exit_validateScopeSelection
    End If
Next rCell

For Each rCell In ws.Range("AssetClass_CC")
    If rCell = True Then
        validateScopeSelection = True
        Exit For
    End If
Next rCell

For Each rCell In ws.Range("AssetClass_Tax")
    If rCell = True Then
        validateScopeSelection = True
        Exit For
    End If
Next rCell


Exit_validateScopeSelection:
    Set ws = Nothing
    Exit Function

Err_validateScopeSelection:
    Err.Raise Err.Number, "validateScopeSelection", Err.Description
    Resume Exit_validateScopeSelection

End Function

Public Function validateForm() As Boolean
Dim ws As Worksheet

validateForm = False

Set ws = Worksheets(wsQuote.Name)

    If validateScopeSelection Then

        ' check Primary Operator
        If wsQuote.cboPrimaryOperator = "" Then
            MsgBox "Please select an operator.", vbInformation
            ws.Range("PrimaryOperator_Label").Select
            GoTo Exit_validateForm
        End If
        
        ' check LOE Signatory
        If wsQuote.cboLOESignatory = "" Then
            MsgBox "Please select a signatory.", vbInformation
            ws.Range("LOESignatory_Label").Select
            GoTo Exit_validateForm
        End If
            
        ' check Client
        If wsQuote.Range("SelectedCompany") = "" Then
            MsgBox "Please select a client.", vbInformation
            ws.Range("client_Label").Select
            GoTo Exit_validateForm
        End If
        
        
        ' check Address
        If wsQuote.Range("PFAddress_01") = "Edit_Address" Then
            MsgBox "Please add a valid address.", vbInformation
            ws.Range("PFAddress_01").Select
            GoTo Exit_validateForm
        End If
        
        ' check dropdown boxes
        If wsQuote.cboBC_Purpose.Visible = True Then
            If wsQuote.cboBC_Purpose = "" Then
                MsgBox "Please select a purpose.", vbInformation
                wsQuote.Range("BC_Purpose").Select
                GoTo Exit_validateForm
            End If
        End If

        If wsQuote.cboMakeGood_ScopeOfService.Visible = True Then
            If wsQuote.cboMakeGood_ScopeOfService = "" Then
                MsgBox "Please select the scope of service.", vbInformation
                wsQuote.Range("makeGood_SoS").Select
                GoTo Exit_validateForm
            End If
        End If

        If wsQuote.Range("Allocation_TotalPerc") <> 1 Then
            MsgBox "The allocation of fees are not correct. Please make sure 100% of the fees have been correctly distributed " _
            & "before attempting to generate another email.", vbInformation
            wsQuote.Range("Fee_FeePercentage_PrimOp").Select
            GoTo Exit_validateForm
        End If

        validateForm = True

    Else
        MsgBox "Please select a valid scope.", vbInformation
        ws.Range("BC_Label").Select
        GoTo Exit_validateForm
    End If
    
    ' cursor gets to this point, all conditions have been met
    validateForm = True

Exit_validateForm:
    Set ws = Nothing
    Exit Function

Err_validateForm:
    Err.Raise Err.Number, "validateForm", Err.Description
    Resume Exit_validateForm

End Function

Public Sub populate_cboReportType(isButton As Boolean)
Dim rCell As Range
Dim iCountScopes_byBusLine As Integer
Dim sScopeDescTemp_byBusLine As String

Dim iCountScopes_All As Integer
Dim sScopeDescTemp_All As String

Dim sLastScopeDesc As String


    aScopes = Array("BC", "CC", "Tax")
    
    ' clear combo box
    'wsSubConsultants.Activate
    wsSubConsultants.cboReportType.Clear
    
    iCountScopes_All = 0
    sScopeDescTemp_All = ""
    
    For i = 0 To UBound(aScopes)
    
        iCountScopes_byBusLine = 0
        sScopeDescTemp_byBusLine = ""
        
        For Each rCell In wsLists.Range("AssetClass_" & aScopes(i))
            ' insert only report types/scope items that have been selected by the operator on wsQuote
            If rCell = True And Len(rCell.Offset(0, 1)) > 0 Then
                wsSubConsultants.cboReportType.AddItem rCell.Offset(0, 1).Value
                
                If isButtonFunctional(rCell.Offset(0, -1)) Then
                    sLastScopeDesc = rCell.Offset(0, 1) ' save last scope description in case it is the last of a multiple of scopes
                    iCountScopes_byBusLine = iCountScopes_byBusLine + 1
                    iCountScopes_All = iCountScopes_All + 1
                    
                    ' service descriptions by business line i.e. BC, CC and Tax
                    Select Case iCountScopes_byBusLine
                    Case 1
                        sScopeDescTemp_byBusLine = rCell.Offset(0, 1)
                    Case Is > 1
                        sScopeDescTemp_byBusLine = sScopeDescTemp_byBusLine & ", " & rCell.Offset(0, 1)
                    End Select
                
                    ' description for all services in all business lines
                    Select Case iCountScopes_All
                    Case 1
                        sScopeDescTemp_All = rCell.Offset(0, 1)
                    Case Is > 1
                        sScopeDescTemp_All = sScopeDescTemp_All & ", " & rCell.Offset(0, 1)
                    End Select
                
                
                End If
        
            End If
            
        Next rCell
    
        If isButton Then
            ' if there are multiple scopes, need to insert an "and" between the second last and last scope
            If iCountScopes_byBusLine > 1 Then
                sScopeDescTemp_byBusLine = Left(sScopeDescTemp_byBusLine, Len(sScopeDescTemp_byBusLine) - Len(sLastScopeDesc) - 2) & " and " & sLastScopeDesc
            End If
            
            wsLists.Range(aScopes(i) & "_CoreScopeDescription") = sScopeDescTemp_byBusLine
            wsQuote.Range("AssetClass_" & aScopes(i) & "_Selected") = sScopeDescTemp_byBusLine
        End If
    
    Next i

    If isButton Then
        ' if there are multiple scopes, need to insert an "and" between the second last and last scope
        If iCountScopes_All > 1 Then
            sScopeDescTemp_All = Left(sScopeDescTemp_All, Len(sScopeDescTemp_All) - Len(sLastScopeDesc) - 2) & " and " & sLastScopeDesc
        End If
        
    End If

    'Call populate_InvoiceWording

End Sub

Public Sub updateInvoiceWording_Address()
Dim sPropAddress As String


    ' check whether job is for a portfolio or a single property
    If wsLists.Range("zzPFStatus") = True Then
        ' populate the Invoice Wording cell
        wsQuote.Range("InvoiceWording_Address").Formula = "=PFName"
    Else
        sPropAddress = "=PFAddress_01 & " + chr(34) + " " + chr(34) + " & PFAddress_01_Postcode"
        
        ' populate the Invoice Wording cell
        wsQuote.Range("InvoiceWording_Address") = sPropAddress & chr(10)
        
    End If


End Sub

Public Sub populate_InvoiceWording()
Dim rCell As Range
Dim iCountScopes_All As Integer
Dim sScopeDescTemp_All As String
Dim sLastScopeDesc As String

    aScopes = Array("BC", "CC", "Tax")
    
    iCountScopes_All = 0
    sScopeDescTemp_All = ""
    
    For i = 0 To UBound(aScopes)
    
        For Each rCell In wsLists.Range("AssetClass_" & aScopes(i))
            ' insert only report types/scope items that have been selected by the operator on wsQuote
            If rCell = True And Len(rCell.Offset(0, 1)) > 0 Then
                
                If isButtonFunctional(rCell.Offset(0, -1)) Then
                    sLastScopeDesc = rCell.Offset(0, 1) ' save last scope description in case it is the last of a multiple of scopes
                    iCountScopes_All = iCountScopes_All + 1
                    
                    ' description for all services in all business lines
                    Select Case iCountScopes_All
                    Case 1
                        sScopeDescTemp_All = rCell.Offset(0, 1)
                    Case Is > 1
                        sScopeDescTemp_All = sScopeDescTemp_All & ", " & rCell.Offset(0, 1)
                    End Select
                
                End If
        
            End If
            
        Next rCell
    
    Next i

    ' if there are multiple scopes, need to insert an "and" between the second last and last scope
    If iCountScopes_All > 1 Then
        sScopeDescTemp_All = Left(sScopeDescTemp_All, Len(sScopeDescTemp_All) - Len(sLastScopeDesc) - 2) & " and " & sLastScopeDesc
    End If
    
    sScopeDescTemp_All = "Professional services for the preparation of a " & sScopeDescTemp_All & " in accordance with your written instructions."
    
    wsQuote.Range("InvoiceWording") = sScopeDescTemp_All
    
    Call updateInvoiceWording_Address

End Sub

Public Sub BC_Purpose_SetUpDropdown()
Dim nCell As Name
Dim bEnablePurpose As Boolean

bEnablePurpose = False

For Each nCell In ActiveWorkbook.Names
'        Debug.Print nCell.Name, nCell.RefersTo
    
    If InStr(1, nCell.Name, "btn_BC_status_", vbBinaryCompare) > 0 Then
        Select Case Right(nCell.Name, 1)
        Case 1, 5, 6, 7
             If wsLists.Range(nCell) = True Then
                bEnablePurpose = True
                ' only needs one selected to enable dropdown
                Exit For
            End If
        Case Else
            bEnablePurpose = False
        End Select
    End If
Next nCell

If bEnablePurpose Then
    ' enable Purpose selection Acquisition or Divestment
    wsQuote.cboBC_Purpose.Visible = True
    wsQuote.Range("cboBC_Purpose_Label") = "Purpose"
Else
    
    wsQuote.cboBC_Purpose = ""
    wsQuote.cboBC_Purpose.Visible = False
    wsQuote.Range("cboBC_Purpose_Label") = ""
End If

'Call resetComboBoxPhysicalAttributes(wsQuote.Name, "cboBC_Purpose", 19.5, 103.5, 395.25)

End Sub
