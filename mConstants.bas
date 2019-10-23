Attribute VB_Name = "mConstants"
Option Explicit
' for the disbursements and subconsultants worksheets
Public Const iStartCell = "A8"
Public Const sortCell_K1 = "A9"
Public Const sortCell_K2 = "B9"
Public Const colSubconsultantRow = "C"

' for wsQuote worksheet
Public Const AutoQuote_UserStartRow = 91
Public Const AutoQuote_UserLastRow = 113

Public Const AutoQuote_ScopeStartRow = 62
Public Const AutoQuote_ScopeLastRow = 82

Public Const AutoQuote_SubcontractorStartRow = 64
Public Const AutoQuote_SubcontractorLastRow = 84

Public Const iMaxNoOfUsers = 23

Public aScopes As Variant

'Risk Levels
Public Const iHigh = 12
Public Const iMediumHigh = 9
Public Const iMedium = 6
Public Const iLowMedium = 3
Public Const iLow = 0

'Asset Class Types
'BUILDING CONSULTANTCY
Public btn_BC_Other As Boolean
Public btn_BC_SchedCond As Boolean
Public btn_BC_SchedMakeGood As Boolean
Public btn_BC_TDD_Purchaser As Boolean
Public btn_BC_TDD_Vendor As Boolean

Public Const btn_BC_Count = 6
Public Const BC_AddScope_Count = 18

'COST CONSULTANTCY
Public btn_CC_AR_CostAssess As Boolean
Public btn_CC_Other As Boolean
Public btn_CC_CostPlanning As Boolean
Public btn_CC_ProgressClaim As Boolean
Public btn_CC_IR_QS_VerifyCC As Boolean
Public btn_CC_Ins_ReinsCostAssess_VerifyCC As Boolean
Public btn_CC_BC_LifeCycleCost As Boolean

Public Const btn_CC_Count = 5

'TAX
Public btn_Tax_AcqAssess As Boolean
Public btn_Tax_BalAdj As Boolean
Public btn_Tax_ComplReview As Boolean
Public btn_Tax_ConstAssessDep As Boolean
Public btn_Tax_DepReplCostBookVal As Boolean
Public btn_Tax_FitOutAbanRefurb As Boolean
Public btn_Tax_FitOutAban As Boolean
Public btn_Tax_FixedAssetReg As Boolean
Public btn_Tax_IndDepSched As Boolean
Public btn_Tax_RefurbExtAssessDep As Boolean
Public btn_Tax_StampDutyAssessMkt As Boolean
Public btn_Tax_StampDutyAssessStatDec As Boolean

Public gsClientName As String

Public gsClientID As Integer
Public gsClientFirstName As String
Public gsClientLastName As String
Public gsClientCompany As String
Public gsClientAddressLine1 As String
Public gsClientAddressLine2 As String
Public gsClientSuburb As String
Public gsClientState As String
Public gsClientPostcode As String
Public gsClientPhone As String
Public gsClientEmailAddress As String
Public gsClientPosition As String

Public gsClientSelected As String
Public gsCompanyComboRefresh As Boolean


Public Const btn_Tax_Count = 10
Public Const SpecialCharacters As String = ",| |!|@|#|$|%|^|&|*|(|)|{|[|]|}|?|/|\|'"  'modify as needed

Type ButtonSizeType
    topPosition As Single
    leftPosition As Single
    height As Single
    width As Single
End Type

Public myButton As ButtonSizeType

Sub GetButtonSize(cb As MSForms.CommandButton)
' Save original button size to solve windows bug that changes the button size to
' adjust to screen resolution, when not in native resolution mode of screen
    myButton.topPosition = cb.Top
    myButton.leftPosition = cb.Left
    myButton.height = cb.height
    myButton.width = cb.width
End Sub

Sub SetButtonSize(cb As MSForms.CommandButton)
' Restore original button size to solve windows bug that changes the button size to
' adjust to screen resolution, when not in native resolution mode of screen
    cb.Top = myButton.topPosition
    cb.Left = myButton.leftPosition
    cb.height = myButton.height
    cb.width = myButton.width
End Sub

Sub MScmdButtonFeatureFix(bStartOfProc As Boolean, cmdButton As MSForms.CommandButton)

If bStartOfProc Then
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    ' to resolve the expanding/reducing of the button
    GetButtonSize cmdButton

Else
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' to resolve the expanding/reducing of the button
    SetButtonSize cmdButton

End If

End Sub

Public Function isButtonFunctional(sButtonName As String) As Boolean
Dim i As Integer
Dim aNonFunctionalButtons() As Variant

aNonFunctionalButtons = Array("BC_SchedMakeGoodStage2")
isButtonFunctional = True

' check if button is functional or not
For i = LBound(aNonFunctionalButtons, 1) To UBound(aNonFunctionalButtons)
    If aNonFunctionalButtons(i) = sButtonName Then
        isButtonFunctional = False
        Exit For
    End If
Next i


End Function

Private Sub text()
Dim btest As Boolean

btest = isButtonFunctional("BC_SchedMakeGoodStage1")

Debug.Print btest
End Sub
