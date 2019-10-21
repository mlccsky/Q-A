Option Explicit

Private Sub populateCompanyContact()
Dim rCell As Range
Dim bFound As Boolean
Dim i As Integer

i = 1

' clear all items in the combo box
wsQuote.cboCompanyContact.Clear

Call clearClientDetails

Call settings(False)
For Each rCell In wsList_CommonClients.Range("zzList_CommonClients")
    If rCell = wsQuote.Range("SelectedCompany") Then
        ' add contact name if company match
        wsQuote.cboCompanyContact.AddItem rCell.Offset(0, -1) ' add index
        wsQuote.cboCompanyContact.List(cboCompanyContact.ListCount - 1, 1) = Trim(rCell.Offset(0, 1)) & " " & Trim(rCell.Offset(0, 2)) ' add name
        
        bFound = True
        i = i + 1
    End If
    
    ' exit loop once all contacts added (list of company and contacts list should always be sorted)
    If bFound And rCell <> wsQuote.Range("SelectedCompany") Then
        Exit Sub
    End If
    
Next rCell
Call settings(True)


End Sub

Private Sub cboClients_Change()


If wsQuote.Range("SelectedCompany") <> "Other" Then

    If wsQuote.Range("SelectedCompany") <> gsClientSelected And wsQuote.Range("SelectedCompany") <> "" Then
        
        Call populateCompanyContact
      
        gsClientSelected = wsQuote.Range("SelectedCompany")
        
        
    End If
End If

'Call populate_InvoiceWording

End Sub

Private Sub cboCompanyContact_Change()
On Error GoTo Err_cboCompanyContact_Change

Dim rCell As Range


If wsQuote.cboCompanyContact <> "" Then

    Call storeClientDetails(CInt(wsQuote.cboCompanyContact))

    With wsQuote
        On Error Resume Next
        .Range("ClientID") = gsClientID
        
        .Range("ClientCompany") = gsClientCompany
        .Range("ClientName") = gsClientFirstName & " " & gsClientLastName
        .Range("ClientAddressLine1") = gsClientAddressLine1
        .Range("ClientAddressLine2") = gsClientAddressLine2
        .Range("ClientSuburb") = gsClientSuburb
        .Range("ClientState") = gsClientState
        .Range("ClientPostcode") = gsClientPostcode
        .Range("ClientPhone") = gsClientPhone
        .Range("ClientEmailAddress") = gsClientEmailAddress
    
    End With

End If


'For Each rCell In wsList_CommonClients.Range("zzList_CommonClients_Index")
    ' look for the selected index and assign associated values
'    If wsQuote.cboCompanyContact <> "" Then
'        If rCell = CInt(wsQuote.cboCompanyContact) Then
'            With wsQuote
                'Call settings(False)
'                Application.EnableEvents = False
'                On Error Resume Next
'                .Range("ClientID") = rCell
                
'                .Range("ClientCompany") = rCell.Offset(0, 1)
'                .Range("ClientName") = rCell.Offset(0, 2)
'                .Range("ClientStreet") = rCell.Offset(0, 3)
                
'                .Range("ClientSuburb") = rCell.Offset(0, 4)
'                .Range("ClientState") = rCell.Offset(0, 5)
'                .Range("ClientPostcode") = rCell.Offset(0, 6)
'                .Range("ClientPhone") = rCell.Offset(0, 7)
'                .Range("ClientEmailAddress") = rCell.Offset(0, 8)
'
'                Call storeClientDetails ' store for later use
'
'            End With
'        End If
'    End If
'
'Next rCell

Exit_cboCompanyContact_Change:
    Call settings(True)
    
 '   Application.EnableEvents = True
    Exit Sub

Err_cboCompanyContact_Change:
    MsgBox "Error Number: " & Err.Number & chr(13) & "Procedure: cboCompanyContact_Change" & chr(13) & "Description: " & Err.Description
    Resume Exit_cboCompanyContact_Change

End Sub
Private Sub cboBC_Purpose_Click()
    ResetButton wsQuote.cboBC_Purpose
End Sub
Private Sub cboMakeGood_ScopeOfService_Click()
    ResetButton wsQuote.cboMakeGood_ScopeOfService
    
    wsQuote.Range("AssetClass_BC_Selected") = wsLists.Range("Description_MG_Stage1")
End Sub
Private Sub cboPrimaryOperator_Click()
    ResetButton wsQuote.cboPrimaryOperator
End Sub
Private Sub cboLOESignatory_Click()
    ResetButton wsQuote.cboLOESignatory
End Sub
Private Sub cboClients_Click()
    ResetButton wsQuote.cboClients
End Sub
Private Sub cboCompanyContact_Click()
    ResetButton wsQuote.cboCompanyContact
End Sub


Private Sub cboPrimaryOperator_Change()
Dim rCell As Range
Dim bValidSignatory As Boolean

bValidSignatory = False

Call settings(False)

Call UpdateUserAllocationList

' copy primary Operator to LOE Signatory
If Me.cboPrimaryOperator <> "" And Me.cboLOESignatory = "" Then
    ' before duplicating the LOE signatory, need to check whether operator is a valid signatory
    For Each rCell In wsLists.Range("zzList_LOE_Signatories")
        If Me.cboPrimaryOperator = rCell Then
            bValidSignatory = True
            Exit For
        End If
    Next rCell

End If

' if valid signatory, then duplicate in LOE signatory
If bValidSignatory Then
    Me.cboLOESignatory = Me.cboPrimaryOperator
End If

Call settings(True)

End Sub
Private Sub clearClientDetails()

    Call settings(False)
    
    With wsQuote
        .Range("ClientName") = ""
        .Range("ClientCompany") = ""
        .Range("ClientAddressLine1") = ""
        .Range("ClientAddressLine2") = ""
        .Range("ClientSuburb") = ""
        .Range("ClientState") = ""
        .Range("ClientPostcode") = ""
        .Range("ClientPhone") = ""
        .Range("ClientEmailAddress") = ""
    End With

    Call settings(True)

End Sub

Private Sub storeClientDetails(iClientID As Integer)
On Error Resume Next
Dim rCell As Range

' store values to retrieve if "Other" is selected
gsClientID = iClientID
    
For Each rCell In wsList_CommonClients.Range("zzList_CommonClients_Index")
    ' look for the selected index and assign associated values
    If rCell = gsClientID Then
        gsClientCompany = rCell.Offset(0, 1)
        gsClientFirstName = rCell.Offset(0, 2)
        gsClientLastName = rCell.Offset(0, 3)
        gsClientPosition = rCell.Offset(0, 4)
        gsClientAddressLine1 = rCell.Offset(0, 5)
        gsClientAddressLine2 = rCell.Offset(0, 6)
        gsClientSuburb = rCell.Offset(0, 7)
        gsClientState = rCell.Offset(0, 8)
        gsClientPostcode = rCell.Offset(0, 9)
        gsClientPhone = rCell.Offset(0, 10)
        gsClientEmailAddress = rCell.Offset(0, 11)
        Exit For
    End If
    
Next rCell

End Sub

Private Sub Worksheet_Activate()

'Call fillCombo
'gsClientSelected = wsQuote.Range("SelectedCompany")

Call storeClientDetails(CInt(wsQuote.Range("ClientID")))

End Sub