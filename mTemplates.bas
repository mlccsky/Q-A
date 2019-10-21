Attribute VB_Name = "mTemplates"
Option Explicit

Public Function prepareTemplate(sDocType As String, sFileName As String) As String
On Error GoTo Err_prepareTemplate

Dim wdApp As Word.Application

Dim myDoc As Word.Document
Dim sPropAddress As String
Dim sMID As String
'Dim sClientCompanyAddress As String
Dim sAssetClass As String
Dim sSelectedCountry As String
Dim bAust As Boolean
Dim sFirstName As String
Dim sClientAddressLine3 As String

' ScopeTypes
Dim bBC_MakeGood As Boolean
Dim bBC_SoC As Boolean
Dim bBC_TDD_Divestment As Boolean

' Buttons
Dim iPos As String
Dim sIndex As String
Dim sButtonNamePrefix As String
Dim sRange As String
Dim sColRange As String
Dim sBusLine As String

' Fee Calcs
Dim dFees_GST As Double
Dim dFees_ExGST As Double
Dim dFees_IncGST As Double
Dim dFees_GST_RunningTotal As Double
Dim dFees_ExGST_RunningTotal As Double
Dim dFees_IncGST_RunningTotal As Double
Dim dFees_Discount_GST As Double
Dim dFees_Discount_ExGST As Double
Dim dFees_Discount_IncGST As Double

Dim bAdditionalScope As Boolean
Dim sAttachmentTableRangeSelection As String
Dim i As Integer

bBC_MakeGood = False
bBC_SoC = False
sAttachmentTableRangeSelection = ""

dFees_GST = 0
dFees_ExGST = 0
dFees_IncGST = 0
dFees_GST_RunningTotal = 0
dFees_ExGST_RunningTotal = 0
dFees_IncGST_RunningTotal = 0
dFees_Discount_GST = 0
dFees_Discount_ExGST = 0
dFees_Discount_IncGST = 0

    If Trim(wsQuote.Range("Country_Selected").Value) = "Australia" Or Trim(wsQuote.Range("Country_Selected").Value) = "New Zealand" Then
        
        sSelectedCountry = Trim(wsQuote.Range("Country_Selected").Value)
    
        Set wdApp = New Word.Application
        With wdApp
            .Visible = False
        End With
    
'        sClientCompanyAddress = wsQuote.Range("ClientStreet") & vbCrLf & wsQuote.Range("ClientSuburb") & vbCrLf & wsQuote.Range("ClientState") & " " & wsQuote.Range("ClientPostcode")
        
        ' retrieve the source template
        Select Case sDocType
        Case "TOE_BC"
            Set myDoc = wdApp.Documents.Add(Template:=(wsLists.Range("zzList_TOE_BC")))
            sAssetClass = "BC_"
        Case "TOE_CC"
            Set myDoc = wdApp.Documents.Add(Template:=(wsLists.Range("zzList_TOE_CC")))
            sAssetClass = "CC_"
        Case "TOE_Tax"
            Set myDoc = wdApp.Documents.Add(Template:=(wsLists.Range("zzList_TOE_Tax")))
            sAssetClass = "Tax_"
        End Select
        
        ' holds status of button i.e. true means button pressed
        sRange = "btn_" & sAssetClass & "status_" '& sIndex
        
        dFees_ExGST_RunningTotal = wsQuote.Range("ClientFeeTotal")
        
        'copy Excel ranges to word doc bookmarks
        '************************************************************************************************************************
        With myDoc.Bookmarks
            Select Case sDocType
            Case "TOE_BC"
                
                dFees_GST = 0
                dFees_ExGST = 0
                dFees_IncGST = 0
            
                dFees_GST_RunningTotal = 0
                dFees_ExGST_RunningTotal = 0
                dFees_IncGST_RunningTotal = 0
                
                dFees_Discount_GST = 0
                dFees_Discount_ExGST = 0
                dFees_Discount_IncGST = 0
                
                Application.StatusBar = "Running calcs..."

                ' determines fees that are to be inserted into the template
                For i = 1 To btn_BC_Count
                    If wsLists.Range(sRange & i) = True Then ' insert fee
                        ' if make good is true (id 3), set flag to true
                        If i = 2 Then
                            bBC_SoC = True
                        End If
                        
                        If i = 3 Then
                            bBC_MakeGood = True
                        End If
                        
                        If i = 5 And wsQuote.cboBC_Purpose = "Divestment" Then
                            bBC_TDD_Divestment = True
                        End If
                        
                        ' calculate GST
                        dFees_ExGST = wsQuote.Range("btn_BC_Fee_" & i)
                        dFees_GST = dFees_ExGST * 0.1
                        dFees_IncGST = dFees_ExGST + dFees_GST
                        
                        ' update the running total
                        dFees_ExGST_RunningTotal = dFees_ExGST_RunningTotal + dFees_ExGST
                        dFees_GST_RunningTotal = dFees_GST_RunningTotal + dFees_GST
                        dFees_IncGST_RunningTotal = dFees_IncGST_RunningTotal + dFees_IncGST
                        
                        ' insert figures into template
                        .Item("BC_Fees_ExGST_" & i).Range.InsertAfter Format(dFees_ExGST, "#,###")
                    
                    Else ' remove unused fee row from template
                        .Item("BC_Fees_" & i).Range.Rows.Delete
                    End If
                
                Next i
                
                ' set up paragraphs in the word templates for specific scope of services
                
                Application.StatusBar = "Formatting document..."
                
                If bBC_MakeGood Or bBC_SoC Then
                
                    .Item("Intro_Standard").Range.Delete
                    .Item("theAsset_Standard").Range.Delete
                    .Item("overview_Standard").Range.Delete
                    .Item("scopeOfDescription_Standard_theAsset").Range.Delete
                    .Item("scopeOfService_Standard").Range.Delete
                    .Item("SchedA_ProjLabel_Default").Range.Delete
                    .Item("SchedA_Scope_Default").Range.Rows.Delete
                    .Item("scopeDescription_Standard").Range.Delete
                    .Item("BC_Attachment1_SoS").Range.Delete
                    .Item("tDD_Divestment_Paragraph").Range.Delete
                    
                    .Item("BC_Attachment_ExclQual_2").Range.Delete
                    .Item("BC_Attachment_TC_3").Range.Delete
                    
                    
                    ' ----- MAKE GOOD SCOPE -----
                    If bBC_MakeGood Then
                        .Item("Intro_SoC").Range.Delete
                        .Item("SchedA_Scope_SoC").Range.Rows.Delete
                        .Item("purpose_SoC_Heading").Range.Delete
                        .Item("scopeOfService_Heading").Range.Delete
                        .Item("scopeOfService_SoC").Range.Delete
                        
                        wsMakeGood.Range("description_thePremises_MG").Copy
                        .Item("description_thePremises").Range.Paste
                        
                        .Item("makeGood_SoS_Intro").Range.InsertAfter wsQuote.Range("makeGood_SoS")
                        .Item("makeGood_SoS_thePremises").Range.InsertAfter UCase(wsQuote.Range("makeGood_SoS"))
                        .Item("makeGood_SoS_Fees").Range.InsertAfter wsQuote.Range("makeGood_SoS")
                        .Item("makeGood_SoS_Instruct_S1").Range.InsertAfter wsQuote.Range("makeGood_SoS")
                        .Item("makeGood_SoS_Instruct_S2").Range.InsertAfter wsQuote.Range("makeGood_SoS")
                        
                        wsMakeGood.Range("scopeOfDescription_Stage1_MakeGood").Copy
                        .Item("scopeOfDescription_Stage1_MakeGood").Range.Paste
                        
                        wsMakeGood.Range("scopeOfDescription_Stage2_MakeGood").Copy
                        .Item("scopeOfDescription_Stage2_MakeGood").Range.Paste
                        
    '                    .Item("scopeOfDescription_Stage1_MakeGood").Range.InsertAfter wsMakeGood.Range("scopeOfDescription_Stage1_MakeGood")
    '                    .Item("scopeOfDescription_Stage2_MakeGood").Range.InsertAfter wsMakeGood.Range("scopeOfDescription_Stage2_MakeGood")
                    End If
                    
                    ' ----- SCHEDULE OF CONDITION SCOPE -----
                    If bBC_SoC Then
                        .Item("Intro_MakeGood").Range.Delete
                        .Item("SchedA_Scope_MakeGood").Range.Rows.Delete
                        .Item("scopeOfDescription_MakeGood").Range.Delete
                        .Item("SchedA_ProjLabel_MakeGood").Range.Delete
                    
                        wsSchedOfCondition.Range("description_thePremises_SoC").Copy
                        .Item("description_thePremises").Range.Paste
                        
                        wsSchedOfCondition.Range("scopeOfDescription_Part1_SoC").Copy
                        .Item("purpose_SoC_Description").Range.Paste
                        
                        wsSchedOfCondition.Range("scopeOfDescription_Part2_SoC").Copy
                        .Item("scopeOfService_SoC").Range.Paste
                    End If
                    
                Else
                    ' ----- STANDARD SCOPE -----
                    .Item("Intro_MakeGood").Range.Delete
                    .Item("Intro_SoC").Range.Delete
                    .Item("thePremises_Heading").Range.Delete
                    .Item("purpose_SoC_Heading").Range.Delete
                    .Item("scopeOfDescription_MakeGood").Range.Delete
                    .Item("scopeOfDescription_thePremises").Range.Delete
                    .Item("SchedA_ProjLabel_MakeGood").Range.Delete
                    .Item("SchedA_Scope_MakeGood").Range.Rows.Delete
                    .Item("SchedA_Scope_SoC").Range.Rows.Delete
                    
                    .Item("BC_Attachment_ExclQual_1").Range.Delete
                    .Item("BC_Attachment_TC_2").Range.Delete
                    
                    ' TDD and Divestment
                    If bBC_TDD_Divestment = False Then
                        .Item("tDD_Divestment_Paragraph").Range.Delete
                    End If
                    
                    .Item("Intro_Standard_Purpose").Range.InsertAfter wsQuote.Range("BC_Purpose")
                    
                    wsAssetDesc.Range("assetDescription").Copy
                    .Item("assetDescription").Range.Paste
                    
                    wsScopeDesc.Range("scopeDescription_Standard").Copy
                    .Item("scopeDescription_Standard").Range.Paste
                    
                End If
                
                ' test to see whether there are any additional scopes/services i.e. BCA, Enviro etc.
                bAdditionalScope = False
                
                ' insert subconsultant fees/additional scope
                For i = 1 To BC_AddScope_Count
                    If wsLists.Range("BC_AddScope_Fees_Status_" & i) = True Then
                        .Item("BC_AddScope_Fees_ExGST_" & i).Range.InsertAfter Format(wsLists.Range("BC_AddScope_Fees_Status_" & i).Offset(0, 1), "#,###")
                        bAdditionalScope = True
                    Else ' delete wording from template
                        .Item("BC_AddScope_Fees_" & i).Range.Rows.Delete
                        .Item("BC_AddScope_Instruct_Fees_" & i).Range.Rows.Delete
                    End If
                Next i
                
                ' if there were no additional scope/services, delete the additional scope heading
                If bAdditionalScope = False Then
                    .Item("BC_AddScope_Fees_All").Range.Rows.Delete
                    .Item("BC_AddScope_SchedA_All").Range.Rows.Delete
                End If
            
                ' enter CBRE details
                If sSelectedCountry = "Australia" Then
                    ' remove NZ tags
                    .Item("businessNo_NZ").Range.Delete
                    .Item("cbreLtd_NZ_1").Range.Delete
                    .Item("cbreLtd_NZ_3").Range.Delete
                    .Item("cbreLtd_NZ_4").Range.Delete
                    .Item("officeAddress_website_NZ").Range.Delete
                Else ' New Zealand
                    ' remove Aust tags
                    .Item("businessNo_Aust").Range.Delete
                    .Item("cbreLtd_Aust_1").Range.Delete
                    .Item("cbreLtd_Aust_3").Range.Delete
                    .Item("cbreLtd_Aust_4").Range.Delete
                    .Item("officeAddress_website_Aust").Range.Delete
                End If
            
                ' check for any discounts
                If wsQuote.Range("ClientFeeTotalDiscountPerc") > 0 Then
                    ' calculate the the discount
                    dFees_Discount_ExGST = dFees_ExGST_RunningTotal * wsQuote.Range("ClientFeeTotalDiscountPerc")
                    
                    ' calculate the total minus the discount
                    dFees_ExGST_RunningTotal = dFees_ExGST_RunningTotal - dFees_Discount_ExGST
                    
                    ' insert discount total
                    .Item("BC_Fees_ExGST_Discount").Range.InsertAfter Format(dFees_Discount_ExGST, "#,###")
                Else
                    .Item("BC_Fees_Discount").Range.Rows.Delete
                End If
                
                ' insert total fees
                .Item("BC_Fees_ExGST_Total").Range.InsertAfter Format(dFees_ExGST_RunningTotal, "#,###")
                        
            Case "TOE_CC"
            
                dFees_GST = 0
                dFees_ExGST = 0
                dFees_IncGST = 0
            
                dFees_GST_RunningTotal = 0
                dFees_ExGST_RunningTotal = 0
                dFees_IncGST_RunningTotal = 0
            
                dFees_Discount_GST = 0
                dFees_Discount_ExGST = 0
                dFees_Discount_IncGST = 0
                
                Application.StatusBar = "Running calcs..."
                
                ' determines fees that are to be inserted into the template
                For i = 1 To btn_CC_Count
                    If wsLists.Range(sRange & i) = True Then ' insert fee
                    
                        'bCC_SoS = True
                        
                        ' calculate fees
                        dFees_ExGST = wsQuote.Range("btn_CC_Fee_" & i)
                        dFees_GST = dFees_ExGST * 0.1
                        dFees_IncGST = dFees_ExGST + dFees_GST
                        
                        ' update the running total
                        dFees_ExGST_RunningTotal = dFees_ExGST_RunningTotal + dFees_ExGST
                        dFees_GST_RunningTotal = dFees_GST_RunningTotal + dFees_GST
                        dFees_IncGST_RunningTotal = dFees_IncGST_RunningTotal + dFees_IncGST
                        
                        ' insert figures into template
                        .Item("CC_Fees_ExGST_" & i).Range.InsertAfter Format(dFees_ExGST, "#,###")
                        .Item("CC_Fees_GST_" & i).Range.InsertAfter Format(dFees_GST, "#,###")
                        .Item("CC_Fees_IncGST_" & i).Range.InsertAfter Format(dFees_IncGST, "#,###")
                    
                    Else ' remove unused fee row from template
                        .Item("CC_Fees_" & i).Range.Rows.Delete
                    End If
                
                Next i
            
                ' check for any discounts
                If wsQuote.Range("ClientFeeTotalDiscountPerc") > 0 Then
                    ' calculate the the discount
                    dFees_Discount_ExGST = dFees_ExGST_RunningTotal * wsQuote.Range("ClientFeeTotalDiscountPerc")
                    dFees_Discount_GST = dFees_GST_RunningTotal * wsQuote.Range("ClientFeeTotalDiscountPerc")
                    dFees_Discount_IncGST = dFees_IncGST_RunningTotal * wsQuote.Range("ClientFeeTotalDiscountPerc")
                    
                    ' calculate the total minus the discount
                    dFees_ExGST_RunningTotal = dFees_ExGST_RunningTotal - dFees_Discount_ExGST
                    dFees_GST_RunningTotal = dFees_GST_RunningTotal - dFees_Discount_GST
                    dFees_IncGST_RunningTotal = dFees_IncGST_RunningTotal - dFees_Discount_IncGST
                    
                    ' insert discount total
                    .Item("CC_Fees_ExGST_Discount").Range.InsertAfter Format(dFees_Discount_ExGST, "#,###")
                    .Item("CC_Fees_GST_Discount").Range.InsertAfter Format(dFees_Discount_GST, "#,###")
                    .Item("CC_Fees_IncGST_Discount").Range.InsertAfter Format(dFees_Discount_IncGST, "#,###")
                Else
                    .Item("CC_Fees_Discount").Range.Rows.Delete
                End If
                
                Application.StatusBar = "Formatting document..."
                
                ' insert Scope of Service table
                wsScopeOfService.Range("ScopeOfService_CC").Copy
                .Item("scopeOfService").Range.Paste
                
                ' insert total fees
                .Item("CC_Fees_ExGST_Total").Range.InsertAfter Format(dFees_ExGST_RunningTotal, "#,###")
                .Item("CC_Fees_GST_Total").Range.InsertAfter Format(dFees_GST_RunningTotal, "#,###")
                .Item("CC_Fees_IncGST_Total").Range.InsertAfter Format(dFees_IncGST_RunningTotal, "#,###")
            
                ' enter CBRE details
                If sSelectedCountry = "Australia" Then
                    ' remove NZ tags
                    .Item("businessNo_NZ").Range.Delete
                    .Item("cbreLtd_NZ_1").Range.Delete
                    .Item("cbreLtd_NZ_2").Range.Delete
                    .Item("cbreLtd_NZ_3").Range.Delete
                    .Item("cbreLtd_NZ_4").Range.Delete
                    .Item("officeAddress_website_NZ").Range.Delete
                Else ' New Zealand
                    ' remove Aust tags
                    .Item("businessNo_Aust").Range.Delete
                    .Item("cbreLtd_Aust_1").Range.Delete
                    .Item("cbreLtd_Aust_2").Range.Delete
                    .Item("cbreLtd_Aust_3").Range.Delete
                    .Item("cbreLtd_Aust_4").Range.Delete
                    .Item("officeAddress_website_Aust").Range.Delete
                End If
            
            Case "TOE_Tax"
            
                dFees_GST = 0
                dFees_ExGST = 0
                dFees_IncGST = 0
            
                dFees_GST_RunningTotal = 0
                dFees_ExGST_RunningTotal = 0
                dFees_IncGST_RunningTotal = 0
            
                dFees_Discount_GST = 0
                dFees_Discount_ExGST = 0
                dFees_Discount_IncGST = 0
                
                ' holds status of button i.e. true means button pressed
                sRange = "btn_" & sAssetClass & "status_"
                
                Application.StatusBar = "Running calcs..."
                
                ' determines fees that are to be inserted into the template
                For i = 1 To btn_Tax_Count
                    If wsLists.Range(sRange & i) = True Then ' insert fee
                        ' check Attachment Table status whether to insert table or not
                        If wsAttachment2.Range("includeAttachment2Table") = "Yes" Then
                        
                            ' determine whether attachment table 2 is required. If required, which table to insert
                            ' if refurbishment and construction costs are selected, insert attachment2Table_CapExpAssess i.e
                            ' construction - buttons 3
                            ' refurbishment - buttons 9
                            ' otherwise, default to attachment2Table_InfoReq
                            ' a blank field equates to no attachment table
                            If i = 3 Or i = 9 Then
                                sAttachmentTableRangeSelection = "attachment2Table_CapExpAssess"
                            Else
                                sAttachmentTableRangeSelection = "attachment2Table_InfoReq"
                            End If
                        
                        End If
                        
                        
                        ' calculate GST
                        dFees_ExGST = wsQuote.Range("btn_Tax_Fee_" & i)
                        dFees_GST = dFees_ExGST * 0.1
                        dFees_IncGST = dFees_ExGST + dFees_GST
                        
                        ' update the running total
                        dFees_ExGST_RunningTotal = dFees_ExGST_RunningTotal + dFees_ExGST
                        dFees_GST_RunningTotal = dFees_GST_RunningTotal + dFees_GST
                        dFees_IncGST_RunningTotal = dFees_IncGST_RunningTotal + dFees_IncGST
                        
                        ' insert figures into template
                        .Item("TaxDep_Fees_ExGST_" & i).Range.InsertAfter Format(dFees_ExGST, "#,###")
                        .Item("TaxDep_Fees_GST_" & i).Range.InsertAfter Format(dFees_GST, "#,###")
                        .Item("TaxDep_Fees_IncGST_" & i).Range.InsertAfter Format(dFees_IncGST, "#,###")
                    
                    Else ' remove unused fee row from template
                        .Item("TaxDep_Fees_" & i).Range.Rows.Delete
                    End If
                
                Next i
                
                ' check for any discounts
                If wsQuote.Range("ClientFeeTotalDiscountPerc") > 0 Then
                    ' calculate the the discount
                    dFees_Discount_ExGST = dFees_ExGST_RunningTotal * wsQuote.Range("ClientFeeTotalDiscountPerc")
                    dFees_Discount_GST = dFees_GST_RunningTotal * wsQuote.Range("ClientFeeTotalDiscountPerc")
                    dFees_Discount_IncGST = dFees_IncGST_RunningTotal * wsQuote.Range("ClientFeeTotalDiscountPerc")
                    
                    ' calculate the total minus the discount
                    dFees_ExGST_RunningTotal = dFees_ExGST_RunningTotal - dFees_Discount_ExGST
                    dFees_GST_RunningTotal = dFees_GST_RunningTotal - dFees_Discount_GST
                    dFees_IncGST_RunningTotal = dFees_IncGST_RunningTotal - dFees_Discount_IncGST
                    
                    ' insert discount total
                    .Item("TaxDep_Fees_ExGST_Discount").Range.InsertAfter Format(dFees_Discount_ExGST, "#,###")
                    .Item("TaxDep_Fees_GST_Discount").Range.InsertAfter Format(dFees_Discount_GST, "#,###")
                    .Item("TaxDep_Fees_IncGST_Discount").Range.InsertAfter Format(dFees_Discount_IncGST, "#,###")
                Else
                    .Item("TaxDep_Fees_Discount").Range.Rows.Delete
                End If
                
                Application.StatusBar = "Formatting document..."
                
                ' insert total fees
                .Item("TaxDep_Fees_ExGST_Total").Range.InsertAfter Format(dFees_ExGST_RunningTotal, "#,###")
                .Item("TaxDep_Fees_GST_Total").Range.InsertAfter Format(dFees_GST_RunningTotal, "#,###")
                .Item("TaxDep_Fees_IncGST_Total").Range.InsertAfter Format(dFees_IncGST_RunningTotal, "#,###")
                
                ' insert Attachment 2 table
                If sAttachmentTableRangeSelection <> "" Then
                    wsAttachment2.Range(sAttachmentTableRangeSelection).Copy
                    .Item("attachment2_Table").Range.Paste
                Else ' delete paragraphs and page
                    .Item("Attachment2Ref_TermsOfPayment").Range.Delete
                    .Item("Attachment2Ref_SchedAInstruct").Range.Delete
                    .Item("attachment2_Table_Section").Range.Delete
                End If

                
                ' enter CBRE details
                If sSelectedCountry = "Australia" Then
                    ' remove NZ tags
                    .Item("acquisitionAssessDeprec_NZ").Range.Delete
                    .Item("businessNo_NZ").Range.Delete
                    .Item("cbreLtd_NZ_1").Range.Delete
                    .Item("cbreLtd_NZ_2").Range.Delete
                    .Item("cbreLtd_NZ_3").Range.Delete
                    .Item("cbreLtd_NZ_4").Range.Delete
                    .Item("officeAddress_website_NZ").Range.Delete
                Else ' New Zealand
                    ' remove Aust tags
                    .Item("acquisitionAssessDeprec_Aust").Range.Delete
                    .Item("businessNo_Aust").Range.Delete
                    .Item("cbreLtd_Aust_1").Range.Delete
                    .Item("cbreLtd_Aust_2").Range.Delete
                    .Item("cbreLtd_Aust_3").Range.Delete
                    .Item("cbreLtd_Aust_4").Range.Delete
                    .Item("officeAddress_website_Aust").Range.Delete
                End If
                
            
            End Select
            
            ' COMMON ITEMS IN ALL TOE docs
            ' ----------------------------
            ' populate scope description
             .Item(sAssetClass & "CoreScopeDescription").Range.InsertAfter wsQuote.Range("AssetClass_" & sAssetClass & "Selected")
            
            ' populate primary contact details
            .Item("officeAddress_Line1").Range.InsertAfter wsLists.Range("PrimContact_Name")
            .Item("officeAddress_Line2").Range.InsertAfter wsLists.Range("OfficeAddress_Line2")
            .Item("officeAddress_Line3").Range.InsertAfter wsLists.Range("OfficeAddress_Line3")
            .Item("officeAddress_mobile").Range.InsertAfter wsLists.Range("PrimContact_ContactNo")
            .Item("officeAddress_email").Range.InsertAfter wsLists.Range("PrimContact_Email")
            .Item("officeAddress_contactNo").Range.InsertAfter wsLists.Range("OfficeAddress_contactNo")
            
            ' primary operator
            .Item("SignOff_Name").Range.InsertAfter wsLists.Range("PrimContact_Name")
            .Item("SignOff_Title").Range.InsertAfter wsLists.Range("PrimContact_Title")
            
            ' if the Signatory differs to the Primary Operator, need to add a second signatory at sign off
            If wsLists.Range("PrimContact_Name") <> wsLists.Range("zzList_LOE_Signatory_Name") And wsLists.Range("zzList_LOE_Signatory_Name") <> "" Then
                
                ' Signatory
                .Item("SignatorySignOff_Name").Range.InsertAfter wsLists.Range("zzList_LOE_Signatory_Name")
                .Item("SignatorySignOff_Title").Range.InsertAfter wsLists.Range("zzList_LOE_Signatory_Title")
            Else
                ' delete signatory section
                .Item("Signatory_Section").Range.Delete
            End If
            
            ' client details
            
            '.Item("referenceNo").Range.InsertAfter wsQuote.Range("ClientCompany")
            sFirstName = Left(wsQuote.Range("ClientName"), InStr(wsQuote.Range("ClientName"), " "))
            .Item("salutation").Range.InsertAfter Trim(sFirstName) 'gsClientFirstName
            
            .Item("clientCompany").Range.InsertAfter wsQuote.Range("ClientCompany") 'gsClientCompany
            .Item("clientCompanyAddressLine1").Range.InsertAfter wsQuote.Range("ClientAddressLine1") 'gsClientAddressLine1
            .Item("clientCompanyAddressLine2").Range.InsertAfter wsQuote.Range("ClientAddressLine2") 'gsClientAddressLine2
            
            sClientAddressLine3 = wsQuote.Range("ClientSuburb") & " " & wsQuote.Range("ClientState") & " " & wsQuote.Range("ClientPostcode")
            .Item("clientCompanyAddressLine3").Range.InsertAfter sClientAddressLine3 'gsClientSuburb & " " & gsClientState & " " & gsClientPostcode
            .Item("clientEmailAddress").Range.InsertAfter wsQuote.Range("ClientEmailAddress") 'gsClientEmailAddress
                
            .Item("Signatory_Name").Range.InsertAfter wsLists.Range("PrimContact_Name")
            .Item("Signatory_Email").Range.InsertAfter wsLists.Range("PrimContact_Email")
            
            ' CBRE contact name, number and email
            .Item("contact_Name").Range.InsertAfter wsLists.Range("PrimContact_Name")
            .Item("contact_Number").Range.InsertAfter wsLists.Range("PrimContact_ContactNo")
            .Item("contact_Email").Range.InsertAfter wsLists.Range("PrimContact_Email")
            
            
            ' need to compare client name and company before assigning the client's position
            If wsQuote.Range("ClientName") = Trim(gsClientFirstName) & " " & Trim(gsClientLastName) And _
                wsQuote.Range("ClientCompany") = Trim(gsClientCompany) Then
                
                .Item("clientName").Range.InsertAfter " - " & gsClientPosition
            End If
            .Item("clientName").Range.InsertAfter wsQuote.Range("ClientName") 'gsClientFirstName & " " & gsClientLastName
            
            
            ' check whether job is for a portfolio or a single property
            If wsLists.Range("zzPFStatus") = True Then
                .Item("propertyDetails").Range.InsertAfter wsQuote.Range("PFName")
                .Item("propertyDetails2").Range.InsertAfter wsQuote.Range("PFName")
                sMID = wsQuote.Range("PFAddress_01_MID")
                
                ' populate the Invoice Wording cell
                'Call populate_InvoiceWording
                'wsQuote.Range("InvoiceWording") = wsQuote.Range("PFName") & chr(10)
            Else
                sPropAddress = wsQuote.Range("PFAddress_01") & " " & wsQuote.Range("PFAddress_01_Postcode")
                .Item("propertyDetails").Range.InsertAfter sPropAddress
                .Item("propertyDetails2").Range.InsertAfter sPropAddress
                sMID = wsQuote.Range("PFAddress_01_MID")
                
            End If
            
            sRange = getPortfolioPropertiesRange
            
            wsLists.Visible = xlSheetVisible
            wsLists.Range("zzPropertyList_Header").Copy
            .Item("PropertyList_Header").Range.Paste
            
            ' paste the list of properties in the quoted job
            wsLists.Range(sRange).Copy
            .Item("propertyList").Range.Paste

            wsLists.Visible = xlSheetHidden
        
        End With
        '************************************************************************************************************************
            
        Application.StatusBar = "Saving document..."
            
        Application.DisplayAlerts = False
        If Len(Dir(wsQuote.Range("zzListFilePath") & sAssetClass & sFileName & ".docx")) > 0 Then
            ' if file exists, remove it before saving
            Kill wsQuote.Range("zzListFilePath") & sAssetClass & sFileName & ".docx"
        End If
        
        ' save file to user specified path
        myDoc.SaveAs2 Filename:=wsQuote.Range("zzListFilePath") & sAssetClass & sFileName & ".docx", FileFormat:=wdFormatDocumentDefault
        myDoc.Close
        Application.DisplayAlerts = True
    
        ' pass back the full path name
        prepareTemplate = wsQuote.Range("zzListFilePath") & sAssetClass & sFileName & ".docx"
    Else
        MsgBox "Please select a valid country", vbCritical, "Invalid Country selected"
        GoTo Exit_prepareTemplate
    End If

Exit_prepareTemplate:
    Set myDoc = Nothing
    Set wdApp = Nothing
    Exit Function

Err_prepareTemplate:
    MsgBox "Error Number: " & Err.Number & chr(13) & "Function: prepareTemplate" & chr(13) & "Description: " & Err.Description
    Resume Exit_prepareTemplate


End Function
Public Sub delete_setAssetClassTypeStatus_Single(sButtonName As String)
Dim iPos As String
Dim sIndex As String
Dim sButtonNamePrefix As String
Dim sRange As String
Dim sColRange As String
Dim sBusLine As String

iPos = InStrRev(sButtonName, "_")
sButtonNamePrefix = Left(sButtonName, iPos)

' get the button index
sIndex = Right(sButtonName, Len(sButtonName) - iPos)

' holds status of button i.e. true means button pressed
sRange = sButtonNamePrefix & "status_" & sIndex

' column range name for fee
sColRange = sButtonNamePrefix & "column_" & sIndex

wsQuote.Shapes.Range(Array(sButtonName)).Select
If wsLists.Range(sRange) = True Then
    
    ' format button
    Selection.Font.ColorIndex = 1
    Selection.Font.FontStyle = "Book"
    wsLists.Range(sRange) = False

    ' hide fee column
    wsQuote.Range(sColRange).Select
    Selection.EntireColumn.Hidden = True

Else
    ' format button
    Selection.Font.ColorIndex = 5
    Selection.Font.FontStyle = "Book Bold"
    wsLists.Range(sRange) = True

    ' unhide fee column
    wsQuote.Range(sColRange).Select
    Selection.EntireColumn.Hidden = False
End If


End Sub

Public Function validateTOE_Templates() As String

Dim sErrorLog As String

sErrorLog = ""

' VALIDATE TOE FILES
' ------------------

' TOE - Building Consultancy
If Len(Dir(wsLists.Range("zzList_TOE_BC"))) = 0 Then
    sErrorLog = sErrorLog & " - Building Consultancy TOE document" & chr(13)
End If

' TOE - Cost Consultancy
If Len(wsLists.Range("zzList_TOE_CC")) > 0 Then
        
    ' check if last character is "\" and temporarily remove if true.
    ' if the last character is "\", DIR will also validate the string as a directory.
    ' need to validate as file
    If validateLastCharacter("\", wsLists.Range("zzList_TOE_CC")) Then
        If Dir(Left(wsLists.Range("zzList_TOE_CC"), Len(wsLists.Range("zzList_TOE_CC")) - 1), vbNormal) = "" Then
            sErrorLog = sErrorLog & " - Cost Consultancy TOE document" & chr(13)
        End If
    Else
        If Dir(wsLists.Range("zzList_TOE_CC"), vbNormal) = "" Then
            sErrorLog = sErrorLog & " - Cost Consultancy TOE document" & chr(13)
        End If
    End If

Else
    sErrorLog = sErrorLog & " - Cost Consultancy TOE document" & chr(13)
End If

' TOE - Tax Consultancy
If Len(Dir(wsLists.Range("zzList_TOE_Tax"))) = 0 Then
    sErrorLog = sErrorLog & " - Tax Consultancy TOE document" & chr(13)
End If

If Len(sErrorLog) > 0 Then
    validateTOE_Templates = "Missing Templates" & chr(13) & sErrorLog
End If

End Function

Public Function validateLastCharacter(sChar As String, sCompStr As String) As Boolean

If Right(sCompStr, 1) = sChar Then
    validateLastCharacter = True
Else
    validateLastCharacter = False
End If

End Function

Public Sub test()

Dim rCell As Range


If IsNumeric(wsQuote.Range("ClientName")) Then
    'retrieve contact name
    For Each rCell In wsList_CommonClients.Range("zzList_CommonClients_Index")
        If CInt(wsQuote.Range("ClientName")) = rCell Then
            Application.EnableEvents = False
            wsQuote.cboCompanyContact.Clear
            wsQuote.cboCompanyContact.AddItem rCell.Offset(0, 2)
            Application.EnableEvents = True
            Exit Sub
        End If
    Next rCell
    
End If



End Sub


Public Sub test2()
Dim sTemp As String

'sTemp = getPortfolioPropertiesRange

'Debug.Print sTemp

ActiveSheet.Range("E53") = "Test" & chr(10) & "Test2"


End Sub



