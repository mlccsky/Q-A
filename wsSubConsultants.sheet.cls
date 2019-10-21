Private Sub cboReportType_Click()
    ResetButton wsSubConsultants.cboReportType
End Sub

Private Sub cboSurveyors_Click()
    ResetButton wsSubConsultants.cboSurveyors
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Call resetSubconsultantRoles
End Sub

