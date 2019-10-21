VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalyticsApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This class describes the Analytics Application
Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1

Private Sub App_ProtectedViewWindowOpen(ByVal Pvw As ProtectedViewWindow)
On Error Resume Next
SendAnalytics "protected open"
End Sub

Private Sub App_WorkbookOpen(ByVal wb As Workbook)
On Error Resume Next
SendAnalytics "Open Excel"
End Sub

Private Sub Class_Terminate()
Set App = Nothing
End Sub

Function GetProperty(PropertyName As String, _
    Optional WhatWorkbook As Workbook) As Variant
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GetProperty
' This procedure returns the value of a DocumentProperty named in
' PropertyName. It will examine BuiltinDocumentProperties,
' or CustomDocumentProperties, or both. The parameters are:
'
'   PropertyName        The name of the property to return.
'
'   PropertySet         One of PropertyLocationBuiltIn,
'                       PropertyLocationCustom, or PropertyLocationBoth.
'                       This specifies the property set to search.
'
'   WhatWorkbook        A reference to the workbook whose properties
'                       are to be examined. If omitted or Nothing,
'                       ThisWorkbook is used.
'
' The function will return:
'
'   The value of property named by PropertyName, or
'
'   #VALUE if the PropertySet parameter is not valid (test with IsError), or
'
'   Null if the property could not be found (test with IsNull)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim wb As Workbook
Dim Props1 As Office.DocumentProperties
Dim Props2 As Office.DocumentProperties
Dim Prop As Office.DocumentProperty

'''''''''''''''''''''''''''''''''''''''''
' Set the workbook whose properties we
' will search.
'''''''''''''''''''''''''''''''''''''''''
If WhatWorkbook Is Nothing Then
    Set wb = ThisWorkbook
Else
    Set wb = WhatWorkbook
End If


'''''''''''''''''''''''''''''''''''''''''
' Looking in both Built-in and
' Custom Document Properties
'''''''''''''''''''''''''''''''''''''''''

Set Props1 = wb.BuiltinDocumentProperties
Set Props2 = wb.CustomDocumentProperties


On Error Resume Next
'''''''''''''''''''''''''''''''''''''''''
' Search either BuiltIn or Custom.
'''''''''''''''''''''''''''''''''''''''''
Set Prop = Props1(PropertyName)
If Err.Number <> 0 Then
    ''''''''''''''''''''''''''''''''''
    ' Not found in one set. See if
    ' we need to look in the other.
    ''''''''''''''''''''''''''''''''''
    If Not Props2 Is Nothing Then
        ''''''''''''''''''''''''''''''''''''
        ' We'll get here only if both Custom
        ' and BuiltIn properties are to be
        ' searched.
        ''''''''''''''''''''''''''''''''''''
        Err.Clear
        Set Prop = Props2(PropertyName)
        If Err.Number <> 0 Then
            ''''''''''''''''''''''''''''''''''''
            ' Property not found. Return NULL.
            ''''''''''''''''''''''''''''''''''''
            GetProperty = Null
            Exit Function
        End If
    Else
        ''''''''''''''''''''''''''''''''''''
        ' Property not found. Return NULL.
        ''''''''''''''''''''''''''''''''''''
        GetProperty = Null
        Exit Function
    End If
End If

''''''''''''''''''''''''''''''''''''
' Property found. Return the value.
''''''''''''''''''''''''''''''''''''
GetProperty = Prop.Value

End Function

Function postData(EvtCategory As String, EvtType As String) As String
    Dim sMachineName, sUserName As String
    Dim wkbName, wksName, wkbTitle, wkbSubject, wkbAuthor, wkbCompany As String
    Dim sSpecificUse, sPropClass, sJobNum, sFileNum, sPropertyDescription, sProductCategory, evtTitle As String
    Dim xlVersion, xlOS As String
    

    On Error Resume Next
    ' Set Event Variables for Google Analytics
    sMachineName = CreateObject("WScript.Network").ComputerName
    sUserName = CreateObject("WScript.Network").UserName
    
    'Get Excel Version Information
    xlVersion = App.Version
    xlOS = App.OperatingSystem
    xlCompany = App.OrganizationName

    'Gets workbook name and path of the Workbook
    wkbName = ActiveWorkbook.Name
    wksName = ActiveWorkbook.ActiveSheet
    
    'Get Preset Document Properties
    wkbTitle = GetProperty("Title", ActiveWorkbook)
    wkbSubject = GetProperty("Subject", ActiveWorkbook)
    wkbAuthor = GetProperty("Author", ActiveWorkbook)
    wkbCompany = GetProperty("Company", ActiveWorkbook)

    'Get key facts from Excel file
    
    sJobNum = ActiveWorkbook.Names("Job_No.").RefersToRange.Value
    sFileNum = ActiveWorkbook.Names("File_No.").RefersToRange.Value
    sPropClass = ActiveWorkbook.Names("Property_Class").RefersToRange.Value
    sSpecificUse = ActiveWorkbook.Names("Specific_Use").RefersToRange.Value
    sPropertyDescription = ActiveWorkbook.Names("Property_Description").RefersToRange.Value
    
    sProductCategory = sPropClass + "/" + sSpecificUse
    If sProductCategory = "/" Then
    sProductCategory = "none"
    Else
    'Do Nothing
    End If

    
    'Clean up eventTitle when Workbook doesn't have a title set
    If wkbTitle = "" Then
    evtTitle = "excel"
    Else
        evtTitle = wkbTitle
    End If
    
    
        'Creates the full URL for posting to Google Analytics to track the "Page View"
    'GA ID
    'CID -> Anonymous User ID
    't -> Tracking event type. Set to pageview
    'dp -> Tracked Page set as Workbook Name
    
    If EvtCategory = "event" Then
    'Returns a Google Analytics Event
'    postData = "v=1&tid=UA-149168632-1&cid=" + sMachineName + "&t=event&ec=" + evtTitle + "&ea=" + EvtType + "&el=" + wkbName + "&ua=Excel/(" + xlVersion + ")"
    postData = "v=1&tid=UA-93507010-3&cid=" + sMachineName + "&t=event&ec=" + evtTitle + "&ea=" + EvtType + "&el=" + wkbName + "&ua=Excel/(" + xlVersion + ")"
    Else
    ' Returns a Google Analytics Page View Event
    postData = "v=1&tid=UA-93507010-3&cid=" + sMachineName + "&t=pageview&dp=" + wkbName + "&dt=" + evtTitle + "&cd1=" + sJobNum + "&cd2=" + sFileNum + "&ua=Excel/" + xlVersion + "&pr1nm=" + evtTitle + "&pr1br=" + wkbCompany + "&pa=detail" + "&pr1ca=" + sProductCategory
    End If
   
    
End Function
Sub SendAnalytics(ByVal EvtType As String)

    Dim Result, myURL, sPageView, sEvent As String
    Dim winHttpReq As Object
    
    'Create the Page View HTTP POST Data
    sPageView = postData("pageview", EvtType)
    
   'Create the Event HTTP POST data
    sEvent = postData("event", EvtType)
    
    'Sets up an HTTP Request object for POST
    Set winHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
    'Base URL for Google Analytics
    myURL = "https://www.google-analytics.com/collect"
    
    'Use the below entry for debugging
    'Requires display of result to Messagebox or debugger
    'myURL = "https://www.google-analytics.com/debug/collect"
    
    'Builds the HTTP Request object
    winHttpReq.Open "POST", myURL, False
    winHttpReq.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    
    'Post Page View
    winHttpReq.send (sPageView)
    Debug.Print sPageView
    Result = winHttpReq.responseText
    'MsgBox Result & Chr(10) & sPageView
    
    'Post Event
    winHttpReq.send (sEvent)
    Debug.Print sPageView
    Result = winHttpReq.responseText
    'MsgBox Result & Chr(10) & sPageView
    
End Sub
