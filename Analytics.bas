Attribute VB_Name = "Analytics"
Option Private Module
Option Explicit

'Variable to hold instance of class AnalyticsApp
Dim collectorApp As AnalyticsApp

Public Sub Init()
    'Reset collectorApp in case it is already loaded
    Set collectorApp = Nothing
    'Create a new instance of Analytics App
    Set collectorApp = New AnalyticsApp
    'Pass the Excel object to it so it knows what application
    'it needs to respond to
    Set collectorApp.App = Application
    'Workbook Open, Activate and New Events do not fire the first time through. This is to ensure analytics are always sent.
    'Sending Analytics for first workbook opened
    On Error Resume Next
    collectorApp.SendAnalytics "open"
End Sub
Public Sub TrackEvent(EvtTrigger As String)
    'Reset collectorApp in case it is already loaded
    Set collectorApp = Nothing
    'Create a new instance of Analytics App
    Set collectorApp = New AnalyticsApp
    'Pass the Excel object to it so it knows what application
    'it needs to respond to
    Set collectorApp.App = Application
    'Workbook Open, Activate and New Events do not fire the first time through. This is to ensure analytics are always sent.
    'Sending Analytics for first workbook opened
    On Error Resume Next
    collectorApp.SendAnalytics EvtTrigger
End Sub

Public Sub TestEventPost()
TrackEvent "Import Tracking Test v2"
End Sub

