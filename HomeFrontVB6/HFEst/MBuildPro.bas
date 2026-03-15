Attribute VB_Name = "MBuildPro"
Option Explicit
Const SRCFILE = "MBuildPro::"

Public Sub SendBuildProPOIndexes()
    Call HFApp.RunTask("SendBuildProPOIndexes")
End Sub

Public Sub SendBuildProVendors()
    Call HFApp.RunTask("SendBuildProVendors")
End Sub

Public Sub SendBuildProRemittances()
    Call HFApp.RunTask("SendBuildProRemittances")
End Sub

Public Sub SendBuildProCommunities()
    Call HFApp.RunTask("SendBuildProCommunities")
End Sub

Public Sub SendBuildProJobs(Job As String)
    Call HFApp.RunTask("SendBuildProJobs|" & Job)
End Sub

Public Sub CancelBuildProPO(Job As String, PONumber As String)
    Call HFApp.RunTask("CancelBuildProPO|" & Job & "|" & PONumber)
End Sub

Public Sub SendBuildProPOs(Job As String, POs As String)
    Call HFApp.RunTask("SendBuildProPOs|" & Job & "|" & POs)
End Sub





