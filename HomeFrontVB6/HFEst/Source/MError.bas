Attribute VB_Name = "MError"
Option Explicit

Public Sub errHandler(Source As String, Optional sql As String)
    Dim bSqlError  As Boolean
    Dim sErrNumber As Long
    Dim sErrDesc   As String
    Dim sErrSource As String
    
    Dim s As String
        
    sErrNumber = Err.Number
    sErrSource = Err.Source
    sErrDesc = Err.Description
    
    LogError Source, sErrNumber, sErrDesc
    Screen.MousePointer = vbDefault
        
    s = ""
    s = s & App.ProductName & " has encountered a problem and cannot complete the requested action." & vbCrLf
    
    s = s & vbCrLf
    s = s & "---ERROR----------------------------------------------------" & vbCrLf
    s = s & sErrSource & vbCrLf
    s = s & sErrDesc & "   (" & sErrNumber & ")" & vbCrLf
   
    s = s & vbCrLf
    s = s & "---SOURCE---------------------------------------------------" & vbCrLf
    s = s & Source & "()" & vbCrLf
    
    If sql <> "" Then
        s = s & vbCrLf
        s = s & "---DETAILS----------------------------------------------" & vbCrLf
        s = s & sql & vbCrLf
    End If

    MsgBox s, vbCritical, App.ProductName & " (version " & App.Major & "." & format(App.Minor, "00") & "." & format(App.Revision, "0000") & ")"
    
End Sub


Public Sub LogError(Source As String, Number As Long, Description As String, Optional sql As String)
On Error Resume Next

    Dim fn As Long
    
    fn = FreeFile
    Open App.Path & "\error.txt" For Append As #fn
    
    Print #fn, Now & vbTab & _
               Source & vbTab & _
               Number & vbTab & _
               Description & vbTab & _
               sql
    
    Close fn
    
End Sub
