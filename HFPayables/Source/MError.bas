Attribute VB_Name = "MError"
Option Explicit

Public Sub ErrHandler(Source As String, Optional AdditionalInfo As String = "")
    Dim bSqlError  As Boolean
    Dim sErrNumber As Long
    Dim sErrDesc   As String
    Dim sErrSource As String
    
    Dim s As String
        
    sErrNumber = err.Number
    sErrSource = err.Source
    sErrDesc = err.Description
    
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
    
    If AdditionalInfo <> "" Then
        s = s & vbCrLf
        s = s & "---DETAILS----------------------------------------------" & vbCrLf
        s = s & AdditionalInfo & vbCrLf
    End If

    MsgBox s, vbCritical, App.ProductName & " (version " & App.Major & "." & Format(App.Minor, "00") & "." & Format(App.Revision, "0000") & ")"
        

End Sub


Private Sub LogError(Source As String, Number As Long, Description As String, Optional Sql As String)
On Error Resume Next

    Dim fn As Long
    
    fn = FreeFile
    Open AppWorkingFolder & "\error.txt" For Append As #fn
    
    Print #fn, Now & vbTab & _
               Source & vbTab & _
               Number & vbTab & _
               Description & vbTab & _
               Sql
    
    Close fn
    
End Sub
