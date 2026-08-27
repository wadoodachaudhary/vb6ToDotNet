Attribute VB_Name = "mMain"
Option Explicit
Public Const LINKCOLOR = vbHighlight
Public Const LINKUNDERLINE = True
Public Const CORPORATEPAGE = "http://www.homefrontsoftware.com"
Public Const SUPPORTPAGE = "http://www.homefrontsoftware.com"
Public Const APPKEY = "APDESK"

Private Declare Sub InitCommonControls Lib "COMCTL32.DLL" ()
Private Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long
Private Const SEM_NOGPFAULTERRORBOX = &H2

'do not use the actual values as they screw up on some resolutions
Public Const ScreenTwipsPerPixelX = 15
Public Const ScreenTwipsPerPixelY = 15


Global MouseGrid As VSFlexGrid
Global MouseCol  As Long

Global SQLDateFormat As String

Public App     As New App
Public HFApp As New HFSystem.Application



Public Sub Main()
    InitCommonControls
    If HFApp.Login(App.Title, App.ProductName, App.Major & "." & App.Minor, App.path) Then
        If App.Login() Then
            'Call HFApp.EditOptions
            'Call HFApp.RunTask("EditDocuments")
            
            
            FMain.Show
        End If
    End If
    
    SQLDateFormat = "" & HFApp.Options.ValueByName("SQLDATEFORMAT")
    If SQLDateFormat = "" Then
        SQLDateFormat = "'YYYY-MM-DD'"
    End If
End Sub
Public Sub UnloadApp()
    If Not InIde() Then SetErrorMode SEM_NOGPFAULTERRORBOX
End Sub

Public Sub UpdatePOCompletions()
On Error GoTo eh
    Dim s As String
    Dim rs As Recordset
    If HFApp.Options(AccountingSystem) = asTimberline Then
        s = "select PONumber,CompletedDate,Updateinacct from POMaster where CompletedDate is not null and UpdateinAcct=1"
        Set rs = HFApp.SqlExec(s, dbHomeFront)
        While Not rs.EOF
            Call HFApp.SqlExec("UPDATE master_jcm_record_12 SET sactcd=" & DbQuote(Date, "" & rs("CompletedDate")) & " WHERE sub=" & DbQuote(Str, rs("PONumber")), dbAccountingDictionary)
            Call HFApp.SqlExec("UPDATE POmaster SET Updateinacct = 0 WHERE DivisionID = " & HFApp.DivisionID & " and PONumber = " & DbQuote(Str, rs("PONumber")), dbHomeFront)
            rs.MoveNext
        Wend
    End If
eh: Exit Sub
End Sub
Public Function InvoiceCodeExists(Vendor As String, Invoice As String, Index As Integer, Value As String) As Boolean
On Error Resume Next
    Dim s As String
    Dim rs As Recordset
    
    InvoiceCodeExists = False
    If Value = "" Then Exit Function
    
    s = ""
    s = s & "SELECT *" & vbCrLf
    s = s & "  FROM Invoices" & vbCrLf
    s = s & " WHERE DivisionID = " & HFApp.DivisionID & " and InvoiceCode" & Index & "=" & DbQuote(Str, Value)
    s = s & "   AND (Vendor<>" & DbQuote(Str, Vendor) & vbCrLf
    s = s & "        OR Invoice<>" & DbQuote(Str, Invoice) & ")" & vbCrLf
    Set rs = HFApp.SqlExec(s, dbHomeFront)
    If Not rs.EOF Then
        InvoiceCodeExists = True
        Exit Function
    End If
    
    If TimberlineAccounting Then
        s = ""
        s = s & "SELECT *"
        s = s & "  FROM new_api_record_1"
        s = s & " WHERE OIVND<>" & DbQuote(Str, Vendor) & vbCrLf
        s = s & "   AND OIINV<>" & DbQuote(Str, Invoice) & vbCrLf
        s = s & "   AND OICODE" & Index & "=" & DbQuote(Str, Value)
        Set rs = HFApp.SqlExec(s, dbAccountingDictionary)
        If Not rs.EOF Then
            InvoiceCodeExists = True
            Exit Function
        End If
        
        s = ""
        s = s & "SELECT *"
        s = s & "  FROM master_apm_record_1"
        s = s & " WHERE OIVND<>" & DbQuote(Str, Vendor) & vbCrLf
        s = s & "   AND OIINV<>" & DbQuote(Str, Invoice) & vbCrLf
        s = s & "   AND OICODE" & Index & "=" & DbQuote(Str, Value)
        Set rs = HFApp.SqlExec(s, dbAccountingDictionary)
        If Not rs.EOF Then
            InvoiceCodeExists = True
            Exit Function
        End If
    End If
End Function




Public Function SyncCmd(cmd As String, Optional parameters As String, Optional SkipAccounting As Boolean = False) As String


    If InIde() Then
        SyncCmd = "C:\Program Files (x86)\HomeFront\HFSync.exe " & _
                  HFApp.ConnectionString(dbHomeFront) & Chr(1) & _
                  HFApp.DivisionID & Chr(1) & _
                  IIf(SkipAccounting, 1, 0) & Chr(1) & _
                  cmd & Chr(1) & _
                  parameters
    Else
        SyncCmd = App.path & "\HFSync.exe " & _
                  HFApp.ConnectionString(dbHomeFront) & Chr(1) & _
                  HFApp.DivisionID & Chr(1) & _
                  IIf(SkipAccounting, 1, 0) & Chr(1) & _
                  cmd & Chr(1) & _
                  parameters
    End If

End Function


Public Function MaxInvoiceLength() As Integer
    Select Case HFApp.Options(AccountingSystem)
        Case asTimberline, asMasterBuilder
            MaxInvoiceLength = 15
        Case Else
            MaxInvoiceLength = 20
    End Select
End Function


