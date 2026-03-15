VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LabelEventTrap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mCtrl As Label
Attribute mCtrl.VB_VarHelpID = -1

Private mItem As String

Private Sub mCtrl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim s As String
    If Button <> vbRightButton Then Exit Sub
    If Not HFApp.UserPermission("AdminUser") Then Exit Sub
    s = left(Trim(InputBox(vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Change description from """ & Trim(mCtrl.Caption) & """ to:", SpaceCase(mItem))), 30)
    If s <> "" Then
        
        mCtrl.Caption = s
        On Error Resume Next
        
        Call HFApp.SqlExec("insert into CustomDescriptions(Custom_Description,Item) values(" & DbQuote(Str, s) & "," & DbQuote(Str, mItem) & ")")
        Call HFApp.SqlExec("UPDATE CustomDescriptions SET Custom_Description=" & DbQuote(Str, s) & " WHERE Item=" & DbQuote(Str, mItem))
    End If
End Sub

Public Property Set Ctrl(RHS As Label)
    Dim rs As Recordset
    Set mCtrl = RHS
    mItem = mCtrl.Caption
        Set rs = HFApp.SqlExec("select custom_description from customdescriptions where item=" & DbQuote(Str, mItem))
    If Not rs.EOF Then mCtrl.Caption = "" & rs(0)
End Property
