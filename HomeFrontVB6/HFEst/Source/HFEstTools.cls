VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HFEstTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Sub Class_Initialize()
    UsingToolsClass = True
End Sub

Public Function GetProcessID() As Long
    GetProcessID = GetCurrentProcessId()
End Function

Public Function CreateSinglePO(DivisionID As Long, ConnectionString As String, HFUser As String, Job As String, Vendor As String, PONumber As String) As String
On Error GoTo eh
    
    InitCommonControls
    
    Call HFApp.SetDivision("" & DivisionID)
    Call HFApp.RunTask("TrustedLogin|" & ConnectionString & "|0|" & HFUser)
    If Not HFApp.Login(App.Title, App.ProductName, App.Major & "." & App.Minor, App.Path) Then Exit Function
    Call ReadUserPermissions
    
    Load FMain
    FMain.Visible = False
    Screen.MousePointer = vbDefault
    CreateSinglePO = FPurchaseOrder.CreateSinglePO(Job, Vendor, PONumber)
    Unload FMain
    
Exit Function
eh: MsgBox "CreateSinglePO" & vbCrLf & Err.Description
End Function





