VERSION 5.00
Begin VB.Form FStartupHelper 
   BackColor       =   &H0000FFFF&
   Caption         =   "Form1"
   ClientHeight    =   795
   ClientLeft      =   510
   ClientTop       =   1500
   ClientWidth     =   2460
   LinkTopic       =   "Form1"
   ScaleHeight     =   795
   ScaleWidth      =   2460
   Begin VB.Timer Timer1 
      Left            =   345
      Top             =   150
   End
End
Attribute VB_Name = "FStartupHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Timer1.Interval = 100
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Unload Me
    If Not UsingToolsClass Then

        'normal app launch
        InitCommonControls
        If HFApp.Login(App.Title, App.ProductName, App.Major & "." & App.Minor, App.Path) Then
'        If HFApp.Login("UNSECUREDAPP", App.ProductName, App.Major & "." & App.Minor , App.Path) Then
            Call ReadUserPermissions
            FMain.Show
            Call RollPrices(True)
        End If

    End If
End Sub
