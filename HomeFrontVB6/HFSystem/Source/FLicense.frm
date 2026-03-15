VERSION 5.00
Begin VB.Form FLicense 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Software Activation"
   ClientHeight    =   3345
   ClientLeft      =   24150
   ClientTop       =   4230
   ClientWidth     =   5925
   Icon            =   "FLicense.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   5925
   Begin VB.TextBox txtActivationCode 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2340
      TabIndex        =   2
      Top             =   2400
      Width           =   3375
   End
   Begin VB.TextBox txtClientID 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2340
      TabIndex        =   0
      Top             =   1800
      Width           =   3375
   End
   Begin VB.TextBox txtCompanyName 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2340
      TabIndex        =   1
      Top             =   2100
      Width           =   3375
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   3180
      TabIndex        =   3
      Top             =   2820
      Width           =   1215
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   4500
      TabIndex        =   4
      Top             =   2820
      Width           =   1215
   End
   Begin VB.TextBox txtZybPswd 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000011&
      Height          =   195
      IMEMode         =   3  'DISABLE
      Left            =   780
      PasswordChar    =   "X"
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2940
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      ForeColor       =   &H80000011&
      Height          =   195
      Left            =   30
      TabIndex        =   11
      Top             =   2940
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblLicenseFile 
      AutoSize        =   -1  'True
      Caption         =   "Z:\Homefront\System\hfest.zlc"
      ForeColor       =   &H80000011&
      Height          =   195
      Left            =   30
      TabIndex        =   10
      Top             =   3150
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   600
      Picture         =   "FLicense.frx":000C
      Top             =   420
      Width           =   480
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Company Name"
      Height          =   255
      Index           =   5
      Left            =   780
      TabIndex        =   9
      Top             =   2130
      Width           =   1515
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Client ID"
      Height          =   255
      Index           =   6
      Left            =   780
      TabIndex        =   8
      Top             =   1830
      Width           =   1515
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Activation Code"
      Height          =   255
      Index           =   8
      Left            =   780
      TabIndex        =   7
      Top             =   2430
      Width           =   1515
   End
   Begin VB.Label lblMessage 
      AutoSize        =   -1  'True
      Caption         =   "This software requires activation before it can be used..."
      Height          =   195
      Left            =   1320
      TabIndex        =   6
      Top             =   180
      Width           =   3960
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   300
      Picture         =   "FLicense.frx":08D6
      Top             =   240
      Width           =   660
   End
End
Attribute VB_Name = "FLicense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mCancel   As Boolean
Private mFilename As String
Private mProduct  As Long
Private mModulus  As Long

Private Sub cmdNav_Click(Index As Integer)
    Dim rc As Long
    Dim s  As String
    
    Dim CreateFile As Boolean
    
    Select Case Index
        Case 0 'ok
            mCancel = False
            rc = LLValidateLicenseNo("R", mModulus, mProduct, Replace(txtActivationCode.Text, "-", ""), txtClientID.Text, txtCompanyName.Text)
            CreateFile = txtZybPswd.Visible And txtZybPswd.Text = "Zyb3r*2"
            Select Case rc
                Case 0:     If WriteLicense(mFilename, mProduct, mModulus, Replace(txtActivationCode.Text, "-", ""), txtClientID.Text, txtCompanyName.Text, CreateFile) Then Unload Me
                Case 43:    VBA.MsgBox "Invalid Client ID", vbExclamation, "Activation Error"
                Case 38:    VBA.MsgBox "Invalid Activation Code", vbExclamation, "Activation Error"
                Case Else:  VBA.MsgBox GetLicenseErrorMsg(rc), vbExclamation, "Activation Error"
            End Select
            
        Case 1 'cancel
            mCancel = True
            Unload Me
            
    End Select
End Sub

Public Function EnterLicense(FileName As String, Product As Long, Modulus As Long, Message As String) As Boolean
On Error Resume Next
    mFilename = FileName
    mProduct = Product
    mModulus = Modulus
    Call CenterForm(Me)
    lblMessage = Message
'    txtActivationCode.Text = GetLicense
    txtCompanyName.Text = GetLicenseCompany
    txtClientID.Text = GetLicenseUser
    Me.Show vbModal
    EnterLicense = Not mCancel
End Function

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    If Shift = vbCtrlMask And Button = vbRightButton Then
        txtZybPswd.Visible = True
        txtZybPswd.SetFocus
    
        lblPassword.Visible = True
        lblLicenseFile.Visible = True
        lblLicenseFile.Caption = mFilename
    End If
End Sub

Private Sub txtActivationCode_GotFocus()
    SelectAll txtActivationCode
End Sub

Private Sub txtActivationCode_Validate(Cancel As Boolean)
    txtActivationCode.Text = FormatLicense(txtActivationCode.Text)
End Sub

Private Sub txtCompanyName_GotFocus()
    SelectAll txtCompanyName
End Sub
Private Sub txtClientID_GotFocus()
    SelectAll txtClientID
End Sub

