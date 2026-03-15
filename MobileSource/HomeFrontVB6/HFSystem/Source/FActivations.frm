VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FActivations 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Activation Codes"
   ClientHeight    =   4680
   ClientLeft      =   3195
   ClientTop       =   1785
   ClientWidth     =   6195
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FActivations.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtZybPswd 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000011&
      Height          =   195
      IMEMode         =   3  'DISABLE
      Left            =   750
      PasswordChar    =   "X"
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4260
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.TextBox txtCompanyName 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2340
      TabIndex        =   4
      Top             =   1620
      Width           =   3375
   End
   Begin VB.TextBox txtClientID 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2340
      TabIndex        =   3
      Top             =   1320
      Width           =   3375
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   3540
      Picture         =   "FActivations.frx":000C
      TabIndex        =   2
      ToolTipText     =   "Login"
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   4845
      Picture         =   "FActivations.frx":0596
      TabIndex        =   1
      ToolTipText     =   "Cancel"
      Top             =   4200
      Width           =   1215
   End
   Begin VSFlex8Ctl.VSFlexGrid gData 
      Height          =   2025
      Left            =   180
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2070
      Width           =   5835
      _cx             =   10292
      _cy             =   3572
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   4194304
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483633
      TreeColor       =   -2147483633
      FloodColor      =   -2147483633
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   240
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FActivations.frx":0B20
      ScrollTrack     =   0   'False
      ScrollBars      =   0
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   1
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   -2147483633
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      ForeColor       =   &H80000011&
      Height          =   195
      Left            =   0
      TabIndex        =   10
      Top             =   4260
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblLicenseFile 
      AutoSize        =   -1  'True
      Caption         =   "Z:\Homefront\System\hfest.zlc"
      ForeColor       =   &H80000011&
      Height          =   195
      Left            =   0
      TabIndex        =   8
      Top             =   4470
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.Label lblMessage 
      Caption         =   $"FActivations.frx":0C1A
      Height          =   915
      Left            =   1320
      TabIndex        =   7
      Top             =   180
      Width           =   4815
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Client ID"
      Height          =   255
      Index           =   6
      Left            =   780
      TabIndex        =   6
      Top             =   1350
      Width           =   1515
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Company Name"
      Height          =   255
      Index           =   5
      Left            =   780
      TabIndex        =   5
      Top             =   1650
      Width           =   1515
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   600
      Picture         =   "FActivations.frx":0CF4
      Top             =   420
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   300
      Picture         =   "FActivations.frx":15BE
      Top             =   240
      Width           =   660
   End
End
Attribute VB_Name = "FActivations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim hff As Object
    Call IniGetForm(Me)
    
    
    Call LoadModuleInfo(1, "hfsales")
    Call LoadModuleInfo(2, "hfest")
    Call LoadModuleInfo(3, "hfsched")
'    Call LoadModuleInfo(4, "PayablesDesk")

    Call WindowOnTop(Me, True)
    
End Sub

Private Sub cmdNav_Click(Index As Integer)
    Dim b As Boolean
    If Index = 0 Then
        b = True
        If b Then b = SaveModuleInfo(1)
        If b Then b = SaveModuleInfo(2)
        If b Then b = SaveModuleInfo(3)
        'If b Then b = SaveModuleInfo(4)
        If b Then Unload Me
    Else
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
End Sub

Private Sub gData_BeforeEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
    Cancel = col <> 3
End Sub

Private Sub gData_ValidateEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
    gData.EditText = FormatLicense(gData.EditText)
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    If Shift = vbCtrlMask And Button = vbRightButton Then
        txtZybPswd.Visible = True
        txtZybPswd.SetFocus
    
        lblPassword.Visible = True
        lblLicenseFile.Visible = True
        lblLicenseFile.Caption = PathAppend(HFApp.SystemFolder, "System\")
    End If
End Sub

Private Sub LoadModuleInfo(Row As Long, mProductName As String)
    Dim rc As Integer
    Dim FileName As String
    
    FileName = PathAppend(HFApp.SystemFolder, "System", mProductName & ".zlc")
    rc = LLSetLibraryLicense("R110050010000267")
    Call LLSetLicenseFile(FileName)
    Call LLSetMachineIdDrive(FileDrive(FileName))
    rc = LLOpenLicenseFile(False)
    Select Case rc
        Case LLER_OK
            txtClientID.Text = GetLicenseUser
            txtCompanyName.Text = GetLicenseCompany
            gData.TextMatrix(Row, 2) = LLGetLicenseUsers
            gData.TextMatrix(Row, 3) = GetLicense
        Case LLER_LICENSEFILEMISSING
            gData.TextMatrix(Row, 2) = 0
            gData.TextMatrix(Row, 3) = ""
        Case Else
            gData.TextMatrix(Row, 2) = 0
            gData.TextMatrix(Row, 3) = GetLicenseErrorMsg(rc)
    End Select
    Call LLCloseLicenseFile
    
End Sub

Private Function SaveModuleInfo(Row As Long) As Boolean
    Dim rc As Integer
    Dim CreateFile As Boolean
    Dim ClientID   As String
    Dim CompanyName    As String
    Dim mProductName  As String
    Dim FileName As String
    Dim ProductID  As Long
    Dim Modulus    As Long
    Dim ActivationCode As String
    
    ClientID = txtClientID.Text
    CompanyName = txtCompanyName.Text
    mProductName = gData.TextMatrix(Row, 0)
    FileName = PathAppend(HFApp.SystemFolder, "System", mProductName & ".zlc")
    Select Case mProductName
        Case "hfsales":          ProductID = 2:    Modulus = 8
        Case "serviceMgr":       ProductID = 3:    Modulus = 8
        Case "PayablesDesk":     ProductID = 4:    Modulus = 8
        Case "hfest":            ProductID = 5:    Modulus = 8
    End Select
    ActivationCode = Replace(gData.TextMatrix(Row, 3), "-", "")
    CreateFile = txtZybPswd.Visible And txtZybPswd.Text = "Zyb3r*2"
    
    
    rc = LLValidateLicenseNo("R", Modulus, ProductID, ActivationCode, ClientID, CompanyName)
    If rc = 0 Then rc = WriteLicense(FileName, ProductID, Modulus, ActivationCode, ClientID, CompanyName, CreateFile)
    Select Case rc
        Case 0 ' file not found
        Case -1
            SaveModuleInfo = True 'success
        Case 43, 44
            VBA.MsgBox "Invalid client id or company name", vbExclamation, "Activation Error"
        Case 38, 7, 8
            VBA.MsgBox "Invalid " & gData.TextMatrix(Row, 1) & " activation code", vbExclamation, "Activation Error"
        Case Else
            VBA.MsgBox GetLicenseErrorMsg(rc) & " on " & gData.TextMatrix(Row, 1), vbExclamation, "Activation Error"
    End Select
    
End Function



