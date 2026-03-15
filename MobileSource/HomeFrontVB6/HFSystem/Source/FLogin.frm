VERSION 5.00
Begin VB.Form FLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   4095
   ClientLeft      =   3840
   ClientTop       =   4905
   ClientWidth     =   7230
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FLogin.frx":000C
   ScaleHeight     =   4095
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lDivisions 
      Appearance      =   0  'Flat
      Height          =   1545
      IntegralHeight  =   0   'False
      Left            =   3825
      TabIndex        =   22
      Top             =   1770
      Width           =   3315
   End
   Begin VB.TextBox txtCheat 
      Height          =   405
      Left            =   5910
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   1245
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   5865
      Picture         =   "FLogin.frx":60876
      TabIndex        =   7
      ToolTipText     =   "Cancel"
      Top             =   3570
      Width           =   1215
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   4560
      Picture         =   "FLogin.frx":60E00
      TabIndex        =   6
      ToolTipText     =   "Login"
      Top             =   3570
      Width           =   1215
   End
   Begin VB.ListBox lConnections 
      Appearance      =   0  'Flat
      Height          =   1545
      IntegralHeight  =   0   'False
      Left            =   3825
      TabIndex        =   0
      Top             =   1770
      Width           =   3315
   End
   Begin VB.Frame FrameUID 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   1545
      Left            =   3825
      TabIndex        =   9
      Top             =   1770
      Visible         =   0   'False
      Width           =   3315
      Begin VB.TextBox txtUID 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1545
         TabIndex        =   1
         Top             =   90
         Width           =   1695
      End
      Begin VB.TextBox txtPWD 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1545
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   390
         Width           =   1695
      End
      Begin VB.TextBox txtNewPWD 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1545
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   690
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtConfirmPWD 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1545
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   990
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblConfirmPWD 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Confirm Password"
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   1050
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Label lblNewPWD 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "New Password"
         Height          =   195
         Left            =   375
         TabIndex        =   13
         Top             =   750
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Password"
         Height          =   195
         Index           =   1
         Left            =   750
         TabIndex        =   12
         Top             =   450
         Width           =   690
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "User ID"
         Height          =   195
         Index           =   0
         Left            =   900
         TabIndex        =   11
         Top             =   150
         Width           =   540
      End
      Begin VB.Label lblChangePassword 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Change Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   1935
         TabIndex        =   10
         Top             =   1290
         Width           =   1290
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   26595
      Y1              =   3405
      Y2              =   3405
   End
   Begin VB.Label lblLicensedTo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00504B4A&
      Height          =   195
      Left            =   210
      TabIndex        =   21
      Top             =   2295
      UseMnemonic     =   0   'False
      Width           =   1320
   End
   Begin VB.Label lblLicenseInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "L1C3NC3N0H343"
      ForeColor       =   &H00400000&
      Height          =   690
      Left            =   210
      TabIndex        =   20
      Top             =   2610
      UseMnemonic     =   0   'False
      Width           =   3315
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "This software is licensed to:"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   195
      TabIndex        =   19
      Top             =   2100
      UseMnemonic     =   0   'False
      Width           =   1950
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " "
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   90
      TabIndex        =   18
      Top             =   3660
      UseMnemonic     =   0   'False
      Width           =   45
   End
   Begin VB.Image Image1 
      Height          =   4095
      Left            =   10140
      Picture         =   "FLogin.frx":6138A
      Top             =   1170
      Visible         =   0   'False
      Width           =   7230
   End
   Begin VB.Label lblAppName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Homefront <AppName>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00504B4A&
      Height          =   240
      Left            =   4080
      TabIndex        =   17
      Top             =   420
      UseMnemonic     =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblAppVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 4.01.1234"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   4080
      TabIndex        =   16
      Top             =   690
      UseMnemonic     =   0   'False
      Width           =   1290
   End
   Begin VB.Label lblAbout 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "About Homefront"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   4080
      TabIndex        =   15
      Top             =   1140
      UseMnemonic     =   0   'False
      Width           =   1200
   End
   Begin VB.Label lblAppCopyright 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2007 HomeFront Software"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   4080
      TabIndex        =   8
      Top             =   900
      UseMnemonic     =   0   'False
      Width           =   2745
   End
End
Attribute VB_Name = "FLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private mLicensed As Boolean
Private mCancel As Boolean
Private mConnections() As String
Private mStep As Integer
Private mChangePswd As Boolean


Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function Login() As Boolean
    mCancel = False
    Me.Show vbModal
    Screen.MousePointer = vbDefault
    Login = Not mCancel
End Function

Private Sub Form_Activate()
On Error Resume Next
    Me.SetFocus
    lConnections.SetFocus
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    
    'did they hit ctrl+shift+Z
    If KeyCode = vbKeyZ And Shift = vbShiftMask + vbCtrlMask Then
        txtCheat.Visible = True
        txtCheat.Text = ""
        txtCheat.SetFocus
    End If
    
End Sub



Private Sub Form_Load()
    
    txtCheat.left = -9999
    lConnections.Move FrameUID.left, FrameUID.Top
    lDivisions.Move FrameUID.left, FrameUID.Top
    Call IniGetForm(Me)
        
    lblAppName.Caption = mModuleName
    lblAppVersion.Caption = "Version " & mSystemVersion
    lblAppCopyright.Caption = Year(VBA.Date()) & " © Copyright Hyphen Solutions"
    
    Call LoadLicense
    
    If JobSimplicity Then
        lblAbout = "About Job Simplicity"
        Me.Picture = Image1.Picture
    End If
    Set StatusCtrl = lblStatus
    
    Call WindowOnTop(Me, True)
    
    On Error Resume Next
    txtUID.Text = IniGet(AppIni(), "FLogin", "LastLogin", "")
    On Error GoTo 0
    
    step = 0
    
End Sub

Private Sub LoadDivisions()
On Error GoTo eh

    Dim s As String
    Dim rs As Recordset
    
    lDivisions.Clear
    
    s = "Select DivisionID,DivisionCode,DivisionName from Divisions"
    Set rs = HFApp.SqlExec(s)
    With lDivisions
        While Not rs.EOF
            .AddItem "" & rs("divisioncode") & vbTab & rs("divisionname")
            .ItemData(.NewIndex) = "" & rs("divisionid")
            rs.MoveNext
        Wend
        .ListIndex = 0
    End With
    
    On Error Resume Next
    lDivisions.ListIndex = Val(IniGet(AppIni(), "FLogin", "LastDivision", 0))
    
Exit Sub
eh:
MsgBox Err.Description, vbOKOnly + vbInformation, App.ProductName
End Sub

Private Sub EncryptPwds()
On Error Resume Next

    Dim file As String
    Dim sections As String
    Dim section As String
    Dim pwd As String
    Dim i As Long
    
    'encrypt passwords in local ini
    file = PathAppend(mExePath, "DBConnections.ini")
    sections = IniGetSectionNames(file)
    For i = 1 To Parse(sections)
        section = Parse(sections, i)
        If section <> "" Then
            pwd = IniGet(file, section, "pwd")
            If pwd <> "" Then
                Call IniPut(file, section, "pwe", HFApp.Encrypt(pwd))
                Call IniRemove(file, section, "pwd")
            End If
        End If
    Next

    'encrypt passwords in system ini
    file = PathAppend(HFApp.SystemFolder, "System", "DBConnections.ini")
    sections = IniGetSectionNames(file)
    For i = 1 To Parse(sections)
        section = Parse(sections, i)
        If section <> "" Then
            pwd = IniGet(file, section, "pwd")
            If pwd <> "" Then
                Call IniPut(file, section, "pwe", HFApp.Encrypt(pwd))
                Call IniRemove(file, section, "pwd")
            End If
        End If
    Next

End Sub

Private Sub LoadConnections()
    Dim i As Long
    Dim sIniFile As String
    Dim CompanyName   As String
    Dim sections As String
    Dim Dsn As String
    Dim uid As String
    Dim pwe As String
    Dim pwd As String
    Dim hff As Object
    Set hff = CreateObject("ZYBFunctions.HomeFrontFunctions")
    
    
    'encrypt any unencrypted pwds
    Call EncryptPwds
    
    'first look in local folder
    sIniFile = PathAppend(mExePath, "DBConnections.ini")
    If Not FileExists(sIniFile) Then
        'then look in system folder
        sIniFile = PathAppend(HFApp.SystemFolder, "System", "DBConnections.ini")
        If Not FileExists(sIniFile) Then
            'not found. load dsns from registry
            sIniFile = ""
        End If
    End If
    
    
    
    If sIniFile <> "" Then
        'inifile exists so use it
        ReDim mConnections(0)
        With lConnections
            .Clear
            sections = IniGetSectionNames(sIniFile)
            For i = 1 To Parse(sections)
            
                CompanyName = Parse(sections, i)
                Dsn = IniGet(sIniFile, CompanyName, "dsn")
                uid = IniGet(sIniFile, CompanyName, "uid")
                pwe = IniGet(sIniFile, CompanyName, "pwe")
                pwd = IniGet(sIniFile, CompanyName, "pwd")
                If pwe <> "" Then pwd = hff.MyDeCrypt(pwe, "Fazlul")

                .AddItem CompanyName
                ReDim Preserve mConnections(UBound(mConnections) + 1)
                mConnections(UBound(mConnections)) = "dsn=" & Dsn & ";uid=" & uid & ";pwd=" & pwd & ";Persist Security Info=True"
            
            Next
            .Visible = True
        End With
    Else
        'no inifile found. use available dsn's
        Dim s()    As String
        With lConnections
            .Clear
            
            If RegEnumKeys(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", s) Then
            
                Call Sort(s)
                
                For i = 1 To UBound(s)
                    If RegGetKey(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", s(i)) = "SQL Server" Then
                        .AddItem s(i)
                    End If
                Next
                
                ReDim mConnections(.ListCount)
                For i = 0 To .ListCount - 1
                    mConnections(i + 1) = "dsn=" & .list(i) & ";Persist Security Info=True"
                Next
                
            End If
            .Visible = True
        End With
    End If

    On Error Resume Next
    lConnections.ListIndex = Val(IniGet(AppIni(), "FLogin", "LastDB", 0))

End Sub
Private Sub LoadLicense()
Dim i As Long
Dim settingsFile As String
Dim s As String
Dim expiryDate As Date
Dim notice As String
Dim tokenPeriod As Long
Dim suspended As Boolean

    With HFApp.License
        mLicensed = True
    
        Select Case True
        
            Case Not FileExists(HFApp.LicenseFile)
                lblLicensedTo.Caption = ""
                lblLicenseInfo.Caption = "No license could be found. Contact your system administrator for assistance."
                lblLicenseInfo.ForeColor = vbRed
                mLicensed = False
                cmdNav(0).Enabled = False
                lConnections.Enabled = False
        
            Case Not .ReadLicense(HFApp.LicenseFile)
                lblLicensedTo.Caption = ""
                lblLicenseInfo.Caption = "Invalid license. Contact your HomeFront solution provider for assistance."
                lblLicenseInfo.ForeColor = vbRed
                mLicensed = False
                cmdNav(0).Enabled = False
                lConnections.Enabled = False
    
            Case Else
                lblLicensedTo.Caption = .CompanyName
'                lblLicenseInfo.Caption = .SerialNumber
'                If .ModuleDemo(.ModuleIndex(mProductCode)) Then
'                    i = .ModuleDemoExpiry(.ModuleIndex(mProductCode)) - VBA.Date()
'                    If i < 1 Then
'                        lblLicenseInfo.Caption = "The evaluation period has expired. Contact your HomeFront solution provider for assistance."
'                        cmdNav(0).Enabled = False
'                        lConnections.Enabled = False
'                    Else
'                        lblLicenseInfo.Caption = i & " days remaining in evaluation period"
'                    End If
'                    lblLicenseInfo.ForeColor = vbRed
'                Else
'                    If .Leased And .LeaseExpiryDate < VBA.Date Then
'                        lblLicenseInfo.Caption = "Your software lease has expired. Contact your HomeFront solution provider for assistance."
'                        lblLicenseInfo.ForeColor = vbRed
'                        cmdNav(0).Enabled = False
'                        lConnections.Enabled = False
'                    End If
'                End If
        
                Dim xx As New HyphenSys.DesktopAuth
                
                settingsFile = PathAppend(HFApp.SystemFolder, "System\Reports\Estimating\rpt.dat")
                notice = IniGet(settingsFile, "LastRun", "notice")
                If notice = "" Then
                    lblLicenseInfo.Caption = .SerialNumber
                    lblLicenseInfo.ForeColor = &H400000
                Else
                    lblLicenseInfo.ForeColor = &H80&
                    lblLicenseInfo.Caption = notice
                End If
                s = IniGet(settingsFile, "LastRun", "date")
                If Not IsDate(s) Then s = ""
                If s = "" Then s = "2019-01-01"
                expiryDate = Int(DateValue(s))
                                
                If expiryDate < Now() Then
                    'try to get new date
                    'if success move expiry date to tomorrow
                    'if expirydate is more than a week past due then deny access
                    
                    If xx.AuthorizeClient(.ClientID, .CompanyName, App.Major & "." & App.Minor & "." & App.Revision, suspended, notice) Then
                        If suspended Then
                            lblLicenseInfo.Caption = "Your software lease has expired. Contact accounting@hyphensolutions.com for assistance."
                            lblLicenseInfo.ForeColor = vbRed
                            mLicensed = False
                            cmdNav(0).Enabled = False
                            lConnections.Enabled = False
                        Else
                            expiryDate = Int(Now()) + 1
                            Call IniPut(settingsFile, "LastRun", "notice", notice)
                            Call IniPut(settingsFile, "LastRun", "date", expiryDate)
                            If notice = "" Then
                                lblLicenseInfo.Caption = .SerialNumber
                                lblLicenseInfo.ForeColor = &H400000
                            Else
                                lblLicenseInfo.Caption = notice
                                lblLicenseInfo.ForeColor = &H80&
                            End If
                        End If
                    Else
                        lblLicenseInfo.Caption = "Unable to reach the license server"
                        lblLicenseInfo.ForeColor = &H80&
                    End If
                    
                End If
                
                If expiryDate < Now() - 7 Then
                    lblLicenseInfo.Caption = "Unable to reach the license server. Contact support@hyphensolutions.com for assistance."
                    lblLicenseInfo.ForeColor = vbRed
                    mLicensed = False
                    cmdNav(0).Enabled = False
                    lConnections.Enabled = False
                End If
                
                
        End Select

    End With

End Sub


Private Function Clean(s As String) As String
    'parse variablevalue out of "variablename=variablevalue;"
    s = Mid(s, InStr(1, s, "=") + 1)
    Clean = left(s, Len(s) - 1)
End Function







Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton And Shift = vbCtrlMask Then
        Call ShellFile(Me.hwnd, mExePath & "\" & "HomeFront.ini")
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
End Sub



Private Sub lblAbout_Click()
On Error Resume Next

    Dim sUser As String
    Dim sPswd As String
    Dim rs As Recordset


    HFApp.Databases(dbHomefront).Close
    HFApp.Databases(dbHomefront).Open mConnections(lConnections.ListIndex + 1)
    HFApp.ConnectionString(dbHomefront) = mConnections(lConnections.ListIndex + 1)
    
    
    Call HFApp.SetLoginID("", "")
    If HFApp.Databases(dbHomefront).State = adStateOpen Then
        sUser = txtUID.Text
        sPswd = HFApp.Encrypt(txtPWD.Text)
        Set rs = HFApp.SqlExec("SELECT user_access FROM User_Manager WHERE User_Id=" & DbQuote(Str, sUser) & " AND User_Password=" & DbQuote(Str, sPswd))
        If Not rs.EOF Then
            If "" & rs(0) = "1" Then
                Call HFApp.SetLoginID(sUser, "")
            End If
        End If
    End If
    
    HFApp.About
End Sub

Private Sub lblChangePassword_Click()
    txtNewPWD.Visible = Not txtNewPWD.Visible
    lblNewPWD.Visible = txtNewPWD.Visible
    txtConfirmPWD.Visible = txtNewPWD.Visible
    lblConfirmPWD.Visible = txtNewPWD.Visible
End Sub



Private Sub lConnections_Click()
    cmdNav(0).Enabled = mLicensed And lConnections.ListIndex <> -1
End Sub

Private Sub lConnections_DblClick()
    Call cmdNav_Click(0)
End Sub

Private Sub lDivisions_Click()
    cmdNav(0).Enabled = mLicensed And lDivisions.ListIndex <> -1
End Sub

Private Sub lDivisions_DblClick()
    Call cmdNav_Click(0)
End Sub

Private Sub txtUID_GotFocus()
    SelectAll txtUID
End Sub
Private Sub txtPWD_GotFocus()
    SelectAll txtPWD
End Sub
Private Sub txtNewPWD_GotFocus()
    SelectAll txtNewPWD
End Sub
Private Sub txtConfirmPWD_GotFocus()
    SelectAll txtConfirmPWD
End Sub


Private Property Let step(RHS As Integer)


    Select Case RHS
        Case -1 ' cancel
            mCancel = True
            Unload Me
            
        Case 0 ' select connection
            If mStep = 0 Then Call LoadConnections
            mStep = RHS
            lConnections.Visible = mStep < 1
            FrameUID.Visible = mStep = 1
            lDivisions.Visible = mStep > 1
            cmdNav(0).Enabled = mLicensed And lConnections.ListIndex <> -1
            cmdNav(1).Enabled = True
            Call SetCtrlFocus(lConnections)
            
        Case 1 ' enter user/pswd
            If DoOpenDB Then
                mStep = RHS
                lConnections.Visible = mStep < 1
                FrameUID.Visible = mStep = 1
                lDivisions.Visible = mStep > 1
                Call SetCtrlFocus(txtUID)
            End If

        Case 2 ' select division
            mChangePswd = txtNewPWD.Visible
            Call LoadDivisions
            Select Case lDivisions.ListCount
                Case 0 'no divisions??
                    Exit Property
                    
                Case 1 'only one division defined
                    step = 3
                    Exit Property
                    
                Case Else
                    mStep = RHS
                    
            End Select
            
            lConnections.Visible = mStep < 1
            FrameUID.Visible = mStep = 1
            lDivisions.Visible = mStep > 1
            cmdNav(0).Enabled = mLicensed And lDivisions.ListIndex <> -1
            cmdNav(1).Enabled = True
            Call SetCtrlFocus(lDivisions)
                        
        Case 3 ' login
            Call DoLogin
            
    End Select

End Property

Private Property Get step() As Integer
    step = mStep
End Property



Private Sub cmdNav_Click(Index As Integer)
    If Index = 0 Then
        step = step + 1
    Else
        step = step - 1
    End If
End Sub

Private Function DoOpenDB() As Boolean
On Error Resume Next
    Dim s As String
    Dim constr As String
    
    constr = mConnections(lConnections.ListIndex + 1)
    
    If Not CBool(InStr(1, constr, "DSN=", vbTextCompare)) Then
        Screen.MousePointer = vbDefault
        Call MsgBox("Your connection parameters are incorrect." & vbCrLf & "Please notify your system administrator.", vbExclamation, "HomeFront")
        Exit Function
    End If

    Screen.MousePointer = vbHourglass

    StatusLabel = "opening database"
    'open databases
    HFApp.ConnectionString(dbHomefront) = constr
    HFApp.Databases(dbHomefront).Close
    
    Set HFApp.Databases(dbHomefront) = New Connection
    HFApp.Databases(dbHomefront).Open HFApp.ConnectionString(dbHomefront)

    StatusLabel = ""

    If HFApp.Databases(dbHomefront).State <> adStateOpen Then
        Screen.MousePointer = vbDefault
        MsgBox "Unable to open the Homefront database" & IIf(Err.Number <> 0, vbCrLf & Parse(Err.Description, Parse(Err.Description, , "]"), "]"), ""), vbCritical, App.ProductName
        Exit Function
    End If

    If Not HFApp.CreateDivisionUsers Then
        Screen.MousePointer = vbDefault
        Exit Function
    End If


    Screen.MousePointer = vbDefault
    DoOpenDB = True

End Function


Private Sub DoLogin()
On Error GoTo eh

    Dim hff   As Object
    Dim sUser As String
    Dim sPswd As String
    Dim rs    As Recordset
    Dim s As String
    Dim tmr As Single
    Dim noSync As Boolean
    Dim canLaunch As Boolean
    
    Screen.MousePointer = vbHourglass
    
    Call IniPut(AppIni(), "FLogin", "LastDB", lConnections.ListIndex)
    Call IniPut(AppIni(), "FLogin", "LastDivision", lDivisions.ListIndex)
    Call IniPut(AppIni(), "FLogin", "LastLogin", txtUID.Text)

    
    'acquire license
'MsgBox "1 acquire license"
    StatusLabel = "aquiring license"
    If txtCheat.Text <> "CHEAT" Then
        If Not HFApp.AcquireLicense() Then
            Screen.MousePointer = vbDefault
            StatusLabel = ""
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If
    StatusLabel = ""
    
    'validate user
'MsgBox "2 validate user"
    StatusLabel = "validating user"
    Call HFApp.SetDivisionID(lDivisions.ItemData(lDivisions.ListIndex))
    
    sUser = txtUID.Text
    sPswd = txtPWD.Text
    Call HFApp.SetLoginID(sUser, sPswd)
    
    Set hff = CreateObject("ZYBFunctions.HomeFrontFunctions")
    sPswd = hff.MyCrypt(sPswd, "Fazlul")
    
    s = ""
    s = s & "SELECT * " & vbCrLf
    s = s & "FROM User_Manager u" & vbCrLf
    s = s & "join Divisionusers d on u.User_ID = d.UserID and DivisionID =" & HFApp.DivisionID & vbCrLf
    s = s & "WHERE u.User_Id=" & DbQuote(Str, sUser) & vbCrLf
    If txtCheat.Text <> "CHEAT" Then
        s = s & "AND ISNULL(u.User_Password,'')=" & DbQuote(Str, sPswd) & " COLLATE Latin1_General_CS_AS"
    End If
    Set rs = HFApp.SqlExec(s)
    If Not rs.EOF Then Call HFApp.SetUserLevel(Val("" & rs("user_access")))
    StatusLabel = ""
    
    
    If rs.EOF Then
        Screen.MousePointer = vbDefault
        If txtCheat.Text = "CHEAT" Then
            MsgBox "Login """ & sUser & """ does not exist.", vbCritical, "HomeFront - Login Manager"
        Else
            MsgBox "Access denied" & vbCrLf & "Invalid login or password", vbCritical, "HomeFront - Login Manager"
        End If
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        Select Case mModuleCode
            Case "hfsched":         s = "LaunchScheduling"
            Case "hfest":           s = "LaunchEstimating"
            Case "hfsales":         s = "LaunchSales"
            Case "hfpayables":      s = "LaunchPayables"
            Case "hfwarranty":      s = "LaunchWarranty"
            Case "workticket":      s = "LaunchWorkticket"
            Case "workflowbuilder": s = "LaunchWorkflow"
            Case "hfstaging":       s = "LaunchStaging"
            Case "HFSalesMgmt":     s = "LaunchSalesMgmt"
            Case Else: s = ""
        End Select
        On Error Resume Next
        canLaunch = "" & rs(s) = "True"
        If Err.Number = 3265 Then
            'Item cannot be found in the collection corresponding to the requested name or ordinal.
            'Assume upgrade has not been run yet
        Else
            If Not canLaunch Then
                Screen.MousePointer = vbDefault
                MsgBox "Access denied - Insufficient Privilege" & vbCrLf & vbCrLf & "You have not been granted access to this module.", vbCritical, "HomeFront"
                Exit Sub
            End If
        End If
        On Error GoTo eh
        
'MsgBox "3 write stats"
        Call WriteStats
'MsgBox "3.1 done writting stats"
        
    End If
    
    

    'update password if needed
'MsgBox "4 update password"
    If txtNewPWD.Text = txtPWD.Text Then mChangePswd = False
    If mChangePswd Then
        If txtNewPWD.Text <> txtConfirmPWD.Text Then
            Screen.MousePointer = vbDefault
            Call MsgBox("Unable to change password." & vbCrLf & vbCrLf & "Your new passwords don't match.", vbExclamation, "HomeFront")
            Exit Sub
        End If
        If txtNewPWD.Text = "" Then
            If vbNo = MsgBox("Are you sure you want to set your password to nothing?", vbExclamation + vbYesNo, "HomeFront") Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
        
        s = "select dbo.ZYB_IsPswdComplex(" & DbQuote(Num, HFApp.DivisionID) & "," & DbQuote(Str, sUser) & "," & DbQuote(Str, txtNewPWD.Text) & ")"
        s = "" & HFApp.SqlExec(s)(0)
        If s <> "" Then
            Screen.MousePointer = vbDefault
            Call MsgBox("Unable to change password. It does not meet complexity rules." & vbCrLf & vbCrLf & s, vbExclamation, "HomeFront")
            Exit Sub
        End If
        
        sPswd = hff.MyCrypt(txtNewPWD.Text, "Fazlul")
        
        s = ""
        s = s & "UPDATE User_Manager SET User_Password=" & DbQuote(Str, sPswd) & " WHERE User_Id=" & DbQuote(Str, sUser) & vbCrLf
        s = s & "UPDATE tblProjectManager set password=" & DbQuote(Str, txtNewPWD.Text) & " WHERE PM=" & DbQuote(Str, sUser)
        Call HFApp.SqlExec(s)
        
        Call MsgBox("Your password has been changed.", vbInformation, "HomeFront")
        txtPWD.Text = txtNewPWD.Text
        txtNewPWD.Text = ""
        txtConfirmPWD.Text = ""
        mChangePswd = False
    End If
    
    
    'upgrade
'MsgBox "5 upgrade db"
    tmr = Timer
    If IsIn(mModuleCode, "hfsales", "hfest", "hfwarranty", "hfpayables", "Workticket", "hfsched", "hfstaging", "HFSalesMgmt") Then
        StatusLabel = "checking database..."
        If Not FDbUpgrade.Upgrade(HFApp.ConnectionString(dbHomefront)) Then
            StatusLabel = ""
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        StatusLabel = ""
    End If

    
    'write client name and id to license table
'MsgBox "6 update license table"
    s = ""
    s = s & "delete license" & vbCrLf
    s = s & "insert license(clientid,companyname)" & vbCrLf
    s = s & "values(" & DbQuote(Str, HFApp.License.ClientID) & "," & DbQuote(Str, HFApp.License.CompanyName) & ")"
    On Error Resume Next
    Call HFApp.SqlExec(s)
    On Error GoTo eh
    
    
    'load options
'MsgBox "7 load options"
    HFApp.Options.ReadData
 
    
    
'MsgBox "8 start synch"
    mCancel = False
    
    noSync = Not IsIn(mModuleCode, "hfest", "hfpayables")
    noSync = noSync Or HFApp.Options.ValueByName("ManualHFSync") = "true"
    noSync = noSync Or RegGetKey(HKEY_LOCAL_MACHINE, "SOFTWARE\HomeFront", "Startup Sync") = "Off"
    If Not noSync Then
        Call Shell(SyncCmd("Start"), vbHide)
    End If
    
    Screen.MousePointer = vbDefault
    Unload Me


Exit Sub
eh:
    MsgBox Err.Description, vbCritical, "Error at DoLogin()"
    Unload Me
End Sub

Private Sub WriteStats()
On Error GoTo eh
    Dim xx As New HyphenSys.DesktopAuth
    Dim s As String
    Dim xml As String
    Dim rs As Recordset
    Dim LicenseText As String
    Dim IsProduction As String
    
    LicenseText = ReadFileToString(HFApp.LicenseFile)
      
Dim i As Long
    
    s = ""
    s = s & "select --- CLIENT ---" & vbCrLf
    s = s & " 1 Tag" & vbCrLf
    s = s & ",null Parent" & vbCrLf
    s = s & "," & DbQuote(Str, HFApp.License.ClientID) & " [Client!1!ID!element]" & vbCrLf
    s = s & "," & DbQuote(Str, HFApp.License.CompanyName) & " [Client!1!Name!element]" & vbCrLf
    s = s & "," & DbQuote(Str, App.Major & "." & App.Minor & "." & App.Revision) & " [Client!1!AppVersion!element]" & vbCrLf
    s = s & "," & DbQuote(Str, LicenseText) & " [Client!1!License!element]" & vbCrLf
    s = s & ",db_name()+'@'+@@servername [Client!1!Database!element]" & vbCrLf
    s = s & ",'' [Divisions!2!Code!element]" & vbCrLf
    s = s & ",'' [Divisions!2!Name!element]" & vbCrLf
    s = s & ",'' [Divisions!2!BuildProCompanyCode!element]" & vbCrLf
    s = s & ",'' [Divisions!2!BuildProEnvironment!element]" & vbCrLf
    s = s & ",'' [Divisions!2!WalletPartyID!element]" & vbCrLf
    s = s & ",'' [Divisions!2!AccountingSystem!element]" & vbCrLf
    s = s & ",'' [Divisions!2!EstimatingSystem!element]" & vbCrLf
    s = s & ",'' [Divisions!2!TakeoffSystem!element]" & vbCrLf
    s = s & ",'' [Stats!3!Year!element]" & vbCrLf
    s = s & ",'' [Stats!3!Month!element]" & vbCrLf
    s = s & ",'' [Stats!3!JobCreates!element]" & vbCrLf
    s = s & ",'' [Stats!3!JobStarts!element]" & vbCrLf
    s = s & ",'' [Stats!3!POsIssued!element]" & vbCrLf
    s = s & ",'' [Stats!3!SchedulesCreated!element]" & vbCrLf
    s = s & ",'' [Stats!3!WarrantyWorkordersCreated!element]" & vbCrLf
    s = s & ",'' [Stats!3!WorkticketsCreated!element]" & vbCrLf
    s = s & ",'' [Suppliers!4!Name!element]" & vbCrLf
    s = s & ",'' [Suppliers!4!Address1!element]" & vbCrLf
    s = s & ",'' [Suppliers!4!Address2!element]" & vbCrLf
    s = s & ",'' [Suppliers!4!City!element]" & vbCrLf
    s = s & ",'' [Suppliers!4!State!element]" & vbCrLf
    s = s & ",'' [Suppliers!4!PostalCode!element]" & vbCrLf
    s = s & ",'' [Suppliers!4!Phone!element]" & vbCrLf
    s = s & ",'' [Suppliers!4!GeneralEmail!element]" & vbCrLf
    s = s & ",'' [Suppliers!4!PurchEmail!element]" & vbCrLf
    s = s & ",'' [Suppliers!4!SchedEmail!element]" & vbCrLf
    s = s & ",'' [Suppliers!4!InvoiceAmt365!element]" & vbCrLf
    s = s & "Union all" & vbCrLf
    s = s & "select --- DIVISION ---" & vbCrLf
    s = s & " 2 Tag" & vbCrLf
    s = s & ",null Parent" & vbCrLf
    s = s & ",'' [Client!1!ID!element]" & vbCrLf
    s = s & ",'' [Client!1!Name!element]" & vbCrLf
    s = s & ",'' [Client!1!AppVersion!element]" & vbCrLf
    s = s & ",'' [Client!1!License!element]" & vbCrLf
    s = s & ",'' [Client!1!Database!element]" & vbCrLf
    s = s & ",d.DivisionCode  [Divisions!2!Code!element]" & vbCrLf
    s = s & ",d.DivisionName  [Divisions!2!Name!element]" & vbCrLf
    s = s & ",bpc.OptionValue [Divisions!2!BuildProCompanyCode!element]" & vbCrLf
    s = s & ",bpe.OptionValue [Divisions!2!BuildProEnvironment!element]" & vbCrLf
    s = s & ",cast(ds.DataSourceID as varchar(40)) [Divisions!2!WalletPartyID!element]" & vbCrLf
    s = s & ",case oas.OptionValue when 1 then 'Sage 300 CRE' when 2 then 'Sage 100' when 3 then 'QuickBooks' when 4 then 'Sage 50' when 5 then 'MYOB' when 6 then 'Peachtree' when 7 then 'Xero' when 8 then 'Spectrum' when 9 then 'Intacct' when 10 then 'QBO' else 'none' end [Divisions!2!AccountingSystem!element]" & vbCrLf
    s = s & ",case oes.OptionValue when 1 then 'Pipeline' when 2 then 'Sage 300' else 'none' end [Divisions!2!EstimatingSystem!element]" & vbCrLf
    s = s & ",ots.OptionValue [Divisions!2!TakeoffSystem!element]" & vbCrLf
    s = s & ",'' [Stats!3!Year!element]" & vbCrLf
    s = s & ",'' [Stats!3!Month!element]" & vbCrLf
    s = s & ",'' [Stats!3!JobCreates!element]" & vbCrLf
    s = s & ",'' [Stats!3!JobStarts!element]" & vbCrLf
    s = s & ",'' [Stats!3!POsIssued!element]" & vbCrLf
    s = s & ",'' [Stats!3!SchedulesCreated!element]" & vbCrLf
    s = s & ",'' [Stats!3!WarrantyWorkordersCreated!element]" & vbCrLf
    s = s & ",'' [Stats!3!WorkticketsCreated!element]" & vbCrLf
    s = s & ",'' [Suppliers!4!Name!element]" & vbCrLf
    s = s & ",'' [Suppliers!4!Address1!element]" & vbCrLf
    s = s & ",'' [Suppliers!4!Address2!element]" & vbCrLf
    s = s & ",'' [Suppliers!4!City!element]" & vbCrLf
    s = s & ",'' [Suppliers!4!State!element]" & vbCrLf
    s = s & ",'' [Suppliers!4!PostalCode!element]" & vbCrLf
    s = s & ",'' [Suppliers!4!Phone!element]" & vbCrLf
    s = s & ",'' [Suppliers!4!GeneralEmail!element]" & vbCrLf
    s = s & ",'' [Suppliers!4!PurchEmail!element]" & vbCrLf
    s = s & ",'' [Suppliers!4!SchedEmail!element]" & vbCrLf
    s = s & ",'' [Suppliers!4!InvoiceAmt365!element]" & vbCrLf
    s = s & "from system_setup s" & vbCrLf
    s = s & "join divisions d on s.id=d.divisionid" & vbCrLf
    s = s & "left outer join datasources ds on d.bookofaccount=ds.bookofaccount" & vbCrLf
    s = s & "left outer join appoptions oas on d.divisionid=oas.divisionid and oas.optionname='AccountingSystem'" & vbCrLf
    s = s & "left outer join appoptions bpc on d.divisionid=bpc.divisionid and bpc.optionname='BuildProCompanyCode'" & vbCrLf
    s = s & "left outer join appoptions bpe on d.divisionid=bpe.divisionid and bpe.optionname='BuildProEnvironment'" & vbCrLf
    s = s & "left outer join appoptions oes on d.divisionid=oes.divisionid and oes.optionname='EstimatingSystem'" & vbCrLf
    s = s & "left outer join appoptions ots on d.divisionid=ots.divisionid and ots.optionname='TakeoffSystem'" & vbCrLf
    s = s & "Union all" & vbCrLf
    s = s & "select --- STATISTICS ---" & vbCrLf
    s = s & " 3 Tag" & vbCrLf
    s = s & ",2 Parent" & vbCrLf
    s = s & ",'' [Client!1!ID!element]" & vbCrLf
    s = s & ",'' [Client!1!Name!element]" & vbCrLf
    s = s & ",'' [Client!1!AppVersion!element]" & vbCrLf
    s = s & ",'' [Client!1!License!element]" & vbCrLf
    s = s & ",'' [Client!1!Database!element]" & vbCrLf
    s = s & ",DivisionCode  [Divisions!2!Code!element]" & vbCrLf
    s = s & ",'' [Divisions!2!Name!element]" & vbCrLf
    s = s & ",'' [Divisions!2!BuildProCompanyCode!element]" & vbCrLf
    s = s & ",'' [Divisions!2!BuildProEnvironment!element]" & vbCrLf
    s = s & ",'' [Divisions!2!WalletPartyID!element]" & vbCrLf
    s = s & ",'' [Divisions!2!AccountingSystem!element]" & vbCrLf
    s = s & ",'' [Divisions!2!EstimatingSystem!element]" & vbCrLf
    s = s & ",'' [Divisions!2!TakeoffSystem!element]" & vbCrLf
    s = s & ",Year  [Stats!3!Year!element]" & vbCrLf
    s = s & ",Month [Stats!3!Month!element]" & vbCrLf
    s = s & ",sum(CreateCount)    [Stats!3!JobCreates!element]" & vbCrLf
    s = s & ",sum(StartCount)     [Stats!3!JobStarts!element]" & vbCrLf
    s = s & ",sum(POCount)        [Stats!3!POsIssued!element]" & vbCrLf
    s = s & ",sum(SchedCount)     [Stats!3!SchedulesCreated!element]" & vbCrLf
    s = s & ",sum(WorkorderCount) [Stats!3!WarrantyWorkordersCreated!element]" & vbCrLf
    s = s & ",sum(WorkTicketCount)[Stats!3!WorkticketsCreated!element]" & vbCrLf
    s = s & ",'' [Suppliers!4!Name!element]" & vbCrLf
    s = s & ",'' [Suppliers!4!Address1!element]" & vbCrLf
    s = s & ",'' [Suppliers!4!Address2!element]" & vbCrLf
    s = s & ",'' [Suppliers!4!City!element]" & vbCrLf
    s = s & ",'' [Suppliers!4!State!element]" & vbCrLf
    s = s & ",'' [Suppliers!4!PostalCode!element]" & vbCrLf
    s = s & ",'' [Suppliers!4!Phone!element]" & vbCrLf
    s = s & ",'' [Suppliers!4!GeneralEmail!element]" & vbCrLf
    s = s & ",'' [Suppliers!4!PurchEmail!element]" & vbCrLf
    s = s & ",'' [Suppliers!4!SchedEmail!element]" & vbCrLf
    s = s & ",'' [Suppliers!4!InvoiceAmt365!element]" & vbCrLf
    s = s & "from" & vbCrLf
    s = s & "(" & vbCrLf
    s = s & "   select" & vbCrLf
    s = s & "    d.divisioncode" & vbCrLf
    s = s & "   ,year(x.CreatedDate) Year" & vbCrLf
    s = s & "   ,datename(month,x.CreatedDate) Month" & vbCrLf
    s = s & "   ,count(*) CreateCount" & vbCrLf
    s = s & "   ,0        StartCount" & vbCrLf
    s = s & "   ,0        POCount" & vbCrLf
    s = s & "   ,0        SchedCount" & vbCrLf
    s = s & "   ,0        WorkorderCount" & vbCrLf
    s = s & "   ,0        WorkTicketCount" & vbCrLf
    s = s & "   from system_setup s join divisions d on s.id=d.divisionid" & vbCrLf
    s = s & "   join tbljobs x on d.divisionid=x.divisionid " & vbCrLf
    s = s & "   group by d.divisioncode,year(x.CreatedDate),datename(month,x.CreatedDate)" & vbCrLf
    s = s & "   union all -------------------------------" & vbCrLf
    s = s & "   select" & vbCrLf
    s = s & "    d.divisioncode" & vbCrLf
    s = s & "   ,year(x.JobDate) Year" & vbCrLf
    s = s & "   ,datename(month,x.JobDate) Month" & vbCrLf
    s = s & "   ,0        CreateCount" & vbCrLf
    s = s & "   ,count(x.job_no) StartCount" & vbCrLf
    s = s & "   ,0        POCount" & vbCrLf
    s = s & "   ,0        SchedCount" & vbCrLf
    s = s & "   ,0        WorkorderCount" & vbCrLf
    s = s & "   ,0        WorkTicketCount" & vbCrLf
    s = s & "   from system_setup s join divisions d on s.id=d.divisionid" & vbCrLf
    s = s & "   join tbljobs x on d.divisionid=x.divisionid and x.JobDate is not null" & vbCrLf
    s = s & "   group by d.divisioncode,year(x.JobDate),datename(month,x.JobDate)" & vbCrLf
    s = s & "   union all -------------------------------" & vbCrLf
    s = s & "   select" & vbCrLf
    s = s & "    d.divisioncode" & vbCrLf
    s = s & "   ,year(x.podate) Year" & vbCrLf
    s = s & "   ,datename(month,x.PODate) Month" & vbCrLf
    s = s & "   ,0        CreateCount" & vbCrLf
    s = s & "   ,0        StartCount" & vbCrLf
    s = s & "   ,count(*) POCount" & vbCrLf
    s = s & "   ,0        SchedCount" & vbCrLf
    s = s & "   ,0        WorkorderCount" & vbCrLf
    s = s & "   ,0        WorkTicketCount" & vbCrLf
    s = s & "   from system_setup s join divisions d on s.id=d.divisionid" & vbCrLf
    s = s & "   join POMaster x on d.divisionid=x.divisionid and x.podate is not null and x.cancelleddate is null" & vbCrLf
    s = s & "   group by d.divisioncode,year(x.podate),datename(month,x.PODate)" & vbCrLf
    s = s & "   union all -------------------------------" & vbCrLf
    s = s & "   select" & vbCrLf
    s = s & "    d.divisioncode" & vbCrLf
    s = s & "   ,year(x.baselinestart) Year" & vbCrLf
    s = s & "   ,datename(month,x.baselinestart) Month" & vbCrLf
    s = s & "   ,0        CreateCount" & vbCrLf
    s = s & "   ,0        StartCount" & vbCrLf
    s = s & "   ,0        POCount" & vbCrLf
    s = s & "   ,count(*) SchedCount" & vbCrLf
    s = s & "   ,0        WorkorderCount" & vbCrLf
    s = s & "   ,0        WorkTicketCount" & vbCrLf
    s = s & "   from divisions d " & vbCrLf
    s = s & "   left outer join schedule x on d.divisionid=x.divisionid and x.baselinestart is not null" & vbCrLf
    s = s & "   group by d.divisioncode,year(x.baselinestart),datename(month,x.baselinestart)" & vbCrLf
    s = s & "   union all -------------------------------" & vbCrLf
    s = s & "   select" & vbCrLf
    s = s & "    d.divisioncode" & vbCrLf
    s = s & "   ,year(x.createddate) Year" & vbCrLf
    s = s & "   ,datename(month,x.createddate) Month" & vbCrLf
    s = s & "   ,0        CreateCount" & vbCrLf
    s = s & "   ,0        StartCount" & vbCrLf
    s = s & "   ,0        POCount" & vbCrLf
    s = s & "   ,0        SchedCount" & vbCrLf
    s = s & "   ,count(*) WorkorderCount" & vbCrLf
    s = s & "   ,0        WorkTicketCount" & vbCrLf
    s = s & "   from divisions d " & vbCrLf
    s = s & "   left outer join workorders x on d.divisionid=x.divisionid and x.createddate is not null" & vbCrLf
    s = s & "   group by d.divisioncode,year(x.createddate),datename(month,x.createddate)" & vbCrLf
    s = s & "   union all -------------------------------" & vbCrLf
    s = s & "   select" & vbCrLf
    s = s & "    d.divisioncode" & vbCrLf
    s = s & "   ,year(x.ticketdate) Year" & vbCrLf
    s = s & "   ,datename(month,x.ticketdate) Month" & vbCrLf
    s = s & "   ,0        CreateCount" & vbCrLf
    s = s & "   ,0        StartCount" & vbCrLf
    s = s & "   ,0        POCount" & vbCrLf
    s = s & "   ,0        SchedCount" & vbCrLf
    s = s & "   ,0        WorkorderCount" & vbCrLf
    s = s & "   ,count(*) WorkTicketCount" & vbCrLf
    s = s & "   from divisions d " & vbCrLf
    s = s & "   left outer join worktickets x on d.divisionid=x.divisionid and x.ticketdate is not null" & vbCrLf
    s = s & "   group by d.divisioncode,year(x.ticketdate),datename(month,x.ticketdate)" & vbCrLf
    s = s & ") tt" & vbCrLf
    s = s & "group by divisioncode,Year,Month" & vbCrLf
    s = s & "union all" & vbCrLf

    s = s & "select --- RECENT VENDORS ---" & vbCrLf
    s = s & " 4 Tag" & vbCrLf
    s = s & ",null Parent" & vbCrLf
    s = s & ",'' [Client!1!ID!element]" & vbCrLf
    s = s & ",'' [Client!1!Name!element]" & vbCrLf
    s = s & ",'' [Client!1!AppVersion!element]" & vbCrLf
    s = s & ",'' [Client!1!License!element]" & vbCrLf
    s = s & ",'' [Client!1!Database!element]" & vbCrLf
    s = s & ",'' [Divisions!2!Code!element]" & vbCrLf
    s = s & ",'' [Divisions!2!Name!element]" & vbCrLf
    s = s & ",'' [Divisions!2!BuildProCompanyCode!element]" & vbCrLf
    s = s & ",'' [Divisions!2!BuildProEnvironment!element]" & vbCrLf
    s = s & ",'' [Divisions!2!WalletPartyID!element]" & vbCrLf
    s = s & ",'' [Divisions!2!AccountingSystem!element]" & vbCrLf
    s = s & ",'' [Divisions!2!EstimatingSystem!element]" & vbCrLf
    s = s & ",'' [Divisions!2!TakeoffSystem!element]" & vbCrLf
    s = s & ",'' [Stats!3!Year!element]" & vbCrLf
    s = s & ",'' [Stats!3!Month!element]" & vbCrLf
    s = s & ",'' [Stats!3!JobCreates!element]" & vbCrLf
    s = s & ",'' [Stats!3!JobStarts!element]" & vbCrLf
    s = s & ",'' [Stats!3!POsIssued!element]" & vbCrLf
    s = s & ",'' [Stats!3!SchedulesCreated!element]" & vbCrLf
    s = s & ",'' [Stats!3!WarrantyWorkordersCreated!element]" & vbCrLf
    s = s & ",'' [Stats!3!WorkticketsCreated!element]" & vbCrLf
    s = s & ",isnull(v.Vendor_name,'') [Suppliers!4!Name!element]" & vbCrLf
    s = s & ",isnull(v.addr1,'')       [Suppliers!4!Address1!element]" & vbCrLf
    s = s & ",isnull(v.addr2,'')       [Suppliers!4!Address2!element]" & vbCrLf
    s = s & ",isnull(v.city,'')        [Suppliers!4!City!element]" & vbCrLf
    s = s & ",isnull(v.state,'')       [Suppliers!4!State!element]" & vbCrLf
    s = s & ",isnull(v.zip,'')         [Suppliers!4!PostalCode!element]" & vbCrLf
    s = s & ",isnull(v.Phone,'')       [Suppliers!4!Phone!element]" & vbCrLf
    s = s & ",isnull(v.email,'')       [Suppliers!4!GeneralEmail!element]" & vbCrLf
    s = s & ",isnull(v.PurchEmail,'')  [Suppliers!4!PurchEmail!element]" & vbCrLf
    s = s & ",isnull(v.SchedEmail,'')  [Suppliers!4!SchedEmail!element]" & vbCrLf
    s = s & ",sum(i.pretax+i.tax)      [Suppliers!4!InvoiceAmt365!element]" & vbCrLf
    s = s & "from tblvendors v" & vbCrLf
    s = s & "join invoices i on v.divisionid=i.divisionid and v.vendor_id=i.vendor" & vbCrLf
    s = s & "where i.invoicedate > getdate()-365  " & vbCrLf
    s = s & "group by " & vbCrLf
    s = s & " isnull(v.Vendor_name,'') " & vbCrLf
    s = s & ",isnull(v.addr1,'')      " & vbCrLf
    s = s & ",isnull(v.addr2,'')      " & vbCrLf
    s = s & ",isnull(v.city,'')       " & vbCrLf
    s = s & ",isnull(v.state,'')      " & vbCrLf
    s = s & ",isnull(v.zip,'')        " & vbCrLf
    s = s & ",isnull(v.Phone,'')      " & vbCrLf
    s = s & ",isnull(v.email,'')      " & vbCrLf
    s = s & ",isnull(v.PurchEmail,'') " & vbCrLf
    s = s & ",isnull(v.SchedEmail,'')" & vbCrLf
    
    s = s & "order by [Divisions!2!Code!element], [Parent], [Stats!3!Year!element], [Stats!3!Month!element]" & vbCrLf
    s = s & "for xml explicit" & vbCrLf
    
    Set rs = HFApp.SqlExec(s)
    xml = ""
    While Not rs.EOF
        xml = xml & rs(0)
        rs.MoveNext
    Wend
    
    
    'now get wallet payments
    s = ""
    s = s & "select " & vbCrLf
    s = s & " ap.DataSourceID WalletPartyID" & vbCrLf
    s = s & ",datepart(year, ap.PaymentDate) Year" & vbCrLf
    s = s & ",datepart(Month, ap.PaymentDate) Month" & vbCrLf
    s = s & ",datepart(week, ap.PaymentDate) Week" & vbCrLf
    s = s & ",count(distinct isnull(nullif(ap.CheckNumber,''),ap.paymentnumber)) Payments" & vbCrLf
    s = s & ",sum(ap.NetAmount) Total" & vbCrLf
    s = s & ",count(distinct case when wp.WalletBatchID is null then null else isnull(nullif(ap.CheckNumber,''),ap.paymentnumber) end) WalletPayments" & vbCrLf
    s = s & ",sum(case when wp.WalletBatchID is null then 0 else ap.NetAmount end) WalletTotal" & vbCrLf
    s = s & ",count(distinct wp.WalletBatchID) WalletBatches" & vbCrLf
    s = s & "from accountingap ap" & vbCrLf
    s = s & "left join Wallet_Payments wp on ap.DataSourceID=wp.DataSourceID and ap.TransactionID=wp.TransactionID" & vbCrLf
    s = s & "where ap.PaymentDate > getdate()-120" & vbCrLf
    s = s & "group by" & vbCrLf
    s = s & " ap.DataSourceID" & vbCrLf
    s = s & ",datepart(year, ap.PaymentDate) " & vbCrLf
    s = s & ",datepart(Month, ap.PaymentDate) " & vbCrLf
    s = s & ",datepart(week, ap.PaymentDate) " & vbCrLf
    s = s & "for xml raw('Wallet'),elements" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    While Not rs.EOF
        xml = xml & rs(0)
        rs.MoveNext
    Wend
    
    
    xml = "<HFDesktopAuth>" & xml & "</HFDesktopAuth>"
    
    Call xx.SetStats(xml, IsProduction)
    
    'if no response show different warning
    IsProductionVerified = IsProduction <> ""
    IsProductionDatabase = IsProduction = "True"

Exit Sub
eh:
End Sub

Private Function SyncCmd(cmd As String, Optional parameters As String, Optional SkipAccounting As Boolean = False) As String
    SyncCmd = App.Path & "\HFSync.exe " & _
              HFApp.ConnectionString(dbHomefront) & Chr(1) & _
              HFApp.DivisionID & Chr(1) & _
              IIf(SkipAccounting, 1, 0) & Chr(1) & _
              cmd & Chr(1) & _
              parameters
End Function




