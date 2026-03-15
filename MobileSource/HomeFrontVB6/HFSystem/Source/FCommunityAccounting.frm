VERSION 5.00
Begin VB.Form FCommunityAccounting 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Accounting"
   ClientHeight    =   5025
   ClientLeft      =   4650
   ClientTop       =   4950
   ClientWidth     =   6375
   Icon            =   "FCommunityAccounting.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox text1 
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   17
      Left            =   1995
      MaxLength       =   50
      TabIndex        =   13
      Top             =   4320
      Width           =   2475
   End
   Begin VB.TextBox text1 
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   16
      Left            =   1995
      MaxLength       =   50
      TabIndex        =   11
      Text            =   " "
      Top             =   3465
      Width           =   2790
   End
   Begin VB.TextBox text1 
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   15
      Left            =   1995
      MaxLength       =   50
      TabIndex        =   10
      Text            =   " "
      Top             =   3210
      Width           =   2790
   End
   Begin VB.TextBox text1 
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   0
      Left            =   1995
      MaxLength       =   50
      TabIndex        =   12
      Top             =   4065
      Width           =   2475
   End
   Begin VB.TextBox text1 
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   9
      Left            =   4200
      MaxLength       =   50
      TabIndex        =   8
      Text            =   " "
      Top             =   2475
      Width           =   945
   End
   Begin VB.TextBox text1 
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   10
      Left            =   4200
      MaxLength       =   50
      TabIndex        =   9
      Top             =   2730
      Width           =   945
   End
   Begin VB.TextBox text1 
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   7
      Left            =   1995
      MaxLength       =   50
      TabIndex        =   6
      Text            =   " "
      Top             =   2730
      Width           =   945
   End
   Begin VB.TextBox text1 
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   8
      Left            =   4200
      MaxLength       =   50
      TabIndex        =   7
      Top             =   2220
      Width           =   945
   End
   Begin VB.TextBox text1 
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   5
      Left            =   1995
      MaxLength       =   50
      TabIndex        =   4
      Text            =   " "
      Top             =   2220
      Width           =   945
   End
   Begin VB.TextBox text1 
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   6
      Left            =   1995
      MaxLength       =   50
      TabIndex        =   5
      Top             =   2475
      Width           =   945
   End
   Begin VB.Frame frmAccounts 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1980
      Index           =   1
      Left            =   15
      TabIndex        =   28
      Top             =   0
      Width           =   5790
      Begin VB.TextBox text1 
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   14
         Left            =   1740
         MaxLength       =   50
         TabIndex        =   3
         Text            =   " "
         Top             =   1485
         Width           =   3135
      End
      Begin VB.TextBox text1 
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   13
         Left            =   1740
         MaxLength       =   50
         TabIndex        =   2
         Text            =   " "
         Top             =   1230
         Width           =   3135
      End
      Begin VB.TextBox text1 
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   11
         Left            =   1740
         MaxLength       =   50
         TabIndex        =   0
         Text            =   " "
         Top             =   720
         Width           =   3135
      End
      Begin VB.TextBox text1 
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   12
         Left            =   1740
         MaxLength       =   50
         TabIndex        =   1
         Text            =   " "
         Top             =   975
         Width           =   3135
      End
      Begin VB.Image cmdBrowse 
         Height          =   240
         Index           =   14
         Left            =   4890
         Picture         =   "FCommunityAccounting.frx":000C
         Top             =   1485
         Width           =   240
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Department"
         Height          =   195
         Left            =   795
         TabIndex        =   38
         Top             =   1500
         Width           =   825
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Parent Job"
         Height          =   195
         Left            =   855
         TabIndex        =   37
         Top             =   1245
         Width           =   765
      End
      Begin VB.Image cmdBrowse 
         Height          =   240
         Index           =   13
         Left            =   4890
         Picture         =   "FCommunityAccounting.frx":0156
         Top             =   1245
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Entity"
         Height          =   195
         Index           =   15
         Left            =   1230
         TabIndex        =   36
         Top             =   720
         Width           =   390
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Location"
         Height          =   195
         Left            =   1005
         TabIndex        =   35
         Top             =   990
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Organization"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   0
         Left            =   585
         TabIndex        =   34
         Top             =   390
         Width           =   1110
      End
      Begin VB.Image cmdBrowse 
         Height          =   240
         Index           =   11
         Left            =   4890
         Picture         =   "FCommunityAccounting.frx":02A0
         Top             =   720
         Width           =   240
      End
      Begin VB.Image cmdBrowse 
         Height          =   240
         Index           =   12
         Left            =   4890
         Picture         =   "FCommunityAccounting.frx":03EA
         Top             =   990
         Width           =   240
      End
   End
   Begin VB.Frame frmAccounts 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1875
      Index           =   0
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   5790
      Begin VB.TextBox text1 
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   1
         Left            =   1740
         MaxLength       =   50
         TabIndex        =   14
         Text            =   " "
         Top             =   480
         Width           =   3135
      End
      Begin VB.TextBox text1 
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   2
         Left            =   1740
         MaxLength       =   50
         TabIndex        =   15
         Text            =   " "
         Top             =   735
         Width           =   3135
      End
      Begin VB.TextBox text1 
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   3
         Left            =   1740
         MaxLength       =   50
         TabIndex        =   16
         Text            =   " "
         Top             =   1110
         Width           =   3135
      End
      Begin VB.TextBox text1 
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   4
         Left            =   1740
         MaxLength       =   50
         TabIndex        =   17
         Top             =   1365
         Width           =   3135
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Income Stmt Prefix"
         Height          =   195
         Index           =   13
         Left            =   300
         TabIndex        =   33
         Top             =   480
         Width           =   1320
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Balance Sheet Prefix"
         Height          =   195
         Index           =   14
         Left            =   135
         TabIndex        =   32
         Top             =   750
         Width           =   1485
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Lot Inventory Debit"
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   31
         Top             =   1110
         Width           =   1350
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Lot Inventory Credit"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   30
         Top             =   1380
         Width           =   1380
      End
      Begin VB.Label Label1 
         Caption         =   "GL Accounts"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   10
         Left            =   585
         TabIndex        =   29
         Top             =   240
         Width           =   1110
      End
      Begin VB.Image cmdBrowse 
         Height          =   240
         Index           =   1
         Left            =   4905
         Picture         =   "FCommunityAccounting.frx":0534
         Top             =   480
         Width           =   240
      End
      Begin VB.Image cmdBrowse 
         Height          =   240
         Index           =   2
         Left            =   4905
         Picture         =   "FCommunityAccounting.frx":067E
         Top             =   735
         Width           =   240
      End
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Tarion Builder No"
      Height          =   195
      Index           =   17
      Left            =   645
      TabIndex        =   41
      Top             =   4335
      Width           =   1230
   End
   Begin VB.Image cmdBrowse 
      Height          =   240
      Index           =   16
      Left            =   4800
      Picture         =   "FCommunityAccounting.frx":07C8
      Top             =   3480
      Width           =   240
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Bank Account"
      Height          =   195
      Left            =   855
      TabIndex        =   40
      Top             =   3480
      Width           =   1020
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Loan Draw Liability Acct"
      Height          =   195
      Left            =   -165
      TabIndex        =   39
      Top             =   3225
      Width           =   2040
   End
   Begin VB.Image cmdBrowse 
      Height          =   240
      Index           =   15
      Left            =   4800
      Picture         =   "FCommunityAccounting.frx":0912
      Top             =   3210
      Width           =   240
   End
   Begin VB.Image cmdBrowse 
      Height          =   240
      Index           =   10
      Left            =   5160
      Picture         =   "FCommunityAccounting.frx":0A5C
      Top             =   2745
      Width           =   240
   End
   Begin VB.Image cmdBrowse 
      Height          =   240
      Index           =   9
      Left            =   5160
      Picture         =   "FCommunityAccounting.frx":0BA6
      Top             =   2490
      Width           =   240
   End
   Begin VB.Image cmdBrowse 
      Height          =   240
      Index           =   8
      Left            =   5160
      Picture         =   "FCommunityAccounting.frx":0CF0
      Top             =   2235
      Width           =   240
   End
   Begin VB.Image cmdBrowse 
      Height          =   240
      Index           =   7
      Left            =   2955
      Picture         =   "FCommunityAccounting.frx":0E3A
      Top             =   2745
      Width           =   240
   End
   Begin VB.Image cmdBrowse 
      Height          =   240
      Index           =   6
      Left            =   2955
      Picture         =   "FCommunityAccounting.frx":0F84
      Top             =   2490
      Width           =   240
   End
   Begin VB.Image cmdBrowse 
      Height          =   240
      Index           =   5
      Left            =   2955
      Picture         =   "FCommunityAccounting.frx":10CE
      Top             =   2235
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Warranty"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   12
      Left            =   1155
      TabIndex        =   26
      Top             =   3825
      Width           =   1110
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Warranty Job"
      Height          =   195
      Index           =   11
      Left            =   600
      TabIndex        =   25
      Top             =   4080
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tax Groups"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   9
      Left            =   945
      TabIndex        =   24
      Top             =   1980
      Width           =   1320
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Overhead"
      Height          =   195
      Index           =   8
      Left            =   3045
      TabIndex        =   23
      Top             =   2460
      Width           =   1035
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Other"
      Height          =   195
      Index           =   7
      Left            =   3360
      TabIndex        =   22
      Top             =   2730
      Width           =   720
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Sub Contracts"
      Height          =   195
      Index           =   6
      Left            =   540
      TabIndex        =   21
      Top             =   2730
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Equipment"
      Height          =   195
      Index           =   5
      Left            =   3000
      TabIndex        =   20
      Top             =   2220
      Width           =   1080
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Labour"
      Height          =   195
      Index           =   4
      Left            =   1050
      TabIndex        =   19
      Top             =   2220
      Width           =   825
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Material"
      Height          =   195
      Index           =   3
      Left            =   990
      TabIndex        =   18
      Top             =   2490
      Width           =   885
   End
End
Attribute VB_Name = "FCommunityAccounting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const SRCFILE = "FCommunityAccounting::"

Private field(17) As String
Private IntacctEntity As String
Private IntacctJobsSQL As String
Private IntacctEntitySQL As String


Public Sub ShowForm(Grid As VSFlexGrid)
    With Grid
    
        text1(0) = .TextMatrix(.Row, .ColIndex("WarrantyJob"))
        text1(1) = .TextMatrix(.Row, .ColIndex("IncomePrefix"))
        text1(2) = .TextMatrix(.Row, .ColIndex("BalanceSheet"))
        text1(3) = .TextMatrix(.Row, .ColIndex("LotInventoryDebit"))
        text1(4) = .TextMatrix(.Row, .ColIndex("LotInventoryCredit"))
        
        text1(5) = .TextMatrix(.Row, .ColIndex("LabourTaxGroup"))
        text1(6) = .TextMatrix(.Row, .ColIndex("MaterialTaxGroup"))
        text1(7) = .TextMatrix(.Row, .ColIndex("SubContractTaxGroup"))
        text1(8) = .TextMatrix(.Row, .ColIndex("EquipmentTaxGroup"))
        text1(9) = .TextMatrix(.Row, .ColIndex("OverheadTaxGroup"))
        text1(10) = .TextMatrix(.Row, .ColIndex("OtherTaxGroup"))
        
        text1(11) = .TextMatrix(.Row, .ColIndex("IntacctEntity"))
        text1(12) = .TextMatrix(.Row, .ColIndex("BalanceSheet")) 'aka location
        text1(13) = .TextMatrix(.Row, .ColIndex("IntacctParentJob"))
        text1(14) = .TextMatrix(.Row, .ColIndex("IntacctDepartment"))
        
        text1(15) = .TextMatrix(.Row, .ColIndex("Mortgage_Credit"))
        text1(16) = .TextMatrix(.Row, .ColIndex("BankAccount"))
        
        text1(17) = .TextMatrix(.Row, .ColIndex("TarionBuilderNumber"))
        Dim b As Boolean
        b = HFApp.Options.ValueByName("EnableTarionFields") = "true"
        text1(17).Visible = b
        Label1(17).Visible = b
        
        
        IntacctEntity = .TextMatrix(.Row, .ColIndex("IntacctEntity"))
        If IntacctEntity = "" Then IntacctEntity = HFApp.Options.ValueByName("IntacctEntity")
        
        
        
        
        
        Me.Show vbModal
        
        .TextMatrix(.Row, .ColIndex("WarrantyJob")) = field(0)
        
        If HFApp.Options(AccountingSystem) = asIntacct Then
            .TextMatrix(.Row, .ColIndex("IntacctEntity")) = field(11)
            .TextMatrix(.Row, .ColIndex("IncomePrefix")) = field(12) 'location is both prefixes
            .TextMatrix(.Row, .ColIndex("BalanceSheet")) = field(12)
            .TextMatrix(.Row, .ColIndex("IntacctParentJob")) = field(13)
            .TextMatrix(.Row, .ColIndex("IntacctDepartment")) = field(14)
        Else
            .TextMatrix(.Row, .ColIndex("IncomePrefix")) = field(1)
            .TextMatrix(.Row, .ColIndex("BalanceSheet")) = field(2)
        End If
            
        .TextMatrix(.Row, .ColIndex("Mortgage_Credit")) = field(15)
        .TextMatrix(.Row, .ColIndex("BankAccount")) = field(16)
        .TextMatrix(.Row, .ColIndex("TarionBuilderNumber")) = field(17)
        
        .TextMatrix(.Row, .ColIndex("LotInventoryDebit")) = field(3)
        .TextMatrix(.Row, .ColIndex("LotInventoryCredit")) = field(4)
        
        .TextMatrix(.Row, .ColIndex("LabourTaxGroup")) = field(5)
        .TextMatrix(.Row, .ColIndex("MaterialTaxGroup")) = field(6)
        .TextMatrix(.Row, .ColIndex("SubContractTaxGroup")) = field(7)
        .TextMatrix(.Row, .ColIndex("EquipmentTaxGroup")) = field(8)
        .TextMatrix(.Row, .ColIndex("OverheadTaxGroup")) = field(9)
        .TextMatrix(.Row, .ColIndex("OtherTaxGroup")) = field(10)
    
    End With
End Sub


Private Sub cmdBrowse_Click(Index As Integer)
    Call text1_KeyDown(Index, vbKeyF4, 0)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub
Private Sub Form_Load()
    Call IniGetForm(Me)

    frmAccounts(0).Visible = HFApp.Options(AccountingSystem) <> asIntacct
    frmAccounts(1).Visible = HFApp.Options(AccountingSystem) = asIntacct
    frmAccounts(1).Move 0, 0


End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
End Sub

Private Sub text1_Change(Index As Integer)
    field(Index) = text1(Index)
End Sub
Private Sub text1_GotFocus(Index As Integer)
    Call SelectAll(text1(Index))
End Sub


Private Sub text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo eh
    Dim s As String
    
    If KeyCode <> vbKeyF4 Or Shift <> 0 Then Exit Sub
    
    Select Case Index
        Case 15
            s = "select Account, Description,AccountType from glaccounts where divisionid=" & DbQuote(Num, HFApp.DivisionID)
            If FPickList.Choose(HFApp.Databases(dbHomefront), "GLAccount", s, text1(Index).Text) Then
                text1(Index).Text = FPickList.SelectedItem("Account")
            End If
         
        
        
        Case 1, 2, 12
        s = text1(Index).Text
        If HFApp.Options(AccountingSystem) = AccountingSystems.asTimberline Then
            Dim sqlstmt As String
            Dim MyRec As New adodb.Recordset
            Set MyRec = HFApp.SqlExec("select * from glm_master__account_format", dbAccounting)
            MyRec.MoveFirst
            If Not MyRec.EOF Then
                If MyRec!account_prefix_abc_length <> 0 Then
                    sqlstmt = "select account_prefix_abc as Prefix,Account_Prefix_ABC_Description as Description from glm_master__account_prefix_abc_1"
                ElseIf MyRec!account_prefix_ab_length <> 0 Then
                    sqlstmt = "select account_prefix_ab as Prefix,Account_Prefix_AB_Description as Description from glm_master__account_prefix_ab_1"
                ElseIf MyRec!account_prefix_a_length <> 0 Then
                    sqlstmt = "select account_prefix_a as Prefix,Account_Prefix_A_Description as Description from glm_master__account_prefix_a_1"
                Else
                    sqlstmt = "select null as Prefix,null as Description from glm_master__account_prefix_a_1 where 1=2"
                End If
                If Not MyRec.EOF Then
                   MyRec.Close
                   If FPickList.Choose(HFApp.Databases(dbAccounting), "Prefix List", sqlstmt, s) Then
                       text1(Index).Text = FPickList.SelectedItem("Prefix")
                   End If
                End If
            End If
        ElseIf HFApp.Options(AccountingSystem) = AccountingSystems.asIntacct Then
            s = "select description,gl_prefix Location,'' from gl_prefix where divisionid=" & DbQuote(Num, HFApp.DivisionID)
            If FPickList.Choose(HFApp.Databases(dbHomefront), "Locations", s, text1(Index).Text) Then
                text1(Index).Text = FPickList.SelectedItem("Location")
            End If
         
        End If
        Case 5, 6, 7, 8, 9, 10
            s = text1(Index).Text
            If FPickList.Choose(HFApp.Databases(dbHomefront), "TaxGroup", "select TaxGroup,Description from taxgroups where DivisionID = " & HFApp.DivisionID, s) Then
                text1(Index).Text = FPickList.SelectedItem("TaxGroup")
            End If
        
        Case 11
            If IntacctEntitySQL = "" Then
                IntacctEntitySQL = Intacct_EntityList(HFApp.Options.ValueByName("IntacctCompanyID"), _
                                                       HFApp.Options.ValueByName("IntacctUID"), _
                                                       HFApp.Options.ValueByName("IntacctPWD"), "")
            End If
            If FPickList.Choose(HFApp.Databases(dbHomefront), "Entity", IntacctEntitySQL) Then
                text1(Index).Text = FPickList.SelectedItem("EntityID")
            End If
        
        Case 14
            If FPickList.Choose(HFApp.Databases(dbHomefront), "Department", "select id,name from IntacctDepartments where DivisionID = " & HFApp.DivisionID & " order by 1") Then
                text1(Index).Text = FPickList.SelectedItem("id")
            End If
            
        Case 16
            If FPickList.Choose(HFApp.Databases(dbHomefront), "Bank Account", "select BankAccount,Description from BankAccounts where DivisionID = " & HFApp.DivisionID & " order by 1") Then
                text1(Index).Text = FPickList.SelectedItem("BankAccount")
            End If
            
            
        Case 13
            Dim Intacct As IntacctWrapper.IntacctWrapper
            If IntacctJobsSQL = "" Then
                Set Intacct = New IntacctWrapper.IntacctWrapper
                Call Intacct.OpenMessage(HFApp.Options.ValueByName("IntacctCompanyID"), HFApp.Options.ValueByName("IntacctUID"), HFApp.Options.ValueByName("IntacctPWD"), IntacctEntity)
                Call Intacct.GetJobs("")
                Call Intacct.CloseMessage
                Call WriteLogFile("intacct.getjobs.req.xml", Intacct.xml())
                s = Intacct.PostMessage(True)
                Call WriteLogFile("intacct.getjobs.res.xml", s)
                IntacctJobsSQL = "exec Intacct_Jobs " & DbQuote(Num, HFApp.DivisionID) & "," & DbQuote(Str, s) & " -- no order by" '<-- no order by is required by fpicklist
            End If
            If FPickList.Choose(HFApp.Databases(dbHomefront), "Parent Job", IntacctJobsSQL) Then
                text1(Index).Text = FPickList.SelectedItem("Job")
            End If
        
        
    End Select
Exit Sub
eh: Call errHandler(SRCFILE & "cmdBrowse", s)
End Sub

Private Sub text1_Validate(Index As Integer, Cancel As Boolean)
    Dim s  As String
    Dim rs As Recordset
    Select Case Index
        Case 5, 6, 7, 8, 9, 10
            If text1(Index).Text <> "" Then
                s = "select * from taxgroups where DivisionID = " & HFApp.DivisionID & " and taxgroup=" & DbQuote(Str, text1(Index).Text)
                Set rs = HFApp.SqlExec(s, dbHomefront)
                If rs.EOF Then
                    Cancel = True
                    MsgBox "Invalid Tax Group", vbExclamation, App.ProductName
                Else
                    text1(Index).Text = "" & rs("taxgroup")
                End If
            End If
    End Select
End Sub
