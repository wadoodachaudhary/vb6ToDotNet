VERSION 5.00
Begin VB.Form FCompany 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Company"
   ClientHeight    =   3690
   ClientLeft      =   5055
   ClientTop       =   2100
   ClientWidth     =   5310
   Icon            =   "FCompany.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox text1 
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   11
      Left            =   1695
      MaxLength       =   50
      TabIndex        =   7
      Text            =   " "
      Top             =   2145
      Width           =   3375
   End
   Begin VB.TextBox text1 
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   7
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   6
      Text            =   " "
      Top             =   1755
      Width           =   3375
   End
   Begin VB.TextBox text1 
      BorderStyle     =   0  'None
      Height          =   420
      Index           =   2
      Left            =   1680
      MaxLength       =   102
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   570
      Width           =   3375
   End
   Begin VB.TextBox text1 
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   3
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   2
      Top             =   1005
      Width           =   2565
   End
   Begin VB.TextBox text1 
      BorderStyle     =   0  'None
      Height          =   230
      Index           =   5
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   4
      Top             =   1260
      Width           =   1395
   End
   Begin VB.TextBox text1 
      BorderStyle     =   0  'None
      Height          =   240
      IMEMode         =   3  'DISABLE
      Index           =   8
      Left            =   1695
      MaxLength       =   30
      TabIndex        =   8
      Top             =   2400
      Width           =   1755
   End
   Begin VB.TextBox text1 
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   1
      Left            =   1680
      TabIndex        =   0
      Text            =   " "
      Top             =   180
      Width           =   3375
   End
   Begin VB.TextBox text1 
      BorderStyle     =   0  'None
      Height          =   240
      IMEMode         =   3  'DISABLE
      Index           =   9
      Left            =   1695
      MaxLength       =   30
      TabIndex        =   9
      Top             =   2655
      Width           =   1755
   End
   Begin VB.TextBox text1 
      BorderStyle     =   0  'None
      Height          =   240
      IMEMode         =   3  'DISABLE
      Index           =   10
      Left            =   1695
      MaxLength       =   20
      TabIndex        =   10
      Top             =   3060
      Width           =   1575
   End
   Begin HFSystem.VBCombo cboCompanyCountry 
      Height          =   240
      Left            =   1680
      TabIndex        =   5
      Top             =   1500
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   423
   End
   Begin HFSystem.VBCombo cboCompanyProvince 
      Height          =   240
      Left            =   4260
      TabIndex        =   3
      Top             =   1005
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   423
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Email"
      Height          =   195
      Index           =   1
      Left            =   420
      TabIndex        =   20
      Top             =   2190
      Width           =   1155
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "County"
      Height          =   195
      Index           =   0
      Left            =   405
      TabIndex        =   19
      Top             =   1800
      Width           =   1155
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "City/Prov"
      Height          =   195
      Left            =   900
      TabIndex        =   18
      Top             =   1020
      Width           =   660
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Address"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   990
      TabIndex        =   17
      Top             =   570
      Width           =   570
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Postal Code"
      Height          =   195
      Left            =   705
      TabIndex        =   16
      Top             =   1275
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Company Name"
      Height          =   195
      Index           =   51
      Left            =   435
      TabIndex        =   15
      Top             =   180
      Width           =   1125
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Phone"
      Height          =   195
      Index           =   47
      Left            =   420
      TabIndex        =   14
      Top             =   2400
      Width           =   1155
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Fax"
      Height          =   195
      Index           =   45
      Left            =   420
      TabIndex        =   13
      Top             =   2655
      Width           =   1155
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Tax Number"
      Height          =   195
      Index           =   52
      Left            =   180
      TabIndex        =   12
      Top             =   3030
      Width           =   1395
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Country"
      Height          =   195
      Index           =   53
      Left            =   405
      TabIndex        =   11
      Top             =   1515
      Width           =   1155
   End
End
Attribute VB_Name = "FCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private field(11) As String

Public Sub ShowForm(Grid As VSFlexGrid)
    With Grid
        text1(1) = .TextMatrix(.Row, .ColIndex("Company"))
        text1(2) = .TextMatrix(.Row, .ColIndex("Address1")) & vbCrLf & _
                   .TextMatrix(.Row, .ColIndex("Address2"))
        text1(3) = .TextMatrix(.Row, .ColIndex("City"))
        text1(5) = .TextMatrix(.Row, .ColIndex("Postal"))
        Call SetComboBoxListIndex(cboCompanyCountry, .TextMatrix(.Row, .ColIndex("Country")))
        Call SetComboBoxListIndex(cboCompanyProvince, .TextMatrix(.Row, .ColIndex("Province")))
        text1(7) = .TextMatrix(.Row, .ColIndex("County"))
        text1(8) = .TextMatrix(.Row, .ColIndex("Phone"))
        text1(9) = .TextMatrix(.Row, .ColIndex("Fax"))
        text1(10) = .TextMatrix(.Row, .ColIndex("TaxNumber"))
        text1(11) = .TextMatrix(.Row, .ColIndex("Email"))
        Me.Show vbModal
        .TextMatrix(.Row, .ColIndex("Company")) = field(1)
        .TextMatrix(.Row, .ColIndex("Address1")) = Parse(field(2), 1, vbCrLf)
        .TextMatrix(.Row, .ColIndex("Address2")) = Parse(field(2), 2, vbCrLf)
        .TextMatrix(.Row, .ColIndex("City")) = field(3)
        .TextMatrix(.Row, .ColIndex("Province")) = field(4)
        .TextMatrix(.Row, .ColIndex("Postal")) = field(5)
        .TextMatrix(.Row, .ColIndex("Country")) = field(6)
        .TextMatrix(.Row, .ColIndex("County")) = field(7)
        .TextMatrix(.Row, .ColIndex("Phone")) = field(8)
        .TextMatrix(.Row, .ColIndex("Fax")) = field(9)
        .TextMatrix(.Row, .ColIndex("TaxNumber")) = field(10)
        .TextMatrix(.Row, .ColIndex("Email")) = field(11)
    End With
End Sub

Private Sub cboCompanyProvince_Change()
    Call cboCompanyProvince_Click
End Sub
Private Sub cboCompanyProvince_Click()
    field(4) = cboCompanyProvince.Text
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub
Private Sub Form_Load()
    Call IniGetForm(Me)
    Call LoadComboBox(cboCompanyCountry, HFApp.Databases(dbHomefront), "select distinct country,'',0 from countrycodes order by 1")
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

Private Sub cboCompanyCountry_Change()
 '   Call cboCompanyCountry_Click
End Sub
Private Sub cboCompanyCountry_Click()
    field(6) = cboCompanyCountry.Text
    cboCompanyProvince.Clear
    cboCompanyProvince.ListIndex = -1
    Call LoadComboBox(cboCompanyProvince, HFApp.Databases(dbHomefront), "select distinct state,'',0 from countrycodes where country=" & DbQuote(Str, cboCompanyCountry.Text) & "order by 1")
End Sub


Private Sub text1_Validate(Index As Integer, Cancel As Boolean)
    If IsIn(Index, 8, 9) Then
        text1(Index).Text = FormatPhone(text1(Index).Text)
    End If
End Sub
