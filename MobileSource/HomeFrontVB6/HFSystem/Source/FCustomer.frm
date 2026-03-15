VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form FCustomer 
   Caption         =   "Customer"
   ClientHeight    =   6300
   ClientLeft      =   525
   ClientTop       =   1275
   ClientWidth     =   11775
   Icon            =   "FCustomer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6300
   ScaleWidth      =   11775
   Begin VB.TextBox txtComments 
      BorderStyle     =   0  'None
      Height          =   930
      Left            =   450
      MaxLength       =   4000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2100
      Width           =   4800
   End
   Begin VB.TextBox txtCustomer 
      BorderStyle     =   0  'None
      Height          =   230
      Left            =   1350
      TabIndex        =   0
      Text            =   " "
      Top             =   1020
      Width           =   2835
   End
   Begin VB.TextBox txtAddress 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   6960
      MaxLength       =   75
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   1680
      Width           =   2835
   End
   Begin VB.TextBox txtCity 
      BorderStyle     =   0  'None
      Height          =   230
      Left            =   6960
      MaxLength       =   50
      TabIndex        =   7
      Top             =   2385
      Width           =   2835
   End
   Begin VB.TextBox txtProvince 
      BorderStyle     =   0  'None
      Height          =   230
      Left            =   6960
      MaxLength       =   20
      TabIndex        =   8
      Top             =   2625
      Width           =   2835
   End
   Begin VB.TextBox txtPhone 
      BorderStyle     =   0  'None
      Height          =   230
      Left            =   6960
      MaxLength       =   25
      TabIndex        =   4
      Top             =   1020
      Width           =   1755
   End
   Begin VB.TextBox txtFax 
      BorderStyle     =   0  'None
      Height          =   230
      Left            =   6960
      MaxLength       =   25
      TabIndex        =   5
      Top             =   1260
      Width           =   1755
   End
   Begin VB.TextBox txtPostal 
      BorderStyle     =   0  'None
      Height          =   230
      Left            =   6960
      MaxLength       =   10
      TabIndex        =   9
      Top             =   2865
      Width           =   1395
   End
   Begin VB.TextBox txtDescription 
      BorderStyle     =   0  'None
      Height          =   230
      Left            =   1350
      MaxLength       =   60
      TabIndex        =   1
      Top             =   1260
      Width           =   3885
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   11
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   1058
      ButtonWidth     =   1508
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "   Open    "
            Key             =   "Open"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "Save"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList LargeIcons 
         Left            =   5490
         Top             =   30
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   55
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":000C
               Key             =   "EditAssembly"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":08E6
               Key             =   "RFP"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":11C0
               Key             =   "CreateJob"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":1A9A
               Key             =   "takeoffsettings"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":2374
               Key             =   "quote"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":2C4E
               Key             =   "DecreaseDecimals"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":3528
               Key             =   "IncreaseDecimals"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":3E02
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":46DC
               Key             =   "ExcelImport"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":4FB6
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":5890
               Key             =   "UpdatePrices"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":616A
               Key             =   "ExcelExport"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":6A44
               Key             =   "Publish"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":731E
               Key             =   "Forecast"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":7BF8
               Key             =   "ViewPOs"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":84D2
               Key             =   "ViewBudgets"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":8DAC
               Key             =   "Open"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":9686
               Key             =   "Preview"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":9F60
               Key             =   "Send"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":A83A
               Key             =   "TakeoffOneTime"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":B114
               Key             =   "Estimate"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":B9EE
               Key             =   "TakeoffAssembly"
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":C2C8
               Key             =   "NewAssembly"
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":CBA2
               Key             =   "New"
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":D47C
               Key             =   "TakeoffItem"
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":DD56
               Key             =   "TakeoffCustom"
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":E630
               Key             =   "Save"
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":EF0A
               Key             =   "SaveAs"
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":F7E4
               Key             =   "Delete"
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":100BE
               Key             =   "AddPricelist"
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":10998
               Key             =   "RePrice"
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":11272
               Key             =   "Pricebook"
            EndProperty
            BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":11B4C
               Key             =   "PricebookEdit"
            EndProperty
            BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":12426
               Key             =   "PricelistExport"
            EndProperty
            BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":12D00
               Key             =   "PricelistImport"
            EndProperty
            BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":135DA
               Key             =   "NewPricelist"
            EndProperty
            BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":13EB4
               Key             =   "View"
            EndProperty
            BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":1478E
               Key             =   "Vendor1"
            EndProperty
            BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":15068
               Key             =   "Vendor"
            EndProperty
            BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":15942
               Key             =   "Add"
            EndProperty
            BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":1621C
               Key             =   "Attachments"
            EndProperty
            BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":16AF6
               Key             =   ""
            EndProperty
            BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":173D0
               Key             =   ""
            EndProperty
            BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":17CAA
               Key             =   "Design Center Options"
            EndProperty
            BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":18584
               Key             =   "Global Options"
            EndProperty
            BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":18E5E
               Key             =   "Models"
            EndProperty
            BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":19738
               Key             =   "Options"
            EndProperty
            BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":1A012
               Key             =   "Generate"
            EndProperty
            BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":1A8EC
               Key             =   "Timberline"
            EndProperty
            BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":1B1C6
               Key             =   "MBImport"
            EndProperty
            BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":1BAA0
               Key             =   "EEEstimating"
            EndProperty
            BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":1C37A
               Key             =   "EEExport"
            EndProperty
            BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":1CC54
               Key             =   "EEImport"
            EndProperty
            BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":1D52E
               Key             =   "MBExport"
            EndProperty
            BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCustomer.frx":1DE08
               Key             =   "MasterBuilder"
            EndProperty
         EndProperty
      End
   End
   Begin HFSystem.VBCombo cboStatus 
      Height          =   240
      Left            =   1350
      TabIndex        =   2
      Top             =   1500
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   423
      Style           =   2
   End
   Begin VSFlex8Ctl.VSFlexGrid gContacts 
      Height          =   2625
      Left            =   270
      TabIndex        =   10
      Top             =   3420
      Width           =   11295
      _cx             =   1999130067
      _cy             =   1999114774
      Appearance      =   2
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
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FCustomer.frx":1EAE2
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   5
      OutlineCol      =   1
      Ellipsis        =   0
      ExplorerBar     =   7
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   1
      OwnerDraw       =   0
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   1
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label Label103 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Status"
      Height          =   195
      Left            =   795
      TabIndex        =   22
      Top             =   1515
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Comments"
      Height          =   195
      Left            =   450
      TabIndex        =   21
      Top             =   1860
      Width           =   735
   End
   Begin VB.Label Label111 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Customer"
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
      Left            =   450
      TabIndex        =   20
      Top             =   1050
      Width           =   795
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Description"
      Height          =   195
      Left            =   450
      TabIndex        =   19
      Top             =   1290
      Width           =   795
   End
   Begin VB.Label lblContacts 
      AutoSize        =   -1  'True
      Caption         =   "Contacts"
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
      Left            =   270
      TabIndex        =   18
      Top             =   3180
      Width           =   765
   End
   Begin VB.Label Label141 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "State/Province"
      Height          =   195
      Left            =   5775
      TabIndex        =   17
      Top             =   2610
      Width           =   1080
   End
   Begin VB.Label Label140 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "City"
      Height          =   195
      Left            =   6600
      TabIndex        =   16
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label Label50 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Phone"
      Height          =   195
      Left            =   6390
      TabIndex        =   15
      Top             =   1005
      Width           =   465
   End
   Begin VB.Label Label40 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Fax"
      Height          =   195
      Left            =   6600
      TabIndex        =   14
      Top             =   1278
      Width           =   255
   End
   Begin VB.Label Label1333 
      Alignment       =   1  'Right Justify
      Caption         =   "Billing Address "
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
      Height          =   405
      Left            =   6135
      TabIndex        =   13
      Top             =   1680
      Width           =   780
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Postal"
      Height          =   195
      Left            =   6420
      TabIndex        =   12
      Top             =   2895
      Width           =   435
   End
End
Attribute VB_Name = "FCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Public EventTraps As Collection
Const SRCFILE = "FCustomer::"

Private mDirty    As Boolean
Private mCustomer As String







Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case True
        Case Shift = vbCtrlMask And KeyCode = vbKeyO:  Call Toolbar_ButtonClick(Toolbar.Buttons("Open"))
        Case Shift = vbCtrlMask And KeyCode = vbKeyS:  Call Toolbar_ButtonClick(Toolbar.Buttons("Save"))
    End Select
End Sub

Private Sub Form_Load()
    Dim s As String
    
    Call SetToolbarIcons(Toolbar, LargeIcons)
    Call IniGetForm(Me)
    Call IniGetGrid(Me, gContacts, , , , True)
    
    With gContacts
        .ColComboList(.ColIndex("SMSAddress")) = "|" & .BuildComboList(HFApp.SqlExec("select address,name from smscarriers order by 2"), "*name,address", "address")
        .ColComboList(.ColIndex("Role")) = HFApp.Options.ValueByName("CustomerContactRoles")
        .ColComboList(.ColIndex("SendVia")) = "#1;Print|#3;Email|#4;Fax"
    End With
        
    cboStatus.Clear
    cboStatus.AddItem "Active"
    cboStatus.AddItem "Inactive"
    
    Call LoadData
    If mCustomer = "?" Or mCustomer = "" Then
        mCustomer = ""
        Call Toolbar_ButtonClick(Toolbar.Buttons("Open"))
        If mCustomer = "" Then Unload Me
    End If
End Sub

Private Sub Form_Resize()
    With gContacts
        .Move .left, .Top, Me.ScaleWidth - 2 * gContacts.left, Me.ScaleHeight - .left - .Top
    End With
End Sub





Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim f As Form
    Dim s As String
    
    Select Case Button.Key
    
        Case "Open"
            If Not SaveData(True) Then Exit Sub
            
            s = ""
            s = s & "Active Customers" & Chr(1) & "select * from Picklist_ActiveARCustomers where divisionid=" & DbQuote(Num, HFApp.DivisionID) & Chr(0)
            s = s & "Inactive Customers" & Chr(1) & "select* from Picklist_InActiveARCustomers where divisionid=" & DbQuote(Num, HFApp.DivisionID)
            If FPickList.Choose(HFApp.Databases(dbHomefront), "Customer", s, txtCustomer.Text, , True, , "divisionid") Then
                mCustomer = FPickList.SelectedItem("Customer")
                
                If mCustomer = "" Then mCustomer = Chr(1)
                Call LoadData
            End If
            
         Case "Save"
            Call SaveData(False)
            
            
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not SaveData(True) Then
        Cancel = True
        Exit Sub
    End If
    Call IniPutForm(Me)
    Call IniPutGrid(Me, gContacts)
End Sub

Public Function SaveData(prompt As Boolean) As Boolean
On Error GoTo eh
    Dim s As String
    Dim r As Long
    
    If Not mDirty Then
        SaveData = True
        Exit Function
    End If
    If prompt Then
        Select Case MsgBox(Me.Caption & " has changed." & vbCrLf & vbCrLf & "Do you want to save these changes?" & vbCrLf, vbExclamation + vbYesNoCancel, App.ProductName)
            Case vbNo
                SaveData = True
                Exit Function
            Case vbCancel
                SaveData = False
                Exit Function
        End Select
    End If
    
    Screen.MousePointer = vbHourglass
    
    If mCustomer = "" Or mCustomer = Chr(1) Then
        s = ""
        s = s & "INSERT arcustomers(arcustomer,description,comments,billaddr1,billaddr2,billcity,billprovince,billpostalcode,phone1,fax1,divisionid)" & vbCrLf
        s = s & "values(" & DbQuote(Str, txtCustomer.Text) & vbCrLf
        s = s & "      ," & DbQuote(Str, txtDescription.Text) & vbCrLf
        s = s & "      ," & DbQuote(Str, txtComments.Text) & vbCrLf
        s = s & "      ," & DbQuote(Str, Trim(Parse(txtAddress.Text, 1, vbCrLf))) & vbCrLf
        s = s & "      ," & DbQuote(Str, Trim(Mid(txtAddress.Text, Len(Parse(txtAddress.Text, 1, vbCrLf)) + 3))) & vbCrLf
        s = s & "      ," & DbQuote(Str, txtCity.Text) & vbCrLf
        s = s & "      ," & DbQuote(Str, txtProvince.Text) & vbCrLf
        s = s & "      ," & DbQuote(Str, txtPostal.Text) & vbCrLf
        s = s & "      ," & DbQuote(Str, txtPhone.Text) & vbCrLf
        s = s & "      ," & DbQuote(Str, txtFax.Text) & vbCrLf
        s = s & "      ," & DbQuote(Num, HFApp.DivisionID) & ")"
        HFApp.SqlExec s
        mCustomer = txtCustomer.Text
        If HFApp.Options(AccountingSystem) = asSimply Then
            Call HFApp.WriteCustomerToAccounting("", mCustomer)
        End If
    Else
        s = ""
        s = s & "UPDATE arcustomers" & vbCrLf
        s = s & "SET description=" & DbQuote(Str, txtDescription.Text) & vbCrLf
        s = s & "   ,comments=" & DbQuote(Str, txtComments.Text) & vbCrLf
        s = s & "   ,BillAddr1=" & DbQuote(Str, Trim(Parse(txtAddress.Text, 1, vbCrLf))) & vbCrLf
        s = s & "   ,BillAddr2=" & DbQuote(Str, Trim(Mid(txtAddress.Text, Len(Parse(txtAddress.Text, 1, vbCrLf)) + 3))) & vbCrLf
        s = s & "   ,BillCity=" & DbQuote(Str, txtCity.Text) & vbCrLf
        s = s & "   ,BillProvince=" & DbQuote(Str, txtProvince.Text) & vbCrLf
        s = s & "   ,BillPostalCode=" & DbQuote(Str, txtPostal.Text) & vbCrLf
        s = s & "   ,Phone1=" & DbQuote(Str, txtPhone.Text) & vbCrLf
        s = s & "   ,Fax1=" & DbQuote(Str, txtFax.Text) & vbCrLf
        s = s & "WHERE ARCustomer=" & DbQuote(Str, mCustomer) & vbCrLf
        s = s & "and divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
        HFApp.SqlExec s
    End If
    
    Call SaveContacts
    
    
    SaveData = True
    mDirty = False
    Screen.MousePointer = vbDefault
    
Exit Function
eh: Call errHandler(SRCFILE & "SaveData", s)
End Function



Private Sub LoadData()
    Dim s As String
    Dim rs As Recordset
    
    
    s = "select * from arcustomers where divisionid=" & DbQuote(Num, HFApp.DivisionID) & " and arcustomer=" & DbQuote(Str, mCustomer)
    Set rs = HFApp.SqlExec(s)
    If rs.EOF Then
        ReadOnly = True
        txtCustomer.Text = ""
        txtDescription.Text = ""
        txtComments.Text = ""
        txtAddress.Text = ""
        txtCity.Text = ""
        txtProvince.Text = ""
        txtPostal.Text = ""
        txtPhone.Text = ""
        txtFax.Text = ""
        cboStatus.ListIndex = 0
        gContacts.Rows = 1
    Else
        ReadOnly = False
        txtCustomer.Text = "" & rs("arcustomer")
        txtDescription.Text = "" & rs("description")
        txtComments.Text = "" & rs("comments")
        txtAddress.Text = "" & rs("billAddr1") & vbCrLf & rs("billAddr2")
        txtCity.Text = "" & rs("billcity")
        txtProvince.Text = "" & rs("billprovince")
        txtPostal.Text = "" & rs("billpostalcode")
        txtPhone.Text = "" & rs("phone1")
        txtFax.Text = "" & rs("fax1")
        cboStatus.ListIndex = IIf("" & rs("Inactive") = "True", 1, 0)
        Call LoadContacts
        Call SetCtrlFocus(txtDescription)
    End If
    txtCustomer.Enabled = txtCustomer.Text = ""
    mDirty = False
    
End Sub

Private Sub txtComments_Change()
    mDirty = True
End Sub

Private Sub txtCustomer_Change()
    mDirty = True
End Sub
Private Sub txtDescription_Change()
    mDirty = True
End Sub
Private Sub txtAddress_Change()
    mDirty = True
End Sub
Private Sub txtCity_Change()
    mDirty = True
End Sub
Private Sub txtProvince_Change()
    mDirty = True
End Sub
Private Sub txtPostal_Change()
    mDirty = True
End Sub
Private Sub txtPhone_Change()
    mDirty = True
End Sub
Private Sub txtFax_Change()
    mDirty = True
End Sub


Private Sub txtCustomer_GotFocus()
    SelectAll txtCustomer
End Sub
Private Sub txtDescription_GotFocus()
    SelectAll txtDescription
End Sub
Private Sub txtAddress_GotFocus()
    SelectAll txtAddress
End Sub
Private Sub txtCity_GotFocus()
    SelectAll txtCity
End Sub
Private Sub txtProvince_GotFocus()
    SelectAll txtProvince
End Sub
Private Sub txtPostal_GotFocus()
    SelectAll txtPostal
End Sub
Private Sub txtPhone_GotFocus()
    SelectAll txtPhone
End Sub
Private Sub txtFax_GotFocus()
    SelectAll txtFax
End Sub


Private Property Let ReadOnly(RHS As Boolean)
'    txtCustomer.Enabled = False
'    txtDescription.Enabled = Not RHS
'    txtAddress.Enabled = Not RHS
'    txtCity.Enabled = Not RHS
'    txtProvince.Enabled = Not RHS
'    txtPostal.Enabled = Not RHS
'    txtPhone.Enabled = Not RHS
'    txtFax.Enabled = Not RHS
End Property

Public Sub ShowForm(Customer As String)
On Error Resume Next
    mCustomer = Customer
    Me.Show vbModal
End Sub








Private Sub cboStatus_Click()
    mDirty = True
End Sub

Private Sub gContacts_AfterSort(ByVal Col As Long, Order As Integer)
    gContacts.AddItem ""
End Sub

Private Sub gContacts_BeforeSort(ByVal Col As Long, Order As Integer)
    gContacts.RemoveItem gContacts.Rows - 1
End Sub

Private Sub gContacts_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete And Shift <> 0 And gContacts.Rows > 1 And gContacts.Row <> gContacts.Rows - 1 Then
        gContacts.RowHidden(gContacts.Row) = True
        mDirty = True
    End If
End Sub

Private Sub gContacts_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gContacts
        If .Row = .Rows - 1 Then
            .TextMatrix(.Row, .ColIndex("SendVia")) = 1
            .AddItem ""
        End If
    End With
End Sub

Private Sub gContacts_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gContacts
        Select Case .ColKey(Col)
            Case "Cell":    .EditText = FormatPhone(.EditText)
            Case "Phone":   .EditText = FormatPhone(.EditText)
            Case "fax":     .EditText = FormatPhone(.EditText)
        End Select
        .RowData(Row) = "DIRTY"
        mDirty = True
    End With
End Sub

Private Sub gContacts_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gContacts
    .EditMaxLength = 0
    Cancel = Row = .Rows - 1
    Select Case .ColKey(Col)
        Case "Phone":        .EditMaxLength = 35
        Case "Cell":         .EditMaxLength = 35
        Case "Fax":          .EditMaxLength = 35
        Case "Email":        .EditMaxLength = 500
        Case "Name":         .EditMaxLength = 50: Cancel = False
        Case "Role":         .EditMaxLength = 25
        Case "SmsAddress":   .EditMaxLength = 50
    End Select
    End With
End Sub

Private Sub LoadContacts()
    Dim s As String
    Dim r As Long
    Dim rs As Recordset
    
    s = ""
    s = s & "select *" & vbCrLf
    s = s & "  from contacts" & vbCrLf
    s = s & " where contacttypeid=99" & vbCrLf
    s = s & "   and arcustomer=" & DbQuote(Str, mCustomer)
    s = s & "order by firstname" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    With gContacts
        .Rows = 1
        While Not rs.EOF
            .AddItem ""
            r = .Rows - 1
            .TextMatrix(r, .ColIndex("ContactID")) = "" & rs("ContactID")
            .TextMatrix(r, .ColIndex("Role")) = "" & rs("Role")
            .TextMatrix(r, .ColIndex("Name")) = "" & rs("FirstName")
            .TextMatrix(r, .ColIndex("Phone")) = "" & rs("WorkPhone")
            .TextMatrix(r, .ColIndex("Cell")) = "" & rs("CellPhone")
            .TextMatrix(r, .ColIndex("Fax")) = "" & rs("Fax")
            .TextMatrix(r, .ColIndex("Email")) = "" & rs("Email")
            .TextMatrix(r, .ColIndex("SendVia")) = Val("" & rs("CommModeID"))
            .TextMatrix(r, .ColIndex("SmsAddress")) = "" & rs("SmsAddress")
            
            
            rs.MoveNext
        Wend
        .AddItem ""
    End With

End Sub

Private Sub SaveContacts()
    Dim s As String
    Dim r As Long
    
    With gContacts
        For r = .Rows - 2 To 1 Step -1
            If .RowHidden(r) Then
                If .TextMatrix(r, .ColIndex("ContactID")) <> "" Then
                    s = "DELETE FROM Contacts WHERE ContactID=" & DbQuote(Num, .TextMatrix(r, .ColIndex("ContactID")))
                    Call HFApp.SqlExec(s)
                End If
                Call .RemoveItem(r)
            ElseIf .RowData(r) = "DIRTY" Then
                If .TextMatrix(r, .ColIndex("ContactID")) = "" Then
                    s = ""
                    s = s & "INSERT INTO Contacts(ContactTypeID,UserID,ARCustomer,Role,FirstName,WorkPhone,CellPhone,Fax,Email,CommModeID,SmsAddress)" & vbCrLf
                    s = s & "VALUES(99," & DbQuote(Str, mCustomer) & "," & DbQuote(Str, mCustomer) & vbCrLf
                    s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Role"))) & vbCrLf
                    s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Name"))) & vbCrLf
                    s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Phone"))) & vbCrLf
                    s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Cell"))) & vbCrLf
                    s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Fax"))) & vbCrLf
                    s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Email"))) & vbCrLf
                    s = s & "      ," & DbQuote(Num, .TextMatrix(r, .ColIndex("SendVia"))) & vbCrLf
                    s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("SmsAddress"))) & vbCrLf
                    s = s & ")"
                    Call HFApp.SqlExec(s)
                    .TextMatrix(r, .ColIndex("ContactID")) = HFApp.SqlIdentity("Contacts")
                Else
                    s = ""
                    s = s & "UPDATE Contacts" & vbCrLf
                    s = s & "SET Role=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Role"))) & vbCrLf
                    s = s & "   ,FirstName=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Name"))) & vbCrLf
                    s = s & "   ,WorkPhone=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Phone"))) & vbCrLf
                    s = s & "   ,CellPhone=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Cell"))) & vbCrLf
                    s = s & "   ,Fax=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Fax"))) & vbCrLf
                    s = s & "   ,Email=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Email"))) & vbCrLf
                    s = s & "   ,CommModeID=" & DbQuote(Num, .TextMatrix(r, .ColIndex("SendVia"))) & vbCrLf
                    s = s & "   ,SmsAddress=" & DbQuote(Str, .TextMatrix(r, .ColIndex("SmsAddress"))) & vbCrLf
                    s = s & "WHERE ContactID=" & DbQuote(Num, .TextMatrix(r, .ColIndex("ContactID"))) & vbCrLf
                    Call HFApp.SqlExec(s)
                End If
                .RowData(r) = ""
            End If
        Next
    End With
End Sub

