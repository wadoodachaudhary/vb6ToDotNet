VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FRfis 
   Caption         =   "Request For Information"
   ClientHeight    =   10665
   ClientLeft      =   1200
   ClientTop       =   1035
   ClientWidth     =   15720
   Icon            =   "FRfis.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10665
   ScaleWidth      =   15720
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   15720
      _ExtentX        =   27728
      _ExtentY        =   1005
      ButtonWidth     =   1191
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open"
            Key             =   "Open"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            Key             =   "New"
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "saveasquote"
                  Text            =   "Save as New Quote"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "saveasjob"
                  Text            =   "Save as New Job"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "converttojob"
                  Text            =   "Convert Quote to Job"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            Key             =   "Delete"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "s2"
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Preview"
            Key             =   "Preview"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Send"
            Key             =   "Send"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "High"
            Key             =   "High"
            Style           =   1
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Low"
            Key             =   "Low"
            Style           =   1
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList LargeIcons 
         Left            =   13080
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   62
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":000C
               Key             =   "SaveAssembly"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":08E6
               Key             =   "Send"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":11C0
               Key             =   "High"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":1A9A
               Key             =   "Low"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":2374
               Key             =   "AssemblyCosts"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":2C4E
               Key             =   "MassChange"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":3528
               Key             =   "FieldPOs"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":3E02
               Key             =   "NewRFQ"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":46DC
               Key             =   "CreateJob"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":4FB6
               Key             =   "takeoffsettings"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":5890
               Key             =   "quote"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":616A
               Key             =   "DecreaseDecimals"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":6A44
               Key             =   "IncreaseDecimals"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":731E
               Key             =   "ExcelImport"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":7BF8
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":84D2
               Key             =   "UpdatePrices"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":8DAC
               Key             =   "ExcelExport"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":9686
               Key             =   "Publish"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":9F60
               Key             =   "Forecast"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":A83A
               Key             =   "ViewPOs"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":B114
               Key             =   "ViewBudgets"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":B9EE
               Key             =   "Open"
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":C2C8
               Key             =   "Preview"
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":CBA2
               Key             =   "SendPOs"
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":D47C
               Key             =   "SendRFQs"
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":DD56
               Key             =   "TakeoffOneTime"
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":E630
               Key             =   "Estimate"
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":EF0A
               Key             =   "TakeoffAssembly"
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":F7E4
               Key             =   "NewAssembly"
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":100BE
               Key             =   "TakeoffItemChart"
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":10998
               Key             =   "New"
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":11272
               Key             =   "TakeoffItem"
            EndProperty
            BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":11B4C
               Key             =   "TakeoffCustom"
            EndProperty
            BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":12426
               Key             =   "Save"
            EndProperty
            BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":12D00
               Key             =   "SaveAs"
            EndProperty
            BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":135DA
               Key             =   "Delete"
            EndProperty
            BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":13EB4
               Key             =   "AddPricelist"
            EndProperty
            BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":1478E
               Key             =   "RePrice"
            EndProperty
            BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":15068
               Key             =   "Pricebook"
            EndProperty
            BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":15942
               Key             =   "PricebookEdit"
            EndProperty
            BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":1621C
               Key             =   "PricelistExport"
            EndProperty
            BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":16AF6
               Key             =   "PricelistImport"
            EndProperty
            BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":173D0
               Key             =   "NewPricelist"
            EndProperty
            BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":17CAA
               Key             =   "View"
            EndProperty
            BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":18584
               Key             =   "Vendor1"
            EndProperty
            BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":18E5E
               Key             =   "Vendor"
            EndProperty
            BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":19738
               Key             =   "Add"
            EndProperty
            BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":1A012
               Key             =   "Attachments"
            EndProperty
            BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":1A8EC
               Key             =   ""
            EndProperty
            BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":1B1C6
               Key             =   ""
            EndProperty
            BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":1BAA0
               Key             =   "Design Center Options"
            EndProperty
            BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":1C37A
               Key             =   "Global Options"
            EndProperty
            BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":1CC54
               Key             =   "Models"
            EndProperty
            BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":1D52E
               Key             =   "Options"
            EndProperty
            BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":1DE08
               Key             =   "Generate"
            EndProperty
            BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":1E6E2
               Key             =   "Timberline"
            EndProperty
            BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":1EFBC
               Key             =   "MBImport"
            EndProperty
            BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":1F896
               Key             =   "EEEstimating"
            EndProperty
            BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":20170
               Key             =   "EEExport"
            EndProperty
            BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":20A4A
               Key             =   "EEImport"
            EndProperty
            BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":21324
               Key             =   "MBExport"
            EndProperty
            BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":21BFE
               Key             =   "MasterBuilder"
            EndProperty
         EndProperty
      End
      Begin VB.Timer Timer1 
         Left            =   14190
         Top             =   60
      End
   End
   Begin HFSystem.Slider Slider 
      Height          =   10470
      Left            =   2850
      Top             =   660
      Width           =   60
      _ExtentX        =   106
      _ExtentY        =   18468
      Max             =   13425
   End
   Begin VSFlex8Ctl.VSFlexGrid gRFIs 
      Height          =   9615
      Left            =   0
      TabIndex        =   15
      Top             =   600
      Width           =   2805
      _cx             =   1999115092
      _cy             =   1999127104
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
      AllowUserResizing=   3
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   20
      Cols            =   17
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FRfis.frx":228D8
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
      ExplorerBar     =   2
      PicturesOver    =   0   'False
      FillStyle       =   1
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   1
      OwnerDraw       =   0
      Editable        =   0
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
      Begin MSComctlLib.ImageList SmallIcons 
         Left            =   690
         Top             =   2610
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":22B2D
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":230C7
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":23661
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":23BFB
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":24195
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRfis.frx":2472F
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   10515
      Left            =   2910
      TabIndex        =   16
      Top             =   570
      Width           =   12345
      Begin VB.TextBox txtApprovedBy 
         BorderStyle     =   0  'None
         Height          =   230
         Left            =   2235
         MaxLength       =   50
         TabIndex        =   13
         Top             =   3945
         Width           =   2115
      End
      Begin VB.TextBox txtApprovedDate 
         BorderStyle     =   0  'None
         Height          =   230
         Left            =   5865
         MaxLength       =   50
         TabIndex        =   14
         Top             =   3945
         Width           =   1575
      End
      Begin VB.TextBox txtSchedulingImpact 
         BorderStyle     =   0  'None
         Height          =   230
         Left            =   5220
         MaxLength       =   50
         TabIndex        =   12
         Top             =   3690
         Width           =   6075
      End
      Begin VB.TextBox txtCostImpact 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   230
         Left            =   2235
         MaxLength       =   50
         TabIndex        =   11
         Text            =   "34.45"
         Top             =   3690
         Width           =   1125
      End
      Begin VB.TextBox txtDueDate 
         BorderStyle     =   0  'None
         Height          =   230
         Left            =   10200
         MaxLength       =   50
         TabIndex        =   5
         Top             =   435
         Width           =   1575
      End
      Begin VB.TextBox txtClosedDate 
         BorderStyle     =   0  'None
         Height          =   230
         Left            =   10200
         MaxLength       =   50
         TabIndex        =   6
         Top             =   675
         Width           =   1575
      End
      Begin VB.TextBox txtSentDate 
         BorderStyle     =   0  'None
         Height          =   230
         Left            =   10200
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   4
         Text            =   "June 14, 2011"
         Top             =   195
         Width           =   1575
      End
      Begin VB.TextBox txtAnswer 
         BorderStyle     =   0  'None
         Height          =   1050
         Left            =   1080
         MaxLength       =   8000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Text            =   "FRfis.frx":24CC9
         Top             =   2550
         Width           =   10905
      End
      Begin VB.TextBox txtQuestion 
         BorderStyle     =   0  'None
         Height          =   1050
         Left            =   1080
         MaxLength       =   8000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Text            =   "FRfis.frx":24CCB
         Top             =   1455
         Width           =   10935
      End
      Begin VB.TextBox txtRecipientsBCC 
         BorderStyle     =   0  'None
         Height          =   230
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   3
         Top             =   915
         Width           =   8265
      End
      Begin VB.TextBox txtRecipientsCC 
         BorderStyle     =   0  'None
         Height          =   230
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   2
         Top             =   675
         Width           =   8265
      End
      Begin VB.TextBox txtRecipients 
         BorderStyle     =   0  'None
         Height          =   230
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   1
         Top             =   435
         Width           =   8265
      End
      Begin VB.TextBox txtSender 
         BorderStyle     =   0  'None
         Height          =   230
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   0
         Top             =   195
         Width           =   8265
      End
      Begin VB.TextBox txtSubject 
         BorderStyle     =   0  'None
         Height          =   230
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1185
         Width           =   10935
      End
      Begin HFSystem.VBCombo cboStatus 
         Height          =   240
         Left            =   10200
         TabIndex        =   7
         Top             =   915
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   423
      End
      Begin VSFlex8Ctl.VSFlexGrid gMessages 
         Height          =   5355
         Left            =   1080
         TabIndex        =   28
         Top             =   5610
         Width           =   10905
         _cx             =   1999129379
         _cy             =   1999119590
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
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   3
         SelectionMode   =   0
         GridLines       =   4
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   20
         Cols            =   14
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FRfis.frx":24CCD
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
         ExplorerBar     =   3
         PicturesOver    =   0   'False
         FillStyle       =   1
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
      Begin VSFlex8Ctl.VSFlexGrid gAttachments 
         Height          =   1275
         Left            =   1080
         TabIndex        =   31
         Top             =   4260
         Width           =   10905
         _cx             =   1999129379
         _cy             =   1999112393
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
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   3
         GridLines       =   0
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   13
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FRfis.frx":24EAA
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   2
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
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
         Begin VB.PictureBox picCanvas 
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   30
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   35
            Top             =   330
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picFolder 
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   420
            Picture         =   "FRfis.frx":250DA
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   34
            Top             =   300
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picEmbedded 
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   720
            Picture         =   "FRfis.frx":25664
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   33
            Top             =   330
            Visible         =   0   'False
            Width           =   240
         End
      End
      Begin VB.Label Label112 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Approved By"
         Height          =   195
         Index           =   15
         Left            =   1230
         TabIndex        =   39
         Top             =   3960
         Width           =   915
      End
      Begin VB.Image cmdBrowse 
         Height          =   240
         Index           =   8
         Left            =   7455
         Picture         =   "FRfis.frx":25BEE
         Top             =   3945
         Width           =   240
      End
      Begin VB.Label Label112 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Approved Date"
         Height          =   195
         Index           =   14
         Left            =   4725
         TabIndex        =   38
         Top             =   3975
         Width           =   1080
      End
      Begin VB.Image cmdBrowse 
         Height          =   240
         Index           =   7
         Left            =   4365
         Picture         =   "FRfis.frx":25D38
         Top             =   3945
         Width           =   240
      End
      Begin VB.Image cmdBrowse 
         Height          =   240
         Index           =   6
         Left            =   9360
         Picture         =   "FRfis.frx":25E82
         Top             =   915
         Width           =   240
      End
      Begin VB.Image cmdBrowse 
         Height          =   240
         Index           =   5
         Left            =   9360
         Picture         =   "FRfis.frx":25FCC
         Top             =   675
         Width           =   240
      End
      Begin VB.Image cmdBrowse 
         Height          =   240
         Index           =   4
         Left            =   9360
         Picture         =   "FRfis.frx":26116
         Top             =   435
         Width           =   240
      End
      Begin VB.Image cmdBrowse 
         Height          =   240
         Index           =   3
         Left            =   9360
         Picture         =   "FRfis.frx":26260
         Top             =   195
         Width           =   240
      End
      Begin VB.Label Label112 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Scheduling Impact"
         Height          =   195
         Index           =   13
         Left            =   3720
         TabIndex        =   37
         Top             =   3720
         Width           =   1320
      End
      Begin VB.Label Label112 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cost Impact"
         Height          =   195
         Index           =   12
         Left            =   1305
         TabIndex        =   36
         Top             =   3720
         Width           =   840
      End
      Begin VB.Label Label112 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Due"
         Height          =   195
         Index           =   7
         Left            =   9840
         TabIndex        =   32
         Top             =   465
         Width           =   300
      End
      Begin VB.Image cmdBrowse 
         Height          =   240
         Index           =   2
         Left            =   11790
         Picture         =   "FRfis.frx":263AA
         Top             =   420
         Width           =   240
      End
      Begin VB.Label Label112 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Conversation"
         Height          =   195
         Index           =   11
         Left            =   60
         TabIndex        =   29
         Top             =   5700
         Width           =   930
      End
      Begin VB.Label Label112 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Attachments"
         Height          =   195
         Index           =   10
         Left            =   105
         TabIndex        =   27
         Top             =   4320
         Width           =   885
      End
      Begin VB.Image cmdBrowse 
         Height          =   240
         Index           =   1
         Left            =   11790
         Picture         =   "FRfis.frx":264F4
         Top             =   660
         Width           =   240
      End
      Begin VB.Image cmdBrowse 
         Height          =   240
         Index           =   0
         Left            =   11790
         Picture         =   "FRfis.frx":2663E
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         Height          =   195
         Left            =   9690
         TabIndex        =   26
         Top             =   945
         Width           =   450
      End
      Begin VB.Label Label112 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Closed"
         Height          =   195
         Index           =   9
         Left            =   9660
         TabIndex        =   25
         Top             =   690
         Width           =   480
      End
      Begin VB.Label Label112 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sent"
         Height          =   195
         Index           =   8
         Left            =   9810
         TabIndex        =   24
         Top             =   225
         Width           =   330
      End
      Begin VB.Label Label112 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Resolution or Answer"
         Height          =   645
         Index           =   6
         Left            =   120
         TabIndex        =   23
         Top             =   2655
         Width           =   870
      End
      Begin VB.Label Label112 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Issue or Question"
         Height          =   555
         Index           =   5
         Left            =   180
         TabIndex        =   22
         Top             =   1545
         Width           =   810
      End
      Begin VB.Label Label112 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BCC"
         Height          =   195
         Index           =   4
         Left            =   675
         TabIndex        =   21
         Top             =   945
         Width           =   315
      End
      Begin VB.Label Label112 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CC"
         Height          =   195
         Index           =   3
         Left            =   780
         TabIndex        =   20
         Top             =   705
         Width           =   210
      End
      Begin VB.Label Label112 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   195
         Index           =   2
         Left            =   795
         TabIndex        =   19
         Top             =   453
         Width           =   195
      End
      Begin VB.Label Label112 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         Height          =   195
         Index           =   1
         Left            =   645
         TabIndex        =   18
         Top             =   225
         Width           =   345
      End
      Begin VB.Label Label112 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subject"
         Height          =   195
         Index           =   0
         Left            =   450
         TabIndex        =   17
         Top             =   1215
         Width           =   540
      End
   End
End
Attribute VB_Name = "FRfis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SRCFILE = "FRfis::"

Private mDirty As Boolean

Private mJob As String
Private mRFIRow As Long
Private mLoading As Boolean 'see LoadJob() and LoadRFI()


Public Sub ShowForm(Job As String, rfiid As Long)
    mJob = Job
    Call Me.Show(vbModal)
End Sub

Private Sub cmdBrowse_Click(Index As Integer)
    Dim s As String
    Select Case Index
        Case 0:    If Not txtSentDate.Locked Then Call DCalendar.Popup(txtSentDate)
        Case 2:    If Not txtDueDate.Locked Then Call DCalendar.Popup(txtDueDate)
        Case 1:    If Not txtClosedDate.Locked Then Call DCalendar.Popup(txtClosedDate)
        Case 8:    If Not txtApprovedDate.Locked Then Call DCalendar.Popup(txtApprovedDate)
        
        Case 3:    s = PickContacts(): If s <> "" Then txtSender.SelText = s
        Case 4:    s = PickContacts(): If s <> "" Then txtRecipients.SelText = s
        Case 5:    s = PickContacts(): If s <> "" Then txtRecipientsCC.SelText = s
        Case 6:    s = PickContacts(): If s <> "" Then txtRecipientsBCC.SelText = s
        Case 7:    s = PickContacts(): If s <> "" Then txtApprovedBy.SelText = s
    
    End Select
    Dirty = True
End Sub

Private Function PickContacts() As String
    Dim i As Long
    Dim s As String
    
    If FPickList.Choose(HFApp.Databases(dbHomefront), "Contact", "select * from addressbook", , , , , , True) Then
        s = ""
        For i = 1 To FPickList.SelectedItems
            s = s & "; " & FPickList.SelectedItem("name", i)
        Next
        PickContacts = Mid(s, 3)
    End If

End Function

Private Sub Form_Load()
    Call SetToolbarIcons(Toolbar, LargeIcons)
    Call IniGetForm(Me)
    Call IniGetGrid(Me, gRFIs)
    Call IniGetGrid(Me, gAttachments)
    Call IniGetGrid(Me, gMessages)
    
    cboStatus.AddItem "Open"
    cboStatus.AddItem "Awaiting Response"
    cboStatus.AddItem "Requires Attention"
    cboStatus.AddItem "Approved"
    cboStatus.AddItem "Closed"
    
    gRFIs.ColImageList(gRFIs.ColIndex("Priority")) = Me.SmallIcons.hImageList
    gMessages.ColImageList(gMessages.ColIndex("Priority")) = Me.SmallIcons.hImageList
    gMessages.ColImageList(gMessages.ColIndex("Original")) = Me.SmallIcons.hImageList
    gMessages.Cell(flexcpPicture, 0, gMessages.ColIndex("priority")) = SmallIcons.ListImages(6).Picture
    
    
    Call LoadJob(mJob)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not SaveData(True) Then
        Cancel = True
        Exit Sub
    End If
    Call IniPutForm(Me)
    Call IniPutGrid(Me, gRFIs)
    Call IniPutGrid(Me, gAttachments)
    Call IniPutGrid(Me, gMessages)
End Sub

Private Sub gMessages_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gMessages
        If .TextMatrix(Row, .ColIndex("MessageType")) = "Email" Then
            Cancel = True
            Exit Sub
        End If
        
        .ComboList = ""
        Select Case .ColKey(Col)
            Case "Attachments", "Size", "ID", "Original", "Priority"
                Cancel = True
                
            Case "MessageType"
                .ComboList = "Mail|Fax|Phone|Other"
                
            Case "SentDate", "ReceivedDate", "Sender", "Recipients", "RecipientsCC", "RecipientsBCC"
                .ComboList = "|..."
        End Select
        
    End With
End Sub

Private Sub gMessages_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim s As String
    With gMessages
        Select Case .ColKey(Col)
            Case "SentDate", "ReceivedDate"
                Call DCalendar.Popup(gMessages, .RowPos(.Row) + .RowHeight(.Row), .ColPos(.Col))
                .RowData(Row) = "DIRTY"
                Dirty = True
            
            Case "Sender", "Recipients", "RecipientsCC", "RecipientsBCC"
                s = PickContacts()
                If s <> "" Then
                    .Text = s
                    .RowData(Row) = "DIRTY"
                    Dirty = True
                End If
            
        End Select
    End With
            
End Sub

Private Sub gMessages_Click()
    With gMessages
        If .ColKey(.Col) = "Priority" And .TextMatrix(.Row, .ColIndex("MessageType")) <> "Email" Then
            Select Case .ValueMatrix(.Row, .Col)
                Case 0:  .TextMatrix(.Row, .Col) = "1"
                Case 1:  .TextMatrix(.Row, .Col) = "2"
                Case 2:  .TextMatrix(.Row, .Col) = "0"
            End Select
            .RowData(.Row) = "DIRTY"
            Dirty = True
        End If
    End With
End Sub

Private Sub gMessages_DblClick()
    Dim s As String
    Dim ID As String
    
    With gMessages
        If .ColKey(.MouseCol) = "Original" And .TextMatrix(.Row, .ColIndex("MessageType")) = "Email" Then
            ID = .TextMatrix(.Row, .ColIndex("id"))
            
            On Error Resume Next
            s = TempFile("msg")
            Call DBGetFile(s, , HFApp.Databases(dbHomefront), "messages WHERE id=" & DbQuote(Str, ID), "Original")
        
            If Not FileExists(s) Then s = "c:\demo\files\" & ID & ".msg"
            Call ShellFile(Me.hwnd, s)
            
        End If
    End With
End Sub

Private Sub Slider_Move()
    Call Form_Resize
End Sub

Private Sub Form_Resize()
On Error Resume Next
Const margin = 120

    Slider.Min = 960
    Slider.Max = Me.ScaleWidth - 960
    Slider.Move Slider.left, Toolbar.Height, Slider.Width, Me.ScaleHeight - Toolbar.Height
    
    gRFIs.Move -15, Slider.Top - 15, Slider.left + 15, Slider.Height + 15
    Frame1.Move Slider.left + Slider.Width, Slider.Top, Me.ScaleWidth - Slider.left - Slider.Width, Slider.Height
        txtSubject.Width = Frame1.Width - txtSubject.left - margin
        txtQuestion.Width = txtSubject.Width
        txtAnswer.Width = txtSubject.Width
        
        txtSchedulingImpact.Width = Frame1.Width - txtSchedulingImpact.left - margin
        
        gAttachments.Width = txtSubject.Width
        gMessages.Width = txtSubject.Width
        gMessages.Height = Frame1.Height - gMessages.Top - margin

End Sub


Public Function SaveData(prompt As Boolean) As Boolean
On Error GoTo eh
    
    Dim i As Long
    Dim s As String
    Dim rfiid As Long
    
    Call UnLoadRFI
    
    If Not mDirty Then
        SaveData = True
        Exit Function
    End If
    If prompt Then
        Select Case MsgBox("This data has changed." & vbCrLf & vbCrLf & "Do you want to save these changes?" & vbCrLf, vbExclamation + vbYesNoCancel, App.ProductName)
            Case vbNo
                SaveData = True
                Exit Function
            Case vbCancel
                SaveData = False
                Exit Function
        End Select
    End If
    
    
    Screen.MousePointer = vbHourglass
    With gRFIs
        i = mRFIRow
        If .TextMatrix(i, .ColIndex("RFIID")) = "" Then
            
            s = "select isnull(LastRFIID,0)+1 from tbljobs where job_no=" & DbQuote(Str, mJob)
            .TextMatrix(i, .ColIndex("RFIID")) = HFApp.SqlExec(s)(0)
            
            s = "update tbljobs set lastrfiid=isnull(LastRFIID,0)+1 where job_no=" & DbQuote(Str, mJob)
            Call HFApp.SqlExec(s)
            
            s = "insert into rfis(job,rfiid) values(" & DbQuote(Str, mJob) & "," & DbQuote(Num, .TextMatrix(i, .ColIndex("RFIID"))) & ")"
            Call HFApp.SqlExec(s)
            
        End If
        
        rfiid = .TextMatrix(i, .ColIndex("RFIID"))
        
        s = ""
        s = s & "update rfis" & vbCrLf
        s = s & "set priority=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Priority"))) & vbCrLf
        s = s & "   ,subject=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Subject"))) & vbCrLf
        s = s & "   ,status=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Status"))) & vbCrLf
        s = s & "   ,Sender=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Sender"))) & vbCrLf
        s = s & "   ,Recipients=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Recipients"))) & vbCrLf
        s = s & "   ,RecipientsCC=" & DbQuote(Str, .TextMatrix(i, .ColIndex("RecipientsCC"))) & vbCrLf
        s = s & "   ,RecipientsBCC=" & DbQuote(Str, .TextMatrix(i, .ColIndex("RecipientsBCC"))) & vbCrLf
        s = s & "   ,Question=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Question"))) & vbCrLf
        s = s & "   ,Answer=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Answer"))) & vbCrLf
        s = s & "   ,SentDate=" & DbQuote(Date, .TextMatrix(i, .ColIndex("SentDate"))) & vbCrLf
        s = s & "   ,DueDate=" & DbQuote(Date, .TextMatrix(i, .ColIndex("DueDate"))) & vbCrLf
        s = s & "   ,ClosedDate=" & DbQuote(Date, .TextMatrix(i, .ColIndex("ClosedDate"))) & vbCrLf
        s = s & "   ,CostImpact=" & DbQuote(Num, .TextMatrix(i, .ColIndex("CostImpact"))) & vbCrLf
        s = s & "   ,SchedulingImpact=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Answer"))) & vbCrLf
        
        s = s & "   ,ApprovedBy=" & DbQuote(Str, .TextMatrix(i, .ColIndex("ApprovedBy"))) & vbCrLf
        s = s & "   ,ApprovedDate=" & DbQuote(Date, .TextMatrix(i, .ColIndex("ApprovedDate"))) & vbCrLf
        
        s = s & "where job=" & DbQuote(Str, mJob) & vbCrLf
        s = s & "  and rfiid=" & DbQuote(Num, rfiid)
        Call HFApp.SqlExec(s, dbHomefront)
        
    End With
    
    
    With gMessages
        For i = .Rows - 2 To 1 Step -1
            If .RowHidden(i) Then
                s = "delete from messages where id=" & DbQuote(Str, .TextMatrix(i, .ColIndex("id")))
                Call HFApp.SqlExec(s)
                Call .RemoveItem(i)
            End If
        Next
        
        
        For i = 1 To .Rows - 2
            If .RowData(i) = "DIRTY" Then
                If .TextMatrix(i, .ColIndex("id")) = "" Then
                    .TextMatrix(i, .ColIndex("id")) = CreateGUID()
                    s = ""
                    s = s & "insert into messages(job,rfiid,id,priority,messagetype,sender,recipients,recipientscc,recipientsbcc,subject,body,sentdate,receiveddate)" & vbCrLf
                    s = s & "values(" & DbQuote(Str, mJob) & vbCrLf
                    s = s & "      ," & DbQuote(Num, rfiid) & vbCrLf
                    s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("ID"))) & vbCrLf
                    s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("Priority"))) & vbCrLf
                    s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("MessageType"))) & vbCrLf
                    s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("Sender"))) & vbCrLf
                    s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("Recipients"))) & vbCrLf
                    s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("RecipientsCC"))) & vbCrLf
                    s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("RecipientsBCC"))) & vbCrLf
                    s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("Subject"))) & vbCrLf
                    s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("Body"))) & vbCrLf
                    s = s & "      ," & DbQuote(Date, .TextMatrix(i, .ColIndex("SentDate"))) & vbCrLf
                    s = s & "      ," & DbQuote(Date, .TextMatrix(i, .ColIndex("ReceivedDate"))) & vbCrLf
                    s = s & ")"
                Else
                    s = ""
                    s = s & "update messages" & vbCrLf
                    s = s & "set priority=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Priority"))) & vbCrLf
                    s = s & "   ,messagetype=" & DbQuote(Str, .TextMatrix(i, .ColIndex("MessageType"))) & vbCrLf
                    s = s & "   ,sender=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Sender"))) & vbCrLf
                    s = s & "   ,recipients=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Recipients"))) & vbCrLf
                    s = s & "   ,recipientscc=" & DbQuote(Str, .TextMatrix(i, .ColIndex("RecipientsCC"))) & vbCrLf
                    s = s & "   ,recipientsbcc=" & DbQuote(Str, .TextMatrix(i, .ColIndex("RecipientsBCC"))) & vbCrLf
                    s = s & "   ,subject=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Subject"))) & vbCrLf
                    s = s & "   ,body=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Body"))) & vbCrLf
                    s = s & "   ,sentdate=" & DbQuote(Date, .TextMatrix(i, .ColIndex("SentDate"))) & vbCrLf
                    s = s & "   ,receiveddate=" & DbQuote(Date, .TextMatrix(i, .ColIndex("ReceivedDate"))) & vbCrLf
                    s = s & "where id=" & DbQuote(Str, .TextMatrix(i, .ColIndex("id")))
                End If
                Call HFApp.SqlExec(s)
                .RowData(i) = ""
            End If
        Next
        
    End With
    
    
    mDirty = False
    Screen.MousePointer = vbDefault

Exit Function
eh: Call errHandler(SRCFILE & "SaveData", s)
End Function
Private Sub LoadJob(Job As String)
On Error GoTo eh
    
    Dim i  As Long
    Dim s  As String
    Dim rs As Recordset
    
    
    
    If Job = "" Then
        s = ""
        s = s & "SELECT DISTINCT j.Job_No Job" & vbCrLf
        s = s & "      ,j.Description" & vbCrLf
        s = s & "      ,j.Municipal_Address Address" & vbCrLf
        s = s & "      ,j.PM,Purchaser" & vbCrLf
        s = s & "      ,j.Estimator" & vbCrLf
        s = s & "FROM tblJobs j" & vbCrLf
        s = s & "WHERE ISNULL(j.Inactive,0)=0 " & vbCrLf
        If FPickList.Choose(HFApp.Databases(dbHomefront), "Jobs", s, mJob) Then
            mJob = FPickList.SelectedItem("Job")
        Else
            Exit Sub
        End If
    Else
        mJob = Job
    End If
    
    
    s = "select * from rfis where job=" & DbQuote(Str, mJob)
    Set rs = HFApp.SqlExec(s, dbHomefront)
    With gRFIs
        mLoading = True
        .Rows = 1
        i = 0
        While Not rs.EOF
            .AddItem ""
            i = i + 1
            .TextMatrix(i, .ColIndex("Priority")) = "" & rs("Priority")
            .TextMatrix(i, .ColIndex("RFIID")) = "" & rs("RFIID")
            .TextMatrix(i, .ColIndex("Subject")) = "" & rs("Subject")
            .TextMatrix(i, .ColIndex("Status")) = "" & rs("Status")
            .TextMatrix(i, .ColIndex("Sender")) = "" & rs("Sender")
            .TextMatrix(i, .ColIndex("Recipients")) = "" & rs("Recipients")
            .TextMatrix(i, .ColIndex("RecipientsCC")) = "" & rs("RecipientsCC")
            .TextMatrix(i, .ColIndex("RecipientsBCC")) = "" & rs("RecipientsBCC")
            .TextMatrix(i, .ColIndex("Question")) = "" & rs("Question")
            .TextMatrix(i, .ColIndex("Answer")) = "" & rs("Answer")
            
            
            .TextMatrix(i, .ColIndex("CostImpact")) = IIf(Val("" & rs("CostImpact")) = 0, "", Format(Val("" & rs("CostImpact")), "#,##0.00"))
            .TextMatrix(i, .ColIndex("SchedulingImpact")) = "" & rs("SchedulingImpact")
            
            .TextMatrix(i, .ColIndex("ApprovedBy")) = "" & rs("ApprovedBy")
            .TextMatrix(i, .ColIndex("ApprovedDate")) = IIf("" & rs("ApprovedDate") = "", "", Format("" & rs("ApprovedDate"), "medium date"))
            
            .TextMatrix(i, .ColIndex("SentDate")) = IIf("" & rs("SentDate") = "", "", Format("" & rs("SentDate"), "medium date"))
            .TextMatrix(i, .ColIndex("DueDate")) = IIf("" & rs("DueDate") = "", "", Format("" & rs("DueDate"), "medium date"))
            .TextMatrix(i, .ColIndex("ClosedDate")) = IIf("" & rs("ClosedDate") = "", "", Format("" & rs("ClosedDate"), "medium date"))
            
            rs.MoveNext
        Wend
        
        mLoading = False
        mRFIRow = 0
        If .Rows > 1 Then .Row = 1
        
    End With
    
    Dirty = False
Exit Sub
eh: Call errHandler(SRCFILE & "LoadJob")
End Sub


Private Sub gRFIs_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal newrow As Long, ByVal NewCol As Long)
    If SaveData(True) Then Call LoadRFI(newrow)
End Sub



Private Sub LoadRFI(RFIRow As Long)
    Dim i As Long
    Dim s As String
    Dim rs As Recordset
    Dim FInfo As New ClsFileInfo
    
    If mLoading Then Exit Sub
    
    mLoading = True
    
    'load rfi fields
    With gRFIs
        mRFIRow = RFIRow
        i = RFIRow
        Toolbar.Buttons("High").Value = IIf(.ValueMatrix(i, .ColIndex("Priority")) = 2, tbrPressed, tbrUnpressed)
        Toolbar.Buttons("Low").Value = IIf(.ValueMatrix(i, .ColIndex("Priority")) = 0, tbrPressed, tbrUnpressed)
        txtSubject = .TextMatrix(i, .ColIndex("Subject"))
        cboStatus.Text = .TextMatrix(i, .ColIndex("Status"))
        txtSender = .TextMatrix(i, .ColIndex("Sender"))
        txtRecipients = .TextMatrix(i, .ColIndex("Recipients"))
        txtRecipientsCC = .TextMatrix(i, .ColIndex("RecipientsCC"))
        txtRecipientsBCC = .TextMatrix(i, .ColIndex("RecipientsBCC"))
        txtQuestion = .TextMatrix(i, .ColIndex("Question"))
        txtAnswer = .TextMatrix(i, .ColIndex("Answer"))
        txtSentDate = .TextMatrix(i, .ColIndex("SentDate"))
        txtDueDate = .TextMatrix(i, .ColIndex("DueDate"))
        txtClosedDate = .TextMatrix(i, .ColIndex("ClosedDate"))
        txtApprovedBy = .TextMatrix(i, .ColIndex("ApprovedBy"))
        txtApprovedDate = .TextMatrix(i, .ColIndex("ApprovedDate"))
        txtCostImpact = .TextMatrix(i, .ColIndex("CostImpact"))
        txtSchedulingImpact = .TextMatrix(i, .ColIndex("SchedulingImpact"))
        
    End With
    
    
    'load messages
    s = ""
    s = s & "select *" & vbCrLf
    s = s & "  from messages" & vbCrLf
    s = s & " where job=" & DbQuote(Str, mJob) & vbCrLf
    s = s & "   and rfiid=" & DbQuote(Str, gRFIs.TextMatrix(mRFIRow, gRFIs.ColIndex("rfiid")))
    Set rs = HFApp.SqlExec(s, dbHomefront)
    With gMessages
        .Rows = 1
        i = 0
        While Not rs.EOF
            .AddItem ""
            i = i + 1
            .TextMatrix(i, .ColIndex("ID")) = "" & rs("ID")
            .TextMatrix(i, .ColIndex("Original")) = IIf("" & rs("MessageType") = "Email", "4", "")
            .TextMatrix(i, .ColIndex("Priority")) = "" & rs("Priority")
            .TextMatrix(i, .ColIndex("Attachments")) = IIf(Val("" & rs("Attachments")) = 0, "", "3")
            .TextMatrix(i, .ColIndex("MessageType")) = "" & rs("MessageType")
            .TextMatrix(i, .ColIndex("Sender")) = "" & rs("Sender")
            .TextMatrix(i, .ColIndex("Recipients")) = "" & rs("Recipients")
            .TextMatrix(i, .ColIndex("RecipientsCC")) = "" & rs("RecipientsCC")
            .TextMatrix(i, .ColIndex("RecipientsBCC")) = "" & rs("RecipientsBCC")
            .TextMatrix(i, .ColIndex("Subject")) = "" & rs("Subject")
            .TextMatrix(i, .ColIndex("Body")) = "" & rs("Body")
            .TextMatrix(i, .ColIndex("Size")) = IIf(Val("" & rs("Size")) = 0, "", "" & rs("Size") & " KB")
            .TextMatrix(i, .ColIndex("SentDate")) = "" & rs("SentDate")
            .TextMatrix(i, .ColIndex("ReceivedDate")) = "" & rs("ReceivedDate")
            rs.MoveNext
        Wend
        .AddItem ""
    End With
    
    
    'load attachments
    s = ""
    s = s & "select a.filename,a.documentclass,a.filesize,a.filetitle,a.path,a.modifieddate,a.createddate,a.attacheddate,a.embedded,a.webcustomers,a.webvendors,a.weblocation,a.Revision" & vbCrLf
    s = s & "  from attachments a " & vbCrLf
    s = s & " where a.objectid=" & DbQuote(Str, ObjectID) & vbCrLf
    s = s & "order by 1"
    Set rs = HFApp.SqlExec(s, dbHomefront)
    With gAttachments
        .Rows = 1
        i = 0
        While Not rs.EOF
            If "" & rs("FileName") <> "" Then
                .AddItem ""
                i = i + 1
                If "" & rs("Embedded") = "True" Then
                    s = PathAppend(AppWorkingFolder, "tmp." & FileExt("" & rs("FileName")))
                    Call CreateFile(s)
                    FInfo.FullPathName = s
                    .TextMatrix(i, .ColIndex("FileName")) = "" & rs("FileName")
                    .TextMatrix(i, .ColIndex("FileSize")) = FInfo.FormatFileSize(Val("" & rs("FileSize")))
                    .TextMatrix(i, .ColIndex("ModifiedDate")) = Format("" & rs("ModifiedDate"), "general date")
                    .TextMatrix(i, .ColIndex("CreatedDate")) = Format("" & rs("CreatedDate"), "general date")
                    .Cell(flexcpPicture, .Rows - 1, .ColIndex("Path")) = picEmbedded.Picture
                    .Cell(flexcpForeColor, .Rows - 1, .ColIndex("FileTitle")) = &HC00000 '&H00FF0000&  '&H800000
                Else
                    FInfo.FullPathName = "" & rs("FileName")
                    .TextMatrix(i, .ColIndex("FileName")) = "" & rs("FileName")
                    .TextMatrix(i, .ColIndex("FileSize")) = FInfo.FormatFileSize(FInfo.FileSize)
                    .TextMatrix(i, .ColIndex("ModifiedDate")) = Format(FInfo.ModifyTime, "general date")
                    .TextMatrix(i, .ColIndex("CreatedDate")) = Format(FInfo.CreationTime, "general date")
                End If
                
                
                .TextMatrix(i, .ColIndex("AttachedDate")) = Format("" & rs("AttachedDate"), "general date")
                .Cell(flexcpPicture, .Rows - 1, .ColIndex("Filetitle")) = FileIcon(FInfo.hSmlIList, FInfo.hSmlIcon)
                .TextMatrix(i, .ColIndex("FileTitle")) = "" & rs("FileTitle")
                .TextMatrix(i, .ColIndex("Path")) = "" & rs("Path")
                .TextMatrix(i, .ColIndex("Revision")) = "" & rs("Revision")
                .TextMatrix(i, .ColIndex("DocumentClass")) = "" & rs("DocumentClass")
                .TextMatrix(i, .ColIndex("Embedded")) = "" & rs("Embedded")
                
                Select Case True
                    Case "" & rs("WebCustomers") = "True" And "" & rs("WebVendors") = "True":    .TextMatrix(i, .ColIndex("WebPortals")) = "All"
                    Case "" & rs("WebCustomers") = "True":                                       .TextMatrix(i, .ColIndex("WebPortals")) = "Customer"
                    Case "" & rs("WebVendors") = "True":                                         .TextMatrix(i, .ColIndex("WebPortals")) = "Vendor"
                    Case Else:                                                                   .TextMatrix(i, .ColIndex("WebPortals")) = "None"
                End Select
                .TextMatrix(i, .ColIndex("WebLocation")) = "" & rs("WebLocation")
                If .TextMatrix(i, .ColIndex("WebLocation")) = "" Then .TextMatrix(i, .ColIndex("WebLocation")) = "Documents"
            End If
            rs.MoveNext
        Wend
    End With
        
    
    
    mLoading = False
    mDirty = False
    
End Sub

Private Sub UnLoadRFI()
    
    With gRFIs
    
        If mRFIRow < 1 Then Exit Sub
        If mRFIRow > .Rows - 1 Then Exit Sub
    
        .TextMatrix(mRFIRow, .ColIndex("Subject")) = txtSubject
        .TextMatrix(mRFIRow, .ColIndex("Status")) = cboStatus.Text
        .TextMatrix(mRFIRow, .ColIndex("Sender")) = txtSender
        .TextMatrix(mRFIRow, .ColIndex("Recipients")) = txtRecipients
        .TextMatrix(mRFIRow, .ColIndex("RecipientsCC")) = txtRecipientsCC
        .TextMatrix(mRFIRow, .ColIndex("RecipientsBCC")) = txtRecipientsBCC
        .TextMatrix(mRFIRow, .ColIndex("Question")) = txtQuestion
        .TextMatrix(mRFIRow, .ColIndex("Answer")) = txtAnswer
        .TextMatrix(mRFIRow, .ColIndex("SentDate")) = txtSentDate
        .TextMatrix(mRFIRow, .ColIndex("DueDate")) = txtDueDate
        .TextMatrix(mRFIRow, .ColIndex("ClosedDate")) = txtClosedDate
        .TextMatrix(mRFIRow, .ColIndex("CostImpact")) = txtCostImpact
        .TextMatrix(mRFIRow, .ColIndex("SchedulingImpact")) = txtSchedulingImpact
        .TextMatrix(mRFIRow, .ColIndex("ApprovedBy")) = txtApprovedBy
        .TextMatrix(mRFIRow, .ColIndex("ApprovedDate")) = txtApprovedDate
        
    End With

End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim i As Long
'    Dim X As HFPrinter.ReportViewer
    Dim s As String
    Dim rfi As Long

    With gRFIs
    Select Case Button.Key
        Case "Open":    If Not SaveData(True) Then Exit Sub
                        Call LoadJob("")
        Case "Save":    Call SaveData(False)
        Case "New":     Call NewRFI
        Case "Delete"
        Case "Send"
        
        Case "Preview"
            rfi = gRFIs.ValueMatrix(mRFIRow, gRFIs.ColIndex("RFIID"))
            If rfi = 0 Then Exit Sub
            s = PathAppend(HFApp.SystemFolder, "System\Reports\Estimating\RFI.rpt")
            If Not FileExists(s) Then
                MsgBox "file not found" & vbCrLf & _
                       "create a report at" & vbCrLf & _
                       s & vbCrLf & vbCrLf & _
                       "Add a text parameter named 'Job' and a numeric parameter named 'RFI'.", vbInformation, App.ProductName
            Else
                'Set X = New HFPrinter.ReportViewer
                'Call X.ShowReport(HFApp.ConnectionString(dbHomefront), s, rvPreview, "", "", "Job", mJob, "RFI", rfi)
                Dim c As New ZybUtil.Crystal
                Call c.LoadODBCReport(s, HFApp.LoginDSN, HFApp.LoginDBUID, HFApp.LoginDBPWD)
                On Error Resume Next
                Call c.ParameterValue("DivisionID", HFApp.DivisionID)
                Call c.ParameterValue("Job", mJob)
                Call c.ParameterValue("RFI", rfi)
                On Error GoTo 0
                Call c.PrintPreview("Print Preview")
            End If
        
        Case "Low"
            If mRFIRow < 1 Or mRFIRow > .Rows - 1 Then
                Button.Value = tbrUnpressed
                Exit Sub
            End If
            If Button.Value = tbrPressed Then Toolbar.Buttons("High").Value = tbrUnpressed
            .TextMatrix(mRFIRow, .ColIndex("Priority")) = IIf(Button.Value = tbrPressed, 0, 1)
            Dirty = True
            
        Case "High"
            If mRFIRow < 1 Or mRFIRow > .Rows - 1 Then
                Button.Value = tbrUnpressed
                Exit Sub
            End If
            If Button.Value = tbrPressed Then Toolbar.Buttons("Low").Value = tbrUnpressed
            .TextMatrix(mRFIRow, .ColIndex("Priority")) = IIf(Button.Value = tbrPressed, 2, 1)
            Dirty = True
            
    End Select
    End With
    
End Sub


Private Sub NewRFI()
    Dim i As Long
    
    With gRFIs
        .AddItem ""
        i = .Rows - 1
        .TextMatrix(i, .ColIndex("Priority")) = "1"
        .TextMatrix(i, .ColIndex("Subject")) = "untitled"
        .TextMatrix(i, .ColIndex("Status")) = "Open"
        .TextMatrix(i, .ColIndex("Sender")) = ""
        .TextMatrix(i, .ColIndex("Recipients")) = ""
        .TextMatrix(i, .ColIndex("RecipientsCC")) = ""
        .TextMatrix(i, .ColIndex("RecipientsBCC")) = ""
        .TextMatrix(i, .ColIndex("Question")) = ""
        .TextMatrix(i, .ColIndex("Answer")) = ""
        .TextMatrix(i, .ColIndex("SentDate")) = ""
        .TextMatrix(i, .ColIndex("DueDate")) = ""
        .TextMatrix(i, .ColIndex("ClosedDate")) = ""
    
        .Row = i
    End With
    
End Sub



Private Sub txtAnswer_Change()
    Dirty = True
End Sub

Private Sub txtApprovedBy_Change()
    Dirty = True
End Sub

Private Sub txtApprovedBy_GotFocus()
    SelectAll txtApprovedBy
End Sub

Private Sub txtApprovedBy_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 And Shift = 0 Then Call cmdBrowse_Click(7)
End Sub

Private Sub txtApprovedDate_Change()
    Dirty = True
End Sub

Private Sub txtApprovedDate_GotFocus()
    SelectAll txtApprovedDate
End Sub

Private Sub txtClosedDate_Change()
    Dirty = True
End Sub

Private Sub txtCostImpact_Change()
    Dirty = True
End Sub

Private Sub txtCostImpact_GotFocus()
    SelectAll txtCostImpact
End Sub

Private Sub txtCostImpact_Validate(Cancel As Boolean)
    If Val(txtCostImpact) = 0 Then
        txtCostImpact = ""
    Else
        txtCostImpact = Format(Val(txtCostImpact), "#,##0.00")
    End If
End Sub


Private Sub txtDueDate_Change()
    Dirty = True
End Sub
Private Sub txtQuestion_Change()
    Dirty = True
End Sub
Private Sub txtRecipients_Change()
    Dirty = True
End Sub

Private Sub txtRecipients_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 And Shift = 0 Then Call cmdBrowse_Click(4)
End Sub

Private Sub txtRecipientsBCC_Change()
    Dirty = True
End Sub

Private Sub txtRecipientsBCC_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 And Shift = 0 Then Call cmdBrowse_Click(6)
End Sub

Private Sub txtRecipientsCC_Change()
    Dirty = True
End Sub

Private Sub txtRecipientsCC_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 And Shift = 0 Then Call cmdBrowse_Click(5)
End Sub

Private Sub txtSchedulingImpact_Change()
    Dirty = True
End Sub

Private Sub txtSchedulingImpact_GotFocus()
    SelectAll txtSchedulingImpact
End Sub

Private Sub txtSender_Change()
    Dirty = True
End Sub

Private Sub txtSender_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 And Shift = 0 Then Call cmdBrowse_Click(3)
End Sub

Private Sub txtSentDate_Change()
    Dirty = True
End Sub
Private Sub txtSubject_Change()
    Dirty = True
End Sub




Private Property Let Dirty(RHS As Boolean)
    If Not mLoading Then
        mDirty = RHS
    End If
End Property

Private Sub txtSubject_GotFocus()
    SelectAll txtSubject
End Sub
Private Sub txtAnswer_GotFocus()
    SelectAll txtAnswer
End Sub
Private Sub txtClosedDate_GotFocus()
    SelectAll txtClosedDate
End Sub
Private Sub txtDueDate_GotFocus()
    SelectAll txtDueDate
End Sub
Private Sub txtQuestion_GotFocus()
    SelectAll txtQuestion
End Sub
Private Sub txtRecipients_GotFocus()
    SelectAll txtRecipients
End Sub
Private Sub txtRecipientsBCC_GotFocus()
    SelectAll txtRecipientsBCC
End Sub
Private Sub txtRecipientsCC_GotFocus()
    SelectAll txtRecipientsCC
End Sub
Private Sub txtSender_GotFocus()
    SelectAll txtSender
End Sub
Private Sub txtSentDate_GotFocus()
    SelectAll txtSentDate
End Sub
















Public Function AddFile(ObjectID As String, Embedded As Boolean, Optional FileName As String)
    Dim s As String
    Dim FInfo As ClsFileInfo
    
    If FileName = "" Then
        If VBGetOpenFileName(FileName, , , , , True, , , , "Attach File") Then
            If PathIsLocalPath(FileName) And Not Embedded Then
                If MsgBox("This is a local file.  It may not be accessible to other users." & vbCrLf & "Are you sure this is what you want to do?", vbQuestion Or vbYesNo, App.ProductName) = vbNo Then
                    FileName = ""
                End If
            End If
        End If
    End If
    If FileName = "" Then Exit Function

    Set FInfo = New ClsFileInfo
    FInfo.FullPathName = FileName
    
    If Not FInfo.FileExists Then
        MsgBox "file not found"
    Else
        
        s = ""
        s = s & "insert into attachments(objectid,filename,modifieddate,createddate,attacheddate,embedded,weblocation)" & vbCrLf
        s = s & "values(" & DbQuote(Str, ObjectID) & vbCrLf
        s = s & "      ," & DbQuote(Str, FileName) & vbCrLf
        s = s & "      ," & DbQuote(DateTime, FInfo.ModifyTime) & vbCrLf
        s = s & "      ," & DbQuote(DateTime, FInfo.CreationTime) & vbCrLf
        s = s & "      ,getdate()" & vbCrLf
        s = s & "      ," & DbQuote(Bit, Embedded) & vbCrLf
        s = s & "      ,'Documents')"
        Call HFApp.SqlExec(s, dbHomefront)
        
        If Embedded Then
            Call DBPutFile(HFApp.Databases(dbHomefront), "Attachments WHERE objectid=" & DbQuote(Str, ObjectID) & " and filename=" & DbQuote(Str, FileName), "FileImage", FileName)
        End If
                
        Call AddFileToGrid(ObjectID, Embedded, FileName)
    End If

End Function


Private Sub gAttachments_OLEDragDrop(Data As VSFlex8Ctl.VSDataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error GoTo eh
    Dim i As Long
    
    Effect = vbDropEffectNone
    For i = 0 To Data.FileCount - 1
        Call AddFile(ObjectID, False, Data.Files(i))
    Next
    
eh: Exit Sub
End Sub

Private Sub gAttachments_OLEDragOver(Data As VSFlex8Ctl.VSDataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, State As Integer)
    Effect = IIf(Data.GetFormat(vbCFFiles), vbDropEffectCopy, vbDropEffectNone)
End Sub

Private Function ObjectID() As String
    ObjectID = "RFI~" & mJob & ":" & mRFIRow
End Function

Private Sub AddFileToGrid(ObjectID As String, Embedded As Boolean, FileName As String)
    'this only gets called when you right click insert/linkto file.
    Dim r As Long
    Dim FInfo As New ClsFileInfo
    With gAttachments
        FInfo.FullPathName = FileName
        
        Call .AddItem("")
        r = .Rows - 1
        
        .TextMatrix(r, .ColIndex("FileName")) = FileName
        .TextMatrix(r, .ColIndex("FileSize")) = FInfo.FormatFileSize(FInfo.FileSize)
        .TextMatrix(r, .ColIndex("ModifiedDate")) = Format(FInfo.ModifyTime, "general date")
        .TextMatrix(r, .ColIndex("CreatedDate")) = Format(FInfo.CreationTime, "general date")
        If Embedded Then
            .Cell(flexcpPicture, r, .ColIndex("Path")) = picEmbedded.Picture
            .Cell(flexcpForeColor, r, .ColIndex("FileTitle")) = &HC00000 '&H00FF0000&  '&H800000
        End If
        .TextMatrix(r, .ColIndex("AttachedDate")) = Format(Now(), "general date")
        .Cell(flexcpPicture, r, .ColIndex("Filetitle")) = FileIcon(FInfo.hSmlIList, FInfo.hSmlIcon)
        .TextMatrix(r, .ColIndex("FileTitle")) = FileTitle(FileName)
        .TextMatrix(r, .ColIndex("Path")) = FilePath(FileName)
        .TextMatrix(r, .ColIndex("ObjectID")) = ObjectID
        .TextMatrix(r, .ColIndex("DocumentClass")) = ""
        .TextMatrix(r, .ColIndex("Revision")) = ""
        .TextMatrix(r, .ColIndex("Embedded")) = Embedded
        
        If IsIn(left(ObjectID, 2), "C~", "J~") Then
            .TextMatrix(r, .ColIndex("WebPortals")) = "None"
            .TextMatrix(r, .ColIndex("WebLocation")) = "Documents"
        End If
    End With
End Sub




Private Function FileIcon(hList As Long, hIcon As Long) As IPictureDisp
    Set picCanvas.Picture = New StdPicture
    Call ImageList_Draw(hList, hIcon, picCanvas.hDC, 0, 0, ILD_TRANSPARENT)
    Set FileIcon = picCanvas.Image
End Function





Private Sub txtSentDate_Validate(Cancel As Boolean)
    With txtSentDate
    If IsDate(.Text) Or Trim(.Text) = "" Then
        .Text = Format(.Text, "medium date")
    Else
        Cancel = True
    End If
    End With
End Sub

Private Sub txtDueDate_Validate(Cancel As Boolean)
    With txtDueDate
    If IsDate(.Text) Or Trim(.Text) = "" Then
        .Text = Format(.Text, "medium date")
    Else
        Cancel = True
    End If
    End With
End Sub
Private Sub txtClosedDate_Validate(Cancel As Boolean)
    With txtClosedDate
    If IsDate(.Text) Or Trim(.Text) = "" Then
        .Text = Format(.Text, "medium date")
    Else
        Cancel = True
    End If
    End With
End Sub
Private Sub txtApprovedDate_Validate(Cancel As Boolean)
    With txtApprovedDate
    If IsDate(.Text) Or Trim(.Text) = "" Then
        .Text = Format(.Text, "medium date")
    Else
        Cancel = True
    End If
    End With
End Sub




Private Sub CreateFile(FileName As String)
    Dim i As Integer
    i = FreeFile
    Open FileName For Output As #i
    Close #i
End Sub





Private Sub gAttachments_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With gAttachments
        Select Case True
            Case .ColKey(.Col) = "DocumentClass":
                Call HFApp.SqlExec("update attachments set documentclass=" & DbQuote(Str, .EditText, , True) & " where objectid=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("ObjectID"))) & " and filename=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("FileName"))), dbHomefront)
                
             Case .ColKey(.Col) = "Revision":
                Call HFApp.SqlExec("update attachments set revision=" & DbQuote(Str, .EditText, , True) & " where objectid=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("ObjectID"))) & " and filename=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("FileName"))), dbHomefront)
                
            Case .ColKey(.Col) = "WebLocation"
                Call HFApp.SqlExec("update attachments set WebLocation=" & DbQuote(Str, .EditText, , True) & " where objectid=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("ObjectID"))) & " and filename=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("FileName"))), dbHomefront)
                
            Case .ColKey(.Col) = "WebPortals" And .EditText = "All"
                Call HFApp.SqlExec("update attachments set WebCustomers=1,WebVendors=1 where objectid=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("ObjectID"))) & " and filename=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("FileName"))), dbHomefront)
                
            Case .ColKey(.Col) = "WebPortals" And .EditText = "Customer"
                Call HFApp.SqlExec("update attachments set WebCustomers=1,WebVendors=0 where objectid=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("ObjectID"))) & " and filename=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("FileName"))), dbHomefront)
            
            Case .ColKey(.Col) = "WebPortals" And .EditText = "Vendor"
                Call HFApp.SqlExec("update attachments set WebCustomers=0,WebVendors=1 where objectid=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("ObjectID"))) & " and filename=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("FileName"))), dbHomefront)
            
            Case .ColKey(.Col) = "WebPortals"
                Call HFApp.SqlExec("update attachments set WebCustomers=0,WebVendors=0 where objectid=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("ObjectID"))) & " and filename=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("FileName"))), dbHomefront)
                
        End Select
    End With
End Sub

Private Sub gAttachments_AfterMoveColumn(ByVal Col As Long, Position As Long)
    With gAttachments
        .OutlineCol = .ColIndex("FileTitle")
    End With
End Sub

Private Sub gAttachments_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gAttachments
        .ComboList = ""
        .AutoSearch = flexSearchFromCursor
        If .TextMatrix(.Row, .ColIndex("FileName")) = "" Then
            Cancel = True
        Else
            Select Case .ColKey(.Col)
            
                Case "WebPortals", "WebLocation"
                    .AutoSearch = flexSearchNone
                    Cancel = Not IsIn(left(.TextMatrix(Row, .ColIndex("ObjectID")), 2), "C~", "J~")

                Case "Revision":
                    .EditMaxLength = 15
                    .AutoSearch = flexSearchNone

                Case "DocumentClass":
                    .EditMaxLength = 50
                    .AutoSearch = flexSearchNone
                    .ComboList = "| |" & .BuildComboList(HFApp.SqlExec("select distinct documentclass from attachments where isnull(documentclass,'') <>''"), "documentclass")
                    
                Case Else
                    Cancel = True
            End Select
        End If
    End With
End Sub


Private Sub gAttachments_DblClick()
    If gAttachments.MouseRow > 1 Then
        'Call mnuFilesSub_Click(mcFILE_OPEN)
    End If
End Sub


Private Sub gAttachments_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case True
'        Case KeyCode = vbKeyDelete:                   Call mnuFilesSub_Click(mcFILE_REMOVE)
'        Case KeyCode = vbKeyReturn:                   Call mnuFilesSub_Click(mcFILE_OPEN)
'        Case KeyCode = vbKeyP And Shift = vbCtrlMask: Call mnuFilesSub_Click(mcFILE_PRINT)
'        Case KeyCode = vbKeyS And Shift = vbCtrlMask: Call mnuFilesSub_Click(mcFILE_SAVE)
    End Select
End Sub


Private Sub gAttachments_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Error Resume Next
'    With gAttachments
'    If Button = vbRightButton Then
'        If .MouseRow = 0 Then
'            mnuColumnsSub(mcCOL_NAME).Checked = Not .ColHidden(.ColIndex("FileTitle"))
'            mnuColumnsSub(mcCOL_CLASS).Checked = Not .ColHidden(.ColIndex("DocumentClass"))
'            mnuColumnsSub(mcCOL_REVISION).Checked = Not .ColHidden(.ColIndex("Revision"))
'            mnuColumnsSub(mcCOL_LOCATION).Checked = Not .ColHidden(.ColIndex("Path"))
'            mnuColumnsSub(mcCOL_SIZE).Checked = Not .ColHidden(.ColIndex("FileSize"))
'            mnuColumnsSub(mcCOL_MODIFIED).Checked = Not .ColHidden(.ColIndex("ModifiedDate"))
'            mnuColumnsSub(mcCOL_CREATED).Checked = Not .ColHidden(.ColIndex("CreatedDate"))
'            mnuColumnsSub(mcCOL_ATTACHED).Checked = Not .ColHidden(.ColIndex("AttachedDate"))
'            mnuColumnsSub(mcCOL_WEBPORTALS).Checked = Not .ColHidden(.ColIndex("WebPortals"))
'            mnuColumnsSub(mcCOL_WEBLOCATION).Checked = Not .ColHidden(.ColIndex("WebLocation"))
'            PopupMenu Me.mnuColumns
'        Else
'            If .TextMatrix(.Row, .ColIndex("FileName")) <> "" Then
'                mnuFilesSub(mcFILE_PROPERTIES).Enabled = .TextMatrix(.Row, .ColIndex("Embedded")) = "False"
'                PopupMenu Me.mnuFiles
'            Else
'                gAttachments.Row = gAttachments.MouseRow
'                If gAttachments.RowData(gAttachments.Row) <> "" Then PopupMenu Me.mnuFolders
'            End If
'        End If
'    End If
'    End With
End Sub





Private Sub gMessages_BeforeSort(ByVal Col As Long, Order As Integer)
On Error Resume Next
    gMessages.RemoveItem gMessages.Rows - 1
End Sub

Private Sub gMessages_AfterSort(ByVal Col As Long, Order As Integer)
On Error Resume Next
    gMessages.AddItem ""
End Sub


Private Sub gMessages_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error Resume Next
    Dim i As Long
    With gMessages
        If Row <> .Rows - 1 Then Exit Sub
        
        i = Row
        .TextMatrix(i, .ColIndex("Priority")) = "1"
        .TextMatrix(i, .ColIndex("MessageType")) = "Other"
        .TextMatrix(i, .ColIndex("ReceivedDate")) = Format(Now(), "medium date")
    
        'add new last row
        .AddItem ""
    
    
    End With
End Sub


Private Sub gMessages_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo eh

    Dim s  As String
    Dim q  As String
    
    With gMessages
        s = .EditText
        Select Case .ColKey(Col)
                
            Case "SentDate", "ReceivedDate"
                If IsDate(s) Then
                    s = Format(s, "Medium Date")
                Else
                    Cancel = True
                End If
         End Select
        .EditText = s
        
        If Not Cancel Then
            Dirty = True
            .RowData(Row) = "DIRTY"
        End If
    End With
    
Exit Sub
eh: Call errHandler(SRCFILE & "gMessages_ValidateEdit")
End Sub








