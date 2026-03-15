VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form FDirectCosts 
   Caption         =   "Journal Entries"
   ClientHeight    =   6660
   ClientLeft      =   8835
   ClientTop       =   3870
   ClientWidth     =   10890
   Icon            =   "FDirectCosts.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6660
   ScaleWidth      =   10890
   Begin VSFlex8Ctl.VSFlexGrid gData 
      Height          =   3915
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   10665
      _cx             =   1969179004
      _cy             =   1969167098
      Appearance      =   2
      BorderStyle     =   0
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
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   14
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FDirectCosts.frx":000C
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
   Begin VB.Frame frmHead 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1275
      Left            =   0
      TabIndex        =   11
      Top             =   630
      Width           =   10785
      Begin VB.TextBox txtDate 
         BorderStyle     =   0  'None
         Height          =   230
         Left            =   1500
         MaxLength       =   100
         TabIndex        =   2
         Top             =   945
         Width           =   2595
      End
      Begin VB.TextBox txtDescription 
         BorderStyle     =   0  'None
         Height          =   230
         Left            =   1500
         MaxLength       =   200
         TabIndex        =   0
         Text            =   " "
         Top             =   450
         Width           =   2835
      End
      Begin VB.TextBox txtComments 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   4470
         TabIndex        =   3
         Text            =   " "
         Top             =   300
         Width           =   6135
      End
      Begin HFSystem.VBCombo cboBatchType 
         Height          =   240
         Left            =   1500
         TabIndex        =   1
         Top             =   690
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   423
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Accounting Date"
         Height          =   255
         Index           =   3
         Left            =   30
         TabIndex        =   22
         Top             =   930
         Width           =   1395
      End
      Begin VB.Image cmdBrowse 
         Height          =   240
         Left            =   4110
         Picture         =   "FDirectCosts.frx":024F
         Top             =   945
         Width           =   240
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Batch Number"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   1
         Left            =   30
         TabIndex        =   16
         Top             =   210
         Width           =   1395
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Description"
         Height          =   255
         Index           =   0
         Left            =   30
         TabIndex        =   15
         Top             =   450
         Width           =   1395
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Batch Type"
         Height          =   255
         Index           =   2
         Left            =   30
         TabIndex        =   14
         Top             =   690
         Width           =   1395
      End
      Begin VB.Label lblBatchNumber 
         AutoSize        =   -1  'True
         Caption         =   "123456"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1500
         TabIndex        =   13
         Top             =   210
         Width           =   645
      End
      Begin VB.Label Label11 
         Caption         =   "Notes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   10
         Left            =   4500
         TabIndex        =   12
         Top             =   90
         Width           =   1395
      End
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   5
      Top             =   0
      Width           =   10890
      _ExtentX        =   19209
      _ExtentY        =   1058
      ButtonWidth     =   1984
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            Key             =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open"
            Key             =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save As"
            Key             =   "SaveAs"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            Key             =   "Delete"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Post"
            Key             =   "Post"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Attachments"
            Key             =   "Attachments"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList LargeIcons 
         Left            =   9450
         Top             =   30
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   56
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":0399
               Key             =   "EditAssembly"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":0C73
               Key             =   "RFP"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":154D
               Key             =   "Post"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":1E27
               Key             =   "CreateJob"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":2701
               Key             =   "takeoffsettings"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":2FDB
               Key             =   "quote"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":38B5
               Key             =   "DecreaseDecimals"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":418F
               Key             =   "IncreaseDecimals"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":4A69
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":5343
               Key             =   "ExcelImport"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":5C1D
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":64F7
               Key             =   "UpdatePrices"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":6DD1
               Key             =   "ExcelExport"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":76AB
               Key             =   "Publish"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":7F85
               Key             =   "Forecast"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":885F
               Key             =   "ViewPOs"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":9139
               Key             =   "ViewBudgets"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":9A13
               Key             =   "Open"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":A2ED
               Key             =   "Preview"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":ABC7
               Key             =   "Send"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":B4A1
               Key             =   "TakeoffOneTime"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":BD7B
               Key             =   "Estimate"
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":C655
               Key             =   "TakeoffAssembly"
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":CF2F
               Key             =   "NewAssembly"
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":D809
               Key             =   "New"
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":E0E3
               Key             =   "TakeoffItem"
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":E9BD
               Key             =   "TakeoffCustom"
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":F297
               Key             =   "Save"
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":FB71
               Key             =   "SaveAs"
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":1044B
               Key             =   "Delete"
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":10D25
               Key             =   "AddPricelist"
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":115FF
               Key             =   "RePrice"
            EndProperty
            BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":11ED9
               Key             =   "Pricebook"
            EndProperty
            BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":127B3
               Key             =   "PricebookEdit"
            EndProperty
            BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":1308D
               Key             =   "PricelistExport"
            EndProperty
            BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":13967
               Key             =   "PricelistImport"
            EndProperty
            BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":14241
               Key             =   "NewPricelist"
            EndProperty
            BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":14B1B
               Key             =   "View"
            EndProperty
            BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":153F5
               Key             =   "Vendor1"
            EndProperty
            BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":15CCF
               Key             =   "Vendor"
            EndProperty
            BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":165A9
               Key             =   "Add"
            EndProperty
            BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":16E83
               Key             =   "Attachments"
            EndProperty
            BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":1775D
               Key             =   ""
            EndProperty
            BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":18037
               Key             =   ""
            EndProperty
            BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":18911
               Key             =   "Design Center Options"
            EndProperty
            BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":191EB
               Key             =   "Global Options"
            EndProperty
            BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":19AC5
               Key             =   "Models"
            EndProperty
            BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":1A39F
               Key             =   "Options"
            EndProperty
            BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":1AC79
               Key             =   "Generate"
            EndProperty
            BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":1B553
               Key             =   "Timberline"
            EndProperty
            BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":1BE2D
               Key             =   "MBImport"
            EndProperty
            BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":1C707
               Key             =   "EEEstimating"
            EndProperty
            BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":1CFE1
               Key             =   "EEExport"
            EndProperty
            BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":1D8BB
               Key             =   "EEImport"
            EndProperty
            BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":1E195
               Key             =   "MBExport"
            EndProperty
            BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":1EA6F
               Key             =   "MasterBuilder"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList SmallIcons 
         Left            =   8820
         Top             =   30
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   73
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":1F749
               Key             =   "option"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":1FCE3
               Key             =   "RFP"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":2027D
               Key             =   "quote"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":20817
               Key             =   "sendreceive"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":20DB1
               Key             =   "communitystandards"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":2134B
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":21C25
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":224FF
               Key             =   "pricelists"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":22DD9
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":236B3
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":23F8D
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":24867
               Key             =   "MB"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":24E01
               Key             =   "QB"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":2539B
               Key             =   "custom"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":25935
               Key             =   "assembly"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":25ECF
               Key             =   "links"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":26469
               Key             =   "ItemDB"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":26A03
               Key             =   "sendpos"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":26F9D
               Key             =   "HelpSearch"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":27537
               Key             =   "Items"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":27AD1
               Key             =   "New"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":2806B
               Key             =   "Edit"
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":28605
               Key             =   "HelpContents"
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":28B9F
               Key             =   "EditAssembly"
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":29139
               Key             =   "Forecast"
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":296D3
               Key             =   "shrink"
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":29C6D
               Key             =   "preview"
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":2A207
               Key             =   "customer"
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":2A7A1
               Key             =   "close"
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":2AD3B
               Key             =   "expand"
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":2B2D5
               Key             =   "SaveAs"
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":2B86F
               Key             =   "Save"
            EndProperty
            BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":2BE09
               Key             =   "RePrice"
            EndProperty
            BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":2C3A3
               Key             =   ""
            EndProperty
            BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":2C93D
               Key             =   "estimating"
            EndProperty
            BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":2CED7
               Key             =   "jobcost"
            EndProperty
            BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":2D471
               Key             =   "error"
            EndProperty
            BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":2DA0B
               Key             =   "salesworksheet"
            EndProperty
            BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":2DFA5
               Key             =   "ExcelExport"
            EndProperty
            BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":2E53F
               Key             =   ""
            EndProperty
            BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":2EAD9
               Key             =   "newworksheet"
            EndProperty
            BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":2F073
               Key             =   "worksheet"
            EndProperty
            BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":2F60D
               Key             =   "purchaseorder"
            EndProperty
            BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":2FBA7
               Key             =   "ExcelImport"
            EndProperty
            BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":30141
               Key             =   "groupphase"
            EndProperty
            BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":306DB
               Key             =   "costcode"
            EndProperty
            BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":30C75
               Key             =   "job"
            EndProperty
            BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":3120F
               Key             =   "information"
            EndProperty
            BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":317A9
               Key             =   "item"
            EndProperty
            BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":31D43
               Key             =   "itemchecked"
            EndProperty
            BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":322DD
               Key             =   "phase"
            EndProperty
            BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":32877
               Key             =   "vendor"
            EndProperty
            BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":32E11
               Key             =   "warning"
            EndProperty
            BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":333AB
               Key             =   "question"
            EndProperty
            BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":33945
               Key             =   "category"
            EndProperty
            BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":33EDF
               Key             =   "folder"
            EndProperty
            BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":34479
               Key             =   "AddItems"
            EndProperty
            BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":34A13
               Key             =   "model"
            EndProperty
            BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":34FAD
               Key             =   "area"
            EndProperty
            BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":35547
               Key             =   "Delete"
            EndProperty
            BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":35AE1
               Key             =   "Underline"
            EndProperty
            BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":35C3B
               Key             =   "Bold"
            EndProperty
            BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":35D95
               Key             =   "AlignCenter"
            EndProperty
            BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":35EEF
               Key             =   "Italic"
            EndProperty
            BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":36049
               Key             =   "AlignLeft"
            EndProperty
            BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":361A3
               Key             =   "Bullet"
            EndProperty
            BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":362FD
               Key             =   "BulletNumber"
            EndProperty
            BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":36457
               Key             =   "AlignRight"
            EndProperty
            BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":365B1
               Key             =   "printer"
            EndProperty
            BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":36B4B
               Key             =   "defaultprinter"
            EndProperty
            BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":370E5
               Key             =   "costcodes"
            EndProperty
            BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":379BF
               Key             =   "defaultvendors"
            EndProperty
            BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDirectCosts.frx":38299
               Key             =   "editpos"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame frmFoot 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   705
      Left            =   120
      TabIndex        =   6
      Top             =   5820
      Width           =   10650
      Begin VB.Frame frmTotals 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   705
         Left            =   8490
         TabIndex        =   17
         Top             =   0
         Width           =   2175
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Total Debits"
            Height          =   195
            Left            =   90
            TabIndex        =   21
            Top             =   150
            Width           =   855
         End
         Begin VB.Label lblDebit 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "$1,500.00"
            Height          =   195
            Left            =   1305
            TabIndex        =   20
            Top             =   150
            Width           =   720
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Total Credits"
            Height          =   195
            Left            =   60
            TabIndex        =   19
            Top             =   360
            Width           =   885
         End
         Begin VB.Label lblCredit 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "$1,500.00"
            Height          =   195
            Left            =   1305
            TabIndex        =   18
            Top             =   360
            Width           =   720
         End
      End
      Begin VB.Label lblPostedDate 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Height          =   195
         Left            =   795
         TabIndex        =   10
         Top             =   360
         Width           =   45
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Posted"
         Height          =   195
         Index           =   8
         Left            =   225
         TabIndex        =   9
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblModifiedDate 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Height          =   195
         Left            =   795
         TabIndex        =   8
         Top             =   150
         Width           =   45
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Modified"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   7
         Top             =   150
         Width           =   600
      End
   End
End
Attribute VB_Name = "FDirectCosts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Public EventTraps As Collection
Const SRCFILE = "FDirectCosts::"

Private mLocked   As Boolean
Private mDirty    As Boolean


'grid menus
Private Const mcGRID_ASC = 0
Private Const mcGRID_DESC = 1
Private Const mcGRID_GOTO = 3
Private Const mcGRID_HIDE = 4
Private Const mcGRID_INSERT = 5
Private Const mcGRID_RENAME = 6
Private Const mcGRID_PRINT = 8
Private Const mcGRID_SAVEAS = 9
Private MouseCtrl  As Control
Private MouseCol   As Long

Private AccountMask As String
    Private mCancelEdit    As Boolean

Private Sub cboBatchType_Change()
    mDirty = True
    Call ValidateData
End Sub

Private Sub cboBatchType_Click()
    mDirty = True
End Sub

Private Sub cboBatchType_Validate(Cancel As Boolean)
    cboBatchType.Text = left(cboBatchType.Text, 200)
End Sub

Private Sub Form_Load()
    
    Call SetToolbarIcons(Toolbar, LargeIcons)
    Call IniGetForm(Me)
    Call IniGetGrid(Me, gData)
    Call GetFormatMasks
    Call LoadData(0)
    
End Sub


Private Sub Form_Resize()
On Error Resume Next
    
    Const margin = 120
    frmHead.Move 0, Toolbar.Height, Me.ScaleWidth
    txtComments.Width = Me.ScaleWidth - margin - txtComments.left
    gData.Move margin, frmHead.Top + frmHead.Height, Me.ScaleWidth - 2 * margin, Me.ScaleHeight - margin - frmHead.Top - frmHead.Height - frmFoot.Height
    frmFoot.Move margin, gData.Top + gData.Height, gData.Width - Screen.TwipsPerPixelX
    frmTotals.left = frmFoot.Width - frmTotals.Width
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If Not SaveData(True) Then
        Cancel = True
        Exit Sub
    End If
    Call IniPutForm(Me)
    Call IniPutGrid(Me, gData)

End Sub



Public Function SaveData(prompt As Boolean) As Boolean
On Error GoTo eh
    Dim s As String
    Dim r As Long
    
    If Not mDirty Then
        SaveData = True
        Exit Function
    End If
    
    If Not ValidateData() Then
        If prompt = False Then
            Call MsgBox("This data is incomplete and cannot be saved.", vbExclamation, App.ProductName)
            SaveData = False
            Exit Function
        Else
            Select Case MsgBox("This data is incomplete and cannot be saved." & vbCrLf & vbCrLf & "Your changes will be lost.", vbExclamation + vbOKCancel + vbDefaultButton2, App.ProductName)
                Case vbOK
                    SaveData = True
                    Exit Function
                Case Else
                    SaveData = False
                    Exit Function
            End Select
        End If
    End If
    
    If prompt Then
        Select Case MsgBox("This data has changed." & vbCrLf & vbCrLf & "Do you want to save these changes?", vbExclamation + vbYesNoCancel, App.ProductName)
            Case vbNo
                SaveData = True
                Exit Function
            Case vbCancel
                SaveData = False
                Exit Function
        End Select
    End If
    
    Screen.MousePointer = vbHourglass
    
    s = ""
    s = s & "update journalentries" & vbCrLf
    s = s & "   set Description=" & DbQuote(Str, txtDescription.Text) & vbCrLf
    s = s & "      ,BatchType=" & DbQuote(Str, cboBatchType.Text) & vbCrLf
    s = s & "      ,AccountingDate=" & DbQuote(Date, txtDate.Text) & vbCrLf
    s = s & "      ,Comments=" & DbQuote(Str, txtComments.Text) & vbCrLf
    s = s & "where Batchid=" & DbQuote(Num, lblBatchNumber.Caption)
    Call HFApp.SqlExec(s, dbHomeFront)
    
    Call HFApp.SqlExec("delete from journalentrylines where batchid=" & DbQuote(Num, lblBatchNumber.Caption))
    With gData
        For r = 1 To .Rows - 2
            s = ""
            s = s & "insert into journalentrylines(DivisionID,BatchID,Description,Job,Extra,CostCode,Category,Account,DebitAmount,CreditAmount)" & vbCrLf
            s = s & "values(" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
            s = s & "      ," & DbQuote(Num, lblBatchNumber.Caption) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Description"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Job"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Extra"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("CostCode"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Category"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Account"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(r, .ColIndex("Debit"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(r, .ColIndex("Credit"))) & vbCrLf
            s = s & ")"
            HFApp.SqlExec s
        Next
    End With
       
    SaveData = True
    mDirty = False
    Screen.MousePointer = vbDefault
    
Exit Function
eh: Call errHandler(SRCFILE & "SaveData", s)
End Function



Private Sub LoadData(BatchID As Long, Optional CreateNew As Boolean)
On Error GoTo eh
    Dim s As String
    Dim rs As Recordset
    Dim r As Long
    
    
    If BatchID = 0 Then
        On Error Resume Next
        If Not CreateNew Then
            BatchID = Val("" & HFApp.SqlExec("select min(batchid) from journalentries where isnull(postingbatchid,0)=0", dbHomeFront)(0))
        End If
        On Error GoTo eh
        If BatchID = 0 Then
            Call HFApp.SqlExec("insert into journalentries(Description,AccountingDate) values('new batch',convert(varchar(25), getdate(),102))")
            BatchID = HFApp.SqlIdentity("journalentries")
        End If
    End If
    
    Call LoadTypes
    
    s = "select * from journalentries where batchid=" & DbQuote(Num, BatchID)
    Set rs = HFApp.SqlExec(s)
    Locked = Val("" & rs("PostingBatchID")) <> 0
    
    lblBatchNumber.Caption = "" & rs("BatchID")
    txtDescription.Text = "" & rs("Description")
    cboBatchType.Text = "" & rs("BatchType")
    txtDate.Text = Format("" & rs("AccountingDate"), HFApp.Options(DateFormat))
    txtComments.Text = "" & rs("Comments")
    If "" & rs("ModifiedDate") <> "" Then lblModifiedDate.Caption = Format("" & rs("ModifiedDate"), HFApp.Options(DateFormat))
    If "" & rs("PostingDate") <> "" Then Me.lblPostedDate.Caption = Format("" & rs("PostingDate"), HFApp.Options(DateFormat))
    



    s = ""
    s = s & "select l.*" & vbCrLf
    s = s & "      ,j.description JobDesc " & vbCrLf
    s = s & "      ,e.description ExtraDesc " & vbCrLf
    s = s & "      ,c.description CostCodeDesc " & vbCrLf
    s = s & "      ,g.description CategoryDesc " & vbCrLf
    s = s & "      ,a.description AccountDesc " & vbCrLf
    s = s & "  from journalentrylines l" & vbCrLf
    s = s & "       left outer join tbljobs j on(l.DivisionID = j.DivisionID and l.job=j.job_no)" & vbCrLf
    s = s & "       left outer join jobextras e on(l.DivisionID = e.DivisionID and l.job=e.job and l.extra=e.extra)" & vbCrLf
    s = s & "       left outer join standardcostcodes c on(l.DivisionID = c.DivisionID and l.costcode=c.costcode)" & vbCrLf
    s = s & "       left outer join standardcategories g on(l.DivisionID = g.DivisionID and l.category=g.category)" & vbCrLf
    s = s & "       left outer join glaccounts a on(l.DivisionID = a.DivisionID and l.account=a.account)" & vbCrLf
    s = s & " where l.DivisionID = " & HFApp.DivisionID & " and batchid=" & DbQuote(Num, BatchID)
    Set rs = HFApp.SqlExec(s)
    With gData
        .Rows = 1
        r = 0
        While Not rs.EOF
            r = r + 1
            .AddItem ""
            .TextMatrix(r, .ColIndex("Sequence")) = "" & rs("Sequence")
            .TextMatrix(r, .ColIndex("Job")) = "" & rs("Job")
            .TextMatrix(r, .ColIndex("JobDesc")) = "" & rs("JobDesc")
            .TextMatrix(r, .ColIndex("Extra")) = "" & rs("Extra")
            .TextMatrix(r, .ColIndex("ExtraDesc")) = "" & rs("ExtraDesc")
            .TextMatrix(r, .ColIndex("CostCode")) = "" & rs("CostCode")
            .TextMatrix(r, .ColIndex("CostCodeDesc")) = "" & rs("CostCodeDesc")
            .TextMatrix(r, .ColIndex("Category")) = "" & rs("Category")
            .TextMatrix(r, .ColIndex("CategoryDesc")) = "" & rs("CategoryDesc")
            .TextMatrix(r, .ColIndex("Description")) = "" & rs("Description")
            .TextMatrix(r, .ColIndex("Account")) = "" & rs("Account")
            .TextMatrix(r, .ColIndex("AccountDesc")) = "" & rs("AccountDesc")
            If Val("" & rs("DebitAmount")) = 0 Then
                .TextMatrix(r, .ColIndex("Credit")) = Val("" & rs("CreditAmount"))
            Else
                .TextMatrix(r, .ColIndex("Debit")) = Val("" & rs("DebitAmount"))
            End If
            rs.MoveNext
        Wend
        .AddItem ""
    End With
    Call gData_AfterEdit(0, 0)
    Call ValidateData
    
    
    mDirty = False
Exit Sub
eh: Call errHandler(SRCFILE & "LoadData", s)
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case True
        Case Shift = vbCtrlMask And KeyCode = vbKeyO:   KeyCode = 0:   Call Toolbar_ButtonClick(Toolbar.Buttons("Open"))
        Case Shift = vbCtrlMask And KeyCode = vbKeyS:   KeyCode = 0:   Call Toolbar_ButtonClick(Toolbar.Buttons("Save"))
    End Select
End Sub



Private Sub gData_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim s As String
    Dim i As Long
    Dim Description As String
    Dim rs As Recordset
    
    Dim AddCCOK  As Boolean
    Dim AddCatOK As Boolean
    
    mCancelEdit = False
    With gData
        s = .TextMatrix(Row, Col)
        Select Case .ColKey(Col)
            Case "Job"
                If s = "" Then
                    .TextMatrix(Row, .ColIndex("Job")) = ""
                    .TextMatrix(Row, .ColIndex("JobDesc")) = ""
                    .TextMatrix(Row, .ColIndex("Extra")) = ""
                    .TextMatrix(Row, .ColIndex("ExtraDesc")) = ""
                    .TextMatrix(Row, .ColIndex("CostCode")) = ""
                    .TextMatrix(Row, .ColIndex("CostCodeDesc")) = ""
                    .TextMatrix(Row, .ColIndex("Category")) = ""
                    .TextMatrix(Row, .ColIndex("CategoryDesc")) = ""
                Else
                    s = HFApp.FormatJob(s)
                    Set rs = HFApp.SqlExec(ValidateJob(s), AccountingDB)
                    If rs.EOF Then
                        .TextMatrix(Row, .ColIndex("JobDesc")) = ""
                        mCancelEdit = True
                        MsgBox .TextMatrix(0, Col) & " not found." & vbCrLf & "It may have been closed.", vbExclamation, App.ProductName
                    Else
                        s = "" & rs("job")
                        .TextMatrix(Row, .ColIndex("JobDesc")) = "" & rs("Description")
                    End If
                    .TextMatrix(Row, .ColIndex("Extra")) = ""
                    .TextMatrix(Row, .ColIndex("ExtraDesc")) = ""
                    .TextMatrix(Row, .ColIndex("CostCode")) = ""
                    .TextMatrix(Row, .ColIndex("CostCodeDesc")) = ""
                    .TextMatrix(Row, .ColIndex("Category")) = ""
                    .TextMatrix(Row, .ColIndex("CategoryDesc")) = ""
                End If
                
            Case "Extra"
                If s = "" Then
                    .TextMatrix(Row, .ColIndex("Extra")) = ""
                    .TextMatrix(Row, .ColIndex("ExtraDesc")) = ""
                    .TextMatrix(Row, .ColIndex("CostCode")) = ""
                    .TextMatrix(Row, .ColIndex("CostCodeDesc")) = ""
                    .TextMatrix(Row, .ColIndex("Category")) = ""
                    .TextMatrix(Row, .ColIndex("CategoryDesc")) = ""
                Else
                    If HFApp.Options(Use_Timberline) Then
                        Set rs = HFApp.SqlExec("select extra,description from jcm_master__extra where job=" & DbQuote(Str, HFApp.FormatJob(.TextMatrix(Row, .ColIndex("Job")))) & " AND extra=" & DbQuote(Str, s), dbAccounting)
                        If rs.EOF Then
                            Description = InputBox(vbCrLf & "Extra '" & s & "' could not be found." & vbCrLf & "Do you want to add it?" & vbCrLf & vbCrLf & vbCrLf & "Enter a description for the extra. (30 characters or less)", App.ProductName, s)
                            If Description = "" Then
                                mCancelEdit = True
                            Else
                                s = ""
                                s = s & "INSERT INTO jcm_master__extra(Job,Extra,Description,Status)" & vbCrLf
                                s = s & "VALUES(" & DbQuote(Str, HFApp.FormatJob(.TextMatrix(Row, .ColIndex("Job")))) & vbCrLf
                                s = s & "      ," & DbQuote(Str, .EditText) & vbCrLf
                                s = s & "      ," & DbQuote(Str, left(Description, 30)) & vbCrLf
                                s = s & "      ,'In progress')" & vbCrLf
                                Set rs = HFApp.SqlExec(s, dbAccounting)
                                s = .EditText
                            End If
                        Else
                            s = "" & rs("Extra")
                            .TextMatrix(Row, .ColIndex("ExtraDesc")) = "" & rs("Description")
                        End If
                    End If
                    .TextMatrix(Row, .ColIndex("CostCode")) = ""
                    .TextMatrix(Row, .ColIndex("CostCodeDesc")) = ""
                    .TextMatrix(Row, .ColIndex("Category")) = ""
                    .TextMatrix(Row, .ColIndex("CategoryDesc")) = ""
                End If
                
                
            Case "CostCode":
                If s = "" Then
                    .TextMatrix(Row, .ColIndex("CostCode")) = ""
                    .TextMatrix(Row, .ColIndex("CostCodeDesc")) = ""
                    .TextMatrix(Row, .ColIndex("Category")) = ""
                    .TextMatrix(Row, .ColIndex("CategoryDesc")) = ""
                Else
                    Set rs = HFApp.SqlExec(ValidateCostCode(.TextMatrix(Row, .ColIndex("Job")), .TextMatrix(Row, .ColIndex("Extra")), s), AccountingDB)
                    If rs.EOF Then
                        mCancelEdit = True
                        MsgBox "Cost code not found", vbExclamation, App.ProductName
                    Else
                        If Val("" & rs(0)) = 1 Then
                            'found in job cost codes
                            .TextMatrix(Row, .ColIndex("CostCode")) = Trim("" & rs("costcode"))
                            .TextMatrix(Row, .ColIndex("CostCodeDesc")) = Trim("" & rs("description"))
                            .TextMatrix(Row, .ColIndex("Category")) = ""
                            .TextMatrix(Row, .ColIndex("CategoryDesc")) = ""
                            .TextMatrix(Row, .ColIndex("Account")) = Trim("" & rs("DebitAccount"))
                            .TextMatrix(Row, .ColIndex("AccountDesc")) = Trim("" & rs("DebitAccountDesc"))
                        Else
                            'not found in job cost codes
                            If AddCCOK Then
                                .TextMatrix(Row, .ColIndex("CostCode")) = Trim("" & rs("CostCode"))
                                .TextMatrix(Row, .ColIndex("CostCodeDesc")) = Trim("" & rs("description"))
                                .TextMatrix(Row, .ColIndex("Category")) = ""
                                .TextMatrix(Row, .ColIndex("CategoryDesc")) = ""
                                .TextMatrix(Row, .ColIndex("Account")) = Trim("" & rs("DebitAccount"))
                                .TextMatrix(Row, .ColIndex("AccountDesc")) = Trim("" & rs("DebitAccountDesc"))
                            ElseIf MsgBox(vbQuote & FullTrim("" & rs("description")) & vbQuote & " is not setup for job """ & FullTrim("" & .TextMatrix(Row, .ColIndex("Job"))) & " - " & FullTrim(.TextMatrix(Row, .ColIndex("JobDesc"))) & """" & vbCrLf & vbCrLf & "Do you want to add this cost code?", vbQuestion + vbYesNo) = vbYes Then
                                AddCCOK = True
                                .TextMatrix(Row, .ColIndex("costcode")) = Trim("" & rs("costcode"))
                                .TextMatrix(Row, .ColIndex("costcodeDesc")) = Trim("" & rs("description"))
                                .TextMatrix(Row, .ColIndex("Category")) = ""
                                .TextMatrix(Row, .ColIndex("CategoryDesc")) = ""
                                .TextMatrix(Row, .ColIndex("Account")) = Trim("" & rs("DebitAccount"))
                                .TextMatrix(Row, .ColIndex("AccountDesc")) = Trim("" & rs("DebitAccountDesc"))
                            Else
                                mCancelEdit = True
                            End If
                        End If
                    End If
                End If
                
                
            Case "Category"
                If s = "" Then
                    .TextMatrix(Row, .ColIndex("Category")) = ""
                    .TextMatrix(Row, .ColIndex("CategoryDesc")) = ""
                Else
                    mCancelEdit = Not ValidateField(gData, s, "Category not found", "SELECT Category,Description FROM StandardCategories WHERE DivisionID = " & HFApp.DivisionID & " and Category=" & DbQuote(Str, s), "CategoryDesc")
                End If
            
            
            Case "Account"
                If s = "" Then
                    .TextMatrix(Row, .ColIndex("Account")) = ""
                    .TextMatrix(Row, .ColIndex("AccountDesc")) = ""
                Else
                    s = FormatAccount(s)
                    Set rs = HFApp.SqlExec(ValidateAccount(s), AccountingDB)
                    If rs.EOF Then
                        mCancelEdit = True
                        MsgBox "Account not found", vbExclamation, App.ProductName
                    Else
                        s = "" & rs("Account")
                        .TextMatrix(Row, .ColIndex("AccountDesc")) = "" & rs("Description")
                    End If
                End If
                
                
            Case "Debit":      s = Round(Val(s), 2):    .TextMatrix(Row, .ColIndex("Credit")) = ""
            Case "Credit":     s = Round(Val(s), 2):    .TextMatrix(Row, .ColIndex("Debit")) = ""
        End Select
        .TextMatrix(Row, Col) = s
            
        mDirty = True
        For i = Min(Row, .RowSel) To Max(Row, .RowSel)
            If .RowData(i) <> "NEW" Then .RowData(i) = "DIRTY"
        Next
    
    End With
            
    
    Call ValidateData


    On Error Resume Next
    If mCancelEdit Then
        Call gData.EditCell
        Exit Sub
    End If
End Sub

Private Function ValidateData() As Boolean
    Dim valid As Boolean
    
    Dim s As Boolean 'true if batchtype is missing
    Dim ad As Boolean 'true if accountingdate is missing
    
    Dim z As Boolean 'true if any row has no credit and no debit
    Dim j As Boolean 'true if any row is missing job info
    Dim a As Boolean 'true if any row is missing account info
    
    Dim d As Double
    Dim c As Double
    Dim r As Long
    d = 0
    c = 0
    
    s = cboBatchType.Text = ""
    ad = Not IsDate(txtDate.Text)
    z = False
    j = False
    a = False
    With gData
        For r = 1 To .Rows - 2
               
            'change zero to blank
            If .ValueMatrix(r, .ColIndex("Debit")) = 0 Then .TextMatrix(r, .ColIndex("Debit")) = ""
            If .ValueMatrix(r, .ColIndex("Credit")) = 0 Then .TextMatrix(r, .ColIndex("Credit")) = ""
            
            'check for incomplete data
            z = z Or .ValueMatrix(r, .ColIndex("Debit")) = 0 And .ValueMatrix(r, .ColIndex("Credit")) = 0
            'j = j Or .TextMatrix(r, .ColIndex("Job")) = "" Or .TextMatrix(r, .ColIndex("CostCode")) = "" Or .TextMatrix(r, .ColIndex("Category")) = ""
            a = a Or .TextMatrix(r, .ColIndex("Account")) = ""
            
            'calc totals
            d = d + Round(.ValueMatrix(r, .ColIndex("Debit")), 2)
            c = c + Round(.ValueMatrix(r, .ColIndex("Credit")), 2)
        Next
    End With
    lblDebit.Caption = Format(d, "$#,##0.00")
    lblCredit.Caption = Format(c, "$#,##0.00")
    
    
    'valid if debits=credits and all rows complete
    valid = Round(d, 2) = Round(c, 2) And d <> 0 And Not (z Or a) And Not s And Not ad
    
    
    Toolbar.Buttons("Delete").Enabled = Not mLocked
    Toolbar.Buttons("Post").Enabled = valid And Not mLocked
    Toolbar.Buttons("Save").Enabled = valid And Not mLocked
    Toolbar.Buttons("SaveAs").Enabled = valid
    ValidateData = valid
    
    
End Function

Private Sub gData_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gData
        .ComboList = ""
        .EditMaxLength = 0
        Select Case .ColKey(Col)
            Case "Job":           .ComboList = "|..."
            Case "Extra":         .ComboList = "|..."
            Case "CostCode":      .ComboList = "|..."
            Case "Category":      .ComboList = "|..."
            Case "Description":   .EditMaxLength = 200
            Case "Account":       .ComboList = "|..."
            
            
            Case "JobDesc":           .ComboList = "..."
            Case "ExtraDesc":         .ComboList = "..."
            Case "CostCodeDesc":      .ComboList = "..."
            Case "CategoryDesc":      .ComboList = "..."
            Case "AccountDesc":       .ComboList = "..."
            
            Case "Debit"
            Case "Credit"
        End Select
    End With
End Sub



Private Sub gData_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal newrow As Long, ByVal NewCol As Long, Cancel As Boolean)
    Cancel = mCancelEdit
End Sub

Private Sub gData_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim Job      As String
    Dim Extra    As String
    Dim CostCode As String
    Dim s As String
    Dim sHiddenCols As String
    With gData
        Job = .TextMatrix(Row, .ColIndex("Job"))
        Extra = .TextMatrix(Row, .ColIndex("Extra"))
        CostCode = .TextMatrix(Row, .ColIndex("CostCode"))
    
        Select Case .ColKey(Col)
            Case "Job"
                s = SelectJob()
                
            Case "JobDesc"
                Col = .ColIndex("Job")
                s = SelectJob()
                
            Case "Extra"
                s = SelectExtra(Job)
            
            Case "ExtraDesc"
                Col = .ColIndex("Extra")
                s = SelectExtra(Job)
            
            Case "CostCode"
                s = "Job Cost Codes" & Chr(1) & SelectJobExtraCostCode(Job, Extra) & Chr(0) & _
                    "Standard Cost Codes" & Chr(1) & SelectJobExtraCostCode()
                
            Case "CostCodeDesc"
                Col = .ColIndex("CostCode")
                If Not TimberlineAccounting Then sHiddenCols = "CostCode"
                s = "Job Cost Codes" & Chr(1) & SelectJobExtraCostCode(Job, Extra) & Chr(0) & _
                    "Standard Cost Codes" & Chr(1) & SelectJobExtraCostCode()
                
            Case "Category"
                s = "Job Categories" & Chr(1) & SelectJobExtraCostCodeCategory(Job, Extra, CostCode) & Chr(0) & _
                    "Standard Categories" & Chr(1) & SelectJobExtraCostCodeCategory()
                                       
            Case "CategoryDesc"
                Col = .ColIndex("Category")
                If Not TimberlineAccounting Then sHiddenCols = "Category"
                s = "Job Categories" & Chr(1) & SelectJobExtraCostCodeCategory(Job, Extra, CostCode) & Chr(0) & _
                    "Standard Categories" & Chr(1) & SelectJobExtraCostCodeCategory()
            
            Case "Account", "AccountDesc"
                Col = .ColIndex("Account")
                s = SelectGLAccount()
                
        End Select
        
        
        If s <> "" Then
            If FPickList.Choose(HFApp.Databases(AccountingDB), .TextMatrix(0, Col), s, .Text, , , , sHiddenCols) Then
                
                s = FPickList.SelectedItem(1)
                .Cell(flexcpText, .Row, Col, .RowSel, Col) = s
                Call gData_AfterEdit(.Row, Col)
            End If
        End If
        
    End With

End Sub

Private Sub gData_KeyDown(KeyCode As Integer, Shift As Integer)
    With gData
    Select Case True
        Case KeyCode = vbKeyDelete And Shift <> 0
            If .Row > 0 And .Row < .Rows - 1 And Not mLocked Then
                mDirty = True
                Call gData.RemoveItem
                Call ValidateData
            End If
    
    
        Case KeyCode = vbKeyDelete
            If .Row > 0 And .Row < .Rows - 1 And Not mLocked Then
                mDirty = True
                gData.Text = ""
                Call gData_AfterEdit(gData.Row, gData.Col)
            End If
    
    
    End Select
    End With
End Sub

Private Sub gData_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gData
        If Row = .Rows - 1 Then
            .AddItem ""
            If Row > 1 Then
                .TextMatrix(Row, .ColIndex("Job")) = .TextMatrix(Row - 1, .ColIndex("Job"))
                .TextMatrix(Row, .ColIndex("JobDesc")) = .TextMatrix(Row - 1, .ColIndex("JobDesc"))
                .TextMatrix(Row, .ColIndex("Extra")) = .TextMatrix(Row - 1, .ColIndex("Extra"))
                .TextMatrix(Row, .ColIndex("ExtraDesc")) = .TextMatrix(Row - 1, .ColIndex("ExtraDesc"))
                .TextMatrix(Row, .ColIndex("CostCode")) = .TextMatrix(Row - 1, .ColIndex("CostCode"))
                .TextMatrix(Row, .ColIndex("CostCodeDesc")) = .TextMatrix(Row - 1, .ColIndex("CostCodeDesc"))
                .TextMatrix(Row, .ColIndex("Category")) = .TextMatrix(Row - 1, .ColIndex("Category"))
                .TextMatrix(Row, .ColIndex("CategoryDesc")) = .TextMatrix(Row - 1, .ColIndex("CategoryDesc"))
                .TextMatrix(Row, .ColIndex("Account")) = .TextMatrix(Row - 1, .ColIndex("Account"))
                .TextMatrix(Row, .ColIndex("AccountDesc")) = .TextMatrix(Row - 1, .ColIndex("AccountDesc"))
            End If
        End If
    End With
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim s As String
    Dim batch As Long
    Select Case Button.Key
        
        Case "New"
            If Not SaveData(True) Then Exit Sub
            Call LoadData(0, True)
        
        Case "Open"
            If Not SaveData(True) Then Exit Sub
            s = "Open Batches" & Chr(1) & "select BatchID,Description,BatchType,AccountingDate from journalentries where isnull(postingbatchid,0)=0" & Chr(0) & _
                "Posted Batches" & Chr(1) & "select BatchID,Description,BatchType,AccountingDate,PostingDate from journalentries where postingbatchid<>0"
            If FPickList.Choose(HFApp.Databases(dbHomeFront), "Batch", s, lblBatchNumber, , True) Then
                batch = Val(FPickList.SelectedItem("BatchID"))
                Call LoadData(batch, True)
            End If
                
        Case "Delete"
            If Not mLocked Then
            If MsgBox("Are you sure you want to delete this batch?", vbYesNo + vbExclamation, App.ProductName) = vbYes Then
                Call HFApp.SqlExec("delete journalentrylines where batchid=" & DbQuote(Num, lblBatchNumber), dbHomeFront)
                Call HFApp.SqlExec("delete journalentries where batchid=" & DbQuote(Num, lblBatchNumber), dbHomeFront)
                Call LoadData(0)
            End If
            End If
        
        Case "Save"
            Call SaveData(False)
        
        Case "SaveAs"
            Call HFApp.SqlExec("insert into journalentries(Description,AccountingDate) values('new batch',convert(varchar(25), getdate(),102))")
            lblBatchNumber.Caption = HFApp.SqlIdentity("journalentries")
            mDirty = True
            Call SaveData(False)
            lblPostedDate = ""
            Locked = False
            Call ValidateData
        
        Case "Post"
            Call PostData
                        
        Case "Attachments"
            Call FDocuments.ShowForm("File Attachments - " & txtDescription.Text, "JE~" & lblBatchNumber.Caption, "Documents")
    End Select
End Sub

Private Sub txtComments_GotFocus()
    SelectAll txtComments
End Sub

Private Sub txtDate_Change()
On Error Resume Next
    mDirty = True
    ValidateData
End Sub

Private Sub txtDate_GotFocus()
    SelectAll txtDate
End Sub

Private Sub txtDate_Validate(Cancel As Boolean)
On Error Resume Next
    If txtDate.Text = "" Then Exit Sub
    If Not IsDate(txtDate.Text) Then
        Cancel = True
    Else
        txtDate.Text = Format(txtDate.Text, HFApp.Options(DateFormat))
    End If
End Sub

Private Sub cmdBrowse_Click()
    Call DCalendar.Popup(txtDate)
End Sub

Private Sub txtDescription_Change()
    mDirty = True
End Sub

Private Sub txtDescription_GotFocus()
    SelectAll txtDescription
End Sub

Private Sub txtComments_Change()
    mDirty = True
End Sub



Private Function TimberlineAccounting() As Boolean
    TimberlineAccounting = HFApp.Options(AccountingSystem) = asTimberline
End Function
Public Function AccountingDB() As Connections
    If TimberlineAccounting Then
        AccountingDB = dbAccountingDictionary
    Else
        AccountingDB = dbHomeFront
    End If
End Function
Private Function SelectJob() As String
    Dim s As String
    
    If TimberlineAccounting Then
        s = ""
        s = s & "SELECT Job" & vbCrLf
        s = s & "      ,Jdesc Description" & vbCrLf
        s = s & "      ,jstatus Status" & vbCrLf
        s = s & "  FROM MASTER_JCM_RECORD_1_1" & vbCrLf
        s = s & " WHERE jstatus<>'Closed'"
    Else
        s = ""
        s = s & "SELECT job_no Job" & vbCrLf
        s = s & "      ,Description" & vbCrLf
        s = s & "      ,case isnull(inactive,0) when 1 then 'Closed' else 'In progress' end Status" & vbCrLf
        s = s & "  FROM tbljobs" & vbCrLf
        s = s & " WHERE isnull(inactive,0)=0" & vbCrLf
        
        If IsIn("" & HFApp.Options(AccountingSystem), "" & asSimply, "" & asQuickBooks) Then
            s = s & " and isnull(ExternalJobID,'')<>''" & vbCrLf
        End If
        s = s & " and DivisionID=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
        
    End If
    
    SelectJob = s
    
End Function


Private Function SelectJobExtraCostCode(Optional Job As String = "", Optional Extra As String = "") As String
    Dim s As String
    If TimberlineAccounting Then
        If Job = "" Then
            'standards
            s = ""
            s = s & "SELECT DISTINCT" & vbCrLf
            s = s & "       LTRIM(sphase) CostCode" & vbCrLf
            s = s & "      ,spdesc Description" & vbCrLf
            s = s & "  FROM MASTER_JCM_Record_16" & vbCrLf
            s = s & " WHERE spgphas=0"
        Else
            'job specific
            s = ""
            s = s & "SELECT DISTINCT" & vbCrLf
            s = s & "       LTRIM(phase) CostCode" & vbCrLf
            s = s & "      ,pdesc Description" & vbCrLf
            s = s & "  FROM MASTER_JCM_Record_3" & vbCrLf
            s = s & " WHERE PJob=" & DbQuote(Str, Job) & vbCrLf
            s = s & "   AND pextra=" & DbQuote(Str, Extra) & vbCrLf
            s = s & "   AND pgphase=0"
        End If
    Else
        If Job = "" Then
            'standards
            s = ""
            s = s & "SELECT CostCode" & vbCrLf
            s = s & "      ,isnull(nullif(Description,''),costcode) Description" & vbCrLf
            s = s & "  FROM standardcostcodes" & vbCrLf
            s = s & " WHERE DivisionID=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
        Else
            'job specific
            s = ""
            s = s & "select CostCode" & vbCrLf
            s = s & "      ,isnull(nullif(Description,''),costcode) Description" & vbCrLf
            s = s & "  from jobextracostcodes" & vbCrLf
            s = s & " WHERE job=" & DbQuote(Str, Job) & vbCrLf
            s = s & "   AND extra=" & DbQuote(Str, Extra) & vbCrLf
            s = s & "   and DivisionID=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
        End If
    End If
    SelectJobExtraCostCode = s
End Function
Private Function SelectJobExtraCostCodeCategory(Optional Job As String = "", Optional Extra As String = "", Optional CostCode As String) As String
    Dim s As String
    If TimberlineAccounting Then
        If Job = "" Then
            'standards
            s = ""
            s = s & "SELECT LTRIM(SCAT) Category" & vbCrLf
            s = s & "      ,SCDESC Description" & vbCrLf
            s = s & "  FROM MASTER_JCM_Record_17" & vbCrLf
        Else
            'job specific
            s = ""
            s = s & "SELECT ltrim(cat) Category" & vbCrLf
            s = s & "      ,cdesc Description" & vbCrLf
            s = s & "      ,ctest Estimated" & vbCrLf
            s = s & "      ,ctuest Units" & vbCrLf
            s = s & "      ,crvcom Committed" & vbCrLf
            s = s & "      ,cjtdc CostToDate" & vbCrLf
            s = s & "      ,cjtdu UnitsToDate" & vbCrLf
            s = s & "      ,ctest-cjtdc EstRemaining" & vbCrLf
            s = s & "      ,crvcom-cjtdc PoRemaining" & vbCrLf
            s = s & "  FROM Master_JCM_Record_4" & vbCrLf
            s = s & " where cjob = " & DbQuote(Str, Job) & vbCrLf
            s = s & "   And cextra = " & DbQuote(Str, Extra) & vbCrLf
            s = s & "   And CPHASE = " & DbQuote(Str, CostCode)
        End If
    Else
        If Job = "" Then
            'standards
            s = ""
            s = s & "SELECT Category" & vbCrLf
            s = s & "      ,Description" & vbCrLf
            s = s & "  FROM standardcategories where DivisionID = " & HFApp.DivisionID & vbCrLf
        Else
            'job specific
            s = ""
            s = s & "select Category" & vbCrLf
            s = s & "      ,Description" & vbCrLf
            s = s & "  from jobextracostcodecategories" & vbCrLf
            s = s & " WHERE DivisionID = " & HFApp.DivisionID & " and job=" & DbQuote(Str, Job) & vbCrLf
            s = s & "   AND extra=" & DbQuote(Str, Extra) & vbCrLf
            s = s & "   AND costcode=" & DbQuote(Str, CostCode) & vbCrLf
        End If
    End If
    SelectJobExtraCostCodeCategory = s
End Function

Private Function SelectExtra(Job As String) As String
    Dim s As String
    If TimberlineAccounting Then
        s = ""
        s = s & "SELECT LTRIM(EXTRA) Extra" & vbCrLf
        s = s & "      ,Xdesc Description" & vbCrLf
        s = s & "  FROM MASTER_JCM_Record_2" & vbCrLf
        s = s & " WHERE XJob=" & DbQuote(Str, Job)
    Else
        s = ""
        s = s & "SELECT Extra" & vbCrLf
        s = s & "      ,Description" & vbCrLf
        s = s & "  FROM JobExtras" & vbCrLf
        s = s & " WHERE Job=" & DbQuote(Str, Job) & vbCrLf
        s = s & " and DivisionID=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
    End If
    SelectExtra = s

End Function

Private Function SelectGLAccount() As String
    Dim s As String
    If TimberlineAccounting Then
        s = "select AAcct Account,Atitle Description FROM master_GLm_record_1"
    Else
        s = "SELECT Account ,Description FROM glaccounts where DivisionID=" & DbQuote(Num, HFApp.DivisionID)
    End If
    SelectGLAccount = s
End Function


Private Function ValidateField(Grid As VSFlexGrid, ByRef EditText As String, WarningMsg As String, sql As String, ParamArray OtherColumns())
On Error GoTo eh
    
    Dim rs As Recordset
    Dim i As Long
    
    With Grid
    If EditText = "" Then
        ValidateField = True
        For i = 0 To UBound(OtherColumns)
            .Cell(flexcpText, .Row, .ColIndex(OtherColumns(i)), .RowSel, .ColIndex(OtherColumns(i))) = ""
        Next
    Else
        Set rs = HFApp.SqlExec(sql)
        EditText = "" & rs(0)
        For i = 0 To UBound(OtherColumns)
            .Cell(flexcpText, .Row, .ColIndex(OtherColumns(i)), .RowSel, .ColIndex(OtherColumns(i))) = "" & rs(i + 1)
        Next
    End If
    End With
    ValidateField = True
    
Exit Function
eh: If WarningMsg <> "" Then MsgBox WarningMsg, vbExclamation, App.ProductName
    ValidateField = False
End Function


Private Sub LoadTypes()
    Dim rs As Recordset
    Set rs = HFApp.SqlExec("select distinct isnull(batchtype,'') from journalentries order by 1", dbHomeFront)
    With cboBatchType
        .Clear
        While Not rs.EOF
            .AddItem "" & rs(0)
            rs.MoveNext
        Wend
    End With
End Sub












Private Sub PostData()
    If Not SaveData(False) Then Exit Sub
    Select Case HFApp.Options(AccountingSystem)
        Case asTimberline:   Call PostToTimberline(Val(lblBatchNumber))
        Case asSimply:       Call PostToSimply(Val(lblBatchNumber))
        Case asQuickBooks:   Call PostToQuickbooks(Val(lblBatchNumber))
        Case asMasterBuilder:
    End Select
End Sub

Private Sub PostToQuickbooks(BatchID As Long)
On Error GoTo eh
    Dim i As Long
    Dim s   As String
    Dim rs  As Recordset
    
    Screen.MousePointer = vbHourglass
    
    s = ""
    s = s & "select b.BatchType,b.description BatchDesc,b.AccountingDate,a.externalid AccountID,j.externaljobid,l.* " & vbCrLf
    s = s & "from journalentries b" & vbCrLf
    s = s & "join journalentrylines l on (b.batchid=l.batchid)" & vbCrLf
    s = s & "left outer join glaccounts a on (l.account=a.account)" & vbCrLf
    s = s & "left outer join tbljobs j on (l.job=j.job_no)" & vbCrLf
    s = s & "where b.batchid=" & DbQuote(Num, BatchID) & vbCrLf
    s = s & "order by l.sequence" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    If Not rs.EOF Then
        s = HFApp.XmlQBStart()
        s = s & "<JournalEntryAddRq>" & vbCrLf
        s = s & "<JournalEntryAdd>" & vbCrLf
        s = s & HFApp.XmlQBAdd(d, 0, "TxnDate", "" & rs("AccountingDate"))
        s = s & HFApp.XmlQBAdd(m, 0, "RefNumber", "" & rs("BatchID"))
        s = s & HFApp.XmlQBAdd(m, 0, "IsAdjustment", "1")
        While Not rs.EOF
            If Val("" & rs("CreditAmount")) = 0 Then
                s = s & "<JournalDebitLine>" & vbCrLf
                s = s & HFApp.XmlQBAdd(m, 0, "AccountRef", "" & rs("AccountID"))
                s = s & HFApp.XmlQBAdd(m, 0, "Amount", Format(Val("" & rs("DebitAmount")), "###.00"))
                s = s & HFApp.XmlQBAdd(m, 0, "Memo", "" & rs("Description"))
                
                If "" & rs("ExternalJobID") <> "" Then s = s & HFApp.XmlQBAdd(m, 0, "EntityRef", "" & rs("ExternalJobID"))
                If "" & rs("Category") <> "" Then s = s & HFApp.XmlQBAdd(m, 0, "ClassRef", "" & rs("Category"))
                
                s = s & "</JournalDebitLine>" & vbCrLf
            Else
                s = s & "<JournalCreditLine>" & vbCrLf
                s = s & HFApp.XmlQBAdd(m, 0, "AccountRef", "" & rs("AccountID"))
                s = s & HFApp.XmlQBAdd(m, 0, "Amount", Format(Val("" & rs("CreditAmount")), "###.00"))
                s = s & HFApp.XmlQBAdd(m, 0, "Memo", "" & rs("Description"))
                
                If "" & rs("ExternalJobID") <> "" Then s = s & HFApp.XmlQBAdd(m, 0, "EntityRef", "" & rs("ExternalJobID"))
                If "" & rs("Category") <> "" Then s = s & HFApp.XmlQBAdd(m, 0, "ClassRef", "" & rs("Category"))
                
                s = s & "</JournalCreditLine>" & vbCrLf
            End If
            rs.MoveNext
        Wend
        s = s & "</JournalEntryAdd>" & vbCrLf
        s = s & "</JournalEntryAddRq>" & vbCrLf
        s = s & HFApp.XmlQBEnd()
        Call HFApp.XmlQBSubmit(s)
    End If
    
    s = "update journalentries set postingdate=getdate(),postingbatchid=batchid where batchid=" & DbQuote(Num, BatchID)
    Call HFApp.SqlExec(s, dbHomeFront)
    Call LoadData(0, True)
    
    Screen.MousePointer = vbDefault
    
    
Exit Sub
eh: Call errHandler(SRCFILE & "PostToQuickbooks")
End Sub


Private Sub PostToSimply(BatchID As Long)
On Error GoTo eh
    Dim es As String
    Dim i As Long
    Dim s   As String
    Dim rs  As Recordset
    Dim X   As SimplyWrapper.JournalBatch
    Dim eNumb As Long
    Dim eDesc As String
    Dim eSrc As String
    
    Screen.MousePointer = vbHourglass
    
es = "opening simply"
    Set X = New SimplyWrapper.JournalBatch
    If Not X.OpenDB(HFApp.Options(SimplyDataFile), HFApp.Options.ValueByName("SimplyUID"), HFApp.Options.ValueByName("SimplyPWD")) Then
        Screen.MousePointer = vbDefault
        MsgBox "Unable to connect to Simply Accounting" & vbCrLf & vbCrLf & "Check that the Simply Accounting security context and password in system settings is correct.", vbCritical, App.ProductName
        Exit Sub
    End If
    
    'query data
es = "querying homefront db"
    s = ""
    s = s & "select j.BatchType,j.description BatchDesc,j.AccountingDate,l.* ,x.ExternalJobID" & vbCrLf
    s = s & "from journalentries j" & vbCrLf
    s = s & "join journalentrylines l on (j.batchid=l.batchid)" & vbCrLf
    s = s & "left outer join tbljobs x on l.job=x.job_no" & vbCrLf
    s = s & "where j.batchid=" & DbQuote(Num, BatchID) & vbCrLf
    s = s & "order by l.sequence" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    
    If Not rs.EOF Then
es = "creating batch"
        Call X.CreateBatch("" & rs("BatchType"), "" & rs("BatchDesc"), "" & rs("AccountingDate"))
        i = 0
        While Not rs.EOF
            i = i + 1
es = "adding line " & i
    
            On Error Resume Next
            Call X.AddLine1("" & rs("Description"), "" & rs("Account"), Val("" & rs("DebitAmount")), Val("" & rs("CreditAmount")), "" & rs("ExternalJobID"))
            eNumb = Err.Number
            eDesc = Err.Description
            eSrc = Err.Source
            On Error GoTo eh
            
            Select Case eNumb
                Case 0    'no error
                Case 430  'doesnt support interface -- client using old version of simplywrapper. use old syntax
                    Call X.AddLine("" & rs("Description"), "" & rs("Account"), Val("" & rs("DebitAmount")), Val("" & rs("CreditAmount")))
                Case Else 're-raise any other error
                    Call Err.Raise(eNumb, eSrc, eDesc)
            End Select
            
            rs.MoveNext
            
        Wend
        
es = "saving batch"
        Call X.SaveBatch
        
es = "closing simply"
        Call X.CloseDB
        
    End If
    
    
    s = "update journalentries set postingdate=getdate(),postingbatchid=batchid where batchid=" & DbQuote(Num, BatchID)
    Call HFApp.SqlExec(s, dbHomeFront)
    Call LoadData(0, True)
    
    Screen.MousePointer = vbDefault
    
    
Exit Sub
eh:
Screen.MousePointer = vbDefault
Select Case True
    
    'ignorable crap -----------------------------------------------------
    Case Err.Description = "Your chequing account is overdrawn. See the Advice topic ""Managing Your Cash Flow"" for suggestions."
        Resume Next
        
    'everything else ----------------------------------------------------
    Case Else
        Call errHandler(SRCFILE & "PostToSimply", es)
End Select
End Sub


Private Sub PostToTimberline(BatchID As Long)
On Error GoTo eh

Const MacroFile = "HomeFrontMacros\JournalEntries.mac"
Const TxtFile = "HomeFrontMacros\JournalEntries.txt"
Const RejectFile = "HomeFrontMacros\JournalEntriesReject.txt"
Const PrintFile = "HomeFrontMacros\JournalEntries.prn"
    
    
    Dim i   As Long
    Dim s   As String
    Dim rs  As Recordset
    
    Screen.MousePointer = vbHourglass
    
    'open the files
    Call CreatePath("export path", FilePath(PathAppend(HFApp.Options(Timberline_Data_Path), TxtFile)))
    i = FreeFile: Open PathAppend(HFApp.Options(Timberline_Data_Path), TxtFile) For Output As #i
    
    'query data
    s = ""
    s = s & "select j.BatchType,j.description BatchDesc,j.AccountingDate,l.* " & vbCrLf
    s = s & "from journalentries j" & vbCrLf
    s = s & "join journalentrylines l on (j.batchid=l.batchid)" & vbCrLf
    s = s & "where j.batchid=" & DbQuote(Num, BatchID) & vbCrLf
    s = s & "order by l.sequence" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    
    'write file
    While Not rs.EOF
        If "" & rs("Job") = "" Then
            Print #i, "GL," & _
                      Quote(Format("" & rs("AccountingDate"), "mm\/dd\/yyyy")) & "," & _
                      Quote(Format("" & rs("AccountingDate"), "mm\/dd\/yyyy")) & "," & _
                      Quote("" & rs("description")) & "," & _
                      Quote(Val("" & rs("DebitAmount")) + Val("" & rs("CreditAmount"))) & "," & _
                      IIf(Val("" & rs("DebitAmount")) <> 0, "" & rs("Account"), "") & "," & _
                      IIf(Val("" & rs("CreditAmount")) <> 0, "" & rs("Account"), "")
        Else
            Print #i, "DC," & _
                      Quote("" & rs("Job")) & "," & _
                      Quote("" & rs("Extra")) & "," & _
                      Quote("" & rs("CostCode")) & "," & _
                      Quote("" & rs("Category")) & "," & _
                      "2," & _
                      Quote(Format("" & rs("AccountingDate"), "mm\/dd\/yyyy")) & "," & _
                      Quote(Format("" & rs("AccountingDate"), "mm\/dd\/yyyy")) & "," & _
                      Quote("" & rs("description")) & "," & _
                      "," & _
                      "," & _
                      Quote(Val("" & rs("DebitAmount")) - Val("" & rs("CreditAmount"))) & "," & _
                      "" & rs("Account")
        
        
        
'                      IIf(Val("" & rs("DebitAmount")) <> 0, "" & rs("Account"), "") & "," & _
'                      IIf(Val("" & rs("CreditAmount")) <> 0, "" & rs("Account"), "")
        End If
        rs.MoveNext
    Wend
    Close i

    
    'launch macro
    Call FTSObject.Run(HFApp.Options(Timberline_Data_Path), _
                       HFApp.Options(Timberline_UID), _
                       HFApp.Options(Timberline_PWD), _
                       PathAppend(HFApp.Options(Timberline_Data_Path), MacroFile), _
                       PathAppend(HFApp.Options(Timberline_Data_Path), PrintFile), _
                       "Processing Journal Entries...")
    
    
    'mark these as posted
    s = "update journalentries set postingdate=getdate(),postingbatchid=batchid where batchid=" & DbQuote(Num, BatchID)
    Call HFApp.SqlExec(s, dbHomeFront)
    Call LoadData(0)
    
    'show processing journal
    s = TempFile(FileExt(PrintFile))
    Call FileCopy(PathAppend(HFApp.Options(Timberline_Data_Path), PrintFile), s)
    Call FTSObject.Run(HFApp.Options(Timberline_Data_Path), HFApp.Options(Timberline_UID), HFApp.Options(Timberline_PWD), s)
    
    
    Screen.MousePointer = vbDefault
    
    
Exit Sub
eh: Call errHandler(SRCFILE & "PostToTimberline")
End Sub




Private Property Let Locked(RHS As Boolean)
    mLocked = RHS
    txtDescription.Locked = mLocked
    cboBatchType.Enabled = Not mLocked
    txtDate.Locked = mLocked
    cmdBrowse.Enabled = Not mLocked
    txtComments.Locked = mLocked
    gData.Editable = IIf(mLocked, flexEDNone, flexEDKbdMouse)
End Property


Private Function ValidateJob(Job As String) As String
    Dim s As String
    
    If TimberlineAccounting Then
        s = ""
        s = s & "SELECT job" & vbCrLf
        s = s & "      ,jdesc description" & vbCrLf
        s = s & "  FROM MASTER_JCM_Record_1_1" & vbCrLf
        s = s & " WHERE job=" & DbQuote(Str, HFApp.FormatJob(Job)) & vbCrLf
    Else
        s = ""
        s = s & "SELECT job_no job" & vbCrLf
        s = s & "      ,description" & vbCrLf
        s = s & "  FROM tbljobs" & vbCrLf
        s = s & " WHERE job_no=" & DbQuote(Str, Job) & vbCrLf
    End If
    
    ValidateJob = s
    
End Function


Private Function ValidateAccount(Account As String) As String
    Dim s As String
    If TimberlineAccounting Then
        s = "select AAcct account,Atitle description FROM master_GLm_record_1 where aacct=" & DbQuote(Str, FormatAccount(Account))
    Else
        s = "select account,description FROM glaccounts where account=" & DbQuote(Str, Account)
    End If
    ValidateAccount = s
End Function



Public Sub GetFormatMasks()
    Dim rs As Recordset
    
    Dim s1 As Long
    Dim s2 As Long
    Dim s3 As Long
    Dim s4 As Long
    Dim s5 As Long
    Dim sep As String
    Dim m As String

    If HFApp.Options(AccountingSystem) = asTimberline And HFApp.Databases(dbAccountingDictionary).State = adStateOpen Then
        Set rs = HFApp.SqlExec("select section,cslen,cnpunct from ts_ctl_record_3 where section>0 and fldcode=9", dbAccountingDictionary)
        If Not rs.EOF Then
            While Not rs.EOF
                Select Case Val("" & rs("section"))
                    Case 1: s1 = Val("" & rs("cslen")):    If "" & rs("cnpunct") <> "" Then sep = "" & rs("cnpunct")
                    Case 2: s2 = Val("" & rs("cslen")):    If "" & rs("cnpunct") <> "" Then sep = "" & rs("cnpunct")
                    Case 3: s3 = Val("" & rs("cslen")):    If "" & rs("cnpunct") <> "" Then sep = "" & rs("cnpunct")
                    Case 4: s4 = Val("" & rs("cslen")):    If "" & rs("cnpunct") <> "" Then sep = "" & rs("cnpunct")
                    Case 5: s5 = Val("" & rs("cslen"))   ' this is the suffix
                End Select
                rs.MoveNext
            Wend
            m = String(s1, "&") & sep & String(s2, "&") & sep & String(s3, "&") & sep & String(s4, "&")
            If sep <> "" Then
                While Right(m, 1) = sep
                    m = left(m, Len(m) - 1)
                Wend
            End If
            If s1 = 0 Then
                m = ""
            End If
            
            m = Replace(m, "----", "-")
            m = Replace(m, "---", "-")
            m = Replace(m, "--", "-")
            
            If s5 <> 0 Then
                m = m & "." & String(s5, "&&")
            End If
            
            AccountMask = m
        End If
    End If

End Sub

Public Function FormatAccount(Account As String) As String
    If TimberlineAccounting Then
        FormatAccount = Format(StripFormating(Trim(Account)), AccountMask)
    Else
        FormatAccount = Account
    End If
End Function

Private Function StripFormating(s As String) As String
    Const FORMATCHRS = "+=_-)(*&^%$#@!~`[]{}\|/?'""<>;:,"
    Dim i As Long
    For i = 1 To Len(FORMATCHRS)
        s = Replace(s, Mid(FORMATCHRS, i, 1), "")
    Next
    StripFormating = s
End Function


Public Function ValidateCostCode(Job As String, Extra As String, Phase As String) As String
    Dim s As String
    
    
    If TimberlineAccounting Then
        s = ""
        s = s & "select 1,phase costcode,pdesc description,'' DebitAccount,'' DebitAccountDesc" & vbCrLf
        s = s & "  from master_jcm_record_3 " & vbCrLf
        s = s & " where pjob=" & DbQuote(Str, HFApp.FormatJob(Job)) & vbCrLf
        s = s & "   and pextra=" & DbQuote(Str, Extra) & vbCrLf
        s = s & "   and phase=" & DbQuote(Str, HFApp.FormatCostCode(Phase)) & vbCrLf
        s = s & "union all" & vbCrLf
        s = s & "select 2,sphase costcode,spdesc description,'' DebitAccount,'' DebitAccountDesc" & vbCrLf
        s = s & "  from master_jcm_record_16" & vbCrLf
        s = s & " where sphase=" & DbQuote(Str, HFApp.FormatCostCode(Phase)) & vbCrLf
        s = s & "order by 1" & vbCrLf
    Else
        s = ""
        s = s & "select 1,c.costcode,c.description,a.account debitaccount,a.description debitaccountdesc" & vbCrLf
        s = s & "  from JobExtraCostCodes c" & vbCrLf
        s = s & "  left outer join standardcostcodes s on(c.DivisionID = s.DivisionID and c.costcode=s.costcode)"
        s = s & "  left outer join glaccounts a on(s.DivisionID = a.DivisionID and s.debitaccount=a.account)"
        s = s & " where c.job=" & DbQuote(Str, Job) & vbCrLf
        s = s & "   and c.extra=" & DbQuote(Str, Extra) & vbCrLf
        s = s & "   and c.costcode=" & DbQuote(Str, Phase) & vbCrLf
        s = s & "   and c.DivisionID = " & HFApp.DivisionID & vbCrLf
        s = s & "union all" & vbCrLf
        s = s & "select 2,s.costcode,s.description,a.account debitaccount,a.description debitaccountdesc" & vbCrLf
        s = s & "  from standardcostcodes s" & vbCrLf
        s = s & "  left outer join glaccounts a on(s.DivisionID = a.DivisionID and s.debitaccount=a.account)"
        s = s & " where s.DivisionID = " & HFApp.DivisionID & " and s.costcode=" & DbQuote(Str, Phase) & vbCrLf
    End If
    ValidateCostCode = s
End Function
Public Function ValidateCategory(Job As String, Extra As String, Phase As String, Category As String) As String
    Dim s As String
    If TimberlineAccounting Then
        s = ""
        s = s & "select 1,cat Category ,cdesc Description" & vbCrLf
        s = s & "  from master_jcm_record_4 " & vbCrLf
        s = s & " where cjob=" & DbQuote(Str, HFApp.FormatJob(Job)) & vbCrLf
        s = s & "   and cextra=" & DbQuote(Str, Extra) & vbCrLf
        s = s & "   and cphase=" & DbQuote(Str, HFApp.FormatCostCode(Phase)) & vbCrLf
        s = s & "   And cat=" & DbQuote(Str, Category) & vbCrLf
        s = s & "union all" & vbCrLf
        s = s & "select 2,scat Category ,scdesc Description" & vbCrLf
        s = s & "  from master_jcm_record_17" & vbCrLf
        s = s & " where scat=" & DbQuote(Str, Category) & vbCrLf
        s = s & "order by 1" & vbCrLf
    Else
        s = ""
        s = s & "select 1,category,description" & vbCrLf
        s = s & "  from JobExtraCostCodeCategories " & vbCrLf
        s = s & " where DivisionID = " & HFApp.DivisionID & " and job=" & DbQuote(Str, Job) & vbCrLf
        s = s & "   and extra=" & DbQuote(Str, Extra) & vbCrLf
        s = s & "   and costcode=" & DbQuote(Str, Phase) & vbCrLf
        s = s & "   and category=" & DbQuote(Str, Category) & vbCrLf
        s = s & "union all" & vbCrLf
        s = s & "select 2,category,description" & vbCrLf
        s = s & "  from standardcategories " & vbCrLf
        s = s & " where DivisionID = " & HFApp.DivisionID & " and category=" & DbQuote(Str, Category) & vbCrLf
    End If
    ValidateCategory = s
End Function

