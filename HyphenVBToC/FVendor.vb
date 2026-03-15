VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FVendor 
   Caption         =   "Vendor"
   ClientHeight    =   10620
   ClientLeft      =   2055
   ClientTop       =   1395
   ClientWidth     =   16395
   Icon            =   "FVendor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10620
   ScaleWidth      =   16395
   Begin VB.TextBox txtTaxID 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1455
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5280
      Width           =   2835
   End
   Begin VB.TextBox txtWalletPayeeID 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1455
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   "{FFB0F0BB-57B9-486F-8654-160314660B36}"
      Top             =   4995
      Width           =   2835
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "Next >>"
      Height          =   510
      Index           =   1
      Left            =   12180
      TabIndex        =   23
      Top             =   9555
      Width           =   1185
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "<< Previous"
      Height          =   510
      Index           =   0
      Left            =   10890
      TabIndex        =   22
      Top             =   9540
      Width           =   1185
   End
   Begin VB.CheckBox chkBuildPro 
      Caption         =   "BuildPro Vendor"
      Height          =   315
      Left            =   1485
      TabIndex        =   3
      Top             =   1785
      Width           =   2700
   End
   Begin VB.CheckBox chkIsTBD 
      Caption         =   "TBD - budget planning only"
      Height          =   315
      Left            =   1485
      TabIndex        =   2
      Top             =   1515
      Width           =   2700
   End
   Begin VB.TextBox txtWebPWD 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1455
      MaxLength       =   20
      TabIndex        =   17
      Top             =   6930
      Width           =   1755
   End
   Begin VB.TextBox txtWebUID 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1455
      MaxLength       =   40
      TabIndex        =   16
      Top             =   6660
      Width           =   1755
   End
   Begin VB.TextBox txtHoldbackPercent 
      BorderStyle     =   0  'None
      Height          =   230
      Left            =   1455
      TabIndex        =   11
      Top             =   4500
      Width           =   405
   End
   Begin VB.TextBox txtVendor 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   230
      Left            =   1470
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   " "
      Top             =   1020
      Width           =   2835
   End
   Begin VB.TextBox txtAddress 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   690
      Left            =   1470
      Locked          =   -1  'True
      MaxLength       =   75
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   2130
      Width           =   2835
   End
   Begin VB.TextBox txtCity 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   230
      Left            =   1470
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   5
      Top             =   2835
      Width           =   2835
   End
   Begin VB.TextBox txtProvince 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   230
      Left            =   1470
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   6
      Top             =   3075
      Width           =   2835
   End
   Begin VB.TextBox txtPhone 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   230
      Left            =   1470
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   8
      Top             =   3555
      Width           =   1755
   End
   Begin VB.TextBox txtFax 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   230
      Left            =   1470
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   9
      Top             =   3795
      Width           =   1755
   End
   Begin VB.TextBox txtPostal 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   230
      Left            =   1470
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   7
      Top             =   3315
      Width           =   1395
   End
   Begin VB.TextBox txtDescription 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   230
      Left            =   1470
      Locked          =   -1  'True
      MaxLength       =   60
      TabIndex        =   1
      Top             =   1260
      Width           =   2835
   End
   Begin VSFlex8Ctl.VSFlexGrid gContacts 
      Height          =   2625
      Left            =   4470
      TabIndex        =   18
      Top             =   1140
      Width           =   11295
      _cx             =   1986416083
      _cy             =   1986400790
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
      Cols            =   13
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FVendor.frx":000C
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
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   27
      Top             =   0
      Width           =   16395
      _ExtentX        =   28919
      _ExtentY        =   1058
      ButtonWidth     =   1508
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "   Open    "
            Key             =   "Open"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "Save"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Import"
            Key             =   "ExcelImport"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Export"
            Key             =   "ExcelExport"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList LargeIcons 
      Left            =   240
      Top             =   2010
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
            Picture         =   "FVendor.frx":01F9
            Key             =   "EditAssembly"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":0AD3
            Key             =   "RFP"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":13AD
            Key             =   "CreateJob"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":1C87
            Key             =   "takeoffsettings"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":2561
            Key             =   "quote"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":2E3B
            Key             =   "DecreaseDecimals"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":3715
            Key             =   "IncreaseDecimals"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":3FEF
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":48C9
            Key             =   "ExcelImport"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":51A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":5A7D
            Key             =   "UpdatePrices"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":6357
            Key             =   "ExcelExport"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":6C31
            Key             =   "Publish"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":750B
            Key             =   "Forecast"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":7DE5
            Key             =   "ViewPOs"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":86BF
            Key             =   "ViewBudgets"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":8F99
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":9873
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":A14D
            Key             =   "Send"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":AA27
            Key             =   "TakeoffOneTime"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":B301
            Key             =   "Estimate"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":BBDB
            Key             =   "TakeoffAssembly"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":C4B5
            Key             =   "NewAssembly"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":CD8F
            Key             =   "New"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":D669
            Key             =   "TakeoffItem"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":DF43
            Key             =   "TakeoffCustom"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":E81D
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":F0F7
            Key             =   "SaveAs"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":F9D1
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":102AB
            Key             =   "AddPricelist"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":10B85
            Key             =   "RePrice"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":1145F
            Key             =   "Pricebook"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":11D39
            Key             =   "PricebookEdit"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":12613
            Key             =   "PricelistExport"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":12EED
            Key             =   "PricelistImport"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":137C7
            Key             =   "NewPricelist"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":140A1
            Key             =   "View"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":1497B
            Key             =   "Vendor1"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":15255
            Key             =   "Vendor"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":15B2F
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":16409
            Key             =   "Attachments"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":16CE3
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":175BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":17E97
            Key             =   "Design Center Options"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":18771
            Key             =   "Global Options"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":1904B
            Key             =   "Models"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":19925
            Key             =   "Options"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":1A1FF
            Key             =   "Generate"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":1AAD9
            Key             =   "Timberline"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":1B3B3
            Key             =   "MBImport"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":1BC8D
            Key             =   "EEEstimating"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":1C567
            Key             =   "EEExport"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":1CE41
            Key             =   "EEImport"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":1D71B
            Key             =   "MBExport"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FVendor.frx":1DFF5
            Key             =   "MasterBuilder"
         EndProperty
      EndProperty
   End
   Begin HFSystem.VBCombo cboStatus 
      Height          =   240
      Left            =   1455
      TabIndex        =   10
      Top             =   4245
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   423
      Style           =   2
   End
   Begin VSFlex8Ctl.VSFlexGrid gTaxGroups 
      Height          =   1455
      Left            =   4440
      TabIndex        =   20
      Top             =   5835
      Width           =   7005
      _cx             =   1986408516
      _cy             =   1986398726
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
      Rows            =   4
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FVendor.frx":1ECCF
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
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
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
   End
   Begin HFSystem.VBCombo cboPOFormat 
      Height          =   240
      Left            =   1455
      TabIndex        =   15
      Top             =   5820
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   423
   End
   Begin HFSystem.VBCombo cboTradeType 
      Height          =   240
      Left            =   1455
      TabIndex        =   14
      Top             =   5565
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   423
   End
   Begin VSFlex8Ctl.VSFlexGrid gInsurance 
      Height          =   1455
      Left            =   4470
      TabIndex        =   21
      Top             =   7740
      Width           =   7005
      _cx             =   1986408516
      _cy             =   1986398726
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
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FVendor.frx":1EE19
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
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
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
   End
   Begin VSFlex8Ctl.VSFlexGrid gPayPoints 
      Height          =   1335
      Left            =   4470
      TabIndex        =   19
      Top             =   4140
      Width           =   8955
      _cx             =   1986411956
      _cy             =   1986398515
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
      Rows            =   4
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FVendor.frx":1EF1D
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
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
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
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Tax ID"
      Height          =   195
      Left            =   900
      TabIndex        =   45
      Top             =   5310
      Width           =   480
   End
   Begin VB.Label lblWalletPayeeID 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Payee ID"
      Height          =   270
      Left            =   720
      TabIndex        =   44
      Top             =   5025
      Width           =   660
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Contract Payment Points"
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
      Left            =   4470
      TabIndex        =   43
      Top             =   3870
      Width           =   2100
   End
   Begin VB.Label Label141 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "State/Province"
      Height          =   195
      Left            =   285
      TabIndex        =   42
      Top             =   3060
      Width           =   1080
   End
   Begin VB.Label Label131 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Web Portal"
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
      Height          =   225
      Left            =   435
      TabIndex        =   41
      Top             =   6405
      Width           =   960
   End
   Begin VB.Label Label411 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   225
      Left            =   675
      TabIndex        =   40
      Top             =   6900
      Width           =   690
   End
   Begin VB.Label Label55 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Login ID"
      Height          =   225
      Left            =   765
      TabIndex        =   39
      Top             =   6660
      Width           =   600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Insurance"
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
      Left            =   4470
      TabIndex        =   38
      Top             =   7470
      Width           =   855
   End
   Begin VB.Label Label1111 
      Alignment       =   1  'Right Justify
      Caption         =   "Retainage %"
      Height          =   255
      Left            =   -30
      TabIndex        =   37
      Top             =   4530
      Width           =   1395
   End
   Begin VB.Label Label1123 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "PO Format"
      Height          =   270
      Left            =   615
      TabIndex        =   36
      Top             =   5865
      Width           =   750
   End
   Begin VB.Label Label113 
      Alignment       =   1  'Right Justify
      Caption         =   "Trade Type"
      Height          =   270
      Left            =   -30
      TabIndex        =   35
      Top             =   5565
      Width           =   1395
   End
   Begin VB.Label Label103 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Status"
      Height          =   195
      Left            =   915
      TabIndex        =   34
      Top             =   4260
      Width           =   450
   End
   Begin VB.Label lblProperties 
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
      Left            =   4470
      TabIndex        =   33
      Top             =   5580
      Width           =   990
   End
   Begin VB.Label Label140 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "City"
      Height          =   195
      Left            =   1110
      TabIndex        =   32
      Top             =   2850
      Width           =   255
   End
   Begin VB.Label Label50 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Phone"
      Height          =   195
      Left            =   900
      TabIndex        =   31
      Top             =   3555
      Width           =   465
   End
   Begin VB.Label Label40 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Fax"
      Height          =   195
      Left            =   1110
      TabIndex        =   30
      Top             =   3795
      Width           =   255
   End
   Begin VB.Label Label1333 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Address"
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
      Left            =   675
      TabIndex        =   29
      Top             =   2145
      Width           =   690
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Postal"
      Height          =   195
      Left            =   930
      TabIndex        =   28
      Top             =   3345
      Width           =   435
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
      Left            =   4500
      TabIndex        =   26
      Top             =   900
      Width           =   765
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Company Name"
      Height          =   255
      Left            =   -30
      TabIndex        =   25
      Top             =   1290
      Width           =   1395
   End
   Begin VB.Label Label111 
      Alignment       =   1  'Right Justify
      Caption         =   "Vendor"
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
      Height          =   255
      Left            =   -30
      TabIndex        =   24
      Top             =   1050
      Width           =   1395
   End
End
Attribute VB_Name = "FVendor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Public EventTraps As Collection
Const SRCFILE = "FVendor::"

Private mDirty    As Boolean
Private mVendor      As String

Private Enum InsuranceEnum
    GeneralRow = 1
    WorkersRow
    UmbrellaRow
    AutomobileRow
End Enum


Private Sub cboDebitAccount_Click()
    mDirty = True
End Sub


Private Sub cboPOFormat_Change()
    mDirty = True
End Sub

Private Sub cboPOFormat_Click()
    mDirty = True
End Sub

Private Sub cboStatus_Click()
    mDirty = True
End Sub

Private Sub cboTradeType_Change()
    mDirty = True
End Sub

Private Sub cboTradeType_Click()
    mDirty = True
End Sub






Private Sub chkBuildPro_Click()
    mDirty = True
End Sub

Private Sub chkIsTBD_Click()
    mDirty = True
End Sub



Private Sub cmdNav_Click(Index As Integer)
On Error GoTo eh
    Dim s As String
    
    If Not SaveData(True) Then Exit Sub
    
    s = "select vendor_id from tblvendors where divisionid=" & DbQuote(Num, HFApp.DivisionID) & " and vendor_id" & IIf(Index = 1, " > ", " < ") & DbQuote(Str, mVendor) & " order by vendor_id " & IIf(Index = 1, " asc ", " desc ")
    mVendor = HFApp.SqlExec(s)(0)
    
    Call LoadData
eh: Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case True
        Case Shift = vbCtrlMask And KeyCode = vbKeyO:  Call Toolbar_ButtonClick(Toolbar.Buttons("Open"))
        Case Shift = vbCtrlMask And KeyCode = vbKeyS:  Call Toolbar_ButtonClick(Toolbar.Buttons("Save"))
    End Select
End Sub

Private Sub Form_Load()
    Dim s As String
    
    ReadOnly = True
    Call SetToolbarIcons(Toolbar, LargeIcons)
    Call IniGetForm(Me)
    Call IniGetGrid(Me, gContacts, , , , True)
    Call IniGetGrid(Me, gInsurance)
    Call IniGetGrid(Me, gPayPoints)
    
    With gContacts
    
        s = .BuildComboList(HFApp.SqlExec("select address,name from smscarriers order by 2"), "*name,address", "address")
        .ColComboList(.ColIndex("SMSAddress")) = "|" & s
        .ColComboList(.ColIndex("Role")) = HFApp.Options.ValueByName("VendorContactRoles")
        .ColComboList(.ColIndex("SendVia")) = "#1;Print|#3;Email|#4;Fax"
    End With
    
    cboStatus.Clear
    cboStatus.AddItem "Active"
    cboStatus.AddItem "Inactive"
        
    
    s = Dir(PathAppend(HFApp.SystemFolder, "Estimating\PO Formats", "*.rpt"), , True)
    While s <> ""
        cboPOFormat.AddItem StripExtension(s)
        s = Dir
    Wend
    
    
    
    Call LoadComboBox(Me.cboTradeType, HFApp.Databases(dbHomefront), "select distinct isnull(VendorGroupID,''),'',0 from tblvendors where divisionid=" & HFApp.DivisionID & " order by 1")
    
    
    If HFApp.Options(AccountingSystem) = AccountingSystems.asNone Then
        txtDescription.Enabled = True
        txtAddress.Enabled = True
        txtCity.Enabled = True
        txtProvince.Enabled = True
        txtPostal.Enabled = True
        txtFax.Enabled = True
    End If
    
    
    
    If mVendor = "" Then
        mDirty = False
        Call Toolbar_ButtonClick(Toolbar.Buttons("Open"))
        'Line below Commented out Nov 9 2011
        'If mVendor = "" Then Unload Me
    Else
        Call LoadData
    End If
End Sub



Private Sub LoadData()
    Dim s As String
    Dim rs As Recordset
    
    
    s = ""
    s = s & "select wp.walletid WalletPayeeID,v.* " & vbCrLf
    s = s & "from tblVendors v" & vbCrLf
    s = s & "join divisions d on v.divisionid=d.divisionid" & vbCrLf
    s = s & "left join datasources ds on d.bookofaccount=ds.bookofaccount" & vbCrLf
    s = s & "left join wallet_payees wp on wp.datasourceid=ds.datasourceid and wp.payeeid=v.vendor_id" & vbCrLf
    s = s & "where v.Vendor_ID=" & DbQuote(Str, mVendor) & vbCrLf
    s = s & "and v.divisionid=" & HFApp.DivisionID & vbCrLf

    
    Set rs = HFApp.SqlExec(s)
    If rs.EOF Then
        ReadOnly = True
        txtVendor.Text = ""
        txtDescription.Text = ""
        txtAddress.Text = ""
        txtCity.Text = ""
        txtProvince.Text = ""
        txtPostal.Text = ""
        txtPhone.Text = ""
        txtFax.Text = ""
        chkIsTBD.Value = vbUnchecked
        chkBuildPro.Value = vbChecked
        
        txtWebUID.Text = ""
        txtWebPWD.Text = ""
        txtWalletPayeeID.Text = ""
        txtTaxID.Text = ""
        
        Call SetCtrlFocus(txtVendor)
        cboStatus.ListIndex = 0
        
        cboTradeType.Text = ""
        cboPOFormat.ListIndex = -1
        txtHoldbackPercent = ""
        
        gContacts.Rows = 1
        gInsurance.Cell(flexcpText, 1, 1, gInsurance.Rows - 1, gInsurance.Cols - 1) = ""
        gInsurance.Editable = flexEDNone
    
    Else
        ReadOnly = False
        txtVendor.Text = "" & rs("Vendor_ID")
        txtDescription.Text = "" & rs("Vendor_Name")
        txtAddress.Text = "" & rs("Addr1") & vbCrLf & rs("Addr2")
        txtCity.Text = "" & rs("city")
        txtProvince.Text = "" & rs("state")
        txtPostal.Text = "" & rs("zip")
        txtPhone.Text = "" & rs("phone")
        txtFax.Text = "" & rs("fax")
        chkIsTBD.Value = IIf("" & rs("isTBD") = "true", vbChecked, vbUnchecked)
        chkBuildPro.Value = IIf("" & rs("BuildProEnabled") = "true", vbChecked, vbUnchecked)
        
        txtWebUID.Text = "" & rs("WebUID")
        txtWebPWD.Text = "" & rs("password")
        txtWalletPayeeID.Text = "" & rs("WalletPayeeID")
        txtTaxID.Text = "" & rs("TaxID")
        
        cboStatus.ListIndex = IIf("" & rs("Inactive") = "True", 1, 0)
        
        cboTradeType.Text = "" & rs("VendorGroupID")
        Call SetListIndex(cboPOFormat, , "" & rs("POFormat"))
        txtHoldbackPercent = Val("" & rs("HoldbackPercentage"))
        
        With gInsurance
            .Editable = flexEDKbdMouse
            .TextMatrix(GeneralRow, .ColIndex("required")) = "" & rs("glinsrequired")
            .TextMatrix(WorkersRow, .ColIndex("required")) = "" & rs("wcinsrequired")
            .TextMatrix(UmbrellaRow, .ColIndex("required")) = "" & rs("umbinsrequired")
            .TextMatrix(AutomobileRow, .ColIndex("required")) = "" & rs("autoinsrequired")
            
            .TextMatrix(GeneralRow, .ColIndex("expirydate")) = "" & rs("glinsexpdate")
            .TextMatrix(WorkersRow, .ColIndex("expirydate")) = "" & rs("wcinsexpdate")
            .TextMatrix(UmbrellaRow, .ColIndex("expirydate")) = "" & rs("umbinsexpdate")
            .TextMatrix(AutomobileRow, .ColIndex("expirydate")) = "" & rs("autoinsexpdate")
            
            .TextMatrix(GeneralRow, .ColIndex("company")) = "" & rs("GLInsCompany")
            .TextMatrix(WorkersRow, .ColIndex("company")) = "" & rs("WCInsCompany")
            .TextMatrix(UmbrellaRow, .ColIndex("company")) = "" & rs("UmbInsCompany")
            .TextMatrix(AutomobileRow, .ColIndex("company")) = "" & rs("AutoInsCompany")
            
            .TextMatrix(GeneralRow, .ColIndex("policynumber")) = "" & rs("GLInsPolicyNumber")
            .TextMatrix(WorkersRow, .ColIndex("policynumber")) = "" & rs("WCInsPolicyNumber")
            .TextMatrix(UmbrellaRow, .ColIndex("policynumber")) = "" & rs("UmbInsPolicyNumber")
            .TextMatrix(AutomobileRow, .ColIndex("policynumber")) = "" & rs("AutoInsPolicyNumber")
            
        End With
        
        With gTaxGroups
            .Rows = 2
            
            'default values
            .TextMatrix(1, .ColIndex("Community")) = ""
            .TextMatrix(1, .ColIndex("CommunityDescription")) = " -- vendor defaults --"
            .TextMatrix(1, .ColIndex("Labour")) = "" & rs("LabourTaxGroup")
            .TextMatrix(1, .ColIndex("Material")) = "" & rs("MaterialTaxGroup")
            .TextMatrix(1, .ColIndex("SubContract")) = "" & rs("SubContractTaxGroup")
            .TextMatrix(1, .ColIndex("Equipment")) = "" & rs("EquipmentTaxGroup")
            .TextMatrix(1, .ColIndex("Overhead")) = "" & rs("OverheadTaxGroup")
            .TextMatrix(1, .ColIndex("Other")) = "" & rs("OtherTaxGroup")
        
            'community specific
            s = ""
            s = s & "select t.community" & vbCrLf
            s = s & "      ,c.description" & vbCrLf
            s = s & "      ,t.MaterialTaxGroup" & vbCrLf
            s = s & "      ,t.LabourTaxGroup" & vbCrLf
            s = s & "      ,t.SubContractTaxGroup" & vbCrLf
            s = s & "      ,t.EquipmentTaxGroup" & vbCrLf
            s = s & "      ,t.OverheadTaxGroup" & vbCrLf
            s = s & "      ,t.OtherTaxGroup" & vbCrLf
            s = s & "  from vendortaxgroups t" & vbCrLf
            s = s & "       join tbllocality c on(t.Community=c.area)" & vbCrLf
            s = s & " where t.vendor=" & DbQuote(Str, mVendor)
            Set rs = HFApp.SqlExec(s, dbHomefront)
            While Not rs.EOF
                .AddItem ""
                .TextMatrix(.Rows - 1, .ColIndex("Community")) = "" & rs("Community")
                .TextMatrix(.Rows - 1, .ColIndex("CommunityDescription")) = "" & rs("Description")
                .TextMatrix(.Rows - 1, .ColIndex("Material")) = "" & rs("MaterialTaxGroup")
                .TextMatrix(.Rows - 1, .ColIndex("Labour")) = "" & rs("LabourTaxGroup")
                .TextMatrix(.Rows - 1, .ColIndex("SubContract")) = "" & rs("SubContractTaxGroup")
                .TextMatrix(.Rows - 1, .ColIndex("Equipment")) = "" & rs("EquipmentTaxGroup")
                .TextMatrix(.Rows - 1, .ColIndex("Overhead")) = "" & rs("OverheadTaxGroup")
                .TextMatrix(.Rows - 1, .ColIndex("Other")) = "" & rs("OtherTaxGroup")
                rs.MoveNext
            Wend
            
            'add "newline" row
            .AddItem ""
      
        End With
        
        Call LoadPayPoints
        Call LoadContacts
        
        Call SetCtrlFocus(txtDescription)
    End If
    
    'hide payee id if not set
    txtWalletPayeeID.Visible = txtWalletPayeeID.Text <> ""
    lblWalletPayeeID.Visible = txtWalletPayeeID.Text <> ""
    
    Me.tag = mVendor
    mDirty = False
    
End Sub


End Sub
Private Sub LoadPayPoints()
    Dim s As String
    Dim r As Long
    Dim rs As Recordset
    
    s = ""
    s = s & "select *" & vbCrLf
    s = s & "  from VendorPaymentPoints" & vbCrLf
    s = s & " where DivisionID = " & HFApp.DivisionID & vbCrLf
    s = s & "   and Vendor=" & DbQuote(Str, mVendor) & vbCrLf
    s = s & "order by POIndex" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    With gPayPoints
        .Rows = 1
        While Not rs.EOF
            .AddItem ""
            r = .Rows - 1
            .TextMatrix(r, .ColIndex("POIndex")) = "" & rs("POIndex")
            
            .TextMatrix(r, .ColIndex("PayPoint1Percent")) = IIf(Val("" & rs("PayPoint1Percent")) = 0, "", Val("" & rs("PayPoint1Percent")))
            .TextMatrix(r, .ColIndex("PayPoint2Percent")) = IIf(Val("" & rs("PayPoint2Percent")) = 0, "", Val("" & rs("PayPoint2Percent")))
            .TextMatrix(r, .ColIndex("PayPoint3Percent")) = IIf(Val("" & rs("PayPoint3Percent")) = 0, "", Val("" & rs("PayPoint3Percent")))
            .TextMatrix(r, .ColIndex("PayPoint4Percent")) = IIf(Val("" & rs("PayPoint4Percent")) = 0, "", Val("" & rs("PayPoint4Percent")))
            .TextMatrix(r, .ColIndex("PayPoint5Percent")) = IIf(Val("" & rs("PayPoint5Percent")) = 0, "", Val("" & rs("PayPoint5Percent")))
            
            rs.MoveNext
        Wend
        .AddItem ""
    End With
End Sub
Private Sub LoadContacts()
    Dim s As String
    Dim r As Long
    Dim rs As Recordset
    
    s = ""
    s = s & "select *" & vbCrLf
    s = s & "  from contacts" & vbCrLf
    s = s & " where DivisionID = " & HFApp.DivisionID & " and contacttypeid=99" & vbCrLf
    s = s & "   and vendorcode=" & DbQuote(Str, mVendor)
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
            
            .TextMatrix(r, .ColIndex("PONotice")) = "" & rs("PONotice")
            .TextMatrix(r, .ColIndex("FPONotice")) = "" & rs("FPONotice")
            .TextMatrix(r, .ColIndex("SchedNotice")) = "" & rs("SchedNotice")
            .TextMatrix(r, .ColIndex("ServiceNotice")) = "" & rs("ServiceNotice")
            
            
            
            rs.MoveNext
        Wend
        .AddItem ""
    End With

End Sub

