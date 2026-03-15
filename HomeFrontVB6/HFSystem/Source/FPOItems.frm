VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form FPOItems 
   Caption         =   "Purchase Order Setup"
   ClientHeight    =   7665
   ClientLeft      =   255
   ClientTop       =   675
   ClientWidth     =   10980
   Icon            =   "FPOItems.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7665
   ScaleWidth      =   10980
   Begin HFSystem.Slider Slider 
      Height          =   6615
      Left            =   5460
      Top             =   420
      Width           =   60
      _ExtentX        =   106
      _ExtentY        =   11668
   End
   Begin VB.ComboBox cboItemView 
      Height          =   315
      ItemData        =   "FPOItems.frx":058A
      Left            =   3300
      List            =   "FPOItems.frx":058C
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   60
      Width           =   2115
   End
   Begin VSFlex8Ctl.VSFlexGrid gItems 
      Height          =   6615
      Left            =   60
      TabIndex        =   1
      Top             =   420
      Width           =   5235
      _cx             =   9234
      _cy             =   11668
      Appearance      =   1
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
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   4
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FPOItems.frx":058E
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
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
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
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8Ctl.VSFlexGrid gPOs 
      Height          =   6615
      Left            =   5580
      TabIndex        =   2
      Top             =   420
      Width           =   5235
      _cx             =   9234
      _cy             =   11668
      Appearance      =   1
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
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   6
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FPOItems.frx":060F
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
      TabBehavior     =   0
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
      Begin VB.Line DropLine 
         BorderColor     =   &H8000000C&
         BorderStyle     =   3  'Dot
         Visible         =   0   'False
         X1              =   0
         X2              =   2880
         Y1              =   3600
         Y2              =   3600
      End
   End
   Begin MSComctlLib.ImageList SmallIcons 
      Left            =   0
      Top             =   0
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
            Picture         =   "FPOItems.frx":06CA
            Key             =   "option"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":0C64
            Key             =   "RFP"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":11FE
            Key             =   "quote"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":1798
            Key             =   "sendreceive"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":1D32
            Key             =   "communitystandards"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":22CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":2BA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":3480
            Key             =   "pricelists"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":3D5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":4634
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":4F0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":57E8
            Key             =   "MB"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":5D82
            Key             =   "QB"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":631C
            Key             =   "custom"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":68B6
            Key             =   "assembly"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":6E50
            Key             =   "links"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":73EA
            Key             =   "ItemDB"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":7984
            Key             =   "sendpos"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":7F1E
            Key             =   "HelpSearch"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":84B8
            Key             =   "Items"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":8A52
            Key             =   "New"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":8FEC
            Key             =   "Edit"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":9586
            Key             =   "HelpContents"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":9B20
            Key             =   "EditAssembly"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":A0BA
            Key             =   "Forecast"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":A654
            Key             =   "shrink"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":ABEE
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":B188
            Key             =   "customer"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":B722
            Key             =   "close"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":BCBC
            Key             =   "expand"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":C256
            Key             =   "SaveAs"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":C7F0
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":CD8A
            Key             =   "RePrice"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":D324
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":D8BE
            Key             =   "estimating"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":DE58
            Key             =   "jobcost"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":E3F2
            Key             =   "error"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":E98C
            Key             =   "salesworksheet"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":EF26
            Key             =   "ExcelExport"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":F4C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":FA5A
            Key             =   "newworksheet"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":FFF4
            Key             =   "worksheet"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":1058E
            Key             =   "purchaseorder"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":10B28
            Key             =   "ExcelImport"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":110C2
            Key             =   "groupphase"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":1165C
            Key             =   "costcode"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":11BF6
            Key             =   "job"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":12190
            Key             =   "information"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":1272A
            Key             =   "item"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":12CC4
            Key             =   "itemchecked"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":1325E
            Key             =   "phase"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":137F8
            Key             =   "vendor"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":13D92
            Key             =   "warning"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":1432C
            Key             =   "question"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":148C6
            Key             =   "category"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":14E60
            Key             =   "folder"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":153FA
            Key             =   "AddItems"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":15994
            Key             =   "model"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":15F2E
            Key             =   "area"
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":164C8
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":16A62
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":16BBC
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":16D16
            Key             =   "AlignCenter"
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":16E70
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":16FCA
            Key             =   "AlignLeft"
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":17124
            Key             =   "Bullet"
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":1727E
            Key             =   "BulletNumber"
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":173D8
            Key             =   "AlignRight"
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":17532
            Key             =   "printer"
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":17ACC
            Key             =   "defaultprinter"
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":18066
            Key             =   "costcodes"
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":18940
            Key             =   "defaultvendors"
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOItems.frx":1921A
            Key             =   "editpos"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblItemView 
      Alignment       =   1  'Right Justify
      Caption         =   "View: "
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblPOs 
      Caption         =   "PO Indexes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5580
      TabIndex        =   4
      Top             =   120
      Width           =   1665
   End
   Begin VB.Label lblItems 
      Caption         =   "Items"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "FPOItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const SRCFILE = "FPOIndexes::"
Private mDirty As Boolean
Private mLoading As Boolean


Private Sub Form_Load()
    Call IniGetForm(Me)
    cboItemView.AddItem "all items"
    cboItemView.AddItem "un-assigned items only"
End Sub
Private Sub Form_Activate()
    cboItemView.ListIndex = Val(HFApp.Options.Value(POIndexView))
End Sub
Private Sub cboItemView_Click()
On Error GoTo eh
    Static inhere As Boolean
    If inhere Then Exit Sub
    inhere = True
    
    If Not SaveData(True) Then
        cboItemView.ListIndex = IIf(cboItemView.ListIndex = 1, 0, 1)
        Exit Sub
    End If

    Me.Enabled = False
    Screen.MousePointer = vbHourglass
    
    Call LoadItems
    Call LoadPOs
    Call gItems.Outline(1)
    Call gItems.Outline(0)
    Call gPOs.Outline(0)
    mDirty = False
    Screen.MousePointer = vbDefault
    Me.Enabled = True
    inhere = False
    
Exit Sub
eh: Call errHandler(SRCFILE & "cboItemView_Click")
    Me.Enabled = True
    inhere = False
End Sub


Public Function SaveData(prompt As Boolean) As Boolean
On Error GoTo eh:

    Dim rc As Long
    If Not mDirty Then
        SaveData = True
        Exit Function
    End If
    rc = vbYes
    If prompt Then rc = MsgBox("This data has changed." & vbCrLf & vbCrLf & "Do you want to save the changes?", vbExclamation + vbYesNoCancel, Me.Caption)
    Select Case rc
        Case vbNo:     SaveData = True:     Exit Function
        Case vbCancel: Exit Function
    End Select
        
    Screen.MousePointer = vbHourglass
    Call SaveItems
    Screen.MousePointer = vbDefault
    SaveData = True
    
    Exit Function
eh: Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function







Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not SaveData(True) Then Cancel = True
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
    HFApp.Options.Value(POIndexView) = cboItemView.ListIndex
End Sub






Private Sub gItems_DblClick()
On Error GoTo eh
    Dim i As Long
    Dim j As Long
    With gPOs
    i = .FindRow(gItems.TextMatrix(gItems.Row, 0), , 0)
    If i < 0 Then Exit Sub
        
        j = .GetNodeRow(i, flexNTParent)
        While j >= 0
            .GetNode(j).Expanded = True
            j = .GetNodeRow(j, flexNTParent)
        Wend
        
        Call .Select(i, 1)
        Call .ShowCell(i, 1)
        Call .SetFocus
    End With
eh: Exit Sub
End Sub


Private Sub gItems_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    With gItems
        If .Row < 0 Then Exit Sub
        Select Case KeyCode

            Case vbKeyReturn
                Call gItems_DblClick
                
            Case vbKeySpace
                .IsCollapsed(.Row) = IIf(.IsCollapsed(.Row) = flexOutlineCollapsed, flexOutlineExpanded, flexOutlineCollapsed)
                
            Case vbKeyLeft
                If .GetNodeRow(.Row, flexNTFirstChild) <> -1 And .IsCollapsed(.Row) <> flexOutlineCollapsed Then
                    .IsCollapsed(.Row) = flexOutlineCollapsed
                Else
                    If .GetNodeRow(.Row, flexNTParent) <> -1 Then .Row = .GetNodeRow(.Row, flexNTParent)
                End If
                
            Case vbKeyRight
                If .GetNodeRow(.Row, flexNTFirstChild) <> -1 Then
                    If .IsCollapsed(.Row) = flexOutlineCollapsed Then
                        .IsCollapsed(.Row) = flexOutlineExpanded
                    Else
                        .Row = .GetNodeRow(.Row, flexNTFirstChild)
                    End If
                End If
                
        End Select
    End With
End Sub


Private Sub gItems_RowColChange()
On Error Resume Next
    gItems.Col = 1
End Sub

Private Sub gPOs_DblClick()
    Dim i As Long
    Dim j As Long
    
    If gPOs.RowOutlineLevel(gPOs.Row) = 0 Then
        
    Else
        With gItems
            i = .FindRow(gPOs.TextMatrix(gPOs.Row, .ColIndex("key")), , .ColIndex("key"))
            If i < 0 Then Exit Sub
            j = .GetNodeRow(i, flexNTParent)
            While j >= 0
                .GetNode(j).Expanded = True
                j = .GetNodeRow(j, flexNTParent)
            Wend
            Call .Select(i, .ColIndex("description"))
            Call .ShowCell(i, .ColIndex("description"))
            Call gItems.SetFocus
        End With
    End If
End Sub

Private Sub gPOs_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    Dim j As Long
    Dim r As Long
    With gPOs
        If .Row < 0 Then Exit Sub
        Select Case KeyCode
            
            Case vbKeyReturn
                Call gPOs_DblClick
            
            Case vbKeySpace
                .IsCollapsed(.Row) = IIf(.IsCollapsed(.Row) = flexOutlineCollapsed, flexOutlineExpanded, flexOutlineCollapsed)
                
            Case vbKeyLeft
                If .GetNodeRow(.Row, flexNTFirstChild) <> -1 And .IsCollapsed(.Row) <> flexOutlineCollapsed Then
                    .IsCollapsed(.Row) = flexOutlineCollapsed
                Else
                    If .GetNodeRow(.Row, flexNTParent) <> -1 Then .Row = .GetNodeRow(.Row, flexNTParent)
                End If
            
            Case vbKeyRight
                If .GetNodeRow(.Row, flexNTFirstChild) <> -1 Then
                    If .IsCollapsed(.Row) = flexOutlineCollapsed Then
                        .IsCollapsed(.Row) = flexOutlineExpanded
                    Else
                        .Row = .GetNodeRow(.Row, flexNTFirstChild)
                    End If
                End If
                
            Case vbKeyDelete:
                For i = .Rows - 1 To 0 Step -1
                    If .IsSelected(i) Then
                        If .RowOutlineLevel(i) = 0 Then
                        
                            'remove my children - first set each as dirty
                            j = .GetNodeRow(i, flexNTFirstChild)
                            While j <> -1
                                r = gItems.FindRow(.TextMatrix(j, 0), , 0)
                                If r <> -1 Then
                                    gItems.RowData(r) = "DIRTY"
                                    mDirty = True
                                    gItems.TextMatrix(r, 2) = ""
                                    gItems.Cell(flexcpPicture, r, 1) = SmallIcons.ListImages("item").Picture
                                    gItems.RowHidden(r) = False
                                End If
                                Call .RemoveItem(j)
                                j = .GetNodeRow(i, flexNTFirstChild)
                            Wend
                            'delete me
                            .RowData(i) = "DELETED"
                            .RowHidden(i) = True
                        
                        ElseIf .RowOutlineLevel(i) = 1 Then
                            'mark gItems.rowdata as dirty
                            r = gItems.FindRow(.TextMatrix(i, 0), , 0)
                            If r > -1 Then
                                gItems.RowData(r) = "DIRTY"
                                
                                'mark parent row in gPOs as dirty (my data is shit so ignore errors)
                                On Error Resume Next
                                gPOs.RowData(gPOs.GetNodeRow(i, flexNTParent)) = "DIRTY"
                                On Error GoTo 0
                                
                                mDirty = True
                                gItems.TextMatrix(r, 2) = ""
                                gItems.Cell(flexcpPicture, r, 1) = SmallIcons.ListImages("item").Picture
                                gItems.RowHidden(r) = False
                                'remove from gPOs list
                                Call .RemoveItem(i)
                            End If
                        End If
                    End If
                Next
                
        End Select
    End With
End Sub

Private Sub gPOs_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DropLine.Visible = False
End Sub

Private Sub gPOs_OLECompleteDrag(Effect As Long)
    DropLine.Visible = False
End Sub

Private Sub Slider_Move()
    Call Form_Resize
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Const margin = 60
    Slider.Move Slider.left, gItems.Top, Slider.Width, Me.ScaleHeight - gItems.Top - margin
    
    gItems.Move gItems.left, gItems.Top, Slider.left - gItems.left, Slider.Height
    lblItems.left = gItems.left
    cboItemView.left = gItems.left + gItems.Width - cboItemView.Width
    lblItemView.left = cboItemView.left - lblItemView.Width
    
    gPOs.Move Slider.left + Slider.Width, gItems.Top, Me.ScaleWidth - Slider.left - Slider.Width - gItems.left, Slider.Height
    DropLine.X2 = gPOs.Width
    lblPOs.left = gPOs.left
End Sub

Private Sub LoadItems()
    Dim i      As Long
    Dim s      As String
    Dim rs     As Recordset
    Dim Group  As String
    Dim Phase  As String
    Dim item   As String
    
    With gItems
        .Redraw = flexRDNone
        .Rows = 0
        i = -1
        
        s = ""
        s = s & "SELECT p.Phase" & vbCrLf
        s = s & "      ,p.Description PhaseDesc" & vbCrLf
        s = s & "      ,p.GroupPhase " & vbCrLf
        s = s & "      ,i.Item" & vbCrLf
        s = s & "      ,i.Description ItemDesc" & vbCrLf
        s = s & "      ,i.POIndex" & vbCrLf
        s = s & "  FROM tblEstPhases p LEFT OUTER JOIN tblPhaseItem i ON i.DivisionID = p.DivisionID and i.Phase=p.Phase" & vbCrLf
        If cboItemView.ListIndex = 1 Then
            'unassigned only
            s = s & "   AND isnull(i.POIndex,'')=''"
        End If
        s = s & " where p.DivisionID = " & HFApp.DivisionID & vbCrLf
        s = s & "ORDER BY p.SortOrder" & vbCrLf
        Set rs = HFApp.SqlExec(s)
        While Not rs.EOF
            
            If rs("GroupPhase") Then
                Group = Trim("" & rs("Phase"))
                i = i + 1
                .AddItem Group & vbTab & _
                         Group & " - " & rs("PhaseDesc")
                .Cell(flexcpPicture, i, 1) = SmallIcons.ListImages("groupphase").Picture
                .IsSubtotal(i) = True
                .RowOutlineLevel(i) = 0
            Else
                If Phase <> Trim("" & rs("Phase")) Then
                    Phase = Trim("" & rs("Phase"))
                    i = i + 1
                    .AddItem Phase & Chr(1) & vbTab & _
                             Phase & " - " & rs("PhaseDesc")
                    .Cell(flexcpPicture, i, 1) = SmallIcons.ListImages("phase").Picture
                    .IsSubtotal(i) = True
                    .RowOutlineLevel(i) = 1
                End If
                
                If "" <> Trim("" & rs("Item")) Then
                    item = Trim("" & rs("Item"))
                    i = i + 1
                    .AddItem Phase & Chr(1) & item & vbTab & _
                             item & " - " & rs("ItemDesc") & vbTab & _
                             "" & rs("POIndex")
                             
                    .Cell(flexcpPicture, i, 1) = IIf("" = "" & rs("POIndex"), SmallIcons.ListImages("item").Picture, SmallIcons.ListImages("itemchecked").Picture)
                    .IsSubtotal(i) = True
                    .RowOutlineLevel(i) = 2
                End If
                
            End If
            rs.MoveNext
        Wend
        
        'if viewing unassigned only then remove empty group and phases nodes
        If cboItemView.ListIndex = 1 Then
            For i = .Rows - 1 To 0 Step -1
                If .RowOutlineLevel(i) < 2 And .GetNodeRow(i, flexNTFirstChild) = -1 Then
                    Call .RemoveItem(i)
                End If
            Next
        End If
        
        Call .AutoSize(1, 2)
        .Redraw = flexRDBuffered
        
    End With
    
End Sub

Private Sub LoadPOs()
    Dim i  As Long
    Dim s  As String
    Dim rs As Recordset
    
    Dim PO     As String
    Dim Phase  As String
    Dim item   As String
    
    With gPOs
        .Redraw = flexRDNone
        .Rows = 0
                
        i = -1
        
        s = ""
        s = s & "SELECT p.POIndex PO" & vbCrLf
        s = s & "      ,i.Phase" & vbCrLf
        s = s & "      ,i.Item" & vbCrLf
        s = s & "      ,i.Description ItemDesc" & vbCrLf
        s = s & "      ,p.Notes" & vbCrLf
        s = s & "      ,p.JCCostCode" & vbCrLf
        s = s & "      ,p.JCCategory" & vbCrLf
        s = s & "  FROM tblPOIndex p LEFT OUTER JOIN tblPhaseItem i ON p.DivisionID = i.DivisionID and p.POIndex=i.POIndex" & vbCrLf
        s = s & " where p.DivisionID = " & HFApp.DivisionID & vbCrLf
        s = s & "ORDER BY 1"
        
        Set rs = HFApp.SqlExec(s)
        PO = Chr(1)
        While Not rs.EOF
            
            If PO <> Trim("" & rs("PO")) Then
                PO = Trim("" & rs("PO"))
                i = i + 1
                .AddItem "" & vbTab & _
                         PO & vbTab & _
                         PO & vbTab & _
                         "" & rs("Notes") & vbTab & _
                         "" & rs("JCCostCode") & vbTab & _
                         "" & rs("JCCategory")
                .Cell(flexcpPicture, i, 1) = SmallIcons.ListImages("purchaseorder").Picture
                .IsSubtotal(i) = True
                .RowOutlineLevel(i) = 0
            End If
            
            If "" <> Trim("" & rs("Item")) Then
                Phase = Trim("" & rs("Phase"))
                item = Trim("" & rs("Item"))
                i = i + 1
                .AddItem Phase & Chr(1) & item & vbTab & _
                         item & " - " & rs("ItemDesc")
                .Cell(flexcpPicture, i, 1) = SmallIcons.ListImages("item").Picture
                .IsSubtotal(i) = True
                .RowOutlineLevel(i) = 1
            End If
                
            rs.MoveNext
        Wend
        
        Call .Outline(0)
        .Redraw = flexRDBuffered
        
    End With
    
End Sub

Private Sub gItems_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo eh
    If Button = vbLeftButton Then
        With gItems
            If .MouseRow > 0 Then
                If .IsSelected(.MouseRow) And Shift = 0 Then
                    .OLEDrag
                    Cancel = True
                End If
            End If
        End With
    End If
    Exit Sub
eh: Call errHandler(SRCFILE & "gItems_BeforeMouseDown")
End Sub

Private Sub gItems_OLEStartDrag(Data As VSFlex8Ctl.VSDataObject, AllowedEffects As Long)
On Error GoTo eh
    Dim s As String
    Dim i As Integer
    AllowedEffects = vbDropEffectMove
    With gItems
    
        'unselect stuff you cant move
        For i = .Rows - 1 To 0 Step -1
            If .IsSelected(i) And .RowOutlineLevel(i) < 2 Then
                .IsSelected(i) = False
            End If
        Next
        
        'now build data package
        For i = 0 To .Rows - 1
            If .IsSelected(i) Then
                s = s & vbCrLf & .TextMatrix(i, 0) & vbTab & .TextMatrix(i, 1)
            End If
        Next
        s = Mid(s, 3)
        
        If s = "" Then AllowedEffects = vbDropEffectNone
        
    End With
    Data.Clear
    Data.SetData s, vbCFText
    Exit Sub
eh: Call errHandler(SRCFILE & "gItems_OLEStartDrag")
End Sub

Private Sub gPOs_OLEDragOver(Data As VSFlex8Ctl.VSDataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, State As Integer)
    With gPOs
        
        'exit if no current row
        If .MouseRow < 0 Then
            Effect = vbNoDrop
            DropLine.Visible = False
            Exit Sub
        End If
        
        Effect = vbDropEffectMove
        DropLine.Visible = True
        
        DropLine.Y1 = .RowPos(.MouseRow) + .RowHeight(.MouseRow)
        DropLine.Y2 = DropLine.Y1
        
    End With
End Sub

Private Sub gPOs_OLEDragDrop(Data As VSFlex8Ctl.VSDataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim list     As String
    Dim Row      As String
    Dim Key      As String
    Dim desc     As String
    Dim POIndex  As String
    Dim i  As Long
    Dim r  As Long
    Dim parentcollapsed As Boolean
    
    DropLine.Visible = False
    
    With gPOs
        For i = 0 To .Rows - 1
            .IsSelected(i) = False
        Next
        
        list = Data.GetData(vbCFText)
        If list = "" Then Exit Sub
        For i = Parse(list, , vbCrLf) To 1 Step -1
            Row = Parse(list, i, vbCrLf)
            Key = Parse(Row, 1, vbTab)
            desc = Parse(Row, 2, vbTab)
            
            'add new row with a fake key
            r = .MouseRow + 1
            If .MouseRow = -1 Then r = .Rows - 1
            
            Call .AddItem("&Ww%$" & vbTab & desc, r)
            
            
            .Cell(flexcpPicture, r, 1) = SmallIcons.ListImages("item").Picture
            .IsSubtotal(r) = True
            .RowOutlineLevel(r) = 1
                                                
            'find and remove other rows with this key
            r = .FindRow(Key, , 0)
            While r >= 0
                Call .RemoveItem(r)
                r = .FindRow(Key, , 0)
            Wend
            
            'find new row, replace fake key with true one
            r = .FindRow("&Ww%$", , 0)
            .TextMatrix(r, 0) = Key
            .IsSelected(r) = True
            .RowHidden(r) = .IsCollapsed(.GetNodeRow(r, flexNTParent)) = flexOutlineCollapsed
            
            'now update gitems
            POIndex = .TextMatrix(.GetNodeRow(r, flexNTParent), 1)
            r = gItems.FindRow(Key, , 0)
            If r > -1 Then
                gItems.RowData(r) = "DIRTY"
                mDirty = True
                gItems.TextMatrix(r, 2) = POIndex
                gItems.Cell(flexcpPicture, r, 1) = IIf(POIndex = "", SmallIcons.ListImages("item").Picture, SmallIcons.ListImages("itemchecked").Picture)
            End If
        Next
        
        .SetFocus
    End With
End Sub

Private Sub gPOs_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo eh
    If Button = vbLeftButton Then
        With gPOs
            If .MouseRow >= 0 Then
                If Shift = 0 And .IsSelected(.MouseRow) And .RowOutlineLevel(.MouseRow) > 0 Then
                    Cancel = True
                    .OLEDrag
                End If
            End If
        End With
    End If
    Exit Sub
eh: Call errHandler(SRCFILE & "gPOs_BeforeMouseDown")
End Sub

Private Sub gPOs_OLEStartDrag(Data As VSFlex8Ctl.VSDataObject, AllowedEffects As Long)
On Error GoTo eh
    Dim s As String
    Dim i As Integer
    AllowedEffects = vbDropEffectMove
    With gPOs
    
        'unselect stuff you cant move
        For i = .Rows - 1 To 0 Step -1
            If .IsSelected(i) And .RowOutlineLevel(i) < 1 Then
                .IsSelected(i) = False
            End If
        Next
        
        'now build data package
        For i = 0 To .Rows - 1
            If .IsSelected(i) Then
                s = s & vbCrLf & .TextMatrix(i, 0) & vbTab & .TextMatrix(i, 1)
            End If
        Next
        s = Mid(s, 3)
        
    End With
    Data.Clear
    Data.SetData s, vbCFText
    Exit Sub
eh: Call errHandler(SRCFILE & "gPOs_OLESetData")
End Sub


Sub SaveItems()
On Error GoTo eh
    Dim b       As Boolean
    Dim s       As String
    Dim i       As Long
    Dim Phase   As String
    Dim item    As String
    Dim POIndex As String


    b = HFApp.Databases(dbEstimating).State = adStateOpen

    With gItems
    For i = 0 To .Rows - 1
        If .RowData(i) = "DIRTY" Then
            
            Phase = Parse(.TextMatrix(i, .ColIndex("key")), 1, Chr(1))
            item = Parse(.TextMatrix(i, .ColIndex("key")), 2, Chr(1))
            POIndex = .TextMatrix(i, .ColIndex("poindex"))
            
            s = ""
            s = s & "UPDATE tblPhaseItem" & vbCrLf
            s = s & "SET POIndex=" & DbQuote(Str, POIndex) & vbCrLf
            If Not b Then
                s = s & "   ,UpdateEstimating=1" & vbCrLf
            End If
            s = s & "WHERE DivisionID = " & HFApp.DivisionID & " and Phase=" & DbQuote(Str, Phase) & vbCrLf
            s = s & "  AND Item=" & DbQuote(Str, item) & vbCrLf
            Call HFApp.SqlExec(s)
            
            
        End If
    Next
    End With
    mDirty = False

    Exit Sub
eh: Call errHandler(SRCFILE & "SaveItems")
End Sub




