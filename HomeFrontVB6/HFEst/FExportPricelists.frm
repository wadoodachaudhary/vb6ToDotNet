VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "vbalProgBar6.ocx"
Begin VB.Form FExportPricelists 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export Wizard"
   ClientHeight    =   5220
   ClientLeft      =   2940
   ClientTop       =   5070
   ClientWidth     =   6495
   ControlBox      =   0   'False
   Icon            =   "FExportPricelists.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   Begin VB.Frame WizFrame 
      Caption         =   "0"
      Height          =   3600
      Index           =   0
      Left            =   120
      TabIndex        =   36
      Top             =   900
      Width           =   6315
      Begin VB.OptionButton optCorporate 
         Caption         =   "Export corporate pricing"
         Height          =   270
         Left            =   1425
         TabIndex        =   38
         Top             =   1200
         Width           =   3840
      End
      Begin VB.OptionButton optDivision 
         Caption         =   "Export divisional price lists"
         Height          =   270
         Left            =   1425
         TabIndex        =   37
         Top             =   930
         Value           =   -1  'True
         Width           =   3840
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   300
         Picture         =   "FExportPricelists.frx":000C
         Top             =   240
         Width           =   480
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid gExportData 
      Height          =   4335
      Left            =   7440
      TabIndex        =   16
      Top             =   1020
      Visible         =   0   'False
      Width           =   7515
      _cx             =   13256
      _cy             =   7646
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
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   18
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FExportPricelists.frx":08D6
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
   Begin VB.Frame WizFrame 
      Caption         =   "5"
      Height          =   3600
      Index           =   5
      Left            =   3840
      TabIndex        =   6
      Top             =   7845
      Visible         =   0   'False
      Width           =   6315
      Begin VB.CheckBox chkExportComments 
         Caption         =   "Include assembly comments and item POIndex"
         Height          =   255
         Left            =   1125
         TabIndex        =   4
         Top             =   3300
         Width           =   4455
      End
      Begin VB.CheckBox chkSortByItem 
         Caption         =   "Sort the list by Item."
         Height          =   255
         Left            =   1125
         TabIndex        =   3
         Top             =   3075
         Width           =   4455
      End
      Begin VB.CheckBox chkSingleFile 
         Caption         =   "Write all pricelists to a single spreadsheet file."
         Height          =   255
         Left            =   1125
         TabIndex        =   0
         Top             =   2355
         Width           =   4455
      End
      Begin VB.CheckBox chkExportNextPrices 
         Caption         =   "Include future prices in export file."
         Height          =   255
         Left            =   1125
         TabIndex        =   1
         Top             =   2595
         Width           =   4455
      End
      Begin VB.CheckBox chkCloseWhenFinished 
         Caption         =   "Close this dialog when all pricelists are completed."
         Height          =   255
         Left            =   1125
         TabIndex        =   2
         Top             =   2835
         Width           =   4455
      End
      Begin VB.TextBox txtPath 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1140
         TabIndex        =   5
         Text            =   "0"
         Top             =   555
         Width           =   3915
      End
      Begin vbalProgBarLib6.vbalProgressBar ProgressBar 
         Height          =   315
         Left            =   1140
         TabIndex        =   7
         Top             =   1230
         Visible         =   0   'False
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   556
         Picture         =   "FExportPricelists.frx":0A1D
         ForeColor       =   0
         BarPicture      =   "FExportPricelists.frx":0A39
         ShowText        =   -1  'True
         TextAlignX      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Segments        =   -1  'True
         XpStyle         =   -1  'True
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   300
         Picture         =   "FExportPricelists.frx":0A55
         Top             =   240
         Width           =   480
      End
      Begin VB.Image cmdChooseFolder 
         Height          =   240
         Index           =   0
         Left            =   5100
         Picture         =   "FExportPricelists.frx":131F
         Top             =   585
         Width           =   240
      End
      Begin VB.Label lblSaving 
         Caption         =   "Saving:"
         Height          =   255
         Left            =   1140
         TabIndex        =   11
         Top             =   990
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label Label3 
         Caption         =   "The selected pricelists will be created in:"
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
         Index           =   1
         Left            =   1140
         TabIndex        =   10
         Top             =   315
         Width           =   3915
      End
      Begin VB.Label lblAssemblyCount 
         AutoSize        =   -1  'True
         Caption         =   "Pricelist 4 of 14"
         Height          =   195
         Left            =   1380
         TabIndex        =   9
         Top             =   1590
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label lblItemDesc 
         AutoSize        =   -1  'True
         Caption         =   "BRCKEN Breckenridge Excavating"
         Height          =   195
         Left            =   1380
         TabIndex        =   8
         Top             =   1830
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   2490
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Options:"
         Height          =   255
         Left            =   1125
         TabIndex        =   15
         Top             =   2145
         Width           =   585
      End
   End
   Begin VB.PictureBox WizFoot 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   6495
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   4635
      Width           =   6495
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Finish"
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   5340
         TabIndex        =   31
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Next >"
         Height          =   375
         Index           =   2
         Left            =   4140
         TabIndex        =   30
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "< &Back"
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   3000
         TabIndex        =   29
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Cancel"
         Height          =   375
         Index           =   0
         Left            =   1845
         TabIndex        =   28
         Top             =   120
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   3
         X1              =   0
         X2              =   26480
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   2
         X1              =   60
         X2              =   26540
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.Frame WizFrame 
      Caption         =   "4"
      Height          =   3600
      Index           =   4
      Left            =   2910
      TabIndex        =   12
      Top             =   7305
      Width           =   6315
      Begin VSFlex8Ctl.VSFlexGrid gPOindexes 
         Height          =   3300
         Left            =   1140
         TabIndex        =   14
         Top             =   240
         Width           =   5175
         _cx             =   9128
         _cy             =   5821
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
         HighLight       =   0
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   3
         SelectionMode   =   1
         GridLines       =   0
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   0
         Cols            =   2
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FExportPricelists.frx":18A9
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   1
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   5
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   2
         PicturesOver    =   0   'False
         FillStyle       =   1
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
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   -2147483624
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.CheckBox chkAllPOIndexes 
         Caption         =   "All PO Indexes"
         Height          =   315
         Left            =   4860
         TabIndex        =   35
         Top             =   -30
         Width           =   1455
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   300
         Picture         =   "FExportPricelists.frx":18DA
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Select Purchase Orders"
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
         Index           =   5
         Left            =   1140
         TabIndex        =   13
         Top             =   0
         Width           =   2025
      End
   End
   Begin VB.Frame WizFrame 
      Caption         =   "3"
      Height          =   3600
      Index           =   3
      Left            =   2175
      TabIndex        =   24
      Top             =   6705
      Width           =   6315
      Begin VSFlex8Ctl.VSFlexGrid gAssemblies 
         Height          =   3300
         Left            =   1140
         TabIndex        =   25
         Top             =   240
         Width           =   5175
         _cx             =   9128
         _cy             =   5821
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
         HighLight       =   0
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   3
         SelectionMode   =   1
         GridLines       =   0
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   0
         Cols            =   2
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FExportPricelists.frx":21A4
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   1
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   5
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   2
         PicturesOver    =   0   'False
         FillStyle       =   1
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
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   -2147483624
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.CheckBox chkAllAssemblies 
         Caption         =   "All Assemblies"
         Height          =   315
         Left            =   4920
         TabIndex        =   34
         Top             =   -30
         Width           =   1455
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   300
         Picture         =   "FExportPricelists.frx":21F5
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Select Assemblies"
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
         Index           =   3
         Left            =   1140
         TabIndex        =   26
         Top             =   0
         Width           =   1545
      End
   End
   Begin VB.Frame WizFrame 
      Caption         =   "2"
      Height          =   3600
      Index           =   2
      Left            =   1485
      TabIndex        =   21
      Top             =   6165
      Width           =   6315
      Begin VSFlex8Ctl.VSFlexGrid gVendors 
         Height          =   3300
         Left            =   1140
         TabIndex        =   22
         Top             =   240
         Width           =   5175
         _cx             =   9128
         _cy             =   5821
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
         HighLight       =   0
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   3
         SelectionMode   =   1
         GridLines       =   0
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   0
         Cols            =   4
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FExportPricelists.frx":2ABF
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   1
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   5
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   2
         PicturesOver    =   0   'False
         FillStyle       =   1
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
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   -2147483624
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.CheckBox chkAllVendors 
         Caption         =   "All Vendors"
         Height          =   315
         Left            =   5100
         TabIndex        =   33
         Top             =   -30
         Width           =   1455
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   300
         Picture         =   "FExportPricelists.frx":2B3C
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Select Vendors"
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
         Index           =   2
         Left            =   1140
         TabIndex        =   23
         Top             =   0
         Width           =   1305
      End
   End
   Begin VB.Frame WizFrame 
      Caption         =   "1"
      Height          =   3600
      Index           =   1
      Left            =   795
      TabIndex        =   18
      Top             =   5175
      Width           =   6315
      Begin VSFlex8Ctl.VSFlexGrid gCommunities 
         Height          =   3300
         Left            =   1140
         TabIndex        =   19
         Top             =   240
         Width           =   5175
         _cx             =   9128
         _cy             =   5821
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
         HighLight       =   0
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   3
         SelectionMode   =   1
         GridLines       =   0
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   0
         Cols            =   2
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FExportPricelists.frx":3406
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   1
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   5
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   2
         PicturesOver    =   0   'False
         FillStyle       =   1
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
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   -2147483624
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.CheckBox chkAllCommunities 
         Caption         =   "All Communities"
         Height          =   315
         Left            =   4800
         TabIndex        =   32
         Top             =   -30
         Width           =   1455
      End
      Begin VB.Image imgError 
         Height          =   480
         Left            =   300
         Picture         =   "FExportPricelists.frx":345A
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Select Communities"
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
         Index           =   0
         Left            =   1140
         TabIndex        =   20
         Top             =   0
         Width           =   1665
      End
   End
   Begin HFEst.WizHead WizHead1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   1588
      Caption         =   "Export Vendor Pricelists"
      Description     =   "Export vendor pricelists to Excel spreadsheets"
      Icon            =   "FExportPricelists.frx":3D24
   End
End
Attribute VB_Name = "FExportPricelists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const SRCFILE = "FExportPricelists::"


Private divCount As Long
Private bInHere As Boolean 'stupid flag see chkAllCommunities_Click() and gCommunities_AfterEdit()


Public Sub ShowForm()
    Me.Show vbModal
End Sub

Private Sub cmdChooseFolder_Click(Index As Integer)
    
    Dim s As String
    Dim t As TextBox
    
    Select Case Index
        Case 0: Set t = txtPath
    End Select
        
    t.SetFocus
    s = t.Text
    If VBChooseFolder(0, BIF_RETURNONLYFSDIRS, s, "Select Export Location", Me.hwnd, s) Then
        t.Text = s
    End If

End Sub

Private Sub Form_Load()
    Dim i As Long
    
    Call IniGetForm(Me)
    Call LoadCustomDescriptions
    chkCloseWhenFinished.value = IIf(HFApp.Options.value(ClosePricelistExportDialog), vbChecked, vbUnchecked)
    chkSingleFile.value = IIf(HFApp.Options.value(PricelistExportSingleFile), vbChecked, vbUnchecked)
    chkSortByItem.value = IIf(HFApp.Options.ValueByName("SortPricelistExportByItem") = "True", vbChecked, vbUnchecked)
    chkExportNextPrices.value = IIf(HFApp.Options.value(PricelistExportNextPrices), vbChecked, vbUnchecked)
    chkExportComments.value = IIf(HFApp.Options.ValueByName("PricelistExportComments") = "True", vbChecked, vbUnchecked)
    txtPath.Text = HFApp.Options(PricelistExportPath)
    
    For i = WizFrame.LBound To WizFrame.UBound
        WizFrame(i).BorderStyle = 0
        WizFrame(i).Move 120, 960
    Next
    
    Call GetDivisionCount
    If divCount < 2 Then
        optDivision.value = True
        CurrentFrame = 1
    End If
    
End Sub

Private Sub cmdNav_Click(Index As Integer)
'On Error Resume Next

    Dim i As Long
    
    Select Case Index
        Case 0: Unload Me
        Case 1: 'back
            i = CurrentFrame - 1
            If optCorporate Then
                'skip comunity and assembly selections
                If i = 1 Then i = 0
                If i = 3 Then i = 2
            End If
            CurrentFrame = i
            
        Case 2: 'forward
            i = CurrentFrame + 1
            If optCorporate Then
                'skip comunity and assembly selections
                If i = 1 Then i = 2
                If i = 3 Then i = 4
            End If
            CurrentFrame = i
        
        Case 3:
            If SaveData Then
                If chkCloseWhenFinished Then
                    Unload Me
                Else
                    cmdNav(0).Enabled = True
                    cmdNav(0).Caption = "Close"
                    cmdNav(1).Enabled = False
                    cmdNav(2).Enabled = False
                    cmdNav(3).Enabled = False
                End If
            End If
    End Select
End Sub

Private Function SaveData() As Boolean
On Error GoTo eh
    Dim rs As Recordset
    Dim s As String
    
    Dim POIndex    As String
    Dim Vendor     As String
    Dim VendorDesc As String
    Dim PriceLink  As Long
    
    Dim lists  As Long
    Dim List   As Long
    Dim items  As Long
    Dim Item   As Long
    Dim r As Long
    Dim colHeads As String
    Dim c As Long

    'options affecting output
    Dim BCorpPricing As Boolean
    Dim BShowNextPricing As Boolean
    Dim BShowComments As Boolean
    BCorpPricing = optCorporate.value
    BShowNextPricing = chkExportNextPrices.value = vbChecked
    BShowComments = chkExportComments.value = vbChecked


    If txtPath.Text = "" Then
        Call MsgBox("You must specify an export path", vbOK + vbExclamation, App.ProductName)
        Exit Function
    End If


    lblSaving.Visible = True
    ProgressBar.Visible = True
    
    
    lblAssemblyCount.Caption = ""
    lblItemDesc.Caption = ""
    lblAssemblyCount.Visible = True
    lblItemDesc.Visible = True
    cmdNav(0).Enabled = False
    cmdNav(1).Enabled = False
    cmdNav(2).Enabled = False
    cmdNav(3).Enabled = False
    
        
        
    s = ""
    s = s & "SELECT COUNT(DISTINCT c.Vendor) lists,COUNT(*) items" & vbCrLf
    s = s & "  FROM tblVendorCost c" & vbCrLf
    s = s & "       JOIN tblVendors v ON(c.Vendor=v.Vendor_ID and " & HFApp.DivisionID & " = v.DivisionID)" & vbCrLf
    s = s & "       LEFT OUTER JOIN tblLocality l ON(c.Community=l.Area AND l.Inactive=0)" & vbCrLf
    s = s & "       LEFT OUTER JOIN tblPhaseItem i ON(v.DivisionID = i.DivisionID and c.Phase=i.Phase AND c.Item=i.Item)" & vbCrLf
    s = s & "       LEFT OUTER JOIN DistinctAssemblies a ON(a.Assembly=c.Assembly)" & vbCrLf
    s = s & " WHERE " & WhereClause
    s = s & " and c.DivisionID = " & IIf(BCorpPricing, 0, HFApp.DivisionID)
    Set rs = HFApp.SqlExec(s)
    lists = Val("" & rs(0))
    items = Val("" & rs(1))
    List = 0
    Item = 1
        
        
        

    
    
    
    s = ""
    s = s & "SELECT c.Vendor" & vbCrLf                                 '1
    s = s & "      ,v.Vendor_Name VendorDesc" & vbCrLf                 '2
    s = s & "      ,c.Community" & vbCrLf                              '3
    s = s & "      ,c.CommunityPhase" & vbCrLf                         '4
    s = s & "      ,l.Description CommunityDesc" & vbCrLf              '5
    s = s & "      ,c.Model" & vbCrLf                                  '6
    s = s & "      ,c.Assembly" & vbCrLf                               '7
    s = s & "      ,a.Description AssemblyDesc" & vbCrLf               '8
    s = s & "      ,'' Phase" & vbCrLf                                 '9
    s = s & "      ,cast(i.PriceLink as varchar) Item" & vbCrLf        '10
    s = s & "      ,0.0 PhaseSortOrder" & vbCrLf                       '11
    s = s & "      ,cast(i.PriceLink as float) ItemSortOrder" & vbCrLf '12
    s = s & "      ,i.PriceLink" & vbCrLf                              '13
    s = s & "      ,MIN(i.Description) ItemDesc" & vbCrLf              '14
    s = s & "      ,MAX(c.PartNumber) PartNumber" & vbCrLf
    s = s & "      ,MAX(i.OrderUOM) OrderUOM" & vbCrLf
    s = s & "      ,MAX(c.Current_Cost) Current_Cost" & vbCrLf
    s = s & "      ,MAX(c.Next_Cost1) Next_Cost1" & vbCrLf
    s = s & "      ,MAX(c.Next_Effective1) Next_Effective1" & vbCrLf
    s = s & "      ,MAX(c.Next_Cost2) Next_Cost2" & vbCrLf
    s = s & "      ,MAX(c.Next_Effective2) Next_Effective2" & vbCrLf
    s = s & "      ,'' AssemblyComments" & vbCrLf
    s = s & "      ,'' POindex" & vbCrLf
    s = s & "  FROM tblVendorCost c" & vbCrLf
    s = s & "       JOIN tblVendors v ON(c.Vendor=v.Vendor_ID and v.DivisionID=" & HFApp.DivisionID & ")" & vbCrLf
    s = s & "       LEFT OUTER JOIN tblLocality l ON(c.Community=l.Area)" & vbCrLf
    s = s & "       LEFT OUTER JOIN tblPhaseItem i ON(v.DivisionID = i.DivisionID and c.Phase=i.Phase AND c.Item=i.Item)" & vbCrLf
    s = s & "       LEFT OUTER JOIN DistinctAssembliesWithComments a ON(a.Assembly=c.Assembly and a.Model = c.Model and a.DivisionID = v.DivisionID)" & vbCrLf
    s = s & " WHERE " & WhereClause & vbCrLf
    s = s & "   and c.DivisionID = " & IIf(BCorpPricing, 0, HFApp.DivisionID) & vbCrLf
    s = s & "   AND ISNULL(i.PriceLink,0)<>0 and c.assembly=''" & vbCrLf
    s = s & "   AND v.inactive=0" & vbCrLf
    s = s & "group by c.Vendor" & vbCrLf
    s = s & "        ,v.Vendor_Name " & vbCrLf
    s = s & "        ,c.Community" & vbCrLf
    s = s & "        ,c.CommunityPhase" & vbCrLf
    s = s & "        ,l.Description" & vbCrLf
    s = s & "        ,c.Model" & vbCrLf
    s = s & "        ,c.Assembly" & vbCrLf
    s = s & "        ,a.Description " & vbCrLf
    s = s & "        ,i.PriceLink " & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT c.Vendor" & vbCrLf
    s = s & "      ,v.Vendor_Name VendorDesc" & vbCrLf
    s = s & "      ,c.Community" & vbCrLf
    s = s & "      ,c.CommunityPhase" & vbCrLf
    s = s & "      ,l.Description CommunityDesc" & vbCrLf
    s = s & "      ,c.Model" & vbCrLf
    s = s & "      ,a.Assembly" & vbCrLf
    s = s & "      ,a.Description AssemblyDesc" & vbCrLf
    s = s & "      ,c.Phase" & vbCrLf
    s = s & "      ,c.Item" & vbCrLf
    s = s & "      ,i.PhaseSortOrder" & vbCrLf
    s = s & "      ,i.ItemSortOrder" & vbCrLf
    s = s & "      ,i.PriceLink" & vbCrLf
    s = s & "      ,i.Description ItemDesc" & vbCrLf
    s = s & "      ,c.PartNumber" & vbCrLf
    s = s & "      ,i.OrderUOM" & vbCrLf
    s = s & "      ,c.Current_Cost" & vbCrLf
    s = s & "      ,c.Next_Cost1" & vbCrLf
    s = s & "      ,c.Next_Effective1" & vbCrLf
    s = s & "      ,c.Next_Cost2" & vbCrLf
    s = s & "      ,c.Next_Effective2" & vbCrLf
    s = s & "      ,a.Comments AssemblyComments" & vbCrLf
    s = s & "      ,i.POIndex" & vbCrLf
    s = s & "  FROM tblVendorCost c" & vbCrLf
    s = s & "       JOIN tblVendors v ON(c.Vendor=v.Vendor_ID and v.DivisionID=" & HFApp.DivisionID & ")" & vbCrLf
    s = s & "       LEFT OUTER JOIN tblLocality l ON(c.Community=l.Area)" & vbCrLf
    s = s & "       LEFT OUTER JOIN tblPhaseItem i ON(v.DivisionID = i.DivisionID and c.Phase=i.Phase AND c.Item=i.Item)" & vbCrLf
    s = s & "       LEFT OUTER JOIN DistinctAssembliesWithComments a ON(a.Assembly=c.Assembly and a.Model = c.Model and a.DivisionID = v.DivisionID)" & vbCrLf
    s = s & " WHERE " & WhereClause & vbCrLf
    s = s & "   and c.DivisionID = " & IIf(BCorpPricing, 0, HFApp.DivisionID) & vbCrLf
    s = s & "   AND (ISNULL(i.PriceLink,0)=0 or c.assembly<>'')" & vbCrLf
    s = s & "   AND v.inactive=0" & vbCrLf
    If chkSortByItem.value = vbChecked Then
        s = s & "ORDER by 1,11,12,6,8,3" & vbCrLf
    Else
        s = s & "ORDER BY 1,3,4,6,8" & vbCrLf
    End If
    Set rs = HFApp.SqlExec(s)
    
    With gExportData
        .Rows = 0
        r = -1
        While Not rs.EOF
            
            If Vendor <> rs("vendor") Then
                List = List + 1
                
                'save vendor file
                If .Rows > 0 And chkSingleFile.value = vbUnchecked Then
                    Call .AutoSize(0, .Cols - 1)
                    .ColWidth(1) = 615
                    s = PathAppend(txtPath, CleanFileName(VendorDesc)) & ".xls"
                    Call CreatePath("Export Path", FilePath(s))
                    
                    If optCorporate Then
                        .Cell(flexcpBackColor, 0, 3, .Rows - 1, 3) = RGB(255, 255, 193)
                        .Cell(flexcpBackColor, 0, 5, .Rows - 1, 5) = RGB(255, 255, 193)
                        If BShowNextPricing Then .Cell(flexcpBackColor, 0, 6, .Rows - 1, 9) = RGB(255, 255, 193)
                        .ColDataType(5) = flexDTCurrency
                        .ColDataType(6) = flexDTCurrency
                        .ColDataType(8) = flexDTCurrency
                    Else
                        If BShowComments Then
                            .Cell(flexcpBackColor, 0, 11, .Rows - 1, 11) = RGB(255, 255, 193)
                            .Cell(flexcpBackColor, 0, 13, .Rows - 1, 13) = RGB(255, 255, 193)
                            If BShowNextPricing Then .Cell(flexcpBackColor, 0, 14, .Rows - 1, 17) = RGB(255, 255, 193)
                            .ColDataType(13) = flexDTCurrency
                            .ColDataType(14) = flexDTCurrency
                            .ColDataType(16) = flexDTCurrency
                        Else
                            .Cell(flexcpBackColor, 0, 9, .Rows - 1, 9) = RGB(255, 255, 193)
                            .Cell(flexcpBackColor, 0, 11, .Rows - 1, 11) = RGB(255, 255, 193)
                            If BShowNextPricing Then .Cell(flexcpBackColor, 0, 12, .Rows - 1, 15) = RGB(255, 255, 193)
                            .ColDataType(11) = flexDTCurrency
                            .ColDataType(12) = flexDTCurrency
                            .ColDataType(14) = flexDTCurrency
                        End If
                    End If
                    Call .SaveGrid(s, flexFileExcel)
                    r = -1
                    .Rows = 0
                End If
                Vendor = rs("Vendor")
                VendorDesc = rs("VendorDesc")
                
                'write vendor heading
                r = r + 1:  .AddItem ""
                r = r + 1:  .AddItem ""
                r = r + 1:  .AddItem "Vendor" & vbTab & Vendor
                r = r + 1:  .AddItem "Company Name" & vbTab & VendorDesc
                .Cell(flexcpFontBold, .Rows - 1, 0, .Rows - 2, 0) = True
                
                'write column headings
                r = r + 1
                If BCorpPricing Then
                    colHeads = "Group__Item__Item Description__Part Number__UOM__Price"
                    If BShowNextPricing Then colHeads = colHeads & "__Next Price 1__Effective 1__Next Price 2__Effective 2"
                Else
                    colHeads = FMain.CD_Community & "__Phase__" & FMain.CD_Community & " Name__Model__Specification__Assembly Description"
                    If BShowComments Then colHeads = colHeads & "__Assembly Comments__POIndex"
                    colHeads = colHeads + "__Group__Item__Item Description__Part Number__UOM__Price"
                    If BShowNextPricing Then colHeads = colHeads & "__Next Price 1__Effective 1__Next Price 2__Effective 2"
                End If
                .AddItem Replace(colHeads, "__", vbTab)
                .Cell(flexcpFontBold, .Rows - 1, 0, .Rows - 1, 17) = True
                
            End If
            
            'show progress
            ProgressBar.value = Item / items * 100
            lblAssemblyCount.Caption = "Pricelist " & List & " of " & lists
            lblAssemblyCount.Refresh
            lblItemDesc.Caption = VendorDesc
            lblItemDesc.Refresh
            

            r = r + 1
            .AddItem ""
            c = -1
            If optCorporate Then
                c = c + 1: .TextMatrix(r, c) = "" & rs("Phase")
                c = c + 1: .TextMatrix(r, c) = "" & rs("Item")
                c = c + 1: .TextMatrix(r, c) = "" & rs("ItemDesc")
                c = c + 1: .TextMatrix(r, c) = "" & rs("PartNumber")
                c = c + 1: .TextMatrix(r, c) = "" & rs("OrderUOM")
                c = c + 1: .TextMatrix(r, c) = "" & rs("Current_Cost")
                If BShowNextPricing Then
                    c = c + 1: .TextMatrix(r, c) = "" & rs("Next_Cost1")
                    c = c + 1: .TextMatrix(r, c) = "" & rs("Next_Effective1")
                    c = c + 1: .TextMatrix(r, c) = "" & rs("Next_Cost2")
                    c = c + 1: .TextMatrix(r, c) = "" & rs("Next_Effective2")
                End If
            Else
                c = c + 1: .TextMatrix(r, c) = "" & rs("Community")
                c = c + 1: .TextMatrix(r, c) = "" & rs("CommunityPhase")
                c = c + 1: .TextMatrix(r, c) = "" & rs("CommunityDesc")
                c = c + 1: .TextMatrix(r, c) = "" & rs("Model")
                c = c + 1: .TextMatrix(r, c) = "" & rs("Assembly")
                c = c + 1: .TextMatrix(r, c) = "" & rs("AssemblyDesc")
                If BShowComments Then
                    c = c + 1: .TextMatrix(r, c) = "" & rs("AssemblyComments")
                    c = c + 1: .TextMatrix(r, c) = "" & rs("POIndex")
                End If
                c = c + 1: .TextMatrix(r, c) = "" & rs("Phase")
                c = c + 1: .TextMatrix(r, c) = "" & rs("Item")
                c = c + 1: .TextMatrix(r, c) = "" & rs("ItemDesc")
                c = c + 1: .TextMatrix(r, c) = "" & rs("PartNumber")
                c = c + 1: .TextMatrix(r, c) = "" & rs("OrderUOM")
                c = c + 1: .TextMatrix(r, c) = "" & rs("Current_Cost")
                If BShowNextPricing Then
                    c = c + 1: .TextMatrix(r, c) = "" & rs("Next_Cost1")
                    c = c + 1: .TextMatrix(r, c) = "" & rs("Next_Effective1")
                    c = c + 1: .TextMatrix(r, c) = "" & rs("Next_Cost2")
                    c = c + 1: .TextMatrix(r, c) = "" & rs("Next_Effective2")
                End If
            End If
            Item = Item + 1
            rs.MoveNext
            
        Wend
    
        'save last file
        If .Rows > 0 Then
            Call .AutoSize(0, .Cols - 1)
            .ColWidth(1) = 615
            s = PathAppend(txtPath, IIf(chkSingleFile, "Vendor Pricelists", CleanFileName(VendorDesc))) & ".xls"
            Call CreatePath("Export Path", FilePath(s))
            
            If optCorporate Then
                .Cell(flexcpBackColor, 0, 3, .Rows - 1, 3) = RGB(255, 255, 193)
                .Cell(flexcpBackColor, 0, 5, .Rows - 1, 5) = RGB(255, 255, 193)
                If BShowNextPricing Then .Cell(flexcpBackColor, 0, 6, .Rows - 1, 9) = RGB(255, 255, 193)
                .ColDataType(5) = flexDTCurrency
                .ColDataType(6) = flexDTCurrency
                .ColDataType(8) = flexDTCurrency
            Else
                If BShowComments Then
                    .Cell(flexcpBackColor, 0, 11, .Rows - 1, 11) = RGB(255, 255, 193)
                    .Cell(flexcpBackColor, 0, 13, .Rows - 1, 13) = RGB(255, 255, 193)
                    If BShowNextPricing Then .Cell(flexcpBackColor, 0, 14, .Rows - 1, 17) = RGB(255, 255, 193)
                    .ColDataType(13) = flexDTCurrency
                    .ColDataType(14) = flexDTCurrency
                    .ColDataType(16) = flexDTCurrency
                Else
                    .Cell(flexcpBackColor, 0, 9, .Rows - 1, 9) = RGB(255, 255, 193)
                    .Cell(flexcpBackColor, 0, 11, .Rows - 1, 11) = RGB(255, 255, 193)
                    If BShowNextPricing Then .Cell(flexcpBackColor, 0, 12, .Rows - 1, 15) = RGB(255, 255, 193)
                    .ColDataType(11) = flexDTCurrency
                    .ColDataType(12) = flexDTCurrency
                    .ColDataType(14) = flexDTCurrency
                End If
            End If
            
            Call .SaveGrid(s, flexFileExcel)
            .Rows = 0
        End If
    
        
        
    End With
        
    SaveData = True
    lblAssemblyCount.Caption = ""
    lblItemDesc.Caption = "Finished!"
    cmdNav(0).Enabled = True
    cmdNav(1).Enabled = True
    cmdNav(2).Enabled = True
    Exit Function
eh: Select Case Err.Number
        Case 70, 75
            lblSaving.Visible = False
            ProgressBar.Visible = False
            lblAssemblyCount.Visible = False
            lblItemDesc.Visible = False
            Call MsgBox("Cannot create the " & s & " file." & vbCrLf & "You may not have access to the folder or the file may be in use.", vbExclamation)
            cmdNav(0).Enabled = True
            cmdNav(1).Enabled = True
            cmdNav(2).Enabled = True
        Case Else
            errHandler (SRCFILE & "SaveData")
    End Select
End Function

Private Function WhereClause() As String
    Dim r As Long
    Dim all As Boolean
    Dim s As String
    Dim w As String

    w = ""
    s = ""
    If chkAllCommunities.value <> vbChecked Then
        With gCommunities
        For r = 0 To .Rows - 1
            If .Cell(flexcpChecked, r, 1) = flexChecked Then
                s = s & "," & DbQuote(Str, .TextMatrix(r, 0))
            End If
        Next
        If s <> "" Then w = w & " AND c.Community IN(" & Mid(s, 2) & ")" & vbCrLf
        End With
    End If

    s = ""
    If chkAllVendors.value <> vbChecked Then
        With gVendors
        For r = 0 To .Rows - 1
            If .Cell(flexcpChecked, r, 1) = flexChecked Then
                s = s & "," & DbQuote(Str, .TextMatrix(r, 0))
            End If
        Next
        If s <> "" Then w = w & " AND c.Vendor IN(" & Mid(s, 2) & ")" & vbCrLf
        End With
    End If
    
    s = ""
    If chkAllAssemblies.value <> vbChecked Then
        With gAssemblies
        For r = 0 To .Rows - 1
            If .Cell(flexcpChecked, r, 0) = flexChecked Then
                s = s & "," & DbQuote(Str, .TextMatrix(r, 0))
            End If
        Next
        If s <> "" Then w = w & " AND c.Assembly IN(" & Mid(s, 2) & ")" & vbCrLf
        End With
    End If
    
    s = ""
    If chkAllPOIndexes.value <> vbChecked Then
        With gPOindexes
        For r = 0 To .Rows - 1
            If .Cell(flexcpChecked, r, 0) = flexChecked Then
                s = s & "," & DbQuote(Str, .TextMatrix(r, 0))
            End If
        Next
        If s <> "" Then w = w & " AND i.POIndex IN(" & Mid(s, 2) & ")" & vbCrLf
        End With
    End If
    If w = "" Then w = " AND 1=1" & vbCrLf
    WhereClause = Mid(w, 6)
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
    HFApp.Options.ValueByName("SortPricelistExportByItem") = chkSortByItem.value = vbChecked
    HFApp.Options.value(PricelistExportSingleFile) = chkSingleFile.value = vbChecked
    HFApp.Options.value(ClosePricelistExportDialog) = chkCloseWhenFinished.value = vbChecked
    HFApp.Options.value(PricelistExportNextPrices) = chkExportNextPrices.value = vbChecked
    HFApp.Options.ValueByName("PricelistExportComments") = chkExportComments.value = vbChecked
    HFApp.Options.value(PricelistExportPath) = txtPath.Text
End Sub



Private Sub gCommunities_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = False
End Sub


Private Sub gVendors_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = False
End Sub
Private Sub gAssemblies_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = Col <> 0
End Sub
Private Sub gPOIndexes_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = False
End Sub



Private Property Get CurrentFrame() As Long
    Dim i As Long
    For i = 0 To WizFrame.UBound
        If WizFrame(i).Visible = True Then
            CurrentFrame = i
            Exit Property
        End If
    Next
    CurrentFrame = WizFrame.LBound
End Property

Private Property Let CurrentFrame(RHS As Long)
    Dim i As Long
    Dim s As String
    
    For i = 0 To WizFrame.UBound
        WizFrame(i).Visible = i = RHS
    Next
    If divCount < 2 Then
        cmdNav(1).Enabled = RHS > 1
    Else
        cmdNav(1).Enabled = RHS > 0
    End If
    cmdNav(2).Enabled = RHS < WizFrame.UBound
    cmdNav(3).Enabled = RHS = WizFrame.UBound
    
    
    
    Dim rs As Recordset
    Dim r As Long
    Dim c As Long
    Select Case RHS
        Case 0 'export type, corp/div
            'trigger reload of vendors
            gVendors.Rows = 0
    
        Case 1 'communities
            With gCommunities
                If .Rows = 0 Then
                    .Redraw = flexRDNone
                    r = -1
                    s = "SELECT Area Community,Description CommunityDesc FROM tblLocality c join DivisionCommunities d on c.Area=d.Community WHERE d.DivisionID = " & HFApp.DivisionID & " and Inactive=0 ORDER BY 2"
                    Set rs = HFApp.SqlExec(s)
                    While Not rs.EOF
                        r = r + 1
                        .AddItem rs(0) & vbTab & rs(1)
                        rs.MoveNext
                    Wend
                    .Cell(flexcpChecked, 0, 1, .Rows - 1, 1) = flexUnchecked
                    chkAllCommunities.value = vbChecked
                    Call .AutoSize(0, .Cols - 1)
                    .Redraw = flexRDBuffered
                End If
            End With
            
            
        Case 2 'vendors
            With gVendors
                If .Rows = 0 Then
                    .Redraw = flexRDNone
                    r = -1
                    If optCorporate Then
                        s = "SELECT DISTINCT v.Vendor_Id Vendor,v.Vendor_name Company, v.City, v.Phone FROM tblVendors v INNER JOIN tblVendorCost c ON(v.Vendor_id=c.Vendor and c.DivisionID=0) WHERE v.inactive=0 and v.DivisionID = " & HFApp.DivisionID & " ORDER BY 2"
                    Else
                        s = "SELECT DISTINCT v.Vendor_Id Vendor,v.Vendor_name Company, v.City, v.Phone FROM tblVendors v INNER JOIN tblVendorCost c ON(v.Vendor_id=c.Vendor and v.DivisionID = c.DivisionID) WHERE v.inactive=0 and v.DivisionID = " & HFApp.DivisionID & " ORDER BY 2"
                    End If
                    Set rs = HFApp.SqlExec(s)
                    While Not rs.EOF
                        r = r + 1
                        .AddItem rs(0) & vbTab & rs(1)
                        rs.MoveNext
                    Wend
                    If .Rows > 1 Then
                        .Cell(flexcpChecked, 0, 1, .Rows - 1, 1) = flexUnchecked
                    End If
                    chkAllVendors.value = vbChecked
                    Call .AutoSize(0, .Cols - 1)
                    .Redraw = flexRDBuffered
                End If
            End With
            
            
        Case 3 'assemblies
            With gAssemblies
                If .Rows = 0 Then
                    .Redraw = flexRDNone
                    r = -1
                    s = "SELECT Assembly,Description AssemblyDesc FROM DistinctAssemblies where DivisionID = " & HFApp.DivisionID & " ORDER BY 1"
                    Set rs = HFApp.SqlExec(s)
                    If Not rs.EOF Then
                        While Not rs.EOF
                            r = r + 1
                            .AddItem rs(0) & vbTab & rs(1)
                            rs.MoveNext
                        Wend
                        .Cell(flexcpChecked, 0, 0, .Rows - 1, 0) = flexUnchecked
                        chkAllAssemblies.value = vbChecked
                        Call .AutoSize(0, .Cols - 1)
                    End If
                    .Redraw = flexRDBuffered
                End If
            End With

            
            
        Case 4 'poindexes
            With gPOindexes
                If .Rows = 0 Then
                    .Redraw = flexRDNone
                    r = -1
                    s = "SELECT POIndex,case when poindex=description then '' else description end FROM tblPOIndex where DivisionID = " & HFApp.DivisionID & " ORDER BY 1"
                    Set rs = HFApp.SqlExec(s)
                    While Not rs.EOF
                        r = r + 1
                        .AddItem rs(0) & vbTab & rs(1)
                        rs.MoveNext
                    Wend
                    .Cell(flexcpChecked, 0, 0, .Rows - 1, 0) = flexUnchecked
                    chkAllPOIndexes.value = vbChecked
                    Call .AutoSize(0, .Cols - 1)
                    .Redraw = flexRDBuffered
                End If
            End With
            
            
        Case 5 'options
    End Select
    
End Property

Private Sub chkAllCommunities_Click()
    If Not bInHere Then
        On Error Resume Next
        gCommunities.Cell(flexcpChecked, 0, 1, gCommunities.Rows - 1, 1) = IIf(chkAllCommunities.value = vbChecked, flexChecked, flexUnchecked)
    End If
End Sub
Private Sub gCommunities_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    bInHere = True
    With gCommunities
    For i = 0 To .Rows - 1
        If .Cell(flexcpChecked, 0, 1) <> .Cell(flexcpChecked, i, 1) Then
            chkAllCommunities.value = vbGrayed
            bInHere = False
            Exit Sub
        End If
    Next
    chkAllCommunities.value = IIf(.Cell(flexcpChecked, 0, 1) = flexChecked, vbChecked, vbUnchecked)
    End With
    bInHere = False
End Sub

Private Sub chkAllVendors_Click()
    If Not bInHere Then
        If gVendors.Rows > 1 Then
            gVendors.Cell(flexcpChecked, 0, 1, gVendors.Rows - 1, 1) = IIf(chkAllVendors.value = vbChecked, flexChecked, flexUnchecked)
        End If
    End If
End Sub
Private Sub gVendors_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    bInHere = True
    With gVendors
    For i = 0 To .Rows - 1
        If .Cell(flexcpChecked, 0, 1) <> .Cell(flexcpChecked, i, 1) Then
            chkAllVendors.value = vbGrayed
            bInHere = False
            Exit Sub
        End If
    Next
    chkAllVendors.value = IIf(.Cell(flexcpChecked, 0, 1) = flexChecked, vbChecked, vbUnchecked)
    End With
    bInHere = False
End Sub

Private Sub chkAllAssemblies_Click()
    If Not bInHere Then
        On Error Resume Next
        gAssemblies.Cell(flexcpChecked, 0, 0, gAssemblies.Rows - 1, 0) = IIf(chkAllAssemblies.value = vbChecked, flexChecked, flexUnchecked)
    End If
End Sub
Private Sub gAssemblies_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    bInHere = True
    With gAssemblies
    For i = 0 To .Rows - 1
        If .Cell(flexcpChecked, 0, 0) <> .Cell(flexcpChecked, i, 0) Then
            chkAllAssemblies.value = vbGrayed
            bInHere = False
            Exit Sub
        End If
    Next
    chkAllAssemblies.value = IIf(.Cell(flexcpChecked, 0, 0) = flexChecked, vbChecked, vbUnchecked)
    End With
    bInHere = False
End Sub

Private Sub chkAllPOIndexes_Click()
    If Not bInHere Then
        On Error Resume Next
        gPOindexes.Cell(flexcpChecked, 0, 0, gPOindexes.Rows - 1, 0) = IIf(chkAllPOIndexes.value = vbChecked, flexChecked, flexUnchecked)
    End If
End Sub

Private Sub gPOIndexes_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    bInHere = True
    With gPOindexes
    For i = 0 To .Rows - 1
        If .Cell(flexcpChecked, 0, 0) <> .Cell(flexcpChecked, i, 0) Then
            chkAllPOIndexes.value = vbGrayed
            bInHere = False
            Exit Sub
        End If
    Next
    chkAllPOIndexes.value = IIf(.Cell(flexcpChecked, 0, 0) = flexChecked, vbChecked, vbUnchecked)
    End With
    bInHere = False
End Sub

Private Sub optCorporate_Click()
    If optCorporate Then
        chkAllCommunities.value = vbChecked
        chkAllAssemblies.value = vbChecked
    End If
End Sub

Private Sub WizHead1_GotFocus()
'Unload Me
End Sub


Private Sub LoadCustomDescriptions()
    Label3(0).Caption = "Select " & FMain.CD_Community
    chkAllCommunities.Caption = "All " & FMain.CD_Community
End Sub

Private Sub GetDivisionCount()
    divCount = Val("" & HFApp.SqlExec("select count(*) from divisions")(0))
End Sub
