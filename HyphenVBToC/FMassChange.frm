VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Begin VB.Form FMassChange 
   Caption         =   "Mass Change Wizard"
   ClientHeight    =   10770
   ClientLeft      =   2355
   ClientTop       =   885
   ClientWidth     =   7650
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FMassChange.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10770
   ScaleWidth      =   7650
   Begin HFEst.WizHead WizHead1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   53
      Top             =   0
      Width           =   7650
      _ExtentX        =   13494
      _ExtentY        =   1588
      Caption         =   "Model/Option Mass Change"
      Description     =   "The mass change wizard will help you update your model/option library"
      Icon            =   "FMassChange.frx":000C
   End
   Begin VB.Frame WizFrame 
      BorderStyle     =   0  'None
      Height          =   4755
      Index           =   3
      Left            =   7770
      TabIndex        =   32
      Top             =   5160
      Visible         =   0   'False
      Width           =   7395
      Begin VSFlex8Ctl.VSFlexGrid gItems 
         Height          =   4095
         Left            =   930
         TabIndex        =   22
         Top             =   390
         Width           =   6165
         _cx             =   1975659130
         _cy             =   1975655479
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
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FMassChange.frx":08E6
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
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   7
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
      Begin VB.Label lblDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Edit item quantities"
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
         Left            =   900
         TabIndex        =   52
         Top             =   150
         UseMnemonic     =   0   'False
         Width           =   1635
      End
      Begin VB.Image imgOK 
         Height          =   480
         Index           =   4
         Left            =   300
         Picture         =   "FMassChange.frx":09BB
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Frame WizFrame 
      BorderStyle     =   0  'None
      Height          =   4755
      Index           =   2
      Left            =   180
      TabIndex        =   28
      Top             =   5220
      Visible         =   0   'False
      Width           =   7395
      Begin VB.TextBox txtPOIndex 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   2865
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1770
         Width           =   2865
      End
      Begin VB.CheckBox chkDefaultPOIndex 
         Caption         =   "Use Default PO Index"
         Height          =   225
         Left            =   1500
         TabIndex        =   13
         Top             =   1470
         Value           =   1  'Checked
         Width           =   3405
      End
      Begin VB.TextBox txtPhaseItem 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   4
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   4200
         Width           =   3705
      End
      Begin VB.TextBox txtPhaseItem 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   3
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   3480
         Width           =   3705
      End
      Begin VB.TextBox txtPhaseItem 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   2
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   3240
         Width           =   3705
      End
      Begin VB.TextBox txtPhaseItem 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   1
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   2490
         Width           =   3705
      End
      Begin VB.TextBox txtTakeoffQty 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2445
         TabIndex        =   12
         Top             =   1140
         Width           =   615
      End
      Begin VB.TextBox txtPhaseItem 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   0
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   900
         Width           =   3705
      End
      Begin VB.OptionButton optTask 
         Caption         =   "&Edit quantities"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   1080
         TabIndex        =   20
         Top             =   3930
         Width           =   3315
      End
      Begin VB.OptionButton optTask 
         Caption         =   "&Substitute one item for another"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   1080
         TabIndex        =   17
         Top             =   2970
         Width           =   3315
      End
      Begin VB.OptionButton optTask 
         Caption         =   "&Remove an item"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   15
         Top             =   2190
         Width           =   3315
      End
      Begin VB.OptionButton optTask 
         Caption         =   "&Add an item"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   44
         Top             =   600
         Value           =   -1  'True
         Width           =   3315
      End
      Begin VB.Image cmdChoosePOindex 
         Enabled         =   0   'False
         Height          =   240
         Left            =   5745
         Picture         =   "FMassChange.frx":1285
         Top             =   1770
         Width           =   240
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "PO Index"
         ForeColor       =   &H80000011&
         Height          =   195
         Index           =   15
         Left            =   2145
         TabIndex        =   56
         Top             =   1800
         Width           =   660
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Phase / Item"
         Height          =   195
         Index           =   12
         Left            =   1410
         TabIndex        =   51
         Top             =   4230
         Width           =   915
      End
      Begin VB.Image cmdChooseItem 
         Height          =   240
         Index           =   4
         Left            =   6165
         Picture         =   "FMassChange.frx":13CF
         Top             =   4200
         Width           =   240
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Replace with"
         Height          =   195
         Index           =   11
         Left            =   1395
         TabIndex        =   50
         Top             =   3510
         Width           =   930
      End
      Begin VB.Image cmdChooseItem 
         Height          =   240
         Index           =   3
         Left            =   6165
         Picture         =   "FMassChange.frx":1519
         Top             =   3480
         Width           =   240
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Remove this"
         Height          =   195
         Index           =   10
         Left            =   1440
         TabIndex        =   49
         Top             =   3270
         Width           =   885
      End
      Begin VB.Image cmdChooseItem 
         Height          =   240
         Index           =   2
         Left            =   6165
         Picture         =   "FMassChange.frx":1663
         Top             =   3240
         Width           =   240
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Phase / Item"
         Height          =   195
         Index           =   9
         Left            =   1410
         TabIndex        =   48
         Top             =   2520
         Width           =   915
      End
      Begin VB.Image cmdChooseItem 
         Height          =   240
         Index           =   1
         Left            =   6165
         Picture         =   "FMassChange.frx":17AD
         Top             =   2490
         Width           =   240
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Takeoff Qty"
         Height          =   195
         Index           =   8
         Left            =   1485
         TabIndex        =   47
         Top             =   1170
         Width           =   840
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Phase / Item"
         Height          =   195
         Index           =   7
         Left            =   1410
         TabIndex        =   46
         Top             =   930
         Width           =   915
      End
      Begin VB.Image cmdChooseItem 
         Height          =   240
         Index           =   0
         Left            =   6165
         Picture         =   "FMassChange.frx":18F7
         Top             =   900
         Width           =   240
      End
      Begin VB.Label lblDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "How would you like to change these models and options?"
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
         Index           =   4
         Left            =   900
         TabIndex        =   45
         Top             =   150
         UseMnemonic     =   0   'False
         Width           =   4905
      End
      Begin VB.Image imgOK 
         Height          =   480
         Index           =   3
         Left            =   300
         Picture         =   "FMassChange.frx":1A41
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblFilename 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   120
         Width           =   45
      End
   End
   Begin VB.Frame WizFrame 
      BorderStyle     =   0  'None
      Height          =   4755
      Index           =   1
      Left            =   7590
      TabIndex        =   29
      Top             =   870
      Visible         =   0   'False
      Width           =   7395
      Begin VSFlex8Ctl.VSFlexGrid gAssemblies 
         Height          =   2895
         Left            =   1020
         TabIndex        =   10
         Top             =   1590
         Width           =   6165
         _cx             =   1975659130
         _cy             =   1975653362
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
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FMassChange.frx":230B
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
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   7
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
      Begin VB.Label lblDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Click next if the list is correct"
         Height          =   192
         Index           =   7
         Left            =   1056
         TabIndex        =   58
         Top             =   1140
         UseMnemonic     =   0   'False
         Width           =   1944
      End
      Begin VB.Image imgOK 
         Height          =   480
         Index           =   2
         Left            =   420
         Picture         =   "FMassChange.frx":23E0
         Top             =   300
         Width           =   480
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   $"FMassChange.frx":2CAA
         Height          =   720
         Index           =   3
         Left            =   1056
         TabIndex        =   43
         Top             =   420
         UseMnemonic     =   0   'False
         Width           =   5808
      End
      Begin VB.Label lblDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Verify your selection"
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
         Left            =   900
         TabIndex        =   42
         Top             =   150
         UseMnemonic     =   0   'False
         Width           =   1740
      End
      Begin VB.Image imgOK 
         Height          =   480
         Index           =   1
         Left            =   240
         Picture         =   "FMassChange.frx":2D9B
         Top             =   180
         Width           =   480
      End
   End
   Begin VB.PictureBox WizFoot 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   7650
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   10185
      Width           =   7650
      Begin VB.CommandButton cmdNav 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Index           =   0
         Left            =   1860
         TabIndex        =   23
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "< &Back"
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   3060
         TabIndex        =   24
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Next >"
         Height          =   375
         Index           =   2
         Left            =   4200
         TabIndex        =   25
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "Commi&t"
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   5400
         TabIndex        =   26
         Top             =   108
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   -60
         X2              =   26420
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   0
         X2              =   26480
         Y1              =   15
         Y2              =   15
      End
   End
   Begin VB.Frame WizFrame 
      BorderStyle     =   0  'None
      Height          =   4755
      Index           =   0
      Left            =   0
      TabIndex        =   31
      Top             =   870
      Visible         =   0   'False
      Width           =   7395
      Begin VB.CheckBox chkActiveOnly 
         Caption         =   "Active assemblies only"
         Height          =   225
         Left            =   2376
         TabIndex        =   9
         Top             =   4044
         Value           =   1  'Checked
         Width           =   3405
      End
      Begin HFEst.VBCombo cboAssemblyType 
         Height          =   216
         Left            =   2400
         TabIndex        =   8
         Top             =   3756
         Width           =   3912
         _ExtentX        =   6906
         _ExtentY        =   423
         Style           =   2
      End
      Begin VB.TextBox txtSeries 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2400
         TabIndex        =   6
         Top             =   3270
         Width           =   3915
      End
      Begin VB.TextBox txtCategory 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2400
         TabIndex        =   7
         Top             =   3510
         Width           =   3915
      End
      Begin VB.TextBox txtStyle 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2400
         TabIndex        =   5
         Top             =   3030
         Width           =   3915
      End
      Begin VB.TextBox txtDescription 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2400
         TabIndex        =   4
         Top             =   2790
         Width           =   3915
      End
      Begin VB.TextBox txtAssembly 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2400
         TabIndex        =   3
         Top             =   2550
         Width           =   3915
      End
      Begin VB.TextBox txtOption 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2400
         TabIndex        =   2
         Top             =   2310
         Width           =   3915
      End
      Begin VB.TextBox txtModel 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2400
         TabIndex        =   1
         Top             =   2070
         Width           =   3915
      End
      Begin VB.TextBox txtCommunity 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2400
         TabIndex        =   0
         Top             =   1824
         Width           =   3915
      End
      Begin VB.Label lblDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Click next to review the records you have selected"
         Height          =   192
         Index           =   6
         Left            =   1056
         TabIndex        =   57
         Top             =   1152
         UseMnemonic     =   0   'False
         Width           =   3540
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Series"
         Height          =   195
         Index           =   14
         Left            =   1860
         TabIndex        =   55
         Top             =   3300
         Width           =   435
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Assembly Type"
         Height          =   195
         Index           =   13
         Left            =   1215
         TabIndex        =   54
         Top             =   3780
         Width           =   1065
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   $"FMassChange.frx":3665
         Height          =   888
         Index           =   1
         Left            =   1056
         TabIndex        =   41
         Top             =   420
         UseMnemonic     =   0   'False
         Width           =   5868
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Option Category"
         Height          =   195
         Index           =   6
         Left            =   1155
         TabIndex        =   40
         Top             =   3540
         Width           =   1140
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Style"
         Height          =   195
         Index           =   5
         Left            =   1950
         TabIndex        =   39
         Top             =   3060
         Width           =   345
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Description"
         Height          =   195
         Index           =   4
         Left            =   1500
         TabIndex        =   38
         Top             =   2820
         Width           =   795
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Assembly"
         Height          =   195
         Index           =   3
         Left            =   1635
         TabIndex        =   37
         Top             =   2580
         Width           =   660
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Option"
         Height          =   195
         Index           =   2
         Left            =   1830
         TabIndex        =   36
         Top             =   2340
         Width           =   465
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Model"
         Height          =   195
         Index           =   1
         Left            =   1860
         TabIndex        =   35
         Top             =   2100
         Width           =   435
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Community"
         Height          =   195
         Index           =   0
         Left            =   1530
         TabIndex        =   34
         Top             =   1860
         Width           =   765
      End
      Begin VB.Label lblDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Choose models and options"
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
         Left            =   900
         TabIndex        =   33
         Top             =   150
         UseMnemonic     =   0   'False
         Width           =   2340
      End
      Begin VB.Image imgOK 
         Height          =   480
         Index           =   0
         Left            =   300
         Picture         =   "FMassChange.frx":3727
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupSub 
         Caption         =   "Remove"
         Index           =   0
      End
   End
End
Attribute VB_Name = "FMassChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const SRCFILE = "FMassChange::"

Private mRecordset As Recordset

Public Function ShowForm()
    cmdNav(2).Enabled = True
    Me.Show vbModal
End Function

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
On Error GoTo eh
    Static AsmblyLoaded As Boolean
    Static ItemsLoaded  As Boolean
    Dim i As Long
    Dim s As String, ra As Long
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Set cn = New ADODB.Connection
    If cn.ConnectionString = "" Then
        Call cn.Open(HFApp.ConnectionString(dbHomefront), HFApp.LoginDBUID, HFApp.LoginDBPWD)
    End If
    
    Screen.MousePointer = vbHourglass
    For i = 0 To WizFrame.UBound
        WizFrame(i).Visible = i = RHS
    Next
    cmdNav(1).Enabled = RHS > 0
    cmdNav(2).Enabled = RHS < WizFrame.UBound
    cmdNav(3).Enabled = RHS = WizFrame.UBound
        
        
    Select Case RHS
        Case 0 'choose assemblies
            Set gAssemblies.DataSource = Nothing
            gAssemblies.Rows = 1
            
        Case 1 'verify selection
            gAssemblies.Rows = 1
            s = ""
            s = s & "select Community,Model,OptionID,Assembly,AssemblyTypeDesc [Assembly Type],Description [Assembly Desc],Style,Category,case Inactive when 1 then 'Inactive' else 'Active' end Active" & vbCrLf
            s = s & "      from tbldbassemblymaster" & vbCrLf
            s = s & "     where DivisionID = " & HFApp.DivisionID & vbCrLf
            s = s & AssemblyWhereClause()
            s = s & " order by 1,2,3,4,5,6,7" & vbCrLf
            Set rs = New Recordset
            Call rs.Open(s, cn)
            gAssemblies.DataMode = flexDMBoundNoRowCount
            Set gAssemblies.DataSource = Nothing
            gAssemblies.Rows = 1
            Set gAssemblies.DataSource = rs
            gAssemblies.Editable = flexEDKbdMouse
            If gAssemblies.ColKey(0) = "" Then
                On Error Resume Next
                For i = 0 To gItems.Cols - 1
                    gAssemblies.ColKey(i) = gAssemblies.TextMatrix(0, i)
                    gAssemblies.ColDataType(i) = flexDTString
                Next
                On Error GoTo eh
                Call IniGetGrid(Me, gAssemblies)
                gAssemblies.TextMatrix(0, gAssemblies.ColIndex("Community")) = FMain.CD_Community
                AsmblyLoaded = True
            End If
            
        Case 2 'choose action
            Call optTask_Click(-1)
            
        Case 3 'edit items
            s = ""
            s = s & "select m.Community" & vbCrLf
            s = s & "      ,m.Model" & vbCrLf
            s = s & "      ,m.OptionID" & vbCrLf
            s = s & "      ,m.Assembly" & vbCrLf
            s = s & "      ,m.AssemblyTypeDesc [Assembly Type]" & vbCrLf
            s = s & "      ,m.Description [Assembly Desc]" & vbCrLf
            s = s & "      ,m.Style" & vbCrLf
            s = s & "      ,m.Category" & vbCrLf
            s = s & "      ,case Inactive when 1 then 'Inactive' else 'Active' end Active" & vbCrLf
            s = s & "      ,d.Phase" & vbCrLf
            s = s & "      ,d.Item" & vbCrLf
            s = s & "      ,i.Description [Item Desc]" & vbCrLf
            s = s & "      ,d.TakeoffQty Quantity" & vbCrLf
            s = s & "      ,i.TakeoffUOM UOM" & vbCrLf
            s = s & "  from tbldbassemblymaster m" & vbCrLf
            s = s & "       join tbldbassemblydetails d on(m.DivisionID = d.DivisionID and m.community=d.community and m.assembly=d.assembly and m.model=d.model and m.optionid=d.optionid)" & vbCrLf
            s = s & "       join tblphaseitem i on(d.DivisionID = i.DivisionID and d.phase=i.phase and d.item=i.item)" & vbCrLf
            s = s & " where m.DivisionID = " & HFApp.DivisionID & vbCrLf
            If txtPhaseItem(4).Tag <> "" Then
                s = s & "   and d.phase=" & DbQuote(Str, Parse(txtPhaseItem(4).Tag, 1, Chr(1))) & vbCrLf
                s = s & "   and d.item=" & DbQuote(Str, Parse(txtPhaseItem(4).Tag, 2, Chr(1))) & vbCrLf
            End If
            s = s & AssemblyWhereClause("m") & vbCrLf
            s = s & "order by 1,2,3,4,7,8" & vbCrLf
            Set rs = New Recordset
            Call rs.Open(s, cn)
            gItems.DataMode = flexDMBoundNoRowCount
            Set gItems.DataSource = Nothing
            gItems.Rows = 1
            Set gItems.DataSource = rs
            gItems.Editable = flexEDKbdMouse
            If gItems.ColKey(0) = "" Then
                On Error Resume Next
                For i = 0 To gItems.Cols - 1
                    gItems.ColKey(i) = gItems.TextMatrix(0, i)
                Next
                On Error GoTo eh
                Call IniGetGrid(Me, gItems)
                gItems.TextMatrix(0, gItems.ColIndex("Community")) = FMain.CD_Community
                ItemsLoaded = True
            End If
        
    End Select
    
    Screen.MousePointer = vbDefault
    
Exit Property
eh: Call errHandler(SRCFILE & "CurrentFrame")
End Property




Private Sub chkDefaultPOIndex_Click()
    Dim b As Boolean
    b = chkDefaultPOIndex.value = vbChecked
    txtPOIndex.Enabled = Not b
    cmdChoosePOindex.Enabled = Not b
    Label2(15).ForeColor = IIf(Not b, vbWindowText, vbGrayText)
    
End Sub

Private Sub cmdChooseItem_Click(Index As Integer)
    If FPickList.Choose(HFApp.Databases(dbHomefront), "Item", "select phase,item,description from tblphaseitem where DivisionID = " & HFApp.DivisionID) Then
        txtPhaseItem(Index).Tag = FPickList.SelectedItem("Phase") & Chr(1) & FPickList.SelectedItem("Item")
        txtPhaseItem(Index).Text = FPickList.SelectedItem("Phase") & "/" & FPickList.SelectedItem("Item") & " - " & FPickList.SelectedItem("Description")
    End If
End Sub

Private Sub cmdChoosePOindex_Click()
    If FPickList.Choose(HFApp.Databases(dbHomefront), "PO Index", "select POIndex,description from tblpoindex where DivisionID = " & HFApp.DivisionID, txtPOIndex.Text) Then
        txtPOIndex.Text = FPickList.SelectedItem("POIndex")
    End If
End Sub

Private Sub cmdNav_Click(Index As Integer)
On Error Resume Next
    Select Case Index
        Case 0: Unload Me
        Case 1: CurrentFrame = CurrentFrame - 1
        Case 2: CurrentFrame = CurrentFrame + 1
        Case 3: Call ApplyChange
    End Select
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Const margin = 60
    Dim i As Long
    For i = 0 To WizFrame.UBound
        WizFrame(i).BorderStyle = 0
        WizFrame(i).Move 0, WizHead1.Height, Me.ScaleWidth, Me.ScaleHeight - WizHead1.Height - WizFoot.Height
    Next
    
    gAssemblies.Move gAssemblies.Left, gAssemblies.Top, WizFrame(0).Width - gAssemblies.Left - 180, WizFrame(0).Height - gAssemblies.Top - 180
    gItems.Move gItems.Left, gItems.Top, WizFrame(0).Width - gItems.Left - 180, WizFrame(0).Height - gItems.Top - 180
    
    
    cmdNav(0).Move Me.ScaleWidth - (4 * (margin + cmdNav(0).Width)), 2 * margin
    cmdNav(1).Move Me.ScaleWidth - (3 * (margin + cmdNav(0).Width)), 2 * margin
    cmdNav(2).Move Me.ScaleWidth - (2 * (margin + cmdNav(0).Width)), 2 * margin
    cmdNav(3).Move Me.ScaleWidth - (1 * (margin + cmdNav(0).Width)), 2 * margin
End Sub

Private Sub Form_Load()
    Call IniGetForm(Me)
    Call IniGetGrid(Me, gItems, "MassChange.ini")
    Call LoadCustomDescriptions
    
    
    cboAssemblyType.AddItem ""
    cboAssemblyType.AddItem "Model"
    cboAssemblyType.AddItem "Model Specific Option"
    cboAssemblyType.AddItem "Global Option"
    cboAssemblyType.AddItem "Design Center Option"

    CurrentFrame = 0
    Call optTask_Click(0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If gAssemblies.ColKey(0) <> "" Then Call IniPutGrid(Me, gAssemblies)
    Call IniPutGrid(Me, gItems)
    Call IniPutForm(Me)
End Sub


Private Sub gAssemblies_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
    If Button = vbRightButton Then
        If gAssemblies.MouseRow = 0 Then
            Call FMain.ShowColumnMenu(gAssemblies, , , , , False)
            If gAssemblies.ColKey(0) <> "" Then Call IniPutGrid(Me, gAssemblies)
        Else
            If gAssemblies.Row > 0 Then Call PopupMenu(mnuPopup)
        End If
    End If
End Sub

Private Sub gItems_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gItems
    Select Case .ColKey(Col)
        Case "Quantity"
            .AutoSearch = flexSearchNone
        Case Else
            .AutoSearch = flexSearchFromCursor
            Cancel = True
    End Select
    End With
End Sub

 Private Sub gItems_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
    If Button = vbRightButton And gItems.MouseRow = 0 Then Call FMain.ShowColumnMenu(gItems, , , , , False)
End Sub

Private Sub gItems_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    gItems.RowData(Row) = "DIRTY"
End Sub

Private Sub mnuPopupSub_Click(Index As Integer)
    Dim i As Long
    With Me.gAssemblies
    For i = Min(.Row, .RowSel) To Max(.Row, .RowSel)
        .RowHidden(i) = True
    Next
    End With
End Sub

Private Sub optTask_Click(Index As Integer)

    
    'add
    txtPhaseItem(0).Enabled = optTask(0).value
    cmdChooseItem(0).Visible = optTask(0).value
    txtTakeoffQty.Enabled = optTask(0).value
    
    'remove
    txtPhaseItem(1).Enabled = optTask(1).value
    cmdChooseItem(1).Visible = optTask(1).value
    
    'substitue
    txtPhaseItem(2).Enabled = optTask(2).value
    txtPhaseItem(3).Enabled = optTask(2).value
    cmdChooseItem(2).Visible = optTask(2).value
    cmdChooseItem(3).Visible = optTask(2).value
    
    'edit
    txtPhaseItem(4).Enabled = optTask(3).value
    cmdChooseItem(4).Visible = optTask(3).value
    
    
    Select Case Index
        Case -1
        Case 0
            txtPhaseItem(1).Text = ""
            txtPhaseItem(1).Tag = ""
            txtPhaseItem(2).Text = ""
            txtPhaseItem(2).Tag = ""
            txtPhaseItem(3).Text = ""
            txtPhaseItem(3).Tag = ""
            txtPhaseItem(4).Text = ""
            txtPhaseItem(4).Tag = ""
        Case 1
            txtPhaseItem(0).Text = ""
            txtPhaseItem(0).Tag = ""
            txtTakeoffQty.Text = ""
            txtPhaseItem(2).Text = ""
            txtPhaseItem(2).Tag = ""
            txtPhaseItem(3).Text = ""
            txtPhaseItem(3).Tag = ""
            txtPhaseItem(4).Text = ""
            txtPhaseItem(4).Tag = ""
        Case 2
            txtPhaseItem(0).Text = ""
            txtPhaseItem(0).Tag = ""
            txtTakeoffQty.Text = ""
            txtPhaseItem(1).Text = ""
            txtPhaseItem(1).Tag = ""
            txtPhaseItem(4).Text = ""
            txtPhaseItem(4).Tag = ""
        Case 3
            txtPhaseItem(0).Text = ""
            txtPhaseItem(0).Tag = ""
            txtTakeoffQty.Text = ""
            txtPhaseItem(1).Text = ""
            txtPhaseItem(1).Tag = ""
            txtPhaseItem(2).Text = ""
            txtPhaseItem(2).Tag = ""
            txtPhaseItem(3).Text = ""
            txtPhaseItem(3).Tag = ""
    End Select
    
    cmdNav(2).Enabled = optTask(3).value
    cmdNav(3).Enabled = Not optTask(3).value

End Sub

Private Sub txtPhaseItem_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 And Shift = 0 Then
        Call cmdChooseItem_Click(Index)
    End If
End Sub

Private Sub txtTakeoffQty_Validate(Cancel As Boolean)
    txtTakeoffQty.Text = Val(txtTakeoffQty.Text)
End Sub

Private Function AssemblyWhereClause(Optional TableAlias As String = "") As String
    Dim s As String
    Dim i As Long
    If TableAlias <> "" Then TableAlias = TableAlias & "."
    
    If txtCommunity.Text <> "" Then s = s & "       and " & TableAlias & "community like " & DbQuote(Str, txtCommunity.Text) & vbCrLf
    If txtModel.Text <> "" Then s = s & "       and " & TableAlias & "Model like " & DbQuote(Str, txtModel.Text) & vbCrLf
    If txtOption.Text <> "" Then s = s & "       and " & TableAlias & "OptionID like " & DbQuote(Str, txtOption.Text) & vbCrLf
    If txtAssembly.Text <> "" Then s = s & "       and " & TableAlias & "Assembly like " & DbQuote(Str, txtAssembly.Text) & vbCrLf
    If txtDescription.Text <> "" Then s = s & "       and " & TableAlias & "Description like " & DbQuote(Str, txtDescription.Text) & vbCrLf
    If txtStyle.Text <> "" Then s = s & "       and " & TableAlias & "Style like " & DbQuote(Str, txtStyle.Text) & vbCrLf
    If txtSeries.Text <> "" Then s = s & "       and " & TableAlias & "Series like " & DbQuote(Str, txtSeries.Text) & vbCrLf
    If txtCategory.Text <> "" Then s = s & "       and " & TableAlias & "Category like " & DbQuote(Str, txtCategory.Text) & vbCrLf
    If Trim(Replace(cboAssemblyType.Text, Chr(0), "")) <> "" Then s = s & "       and " & TableAlias & "AssemblyTypeDesc like " & DbQuote(Str, cboAssemblyType.Text) & vbCrLf
    If chkActiveOnly.value = vbChecked Then s = s & "       and isnull(" & TableAlias & "inactive,0)=0" & vbCrLf
    
    With Me.gAssemblies
    For i = 1 To .Rows - 1
        If .RowHidden(i) Then
            s = s & "       and not (" & TableAlias & "assembly=" & DbQuote(Str, .TextMatrix(i, .ColIndex("assembly"))) & vbCrLf
            s = s & "            and " & TableAlias & "model=" & DbQuote(Str, .TextMatrix(i, .ColIndex("model"))) & vbCrLf
            s = s & "            and " & TableAlias & "optionid=" & DbQuote(Str, .TextMatrix(i, .ColIndex("optionid"))) & vbCrLf
            s = s & "            and " & TableAlias & "community=" & DbQuote(Str, .TextMatrix(i, .ColIndex("community"))) & ")" & vbCrLf
        End If
    Next
    End With
    
    AssemblyWhereClause = s
End Function

Private Sub ApplyChange()
On Error GoTo eh
    Dim r As Long
    Dim i As Long
    Dim s As String, si As String
    Dim phase1 As String
    Dim item1  As String
    Dim phase2 As String
    Dim item2  As String
    
    
    Select Case True
        Case optTask(0).value  'add item
            If txtPhaseItem(0).Tag = "" Or txtTakeoffQty.Text = "" Then
                MsgBox "You must select an item and enter a quantity", vbExclamation, App.ProductName
                Exit Sub
            End If
            phase1 = Parse(txtPhaseItem(0).Tag, 1, Chr(1))
            item1 = Parse(txtPhaseItem(0).Tag, 2, Chr(1))
            
            si = "select count(*) "
            si = si & "  from tbldbassemblymaster m" & vbCrLf
            si = si & "      ,tblphaseitem i" & vbCrLf
            si = si & " where i.DivisionID = " & HFApp.DivisionID & " and i.phase=" & DbQuote(Str, phase1) & vbCrLf
            si = si & "   and i.item=" & DbQuote(Str, item1) & vbCrLf
            si = si & "   and m.DivisionID = " & HFApp.DivisionID & vbCrLf
            si = si & AssemblyWhereClause("m")
            i = HFApp.SqlExec(si, dbHomefront)(0)
            
            s = ""
            s = s & "insert into tbldbassemblydetails(AssemblyID,DivisionID,community,assembly,model,optionid,phase,item,poindex,takeoffqty,orderqty,ustmp,tstmp)" & vbCrLf
            s = s & "select m.assemblyid, " & HFApp.DivisionID & ",m.community" & vbCrLf
            s = s & "      ,m.assembly" & vbCrLf
            s = s & "      ,m.model" & vbCrLf
            s = s & "      ,m.optionid" & vbCrLf
            s = s & "      ,i.phase" & vbCrLf
            s = s & "      ,i.item" & vbCrLf
            If chkDefaultPOIndex.value = vbChecked Then
                s = s & "      ,''" & vbCrLf
            Else
                s = s & "      ," & DbQuote(Str, txtPOIndex.Text) & vbCrLf
            End If
            s = s & "      ," & DbQuote(Num, txtTakeoffQty.Text) & vbCrLf
            s = s & "      ," & DbQuote(Num, txtTakeoffQty.Text) & " * (100 + i.WastePercent) / 100 * isnull(nullif(i.ConversionFactor,0),1)" & vbCrLf
            s = s & "      ," & DbQuote(Str, HFApp.LoginID) & vbCrLf
            s = s & "      ,getdate()" & vbCrLf
            s = s & "  from tbldbassemblymaster m" & vbCrLf
            s = s & "      ,tblphaseitem i" & vbCrLf
            s = s & " where i.DivisionID = " & HFApp.DivisionID & " and i.phase=" & DbQuote(Str, phase1) & vbCrLf
            s = s & "   and i.item=" & DbQuote(Str, item1) & vbCrLf
            s = s & "   and m.DivisionID = " & HFApp.DivisionID & vbCrLf
            s = s & AssemblyWhereClause("m")
            Call HFApp.SqlExec(s, dbHomefront)
            

            MsgBox "The item was added to " & i & " assemblies.", vbInformation, App.ProductName
        
        
        Case optTask(1).value  'remove item
            If txtPhaseItem(1).Tag = "" Then
                MsgBox "You must select an item", vbExclamation, App.ProductName
                Exit Sub
            End If
            phase1 = Parse(txtPhaseItem(1).Tag, 1, Chr(1))
            item1 = Parse(txtPhaseItem(1).Tag, 2, Chr(1))
            
            si = ""
            si = si & "select count(*)" & vbCrLf
            si = si & "from tbldbassemblymaster  m" & vbCrLf
            si = si & "join tbldbassemblydetails d on(m.DivisionID = d.DivisionID and m.community=d.community and m.assembly=d.assembly and m.model=d.model and m.optionid=d.optionid)" & vbCrLf
            si = si & " where d.phase=" & DbQuote(Str, phase1) & vbCrLf
            si = si & "   and d.item=" & DbQuote(Str, item1) & vbCrLf
            si = si & "   and d.DivisionID = " & HFApp.DivisionID & vbCrLf
            si = si & AssemblyWhereClause("m")
            i = HFApp.SqlExec(si, dbHomefront)(0)
            
            s = ""
            s = s & "delete from tbldbassemblydetails" & vbCrLf
            s = s & "from tbldbassemblymaster  m" & vbCrLf
            s = s & "join tbldbassemblydetails d on(m.DivisionID = d.DivisionID and m.community=d.community and m.assembly=d.assembly and m.model=d.model and m.optionid=d.optionid)" & vbCrLf
            s = s & " where d.phase=" & DbQuote(Str, phase1) & vbCrLf
            s = s & "   and d.item=" & DbQuote(Str, item1) & vbCrLf
            s = s & "   and d.DivisionID = " & HFApp.DivisionID & vbCrLf
            s = s & AssemblyWhereClause("m")
            Call HFApp.SqlExec(s, dbHomefront)
            

            MsgBox i & " items removed.", vbInformation, App.ProductName
            
            
        Case optTask(2).value  'substitute item
            If txtPhaseItem(2).Tag = "" Or txtPhaseItem(3).Tag = "" Then
                MsgBox "You must select both items", vbExclamation, App.ProductName
                Exit Sub
            End If
            phase1 = Parse(txtPhaseItem(2).Tag, 1, Chr(1))
            item1 = Parse(txtPhaseItem(2).Tag, 2, Chr(1))
            phase2 = Parse(txtPhaseItem(3).Tag, 1, Chr(1))
            item2 = Parse(txtPhaseItem(3).Tag, 2, Chr(1))
            
            si = ""
            si = si & "select count(*)" & vbCrLf
            si = si & "from tbldbassemblymaster  m" & vbCrLf
            si = si & "join tbldbassemblydetails d on(m.DivisionID = d.DivisionID and m.community=d.community and m.assembly=d.assembly and m.model=d.model and m.optionid=d.optionid)" & vbCrLf
            si = si & " where d.phase=" & DbQuote(Str, phase1) & vbCrLf
            si = si & "   and d.item=" & DbQuote(Str, item1) & vbCrLf
            si = si & "   and d.DivisionID = " & HFApp.DivisionID & vbCrLf
            si = si & AssemblyWhereClause("m")
            i = HFApp.SqlExec(si, dbHomefront)(0)
            
            s = ""
            s = s & "update tbldbassemblydetails" & vbCrLf
            s = s & "   set phase=p.phase" & vbCrLf
            s = s & "      ,item=p.item" & vbCrLf
            s = s & "      ,poindex=p.poindex" & vbCrLf
            s = s & "from tbldbassemblymaster  m" & vbCrLf
            s = s & "join tbldbassemblydetails d on(m.DivisionID = d.DivisionID and m.community=d.community and m.assembly=d.assembly and m.model=d.model and m.optionid=d.optionid)" & vbCrLf
            s = s & "join tblphaseitem p on(d.divisionid=p.divisionid and p.phase=" & DbQuote(Str, phase2) & " and p.item=" & DbQuote(Str, item2) & ")" & vbCrLf
            s = s & " where d.phase=" & DbQuote(Str, phase1) & vbCrLf
            s = s & "   and d.item=" & DbQuote(Str, item1) & vbCrLf
            s = s & "   and d.DivisionID=" & HFApp.DivisionID & vbCrLf
            s = s & AssemblyWhereClause("m")
            Call HFApp.SqlExec(s, dbHomefront)
            

            MsgBox i & " items replaced.", vbInformation, App.ProductName
        
        Case optTask(3).value  'save editted items
            With gItems
            r = 0
            For i = 1 To .Rows - 1
                If .RowData(i) = "DIRTY" Then
                    r = r + 1
                    s = ""
                    s = s & "update tbldbassemblydetails" & vbCrLf
                    s = s & "   set takeoffqty=" & DbQuote(Num, Val(.TextMatrix(i, .ColIndex("quantity")))) & vbCrLf
                    s = s & "      ,orderqty=" & DbQuote(Num, Val(.TextMatrix(i, .ColIndex("quantity")))) & " * (100 + i.WastePercent) / 100 * isnull(nullif(i.ConversionFactor,0),1)" & vbCrLf
                    s = s & "from tbldbassemblydetails d join tblphaseitem i on(d.DivisionID = i.DivisionID and d.phase=i.phase and d.item=i.item)" & vbCrLf
                    s = s & " where d.community=" & DbQuote(Str, .TextMatrix(i, .ColIndex("community"))) & vbCrLf
                    s = s & "   and d.model=" & DbQuote(Str, .TextMatrix(i, .ColIndex("model"))) & vbCrLf
                    s = s & "   and d.optionid=" & DbQuote(Str, .TextMatrix(i, .ColIndex("optionid"))) & vbCrLf
                    s = s & "   and d.assembly=" & DbQuote(Str, .TextMatrix(i, .ColIndex("assembly"))) & vbCrLf
                    s = s & "   and d.phase=" & DbQuote(Str, .TextMatrix(i, .ColIndex("phase"))) & vbCrLf
                    s = s & "   and d.item=" & DbQuote(Str, .TextMatrix(i, .ColIndex("item"))) & vbCrLf
                    s = s & "   and d.DivisionID = " & HFApp.DivisionID
                    Call HFApp.SqlExec(s, dbHomefront)
                End If
            Next
            .Rows = 1
            End With
            If r > 0 Then MsgBox r & " items updated.", vbInformation, App.ProductName
    End Select
    
    
    


Exit Sub
eh: Call errHandler(SRCFILE & "ApplyChange", s)
End Sub

Private Sub LoadCustomDescriptions()
    Label2(0).Caption = FMain.CD_Community
End Sub

