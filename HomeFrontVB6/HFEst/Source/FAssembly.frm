VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{302C5C1A-C2E2-4302-9AF9-BAC87EEFDECE}#1.0#0"; "Panels.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FAssembly 
   Caption         =   "Model & Option Library"
   ClientHeight    =   9000
   ClientLeft      =   1890
   ClientTop       =   2145
   ClientWidth     =   20220
   Icon            =   "FAssembly.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9000
   ScaleWidth      =   20220
   Begin Panels.Slider Slider 
      Height          =   60
      Left            =   105
      Top             =   4395
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   106
      Orientation     =   1
   End
   Begin VSFlex8Ctl.VSFlexGrid gComponents 
      Height          =   1020
      Left            =   105
      TabIndex        =   36
      Top             =   3105
      Width           =   12885
      _cx             =   22728
      _cy             =   1799
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
      BackColorSel    =   14336431
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   255
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FAssembly.frx":000C
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
      FillStyle       =   1
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   2
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
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Width           =   20220
      _ExtentX        =   35666
      _ExtentY        =   1058
      ButtonWidth     =   1667
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            Key             =   "New"
            Style           =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open"
            Key             =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Save"
            Key             =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "One Time"
            Key             =   "TakeoffOneTime"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Item"
            Key             =   "TakeoffItem"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Assembly"
            Key             =   "TakeoffAssembly"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "PlanSwift"
            Key             =   "TakeoffPlanSwift"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Custom"
            Key             =   "TakeoffCustom"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Item Chart"
            Key             =   "TakeoffItemChart"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Create/Edit Chart"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Change"
            Key             =   "MassChange"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pricelist"
            Key             =   "PricelistExport"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Import"
            Key             =   "Import"
            Object.ToolTipText     =   "Import from Pipeline"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ImportModels"
                  Text            =   "Models and Options"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ImportGlobals"
                  Text            =   "Global Options"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.Timer Timer1 
         Left            =   9510
         Top             =   0
      End
   End
   Begin VB.Frame HeaderFrame 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   2280
      Left            =   -45
      TabIndex        =   41
      Top             =   585
      Width           =   21432
      Begin VB.Frame frmBilling 
         Appearance      =   0  'Flat
         Caption         =   " Billing "
         ForeColor       =   &H8000000D&
         Height          =   1245
         Left            =   14880
         TabIndex        =   80
         Top             =   990
         Width           =   3195
         Begin VB.TextBox txtBillingFactor 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   240
            Left            =   945
            MaxLength       =   10
            TabIndex        =   33
            Top             =   270
            Width           =   615
         End
         Begin HFEst.VBCombo cboBillingCode 
            Height          =   240
            Left            =   945
            TabIndex        =   34
            Top             =   525
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   423
         End
         Begin HFEst.VBCombo cboBillingCategory 
            Height          =   240
            Left            =   945
            TabIndex        =   35
            Top             =   780
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   423
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Factor"
            Height          =   195
            Left            =   435
            TabIndex        =   83
            Top             =   255
            Width           =   450
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Category"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   82
            Top             =   810
            Width           =   630
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Code"
            Height          =   195
            Index           =   1
            Left            =   495
            TabIndex        =   81
            Top             =   555
            Width           =   375
         End
      End
      Begin VB.CheckBox chkIsBaseAssembly 
         Caption         =   "Base assembly. Will not be published to sales."
         Height          =   195
         Left            =   11280
         TabIndex        =   31
         Top             =   1410
         Width           =   3705
      End
      Begin VB.Frame frmSeriesFields 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   525
         Left            =   5760
         TabIndex        =   53
         Top             =   0
         Visible         =   0   'False
         Width           =   3912
         Begin VB.TextBox txtElevation 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   1104
            MaxLength       =   20
            TabIndex        =   9
            Top             =   270
            Width           =   2400
         End
         Begin HFEst.VBCombo cboSeries 
            Height          =   240
            Left            =   1104
            TabIndex        =   8
            Top             =   12
            Width           =   2652
            _ExtentX        =   4683
            _ExtentY        =   423
         End
         Begin VB.Label lblElevation 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Elevation"
            Height          =   195
            Left            =   405
            TabIndex        =   59
            Top             =   300
            Width           =   660
         End
         Begin VB.Label lblSeries 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Series"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   192
            Left            =   600
            TabIndex        =   54
            Top             =   36
            Width           =   432
         End
      End
      Begin VB.TextBox txtGraphicPath 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   1344
         MaxLength       =   200
         TabIndex        =   6
         Top             =   1020
         Width           =   4140
      End
      Begin VB.TextBox txtLStyle 
         BorderStyle     =   0  'None
         Height          =   230
         Left            =   18975
         MaxLength       =   30
         TabIndex        =   74
         Top             =   420
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox txtLOther 
         BorderStyle     =   0  'None
         Height          =   230
         Left            =   18975
         MaxLength       =   30
         TabIndex        =   73
         Top             =   900
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox txtLFinish 
         BorderStyle     =   0  'None
         Height          =   230
         Left            =   18975
         MaxLength       =   30
         TabIndex        =   72
         Top             =   660
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox txtLColor 
         BorderStyle     =   0  'None
         Height          =   230
         Left            =   18975
         MaxLength       =   30
         TabIndex        =   71
         Top             =   180
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox txtAStyle 
         BorderStyle     =   0  'None
         Height          =   230
         Left            =   19590
         MaxLength       =   30
         TabIndex        =   70
         Top             =   420
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox txtAOther 
         BorderStyle     =   0  'None
         Height          =   230
         Left            =   19590
         MaxLength       =   30
         TabIndex        =   69
         Top             =   900
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox txtAFinish 
         BorderStyle     =   0  'None
         Height          =   230
         Left            =   19590
         MaxLength       =   30
         TabIndex        =   68
         Top             =   660
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox txtAColor 
         BorderStyle     =   0  'None
         Height          =   230
         Left            =   19590
         MaxLength       =   30
         TabIndex        =   66
         Top             =   180
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Frame frmEstimatorNotes 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   945
         Left            =   11115
         TabIndex        =   64
         Top             =   0
         Width           =   4215
         Begin VB.TextBox txtNotes 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   705
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   28
            Top             =   210
            Width           =   4140
         End
         Begin VB.Label Label111 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Estimator Notes"
            Height          =   192
            Index           =   1
            Left            =   48
            TabIndex        =   65
            Top             =   0
            Width           =   1116
         End
      End
      Begin VB.CheckBox chkActive 
         Caption         =   "Active"
         Height          =   195
         Left            =   3408
         TabIndex        =   4
         Top             =   555
         Width           =   2055
      End
      Begin VB.CheckBox chkUseNormalSalesQtyFactors 
         Caption         =   "Use normal sales quantity factors."
         Height          =   195
         Left            =   11280
         TabIndex        =   30
         Top             =   1185
         Width           =   3075
      End
      Begin VB.CheckBox chkTakeoffRequired 
         Caption         =   "Takeoff is required on every sale."
         Height          =   195
         Left            =   11280
         TabIndex        =   29
         Top             =   960
         Width           =   3075
      End
      Begin VB.TextBox txtComments 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   744
         Left            =   1344
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   1275
         Width           =   4140
      End
      Begin VB.TextBox txtOption 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   3408
         MaxLength       =   20
         TabIndex        =   2
         Top             =   255
         Width           =   2055
      End
      Begin VB.TextBox txtDescription 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   1344
         MaxLength       =   200
         TabIndex        =   5
         Top             =   765
         Width           =   4140
      End
      Begin HFEst.VBCombo cboModel 
         Height          =   240
         Left            =   1350
         TabIndex        =   1
         Top             =   255
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   423
         Style           =   2
      End
      Begin HFEst.VBCombo cboCommunity 
         Height          =   240
         Left            =   1344
         TabIndex        =   0
         Top             =   0
         Width           =   4188
         _ExtentX        =   7382
         _ExtentY        =   423
         Style           =   2
      End
      Begin HFEst.VBCombo cboPriceCommunity 
         Height          =   240
         Left            =   11130
         TabIndex        =   32
         Top             =   1920
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   423
         Style           =   2
      End
      Begin VB.TextBox txtAssembly 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   1344
         MaxLength       =   20
         TabIndex        =   3
         Top             =   510
         Width           =   2055
      End
      Begin VB.Frame frmOptionFields 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Caption         =   "1"
         Height          =   1605
         Left            =   5760
         TabIndex        =   49
         Top             =   525
         Visible         =   0   'False
         Width           =   5280
         Begin VB.CheckBox chkDisplayTotalOnly 
            Caption         =   "Display total only"
            Height          =   215
            Left            =   3420
            TabIndex        =   27
            Top             =   945
            Width           =   1740
         End
         Begin VB.CheckBox chkSelectByRoom 
            Caption         =   "Select by room"
            Height          =   215
            Left            =   3420
            TabIndex        =   26
            Top             =   720
            Width           =   1740
         End
         Begin VB.CheckBox chkDCSalesOnly 
            Caption         =   "DC Sales Only"
            Height          =   215
            Left            =   3420
            TabIndex        =   25
            Top             =   495
            Width           =   1740
         End
         Begin VB.TextBox txtConstCutoff 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   240
            Left            =   1104
            MaxLength       =   10
            TabIndex        =   22
            Top             =   765
            Width           =   300
         End
         Begin VB.TextBox txtQty 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   240
            Left            =   1104
            MaxLength       =   10
            TabIndex        =   20
            Top             =   510
            Width           =   795
         End
         Begin VB.CheckBox chkIncludedOption 
            Caption         =   "Included Option"
            Height          =   215
            Left            =   3420
            TabIndex        =   24
            Top             =   270
            Width           =   1740
         End
         Begin VB.TextBox txtJCExtra 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   240
            Left            =   1104
            MaxLength       =   10
            TabIndex        =   19
            Top             =   255
            Width           =   1155
         End
         Begin HFEst.VBCombo cboUOM 
            Height          =   240
            Left            =   1920
            TabIndex        =   21
            Top             =   510
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   423
         End
         Begin HFEst.VBCombo cboLocation 
            Height          =   240
            Left            =   1110
            TabIndex        =   23
            Top             =   1020
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   423
            Style           =   2
         End
         Begin HFEst.VBCombo cboCategory 
            Height          =   240
            Left            =   1104
            TabIndex        =   18
            Top             =   0
            Width           =   2652
            _ExtentX        =   4683
            _ExtentY        =   423
         End
         Begin VB.Label lblConstCutoff 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Const Cutoff"
            Height          =   195
            Left            =   75
            TabIndex        =   78
            Top             =   780
            Width           =   975
         End
         Begin VB.Label lblMoreAttributes 
            AutoSize        =   -1  'True
            Caption         =   "More Attributes..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   1110
            TabIndex        =   67
            Top             =   1275
            Width           =   1200
         End
         Begin VB.Label lblLocation 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Location"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   420
            TabIndex        =   63
            Top             =   1035
            Width           =   615
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "/"
            Height          =   195
            Index           =   2
            Left            =   375
            TabIndex        =   57
            Top             =   -15
            Width           =   75
         End
         Begin VB.Label lblGroup 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cat"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   120
            TabIndex        =   56
            Top             =   -15
            Width           =   240
         End
         Begin VB.Label Label111 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Qty/UOM"
            Height          =   195
            Index           =   0
            Left            =   345
            TabIndex        =   52
            Top             =   525
            Width           =   690
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "JC Extra"
            Height          =   195
            Index           =   0
            Left            =   450
            TabIndex        =   51
            Top             =   270
            Width           =   585
         End
         Begin VB.Label lblCategory 
            AutoSize        =   -1  'True
            Caption         =   "Sub Cat"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   480
            TabIndex        =   50
            Top             =   -15
            Width           =   570
         End
      End
      Begin VB.Frame frmModelFields 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1605
         Left            =   5760
         TabIndex        =   46
         Top             =   525
         Visible         =   0   'False
         Width           =   5265
         Begin VB.TextBox txtSpecDocument 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   240
            Left            =   1104
            MaxLength       =   200
            TabIndex        =   17
            Top             =   1020
            Width           =   2412
         End
         Begin VB.TextBox txtMaxWidth 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   240
            Left            =   1104
            TabIndex        =   15
            Top             =   765
            Width           =   630
         End
         Begin VB.TextBox txtMaxLength 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   240
            Left            =   1752
            TabIndex        =   16
            Top             =   765
            Width           =   630
         End
         Begin VB.TextBox txtFloorArea 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   240
            Left            =   2400
            TabIndex        =   14
            Top             =   510
            Width           =   930
         End
         Begin VB.TextBox txtBathrooms 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   240
            Left            =   1752
            TabIndex        =   13
            Top             =   510
            Width           =   630
         End
         Begin VB.TextBox txtBedrooms 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   240
            Left            =   1104
            TabIndex        =   12
            Top             =   510
            Width           =   630
         End
         Begin HFEst.VBCombo cboStyle 
            Height          =   240
            Left            =   1104
            TabIndex        =   10
            Top             =   0
            Width           =   2652
            _ExtentX        =   4683
            _ExtentY        =   423
         End
         Begin VB.TextBox txtStyle 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   240
            Left            =   1104
            MaxLength       =   30
            TabIndex        =   38
            Top             =   0
            Width           =   2652
         End
         Begin HFEst.VBCombo cboScheduleTemplate 
            Height          =   240
            Left            =   1110
            TabIndex        =   11
            Top             =   255
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   423
            Style           =   2
         End
         Begin VB.Label lblScheduleTemplate 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Sched Tmpl"
            Height          =   195
            Left            =   195
            TabIndex        =   91
            Top             =   285
            Width           =   855
         End
         Begin VB.Label lblModelDimensions 
            AutoSize        =   -1  'True
            Caption         =   "Edit Dimensions..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   1125
            TabIndex        =   90
            Top             =   1335
            Width           =   1260
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Spec Doc"
            Height          =   195
            Left            =   330
            TabIndex        =   77
            Top             =   1020
            Width           =   720
         End
         Begin VB.Image cmdBrowse 
            Height          =   240
            Index           =   1
            Left            =   3540
            Picture         =   "FAssembly.frx":0134
            Top             =   1035
            Width           =   240
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Width/Length"
            Height          =   195
            Left            =   120
            TabIndex        =   76
            Top             =   780
            Width           =   930
         End
         Begin VB.Label Label177 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Bds/Bths/SqFt"
            Height          =   195
            Left            =   -15
            TabIndex        =   48
            Top             =   525
            Width           =   1065
         End
         Begin VB.Label Label166 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Style"
            Height          =   195
            Left            =   705
            TabIndex        =   47
            Top             =   30
            Width           =   345
         End
      End
      Begin VB.Label lblAssemblyCost 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   1
         Left            =   12465
         TabIndex        =   89
         Top             =   1695
         Width           =   945
      End
      Begin VB.Label lblComponentCost 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   16815
         TabIndex        =   88
         Top             =   345
         Width           =   945
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Component Cost:"
         Height          =   195
         Index           =   3
         Left            =   15465
         TabIndex        =   87
         Top             =   345
         Width           =   1215
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total Cost:"
         Height          =   195
         Index           =   2
         Left            =   15915
         TabIndex        =   86
         Top             =   585
         Width           =   765
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Assembly Cost:"
         Height          =   210
         Index           =   1
         Left            =   15615
         TabIndex        =   85
         Top             =   90
         Width           =   1065
      End
      Begin VB.Label lblTotalCost 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   16815
         TabIndex        =   84
         Top             =   585
         Width           =   945
      End
      Begin VB.Label lblComponents 
         AutoSize        =   -1  'True
         Caption         =   "Components"
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
         Left            =   30
         TabIndex        =   79
         Top             =   2070
         Width           =   1050
      End
      Begin VB.Image cmdBrowse 
         Height          =   240
         Index           =   0
         Left            =   5505
         Picture         =   "FAssembly.frx":027E
         Top             =   1020
         Width           =   240
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Image"
         Height          =   195
         Left            =   810
         TabIndex        =   75
         Top             =   990
         Width           =   450
      End
      Begin VB.Label lblAssemblyCost 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   0
         Left            =   16815
         TabIndex        =   60
         Top             =   90
         Width           =   945
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Show pricing for:"
         Height          =   195
         Index           =   0
         Left            =   11130
         TabIndex        =   58
         Top             =   1695
         Width           =   1185
      End
      Begin VB.Label lblCommunity 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Community"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   360
         TabIndex        =   55
         Top             =   36
         Width           =   912
      End
      Begin VB.Label Label111 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Comments"
         Height          =   210
         Index           =   10
         Left            =   480
         TabIndex        =   45
         Top             =   1245
         Width           =   795
      End
      Begin VB.Label Label133 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Description"
         Height          =   195
         Left            =   480
         TabIndex        =   44
         Top             =   795
         Width           =   795
      End
      Begin VB.Label Label144 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Model/Option"
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
         Left            =   90
         TabIndex        =   43
         Top             =   300
         Width           =   1170
      End
      Begin VB.Label Label131 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Assembly"
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
         Left            =   480
         TabIndex        =   42
         Top             =   555
         Width           =   795
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid gItems 
      Height          =   2655
      Left            =   75
      TabIndex        =   37
      Top             =   4620
      Width           =   13515
      _cx             =   1990876447
      _cy             =   1990857291
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
      BackColorSel    =   14336431
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   255
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   37
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FAssembly.frx":03C8
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
      FillStyle       =   1
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   2
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
      Begin VB.PictureBox picWarningMessages 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2130
         ScaleHeight     =   585
         ScaleWidth      =   3765
         TabIndex        =   61
         Top             =   1050
         Visible         =   0   'False
         Width           =   3795
         Begin VB.TextBox lblWarningMessages 
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   360
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   62
            Text            =   "FAssembly.frx":099F
            Top             =   60
            Width           =   1095
         End
         Begin VB.Image imgTipIcon 
            Height          =   240
            Left            =   60
            Top             =   60
            Width           =   240
         End
      End
      Begin VB.Image imgInfo 
         Height          =   240
         Left            =   0
         Picture         =   "FAssembly.frx":09A5
         Top             =   450
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgWarning 
         Height          =   240
         Left            =   0
         Picture         =   "FAssembly.frx":0F2F
         Top             =   210
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid gExportData 
      Height          =   3285
      Left            =   9975
      TabIndex        =   39
      Top             =   5880
      Visible         =   0   'False
      Width           =   3045
      _cx             =   1990857979
      _cy             =   1990858402
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
      Cols            =   12
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FAssembly.frx":14B9
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
End
Attribute VB_Name = "FAssembly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const SRCFILE = "FAssembly::"


Private mDirty         As Boolean
Private mReadOnly      As Boolean
Private mAssemblyID    As Long
Private mCommunity     As String
Private mModel         As String
Private mOptionID      As String
Private mAssembly      As String
Private mAssemblyDesc  As String
Private mAssemblyType  As AssemblyTypes
Private mDataSource    As Long '0=estimating, 1 = salesdata
Private mTimerTask     As String
Private GroupedColumns As Long 'number of grouped columns
Private mCopiedAssemblyID As Long
Private mCopiedCommunity  As String
Private mCopiedModel      As String
Private mCopiedOption     As String
Private mCopiedAssembly   As String
Private mCopiedKey As String 'concatenation of community/model/optionid/assembly cleared on load, reset on save. make sure user has changed something on "copy of"
Private mSortCol As Long
Private mComponentsEnabled As Boolean 'assembly cant be both a parent and a child
Private mUseComponents As Boolean 'components are not used. hide them.
Private mUseBilling As Boolean

Private mImportGlobals As Boolean 'last import type run. see toolbar_buttonclick and toolbar_menuclick

'new menu constants
Private Const mcNEW_MODEL = 0
Private Const mcNEW_OPTION = 1
Private Const mcNEW_GLOBAL = 2
Private Const mcNEW_DCOPTION = 3
Private Const mcNEW_COPYOF = 5

'components menu constants
Private Const mcCOMP_ADD = 0
Private Const mcCOMP_REMOVE = 1

'items menu constants
Private Const mcITEM_COPY = 0
Private Const mcITEM_SUBSTITUE = 1
Private Const mcITEM_REMOVE = 2
Private Const mcITEM_EDITITEMCHART = 3
Private Const mcITEM_VIEWFILES = 5

Private mOrigPrice As Double
Private mFunnyFlag As Boolean ' used to ignore the first mouse move event see gitems_mousemove and gitems_rowcolchanged

Public Function community() As String
    community = GetComboBoxListKey(cboCommunity)
End Function
Public Function Assembly() As String
    Assembly = txtAssembly.Text
End Function
Public Function OptionID() As String
    OptionID = txtOption.Text
End Function

Private Sub cboBillingCategory_Click()
    mDirty = True
End Sub

Private Sub cboBillingCode_Click()
    mDirty = True
End Sub

Private Sub cboCommunity_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = vbKeyDelete Then
        cboCommunity.ListIndex = -1
        mDirty = True
    End If
End Sub

Private Sub cboLocation_Change()
mDirty = True
Me.Caption = Date
End Sub

Private Sub cboLocation_Click()
mDirty = True
Me.Caption = Date
End Sub

Private Sub cboLocation_GotFocus()
SelectAll cboLocation
End Sub


Private Sub cboPriceCommunity_Click()
    If SaveData(True) Then Call LoadData
End Sub

Private Sub cboScheduleTemplate_Click()
    mDirty = True
End Sub

Private Sub cboStyle_Click()
    mDirty = True
End Sub

Private Sub cboUOM_Click()
    mDirty = True
End Sub

Private Sub cboUOM_Validate(Cancel As Boolean)
    mDirty = True
    cboUOM.Text = Left(cboUOM.Text, 10)
End Sub


Private Sub chkActive_Click()
    mDirty = True
End Sub

Private Sub chkIncludedOption_Click()
    mDirty = True
End Sub


Private Sub chkDCSalesOnly_Click()
    mDirty = True
End Sub
Private Sub chkSelectByRoom_Click()
    mDirty = True
End Sub
Private Sub chkDisplayTotalOnly_Click()
    mDirty = True
End Sub


Private Sub chkTakeoffRequired_Click()
    mDirty = True
End Sub


Private Sub chkUseNormalSalesQtyFactors_Click()
    mDirty = True
End Sub
Private Sub chkIsBaseAssembly_Click()
    mDirty = True
End Sub


Private Sub cmdBrowse_Click(Index As Integer)
    Dim s As String
    
    Select Case Index
    Case 0
        s = txtGraphicPath.Text
        If VBGetOpenFileName(s, , , , , , "All Files (*.*)|*.*", , , , , FMain.hwnd) Then
            txtGraphicPath.Text = s
            mDirty = True
        End If
    Case 1
        s = txtSpecDocument.Text
        If VBGetOpenFileName(s, , , , , , "All Files (*.*)|*.*", , , , , FMain.hwnd) Then
            txtSpecDocument.Text = s
            mDirty = True
        End If
    End Select
        
        
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Select Case True
        Case KeyCode = vbKeyO And Shift = vbCtrlMask:      Call Toolbar_ButtonClick(Toolbar.Buttons("Open"))
        Case KeyCode = vbKeyS And Shift = vbCtrlMask:      Call Toolbar_ButtonClick(Toolbar.Buttons("Save"))
    End Select
End Sub

Private Sub Form_Load()
    Dim s As String
    Dim t As String
    Dim i As Long

    mUseComponents = HFApp.Options.ValueByName("UseComponents") = "true"
    mUseBilling = HFApp.Options.ValueByName("PostAssembliesAsSalesInvoices") = "true"

    HeaderFrame.BackColor = vbButtonFace
    frmEstimatorNotes.BackColor = vbButtonFace
    frmModelFields.BackColor = vbButtonFace
    frmOptionFields.BackColor = vbButtonFace
    frmSeriesFields.BackColor = vbButtonFace

    Call SetToolbarIcons(Toolbar, FMain.LargeIcons)
    Call LoadCustomDescriptions
    Call LoadWBSDescriptions
    
    Toolbar.Buttons("TakeoffPlanSwift").Visible = TakeoffSystem = tsPlanSwift
    
    Call IniGetGrid(Me, gItems, , , , True)
    Call IniGetGrid(Me, gComponents, , , , True)
    Call IniGetForm(Me)
    gItems.Rows = 1
    ReadOnly = True
    Me.Caption = IIf(HFApp.Options.ValueByName("BuilderType") = "Commercial", "Assembly Library ", "Model and Option Library")
    Me.Show
    
    
    If HFApp.Options.ValueByName("HideElevation") = "True" Then
        lblElevation.Visible = False
        txtElevation.Visible = False
    End If
    
    
    If HFApp.Options(SalesSystem) = asBuilder1440 Then
        cboStyle.AddItem "Single-Family"
        cboStyle.AddItem "Condo"
        cboStyle.AddItem "Townhouse"
        cboStyle.Visible = True
        txtStyle.Visible = False
    Else
        cboStyle.Visible = False
        txtStyle.Visible = True
    End If
    
    s = HFApp.Options.ValueByName("ScheduleTemplates")
    cboScheduleTemplate.Clear
    Call cboScheduleTemplate.AddItem("")
    For i = 1 To Parse(s, , "|")
        t = Parse(s, i, "|")
        If t <> "" Then Call cboScheduleTemplate.AddItem(t)
    Next
    
    Call LoadComboBox(cboUOM, HFApp.Databases(dbHomefront), "SELECT DISTINCT ISNULL(AssemblyUOM,''),'',0 FROM tblDBAssemblyMaster where divisionid=" & DbQuote(Num, HFApp.DivisionID) & " ORDER BY 1")
    Call LoadComboBox(cboModel, HFApp.Databases(dbHomefront), "select distinct model,'',0 from tbldbassemblymaster where assemblytype=0 and divisionid=" & DbQuote(Num, HFApp.DivisionID) & " order by 1")
    
    s = "select '','',0 union all SELECT isnull(c.area,'') + ' - ' + isnull(c.description,''),c.area,0 FROM tblLocality c"
    If HFApp.DivisionID <> "" Then
        s = s & " left outer join DivisionCommunities d on d.Community = c.Area" & vbCrLf
    End If
    s = s & " where c.inactive = 0 "
    If HFApp.DivisionID <> "" Then
        s = s & " and d.DivisionID=" & HFApp.DivisionID & vbCrLf
    End If
    s = s & " order by 1"
    Call LoadComboBox(cboCommunity, HFApp.Databases(dbHomefront), s)
    Call LoadComboBox(cboLocation, HFApp.Databases(dbHomefront), "select description,area,0 from tblareas order by description")
     
    
    frmBilling.Visible = HFApp.Options.ValueByName("PostAssembliesAsSalesInvoices") = "true"
    Call LoadComboBox(cboBillingCode, HFApp.Databases(dbHomefront), "select description,costcode,0 from standardcostcodes where divisionid=" & DbQuote(Num, HFApp.DivisionID) & " order by description")
    Call LoadComboBox(cboBillingCategory, HFApp.Databases(dbHomefront), "select description,category,0 from standardcategories where divisionid=" & DbQuote(Num, HFApp.DivisionID) & " order by description")
     
     
    Call LoadSeries
    Call LoadCategories

    i = Val("" & HFApp.Options.ValueByName("EstimatingSystem"))
    Toolbar.Buttons("Import").Visible = i = 1
    
End Sub

Private Sub LoadPricingCommunities(LimitToCommunity As String)
    Dim s As String
    
    s = ""
    s = s & "select '','',0" & vbCrLf
    s = s & "union" & vbCrLf
    s = s & "select c.Description,c.area + char(2),0" & vbCrLf
    s = s & "from tbllocality c " & vbCrLf
    If HFApp.DivisionID <> "" Then
        s = s & " left outer join DivisionCommunities d on d.Community = c.Area" & vbCrLf
    End If
    s = s & "where isnull(c.inactive,0)=0" & vbCrLf
    If HFApp.DivisionID <> "" Then
        s = s & " and d.DivisionID=" & HFApp.DivisionID & vbCrLf
    End If
    If LimitToCommunity <> "" Then
        s = s & " and c.area=" & DbQuote(Str, LimitToCommunity) & vbCrLf
    End If
    s = s & "union" & vbCrLf
    s = s & "select c.description + isnull(' ' + p.description,''),c.area + char(2) + isnull(p.communityphase,''),0" & vbCrLf
    s = s & "from tbllocality c left outer join communityphase p on(c.area=p.community)" & vbCrLf
    If HFApp.DivisionID <> "" Then
        s = s & " left outer join DivisionCommunities d on d.Community = c.Area" & vbCrLf
    End If
    s = s & "where isnull(c.inactive,0)=0" & vbCrLf
    If HFApp.DivisionID <> "" Then
        s = s & " and d.DivisionID=" & HFApp.DivisionID & vbCrLf
    End If
    If LimitToCommunity <> "" Then
        s = s & " and c.area=" & DbQuote(Str, LimitToCommunity) & vbCrLf
    End If
    s = s & "order by 1"
    Call LoadComboBox(cboPriceCommunity, HFApp.Databases(dbHomefront), s)

End Sub


Private Sub LoadSeries()
    
    lblSeries.FontUnderline = HFApp.Options(SalesSystem) <> asHomeFront
    lblSeries.ForeColor = IIf(HFApp.Options(SalesSystem) <> asHomeFront, &HFF0000, vbButtonText)
    
    Call LoadComboBox(cboSeries, HFApp.Databases(dbHomefront), "SELECT isnull(description,series),series,0 FROM tblSeries where divisionid = " & HFApp.DivisionID & " order by 1")
End Sub

Private Sub LoadCategories()
    Call LoadComboBox(cboCategory, HFApp.Databases(dbHomefront), "SELECT isnull(description,category),category,0 FROM tblcategories where isnull(inactive,0)=0 order by 1")
End Sub
Public Function ShowForm(SaveAs As Boolean, Optional community As String, Optional Model As String, Optional OptionID As String, Optional Assembly As String, Optional AssemblyDesc As String, Optional AssemblyType As AssemblyTypes) As Boolean
    
'    mDirty = False
    ShowForm = True
    
'Model = "model"
'OptionID = "optionid"
'Assembly = "assembly"
'community = "community"
    
    If community & Model & OptionID & Assembly = "" And Not SaveAs Then
        mCopiedCommunity = ""
        mCopiedAssembly = ""
        mCopiedAssemblyID = 0
        mCopiedModel = ""
        mCommunity = ""
        mAssembly = ""
        mModel = ""
        mOptionID = ""
'        mDataSource = 1 '-- why is this defaulting to the sales assembly list? dont know. dont like it. changed it back to dbassembly. jan 16, 2015
        mDataSource = 0
        mAssemblyType = atModel
        
        Call LoadPricingCommunities("")
        cboPriceCommunity.ListIndex = -1
        Call LoadData
        ReadOnly = True
        Call Toolbar_ButtonClick(Toolbar.Buttons("Open"))
    Else
    
        mCopiedCommunity = ""
        mCopiedAssembly = ""
        mCopiedAssemblyID = 0
        mCopiedModel = ""
    
        mCommunity = community
        mAssembly = Assembly
        mAssemblyDesc = AssemblyDesc
        mAssemblyType = AssemblyType
        mModel = Model
        mOptionID = OptionID
        
        'if this is a manual takeoff then assume it is a global
        If mAssemblyType = atCustomOption Then mAssemblyType = atGlobal
        
        'ignore model if it's not model specific
        Select Case mAssemblyType
            Case atGlobal, atDesignCenter
                mModel = ""
        End Select
        
        Call LoadData
        If gItems.Rows > 1 Then
            If vbYes = MsgBox("This assembly already exists." & vbCrLf & vbCrLf & "Do you want to replace it with the items from the customer record?", vbExclamation + vbYesNo, App.ProductName) Then
                gItems.Rows = 1
                mDirty = False
            Else
                ShowForm = False
            End If
        End If
        
    End If
    
End Function



Private Sub Form_Resize()
On Error Resume Next

    
    Select Case mAssemblyType
    
        Case atModel
            frmSeriesFields.Visible = True
            frmModelFields.Visible = True
            frmOptionFields.Visible = False
            
        Case atoption
            frmSeriesFields.Visible = True
            frmModelFields.Visible = False
            frmOptionFields.Visible = True
            
        Case atDesignCenter
            frmSeriesFields.Visible = False
            frmModelFields.Visible = False
            frmOptionFields.Visible = True
            
        Case atGlobal
            frmSeriesFields.Visible = False
            frmModelFields.Visible = False
            frmOptionFields.Visible = True
            
    End Select
    
    HeaderFrame.Move 0, Toolbar.Height + 120, Me.ScaleWidth
    Slider.Min = HeaderFrame.Top + HeaderFrame.Height + 200
    Slider.Max = Me.ScaleHeight - 500
    Slider.Move 0, Max(Min(Slider.Top, Slider.Max), Slider.Min), Me.ScaleWidth
    gComponents.Move 0, HeaderFrame.Top + HeaderFrame.Height, Me.ScaleWidth, Slider.Top - (HeaderFrame.Top + HeaderFrame.Height)
    gItems.Move 0, Slider.Top + Slider.Height, Me.ScaleWidth, Me.ScaleHeight - (Slider.Top + Slider.Height)
    
    Slider.Visible = mComponentsEnabled
    gComponents.Visible = mComponentsEnabled
    lblComponents.Visible = mComponentsEnabled
    If Not mComponentsEnabled Then
        gItems.Move 0, gComponents.Top, Me.ScaleWidth, Me.ScaleHeight - gComponents.Top
    End If
    
    'old label
    lblAssemblyCost(1).Visible = Not mComponentsEnabled
    
    'new assembly+component=total labels
    lblAssemblyCost(0).Visible = mComponentsEnabled
    lblComponentCost.Visible = mComponentsEnabled
    lblTotalCost.Visible = mComponentsEnabled
    Label18(3).Visible = mComponentsEnabled
    Label18(2).Visible = mComponentsEnabled
    Label18(1).Visible = mComponentsEnabled
        
    frmBilling.Visible = mUseBilling
    
    With frmModelFields
        frmOptionFields.Move .Left, .Top
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not SaveData(True) Then
        Cancel = True
        Exit Sub
    End If
    Call IniPutGrid(Me, gItems)
    Call IniPutGrid(Me, gComponents)
    Call IniPutForm(Me)
End Sub


Private Sub gComponents_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gComponents
    Select Case .ColKey(Col)
    Case "qty"
    Case Else
        Cancel = True
    End Select
    End With
End Sub

Private Sub gComponents_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
    With gComponents
        If Not .Enabled Then Exit Sub
        If Button = vbRightButton Then
            If .MouseRow < 1 And GridLastVisibleRow(gComponents) <> 0 Then
                Cancel = True
                Call FMain.ShowColumnMenu(gComponents, , , , , False)
            Else
                Call PopupMenu(FMain.mnuAssemblyComponents)
            End If
        End If
    End With
End Sub

Private Sub gComponents_KeyDown(KeyCode As Integer, Shift As Integer)
    With gComponents
    If Not .Enabled Then Exit Sub
    If KeyCode = vbKeyDelete And Shift <> 0 And .Row > 0 Then
        .RowData(.Row) = "delete"
        .RowHidden(.Row) = True
        mDirty = True
    End If
    End With
End Sub

Private Sub gComponents_SelChange()
    gComponents.ColSel = gComponents.Col
End Sub

Private Sub gComponents_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gComponents
        .EditText = Val(.EditText)
        mDirty = True
        If .RowData(Row) = "" Then .RowData(Row) = "dirty"
    End With
End Sub

Private Sub gItems_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    Dim r As Long
    Dim TakeoffQty As Double
    Dim OrderQty As Double
    Dim Conversion As Double
    Dim RoundDir As Long
    Dim RoundUnit As Double
    Dim WastePercent As Long
    Dim cost As Double
    Dim DoRegroup As Boolean
    
    DoRegroup = False
    With gItems
        Select Case .ColKey(Col)
            Case "Invertable"
                If .RowSel = 0 Then .RowSel = Row
                
            Case "Vendor", "VendorDesc"
                Exit Sub
                
            Case "AssemblyConversionFactor"
                
            Case "Price"
                For r = Min(Row, .RowSel) To Max(Row, .RowSel)
                    .TextMatrix(r, .ColIndex("Price")) = Val(.ValueMatrix(r, .ColIndex("Price")))
                    .TextMatrix(r, .ColIndex("TaxAmount")) = .ValueMatrix(r, .ColIndex("Price")) * ((.ValueMatrix(r, .ColIndex("TaxRate")))) * .ValueMatrix(r, .ColIndex("OrderQty"))
                    .TextMatrix(r, .ColIndex("PretaxAmount")) = .ValueMatrix(r, .ColIndex("Price")) * .ValueMatrix(r, .ColIndex("OrderQty"))
                    .TextMatrix(r, .ColIndex("ExtendedAmount")) = .ValueMatrix(r, .ColIndex("Price")) * ((1 + .ValueMatrix(r, .ColIndex("TaxRate")))) * .ValueMatrix(r, .ColIndex("OrderQty"))
                    cost = Val(Replace(Replace(lblAssemblyCost(0).Caption, "$", ""), ",", "")) + (Val(.ValueMatrix(Row, .ColIndex("Price"))) * (1 + Val(.ValueMatrix(Row, .ColIndex("TaxRate")))) * Val(.ValueMatrix(Row, .ColIndex("OrderQty"))))  '(Val("" & .ValueMatrix(Row, .ColIndex("Price"))) * Val("" & .ValueMatrix(Row, .ColIndex("OrderQty"))))
                    .Cell(flexcpData, r, .ColIndex("Price")) = "Dirty"
                    lblAssemblyCost(0).Caption = format(cost, "$###,###.00")
                Next
                DoRegroup = True
                
                
            Case "ItemType"
                For r = Min(Row, .RowSel) To Max(Row, .RowSel)
                    .Cell(flexcpData, r, .ColIndex("ItemType")) = "Dirty"
                    .Cell(flexcpData, r, .ColIndex("Price")) = "Dirty"
                Next
                
            Case "PriceLevel", "Vendor", "VendorDesc"
                For r = Min(Row, .RowSel) To Max(Row, .RowSel)
                    .Cell(flexcpData, r, .ColIndex("Price")) = "Dirty"
                Next
                
            Case "TakeoffQty"
                For r = Min(Row, .RowSel) To Max(Row, .RowSel)
                    TakeoffQty = Val(.ValueMatrix(r, .ColIndex("TakeoffQty")))
                    Conversion = IIf(.ValueMatrix(r, .ColIndex("AssemblyConversionFactor")) = 0, .ValueMatrix(r, .ColIndex("ItemConversionFactor")), .ValueMatrix(r, .ColIndex("AssemblyConversionFactor")))
                    If Conversion = 0 Then Conversion = 1
                    RoundDir = Val(.TextMatrix(r, .ColIndex("RoundDir")))
                    RoundUnit = Val(.ValueMatrix(r, .ColIndex("RoundTo")))
                    WastePercent = Val(.ValueMatrix(r, .ColIndex("WastePercent")))
                    OrderQty = RoundTo(TakeoffQty * (100 + WastePercent) / 100 * Conversion, RoundUnit, RoundDir)
                    .TextMatrix(r, .ColIndex("TakeoffQty")) = TakeoffQty
                    .TextMatrix(r, .ColIndex("OrderQty")) = OrderQty
                    .TextMatrix(r, .ColIndex("TaxAmount")) = .ValueMatrix(r, .ColIndex("Price")) * ((.ValueMatrix(r, .ColIndex("TaxRate")))) * .ValueMatrix(r, .ColIndex("OrderQty"))
                    .TextMatrix(r, .ColIndex("PretaxAmount")) = .ValueMatrix(r, .ColIndex("Price")) * .ValueMatrix(r, .ColIndex("OrderQty"))
                    .TextMatrix(r, .ColIndex("ExtendedAmount")) = .ValueMatrix(r, .ColIndex("Price")) * ((1 + .ValueMatrix(r, .ColIndex("TaxRate")))) * .ValueMatrix(r, .ColIndex("OrderQty"))
                    cost = Val(Replace(Replace(lblAssemblyCost(0).Caption, "$", ""), ",", "")) + (Val(.ValueMatrix(Row, .ColIndex("Price"))) * (1 + Val(.ValueMatrix(Row, .ColIndex("TaxRate")))) * Val(.ValueMatrix(Row, .ColIndex("OrderQty"))))  '(Val("" & .ValueMatrix(Row, .ColIndex("Price"))) * Val("" & .ValueMatrix(Row, .ColIndex("OrderQty"))))
                    
                    lblAssemblyCost(0).Caption = format(cost, "$###,###.00")
                Next
                DoRegroup = True
                
                
            Case "OrderQty"
                For r = Min(Row, .RowSel) To Max(Row, .RowSel)
                    If .RowData(r) <> "new" Then .RowData(r) = "update"
                    
                    OrderQty = Val(.ValueMatrix(r, .ColIndex("OrderQty")))
                    Conversion = IIf(.ValueMatrix(r, .ColIndex("AssemblyConversionFactor")) = 0, .ValueMatrix(r, .ColIndex("ItemConversionFactor")), .ValueMatrix(r, .ColIndex("AssemblyConversionFactor")))
                    If Conversion = 0 Then Conversion = 1
                    WastePercent = Val(.ValueMatrix(r, .ColIndex("WastePercent")))
                    
                    TakeoffQty = Round(OrderQty / ((100 + WastePercent) / 100) / Conversion, 5)
                    .TextMatrix(r, .ColIndex("TakeoffQty")) = TakeoffQty
                    .TextMatrix(r, .ColIndex("OrderQty")) = OrderQty
                    .TextMatrix(r, .ColIndex("TaxAmount")) = .ValueMatrix(r, .ColIndex("Price")) * ((.ValueMatrix(r, .ColIndex("TaxRate")))) * .ValueMatrix(r, .ColIndex("OrderQty"))
                    .TextMatrix(r, .ColIndex("PretaxAmount")) = .ValueMatrix(r, .ColIndex("Price")) * .ValueMatrix(r, .ColIndex("OrderQty"))
                    .TextMatrix(r, .ColIndex("ExtendedAmount")) = .ValueMatrix(r, .ColIndex("Price")) * ((1 + .ValueMatrix(r, .ColIndex("TaxRate")))) * .ValueMatrix(r, .ColIndex("OrderQty"))
                    cost = Val(Replace(Replace(lblAssemblyCost(0).Caption, "$", ""), ",", "")) + (Val(.ValueMatrix(Row, .ColIndex("Price"))) * (1 + Val(.ValueMatrix(Row, .ColIndex("TaxRate")))) * Val(.ValueMatrix(Row, .ColIndex("OrderQty"))))  '(Val("" & .ValueMatrix(Row, .ColIndex("Price"))) * Val("" & .ValueMatrix(Row, .ColIndex("OrderQty"))))
                    
                    lblAssemblyCost(0).Caption = format(cost, "$###,###.00")
                    
                Next
                DoRegroup = True
                
        End Select
    
        For r = Min(Row, .RowSel) To Max(Row, .RowSel)
            If .RowData(r) <> "new" Then .RowData(r) = "update"
        Next
        
        If DoRegroup Then Call GroupGrid
    
    End With
    
    lblAssemblyCost(1).Caption = lblAssemblyCost(0).Caption
    
    cost = Val(Replace(Replace(lblAssemblyCost(0).Caption, "$", ""), ",", "")) + Val(Replace(Replace(lblComponentCost.Caption, "$", ""), ",", ""))
    lblTotalCost.Caption = format(cost, "$###,###.00")
    
    mDirty = True
End Sub


Private Sub gItems_AfterMoveColumn(ByVal Col As Long, Position As Long)
    Call GroupGrid
End Sub

Private Sub gItems_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error Resume Next
    
    With gItems
        If OldRow > 0 Then
            .Cell(flexcpBackColor, OldRow, 0, OldRow, .Cols - 1) = vbWindowBackground
        End If
        If NewRow > 0 Then
            .Cell(flexcpBackColor, NewRow, 0, NewRow, .Cols - 1) = .BackColorSel
        End If
    End With

End Sub

Private Sub gItems_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim s As String
    Dim PriceLevel As String
    
    With gItems
        .ComboList = ""
        .AutoSearch = flexSearchNone
        
        
        If .IsSubtotal(Row) Then
            Cancel = True
            Exit Sub
        End If
        
        Select Case .ColKey(Col)
            Case "AssemblyConversionFactor"
            
            Case "Invertable"
                Cancel = .TextMatrix(Row, .ColIndex("hasInverseItem")) <> "1"
            
            Case "Price"
                Cancel = PricingCommunity = "" Or .TextMatrix(Row, .ColIndex("itemtype")) = "" Or .TextMatrix(Row, .ColIndex("pricelevel")) = ""
                mOrigPrice = .ValueMatrix(Row, .ColIndex("Price"))
                
            Case "PriceLevel"
                Cancel = PricingCommunity = ""
                If .TextMatrix(Row, .ColIndex("Vendor")) = "" Then
                    .ComboList = "#4;Item DB"
                Else
                    If .TextMatrix(Row, .ColIndex("ItemType")) = "Quote" Then
                        If PricingCommunityPhase <> "" Then
                            PriceLevel = "Global (any " & FMain.CD_Community & ")|Item DB|Community Specific|Phase Specific"
                        Else
                            PriceLevel = "Global (any " & FMain.CD_Community & ")|Item DB|Community Specific"
                        End If
                    Else
                        If PricingCommunityPhase <> "" Then
                            PriceLevel = "Corporate (any Division)|Global (any " & FMain.CD_Community & ")|Item DB|Community Specific|Phase Specific"
                        Else
                            PriceLevel = "Corporate (any Division)|Global (any " & FMain.CD_Community & ")-|Item DB|Community Specific"
                        End If
                    End If
                    PriceLevel = Replace(PriceLevel, "Community Specific", FMain.CD_Community & " Specific", 1, 1, vbTextCompare)
                    .ComboList = PriceLevel
                End If
                
            Case "ItemType":      .ComboList = "Quote|Unit Price"
            'Case "Notes":         .ComboList = "|..."
            Case "OrderQty":
            Case "TakeoffQty":
            Case "UseModelCost":    Cancel = mAssemblyType <> atoption
            Case "Notes", "Formula":    .ComboList = "|..."
            Case "POIndex":     .ComboList = "..."
        
            Case "Vendor":        .ComboList = "|...": Cancel = PricingCommunity = ""
            Case "VendorDesc":    .ComboList = "...": Cancel = PricingCommunity = ""
            
            Case "Location", "WBS01", "WBS02", "WBS03", "WBS04", "WBS05", "WBS06", "WBS07", "WBS08", "WBS09", "WBS10", "WBS11", "WBS12", "WBS13", "WBS14", "WBS15", "WBS16", "WBS17", "WBS18", "WBS19", "WBS20", "WBS21", "WBS22", "WBS23", "WBS24", "WBS25", "WBS26", "WBS27", "WBS28", "WBS29", "WBS30", "WBS31", "WBS32", "WBS33", "WBS34", "WBS35", "WBS36", "WBS37", "WBS38", "WBS39", "WBS40"
                .EditMaxLength = 50
            
            Case Else
                Cancel = True
                .AutoSearch = flexSearchFromCursor
        End Select
    End With
End Sub


Private Sub gItems_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim s As String
    Dim f As New FFormulaEditor
    Dim i As Long
    
    With gItems
        Select Case .ColKey(Col)
        
            Case "POIndex"
                s = "SELECT POIndex ID,FullDescription POIndex FROM tblPOIndex where DivisionID = " & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Purchase Order", s, gItems, , , , "ID") Then
                    .Cell(flexcpData, .Row, Col, .RowSel, Col) = FPickList.SelectedItem("ID")
                    .Cell(flexcpText, .Row, Col, .RowSel, Col) = FPickList.SelectedItem("POIndex")
                    
                    For i = Min(.Row, .RowSel) To Max(.Row, .RowSel)
                        If .RowData(i) = "" Then .RowData(i) = "update"
                    Next
                    
                    mDirty = True
                End If
                
                
            Case "Vendor"
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Vendor", "select Vendor_name Company,vendor_id Vendor from tblvendors where DivisionID = " & HFApp.DivisionID & " and isnull(inactive,0)=0") Then
                    .Text = FPickList.SelectedItem("Vendor")
                    .Cell(flexcpText, Min(.Row, .RowSel), .ColIndex("VendorDesc"), Max(.Row, .RowSel), .ColIndex("VendorDesc")) = FPickList.SelectedItem("Company")
                    For i = Min(.Row, .RowSel) To Max(.Row, .RowSel)
                        If .RowData(i) = "" Then .RowData(i) = "update"
                    Next
                    Call gItems_AfterEdit(Row, Col)
                    Call RefreshVendorPrice(Row, Col)
                    Call gItems_AfterEdit(Row, Col)
                End If
                
            Case "VendorDesc"
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Vendor", "select Vendor_name Vendor,vendor_id  from tblvendors where DivisionID = " & HFApp.DivisionID & " and isnull(inactive,0)=0", , , , , "vendor_id") Then
                    .Text = FPickList.SelectedItem("Vendor")
                    .Cell(flexcpText, Min(.Row, .RowSel), .ColIndex("Vendor"), Max(.Row, .RowSel), .ColIndex("Vendor")) = FPickList.SelectedItem("Vendor_ID")
                    For i = Min(.Row, .RowSel) To Max(.Row, .RowSel)
                        If .RowData(i) = "" Then .RowData(i) = "update"
                    Next
                    Call RefreshVendorPrice(Row, Col)
                    Call gItems_AfterEdit(Row, Col)
                End If
                
            Case "Formula"
                s = .Text
                If f.EditFormula("", s) Then
                    .Text = s
                    For i = Min(.Row, .RowSel) To Max(.Row, .RowSel)
                        If .RowData(i) = "" Then .RowData(i) = "update"
                    Next
                    mDirty = True
                End If
                
            Case "Notes"
                s = .Text
                If FComments.Edit(s, gItems, , "Notes", 4000) Then
                    .Text = s
                    For i = Min(.Row, .RowSel) To Max(.Row, .RowSel)
                        If .RowData(i) = "" Then .RowData(i) = "update"
                    Next
                    mDirty = True
                End If
                
        End Select
    End With
End Sub

Private Sub gitems_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, done As Boolean)
    With gItems
        If .ColData(Col) = "GROUPED" And Row > .FixedRows And Not .IsSubtotal(Row) Then
            done = True
        End If
        If .IsSubtotal(Row) And .ColData(Col) = "GROUPED" And .RowOutlineLevel(Row) >= 1 And Col < .RowOutlineLevel(Row) Then
            done = True
        End If
    End With
End Sub



Private Sub gItems_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo eh
    Dim r As Long
    Dim s As String
    Dim Cancel As Boolean

    With gItems
    Select Case True

        Case KeyCode = vbKeyF And Shift = vbCtrlMask
            Call FFind.ShowForm(gItems)

        Case KeyCode = vbKeyDelete And Shift = vbCtrlMask
            If Not ReadOnly Then
            For r = Max(.Row, .RowSel) To Min(.Row, .RowSel) Step -1
                If r < 1 Then Exit Sub
                gItems.RowData(r) = "delete"
                gItems.RowHidden(r) = True
                mDirty = True
            Next
            .Row = .Row
            End If

        Case Shift = vbCtrlMask And KeyCode = vbKeyC
            s = .clip
            s = Replace(s, Chr(10), "")
            s = Replace(s, Chr(13), vbCrLf)
            Call Clipboard.SetText(s)

        Case Shift = vbCtrlMask And KeyCode = vbKeyV And Not ReadOnly And .Row > 0 And .RowSel > 0
            mDirty = True
            .clip = Clipboard.GetText

    End Select
    End With
Exit Sub
eh: Call errHandler(SRCFILE & "gItems_KeyDown")
End Sub


Private Sub gItems_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    If Button <> 0 Then Exit Sub
    With gItems
        If mFunnyFlag Then
            mFunnyFlag = False
            Exit Sub
        End If
        If .MouseRow >= Min(.Row, .RowSel) And .MouseRow <= Max(.Row, .RowSel) And .MouseCol = .Col Then
            'Call gItems_SelChange
        Else
            Call ShowTip(.MouseRow, .MouseCol, MouseX(gItems.hwnd) * Screen.TwipsPerPixelX + 315, MouseY(gItems.hwnd) * Screen.TwipsPerPixelY + 315)
        End If
    End With
End Sub

Private Sub gItems_RowColChange()
On Error Resume Next
    With gItems
        mFunnyFlag = True
        Call ShowTip(.Row, .Col, .colPos(.Col) + 180, .RowPos(.Row) + .RowHeight(.Row) + 180)
    End With
End Sub

Private Sub gItems_SelChange()
   
   'limit selection to one column
    Static bInHere As Boolean
    If bInHere Then Exit Sub
    bInHere = True
    gItems.ColSel = gItems.Col




    Dim r As Long
    
    Dim pretax As Double
    Dim tax As Double
    Dim tip As String
    Dim qty As Double
    
    
    With gItems
        If .Row <> .RowSel Then
            For r = Min(.Row, .RowSel) To Max(.Row, .RowSel)
            If .RowHidden(r) = False Then
                pretax = pretax + (.ValueMatrix(r, .ColIndex("OrderQty")) * .ValueMatrix(r, .ColIndex("Price")))
                tax = tax + .ValueMatrix(r, .ColIndex("Price")) * ((.ValueMatrix(r, .ColIndex("TaxRate")))) * .ValueMatrix(r, .ColIndex("OrderQty"))
                If IsIn(.ColKey(.Col), "TakeoffQty", "OrderQty") Then
                    qty = qty + .ValueMatrix(r, .Col)
                End If
            End If
            Next
            
                

            
            tip = IIf(IsIn(.ColKey(.Col), "TakeoffQty", "OrderQty"), .ColKey(.Col) & ":" & vbTab & qty & vbCrLf & vbCrLf, "") & _
                  "Pretax:" & vbTab & format(pretax, "#,##0.00") & vbCrLf & _
                  "Tax:" & vbTab & format(tax, "#,##0.00") & vbCrLf & _
                  "Total:" & vbTab & format(pretax + tax, "#,##0.00")
            
            Call ShowTip(.MouseRow, .MouseCol, MouseX(gItems.hwnd) * Screen.TwipsPerPixelX + 315, MouseY(gItems.hwnd) * Screen.TwipsPerPixelY + 315, tip)
            
        End If
    End With
    
    
    bInHere = False
End Sub


Private Sub RefreshVendorPrice(Row As Long, Col As Long)
    Dim s As String
    Dim i As Long
    Dim rs As Recordset
    Dim Vendor As String
    Dim r As Long
    Dim taxrate As Double
    
    With gItems
        'retrieve price
        Vendor = .TextMatrix(Row, .ColIndex("Vendor"))
        If PricingCommunity <> "" And Vendor <> "" Then
            For r = Min(.Row, .RowSel) To Max(.Row, .RowSel)
                
                s = ""
                s = s & "SELECT isnull(dbo.Purch_GetItemRate(0,0," & DbQuote(Str, PricingCommunity) & ", " & DbQuote(Str, PricingCommunityPhase) & ", d.Assembly,d.Model,d.OptionID,d.Phase,d.Item,d.Sequence,v.Vendor_id,getdate()," & HFApp.DivisionID & "),0)        Rate" & vbCrLf
                s = s & "      ,dbo.Purch_GetItemRateQuality(0,0," & DbQuote(Str, PricingCommunity) & ", " & DbQuote(Str, PricingCommunityPhase) & ", d.Assembly,d.Model,d.OptionID,d.Phase,d.Item,d.Sequence,v.Vendor_id,getdate()," & HFApp.DivisionID & ") RateQuality" & vbCrLf
                s = s & "  FROM tblDBAssemblyDetails d" & vbCrLf
                s = s & "       LEFT OUTER JOIN tblvendors v ON(v.DivisionID = d.DivisionID and v.vendor_id=" & DbQuote(Str, Vendor) & ")" & vbCrLf
                s = s & " WHERE d.DivisionID = " & HFApp.DivisionID & " and isnull(d.Community,'')=" & DbQuote(Str, mCommunity) & vbCrLf
                s = s & "   AND isnull(d.Model,'')=" & DbQuote(Str, mModel) & vbCrLf
                s = s & "   AND isnull(d.Optionid,'')=" & DbQuote(Str, mOptionID) & vbCrLf
                s = s & "   AND isnull(d.Assembly,'')=" & DbQuote(Str, mAssembly) & vbCrLf
                s = s & "   AND isnull(d.Phase,'')=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Phase"))) & vbCrLf
                s = s & "   AND isnull(d.Item,'')=" & DbQuote(Str, .TextMatrix(r, .ColIndex("item"))) & vbCrLf
                s = s & "ORDER BY d.Phase,d.Item,d.itemchart,d.Sequence" & vbCrLf
                Set rs = HFApp.SqlExec(s)
                If Not rs.EOF Then
                
                    i = Val("" & rs("RateQuality"))
                    Select Case i
                        'assembly specific
                        Case 4:  .TextMatrix(r, .ColIndex("PriceLevel")) = "Phase Specific"
                                 .TextMatrix(r, .ColIndex("ItemType")) = "Quote"
                        Case 5:  .TextMatrix(r, .ColIndex("PriceLevel")) = Replace("Community Specific", "Community", RTrim(FMain.CD_Community), 1, 1, vbTextCompare)
                                 .TextMatrix(r, .ColIndex("ItemType")) = "Quote"
                        Case 6:  .TextMatrix(r, .ColIndex("PriceLevel")) = Replace("Global (any community)", "community", LCase(RTrim(FMain.CD_Community)), 1, 1, vbTextCompare)
                                 .TextMatrix(r, .ColIndex("ItemType")) = "Quote"
                        
                        'not assembly specific
                        Case 9:  .TextMatrix(r, .ColIndex("PriceLevel")) = "Phase Specific"
                                 .TextMatrix(r, .ColIndex("ItemType")) = "Unit Price"
                        Case 10: .TextMatrix(r, .ColIndex("PriceLevel")) = Replace("Community Specific", "Community", RTrim(FMain.CD_Community), 1, 1, vbTextCompare)
                                 .TextMatrix(r, .ColIndex("ItemType")) = "Unit Price"
                        Case 11: .TextMatrix(r, .ColIndex("PriceLevel")) = Replace("Global (any community)", "community", LCase(RTrim(FMain.CD_Community)), 1, 1, vbTextCompare)
                                 .TextMatrix(r, .ColIndex("ItemType")) = "Unit Price"
                        
                        'no vendor rate found
                        Case 12: .TextMatrix(r, .ColIndex("PriceLevel")) = "Item DB"
                    End Select
                    lblAssemblyCost(0).Caption = format(Val(Mid(Replace(lblAssemblyCost(0), ",", ""), 2, 20)) - .ValueMatrix(r, .ColIndex("Price")) * .ValueMatrix(r, .ColIndex("OrderQty")) + (rs("Rate") * .ValueMatrix(r, .ColIndex("OrderQty"))), "$###,###.00")
                    lblAssemblyCost(1).Caption = lblAssemblyCost(0).Caption
                    
                    .TextMatrix(r, .ColIndex("Price")) = "" & rs("Rate")
                    taxrate = .ValueMatrix(r, .ColIndex("TaxRate"))
                    .TextMatrix(r, .ColIndex("TaxAmount")) = (Val("" & rs("Rate")) * taxrate) * .ValueMatrix(r, .ColIndex("OrderQty"))
                    .TextMatrix(r, .ColIndex("PretaxAmount")) = Val("" & rs("Rate")) * .ValueMatrix(r, .ColIndex("OrderQty"))
                    .TextMatrix(r, .ColIndex("ExtendedAmount")) = (Val("" & rs("Rate")) * (1 + taxrate)) * .ValueMatrix(r, .ColIndex("OrderQty"))

                    
                    'save original value so we can remove old cost records
                    .TextMatrix(r, .ColIndex("OldPriceLevel")) = .TextMatrix(r, .ColIndex("PriceLevel"))
                    .TextMatrix(r, .ColIndex("OldItemType")) = .TextMatrix(r, .ColIndex("ItemType"))
                
                End If
                
            Next
        End If
    
                    
    End With
End Sub

Private Sub gItems_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim s As String
    Dim i As Long
    Dim rs As Recordset
    Dim Vendor As String
    Dim r As Long
    Dim cost As Double
    
    With gItems
        Select Case .ColKey(Col)
            Case "AssemblyConversionFactor"
                .EditText = Val(.EditText)
                If Val(.EditText) = 0 Then .EditText = ""
                
            Case "Vendor"
                'validate vendor
                s = "select vendor_name,vendor_id from tblvendors where Divisionid = " & HFApp.DivisionID & " and vendor_id=" & DbQuote(Str, .EditText)
                Set rs = HFApp.SqlExec(s)
                If rs.EOF Then
                    Call MsgBox("Vendor not found.", vbExclamation, App.ProductName)
                    Cancel = True
                    Exit Sub
                Else
                    .EditText = "" & rs(1)
                    .Cell(flexcpText, Min(.Row, .RowSel), .ColIndex("Vendor"), Max(.Row, .RowSel), .ColIndex("Vendor")) = "" & rs(1)
                    .Cell(flexcpText, Min(.Row, .RowSel), .ColIndex("VendorDesc"), Max(.Row, .RowSel), .ColIndex("VendorDesc")) = "" & rs(0)
                End If
                
                'retrieve price
                Call RefreshVendorPrice(Row, Col)
                
            Case "Price"
                If .TextMatrix(Row, .ColIndex("PriceLevel")) = "Item DB" Then
                    Call MsgBox("Item DB prices are not editable. To enter a price here you must first select a vendor then set the pricing level.", vbOKOnly + vbInformation, App.ProductName)
                    Cancel = True
                    Exit Sub
                End If
                For r = Min(Row, .RowSel) To Max(Row, .RowSel)
                    cost = Val(Replace(Replace(lblAssemblyCost(0).Caption, "$", ""), ",", "")) - (Val("" & .ValueMatrix(r, .ColIndex("Price"))) * Val("" & .ValueMatrix(r, .ColIndex("OrderQty"))))
                    lblAssemblyCost(0).Caption = format(cost, "$###,###.00")
                    lblAssemblyCost(1).Caption = format(cost, "$###,###.00")
                Next
                
            Case "OrderQty", "TakeoffQty"
                For r = Min(Row, .RowSel) To Max(Row, .RowSel)
                    cost = Val(Replace(Replace(lblAssemblyCost(0).Caption, "$", ""), ",", "")) - (Val("" & .ValueMatrix(r, .ColIndex("Price"))) * Val("" & .ValueMatrix(r, .ColIndex("OrderQty"))))
                    lblAssemblyCost(0).Caption = format(cost, "$###,###.00")
                    lblAssemblyCost(1).Caption = format(cost, "$###,###.00")
                Next
                
        End Select
        
    End With
End Sub


Private Sub lblCategory_Click()
    Dim i As Long
    'Call FDBGrid.ShowForm("Option Subcategories", "select Category,Subcategory,Description,COMarkup from OptionSubcategories order by 1,2", "OptionSubcategories", "Subcategory", , , , True, False)
    Call FDBGrid.ShowForm("Option Subcategories", "select Category,Subcategory,Description,COMarkup,Inactive from OptionSubcategories order by 1,2", "OptionSubcategories", "Subcategory", , , , True, False)
    i = cboCategory.ListIndex
    Call LoadCategories
On Error Resume Next
    cboCategory.ListIndex = i
End Sub

Private Sub lblGroup_Click()
    'Call FDBGrid.ShowForm("Option Categories", "select Category,Description from OptionCategories order by 1", "OptionCategories", "Category", , , , True, False)
    Call FDBGrid.ShowForm("Option Categories", "select Category,Description,Inactive from OptionCategories order by 1", "OptionCategories", "Category", , , , True, False)
End Sub

Private Sub lblLocation_Click()
    Dim i As Long
    Call FDBGrid.ShowForm("Room Locations", "select Area,Description from tblareas", "tblareas", "area", , , , True)
    
On Error Resume Next
    i = cboLocation.ListIndex
    Call LoadComboBox(cboLocation, HFApp.Databases(dbHomefront), "select description,area,0 from tblareas order by description")
    cboLocation.ListIndex = i

End Sub

Private Sub lblModelDimensions_Click()
    Call FModelDimensions.ShowForm(cboModel.Text)
End Sub

Private Sub lblMoreAttributes_Click()
    mDirty = True
    Call FAssemblyAttributes.EditAttributes(Me)
End Sub

Private Sub lblSeries_Click()
      Dim i As Long
    If HFApp.Options(SalesSystem) = asHomeFront Then Exit Sub
     
    Call FDBGrid.ShowForm(FMain.CD_Series, "select Series,Description from tblseries where divisionid = " & HFApp.DivisionID, "tblSeries", "divisionid,series")
    i = cboSeries.ListIndex
    Call LoadSeries
On Error Resume Next
    cboSeries.ListIndex = i
End Sub



Public Sub mnuAssemblyComponentsSub_Click(Index As Integer)
    Dim s As String
    Dim i As Long
    Dim r As Long
    
    With gComponents
    Select Case Index
        Case mcCOMP_ADD
        
            s = ""
            s = s & "select" & vbCrLf
            s = s & " m.AssemblyTypeDesc AssemblyType" & vbCrLf
            s = s & ",m.Community" & vbCrLf
            s = s & ",m.Assembly" & vbCrLf
            s = s & ",m.AssemblyID" & vbCrLf
            s = s & ",m.Model" & vbCrLf
            s = s & ",m.OptionID" & vbCrLf
            s = s & ",m.Description" & vbCrLf
            s = s & "from tbldbassemblymaster m" & vbCrLf
            s = s & "left outer join tbldbassemblycomponents c on m.assemblyid=c.parentassemblyid" & vbCrLf
            s = s & "where m.divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
            s = s & "and isnull(m.inactive,0)=0" & vbCrLf
            s = s & "and m.assemblyid<>" & DbQuote(Num, mAssemblyID) & " --cant add self as component" & vbCrLf
            Select Case mAssemblyType
            Case atModel
                s = s & "and m.assemblytype in(0,2,3,4) -- if 0-model can be anything" & vbCrLf
            Case atoption
                s = s & "and m.assemblytype in(2,3,4)   -- if 2-opt can not be model" & vbCrLf
            Case atGlobal, atDesignCenter
                s = s & "and m.assemblytype in(3,4)     -- if 3,4-global,dc opt can only be global,dc option" & vbCrLf
            End Select
            s = s & "and c.parentassemblyid is null --cant add other parents as components" & vbCrLf
            If Not FPickList.Choose(HFApp.Databases(dbHomefront), "Component Assembly", s, , , , , "AssemblyID", True) Then Exit Sub
            For i = 1 To FPickList.SelectedItems
                .AddItem ""
                r = .Rows - 1
                .RowData(r) = "new"
                mDirty = True
                .TextMatrix(r, .ColIndex("community")) = FPickList.SelectedItem("community", i)
                .TextMatrix(r, .ColIndex("Model")) = FPickList.SelectedItem("model", i)
                .TextMatrix(r, .ColIndex("option")) = FPickList.SelectedItem("optionid", i)
                .TextMatrix(r, .ColIndex("assemblyid")) = FPickList.SelectedItem("assemblyid", i)
                .TextMatrix(r, .ColIndex("assembly")) = FPickList.SelectedItem("assembly", i)
                .TextMatrix(r, .ColIndex("description")) = FPickList.SelectedItem("description", i)
                .TextMatrix(r, .ColIndex("qty")) = 1
            Next
        
        Case mcCOMP_REMOVE
            If .Row < 1 Then Exit Sub
            .RowData(.Row) = "delete"
            .RowHidden(.Row) = True
            mDirty = True
        End Select
    
    End With
End Sub

Private Sub Slider_Move()
    Call Form_Resize
End Sub

Private Sub Timer1_Timer()
    Call HFApp.RunTask(mTimerTask)
    Timer1.Enabled = False
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo eh
    
    Dim rs As Recordset
    Dim s As String
    Dim view As Long
    Dim i As Long
    
    Dim sCommunity As String
    Dim sModel     As String
    Dim sOptionID  As String
    Dim sAssembly  As String
    
    Select Case Button.key
    
        Case "Import"
            If Not SaveData(True) Then Exit Sub
            If mImportGlobals Then
                Call ImportPipelineGlobalOptions
            Else
                Call ImportPipelineModelsAndOptions
            End If
            
        Case "MassChange"
            Call FMassChange.ShowForm
    
        Case "PricelistExport"
            Call ExportPriceList

        Case "New"
            Call Toolbar_ButtonDropDown(Button)
            
        Case "Save"
            Call SaveData(False)
            
'        Case "Delete"
'            Call DeleteAssembly
            
        Case "Open"
            ' if you make changes here, also do them in FAddPriceList.cmdAssembly_Click
            If SaveData(True) Then
                If HFApp.Options.ValueByName("BuilderType") = "Commercial" Then
                    s = ""
                    s = s & "Master Assemblies" & Chr(1) & "select c.area AreaID,c.description " & FMain.CD_Community & ",a.Model,a.Model modelid,a.OptionID,a.Assembly,a.Description,a.assemblytype,0 Source,a.Series,a.Style,a.Elevation from tbldbassemblymaster a left outer join tbllocality c on(a.community=c.area) where a.DivisionID = " & HFApp.DivisionID & " and isnull(a.inactive,0)=0 and a.assemblytype=0 order by 1,2,3,4,5,6,7" & Chr(0)
                    s = s & "Assembly Specific Extras" & Chr(1) & "select ct.Description Category,c.area AreaID,c.description " & FMain.CD_Community & ",a.Model,dm.description ModelDescription,a.Series,a.Model modelid,a.OptionID,a.OptionID [Option],a.Assembly,a.Description,a.assemblytype,0 Source from tbldbassemblymaster a left outer join tbllocality c on(a.community=c.area) left outer join tblcategories ct on(a.category=ct.category) left outer join distinctmodelsbydivision dm on(a.model=dm.model and a.divisionid=dm.divisionid) where a.DivisionID = " & HFApp.DivisionID & " and isnull(a.inactive,0)=0 and a.assemblytype=2 order by 1,2,3,4,5,6,7" & Chr(0)
                    s = s & "Design Assemblies" & Chr(1) & "select ct.Description Category,c.area AreaID,c.description " & FMain.CD_Community & ",a.Model modelid,a.OptionID,a.OptionID [Option],a.Assembly,a.Description,a.assemblytype,0 Source from tbldbassemblymaster a left outer join tbllocality c on(a.community=c.area) left outer join tblcategories ct on(a.category=ct.category) where a.DivisionID = " & HFApp.DivisionID & " and isnull(a.inactive,0)=0 and a.assemblytype=4  order by 1,2,3,4,5,6,7" & Chr(0)
                    s = s & "Global Assemblies" & Chr(1) & "select ct.Description Category,c.area AreaID,c.description " & FMain.CD_Community & ",a.Model modelid,a.OptionID,a.OptionID [Option],a.Assembly,a.Description,a.assemblytype,0 Source from tbldbassemblymaster a left outer join tbllocality c on(a.community=c.area) left outer join tblcategories ct on(a.category=ct.category) where a.DivisionID = " & HFApp.DivisionID & " and isnull(a.inactive,0)=0 and a.assemblytype=3 order by 1,2,3,4,5,6,7" & Chr(0)
                    s = s & "-------------------------" & Chr(1) & "select '' where 1=0" & Chr(0)
                    s = s & "Sales Assemblies" & Chr(1) & "select community areaID,CommunityDesc " & FMain.CD_Community & ",Model,Series,Model modelid,OptionID,Assembly,Description,assemblytype,1 Source from salesassemblylist where DivisionID = " & HFApp.DivisionID & " and assemblytype=0 order by 1,2,3,4,5,6,7" & Chr(0)
                    s = s & "Sales Assembly Specific Extras" & Chr(1) & "select CategoryDesc Category,community areaID,CommunityDesc " & FMain.CD_Community & ",a.Model,dm.description ModelDescription,Series,a.Model modelid,OptionID,OptionID [Option],Assembly,a.Description,assemblytype,1 Source from salesassemblylist a left outer join distinctmodelsbydivision dm on(a.model=dm.model and a.divisionid=dm.divisionid) where a.DivisionID = " & HFApp.DivisionID & " and assemblytype=2 order by 1,2,3,4,5,6,7" & Chr(0)
                    s = s & "Sales Design Assemblies" & Chr(1) & "select CategoryDesc Category,community areaID,CommunityDesc " & FMain.CD_Community & ",Model modelid,OptionID,OptionID [Option],Assembly,Description,assemblytype,1 Source from salesassemblylist where DivisionID = " & HFApp.DivisionID & " and assemblytype=4 order by 1,2,3,4,5,6,7" & Chr(0)
                    s = s & "Sales Global Assemblies" & Chr(1) & "select CategoryDesc Category,community areaID,CommunityDesc " & FMain.CD_Community & ",Model modelid,OptionID,OptionID [Option],Assembly,Description,assemblytype,1 Source from salesassemblylist where DivisionID = " & HFApp.DivisionID & " and assemblytype=3 order by 1,2,3,4,5,6,7" & Chr(0)
                    s = s & "-------------------------" & Chr(1) & "select '' where 1=0" & Chr(0)
                    s = s & "Inactive Master Assemblies" & Chr(1) & "select c.area AreaID,c.description " & FMain.CD_Community & ",a.Model,a.Model modelid,a.OptionID,a.Assembly,a.Description,a.assemblytype,0 Source,a.Series,a.Style,a.Elevation from tbldbassemblymaster a left outer join tbllocality c on(a.community=c.area) where a.DivisionID = " & HFApp.DivisionID & " and a.inactive=1 and a.assemblytype=0 order by 1,2,3,4,5,6,7" & Chr(0)
                    s = s & "Inactive Assembly Specific Extras" & Chr(1) & "select ct.Description Category,c.area AreaID,c.description Community,a.Model,dm.description ModelDescription,a.Series,a.Model modelid,a.OptionID,a.OptionID [Option],a.Assembly,a.Description,a.assemblytype,0 Source from tbldbassemblymaster a left outer join tbllocality c on(a.community=c.area) left outer join tblcategories ct on(a.category=ct.category) left outer join distinctmodelsbydivision dm on(a.model=dm.model and a.divisionid=dm.divisionid) where a.DivisionID = " & HFApp.DivisionID & " and a.inactive=1 and a.assemblytype=2 order by 1,2,3,4,5,6,7" & Chr(0)
                    s = s & "Inactive Design Assemblies" & Chr(1) & "select ct.Description Category,c.area AreaID,c.description " & FMain.CD_Community & ",a.Model modelid,a.OptionID,a.OptionID [Option],a.Assembly,a.Description,a.assemblytype,0 Source from tbldbassemblymaster a left outer join tbllocality c on(a.community=c.area) left outer join tblcategories ct on(a.category=ct.category) where a.DivisionID = " & HFApp.DivisionID & " and a.inactive=1 and a.assemblytype=4  order by 1,2,3,4,5,6,7" & Chr(0)
                    s = s & "Inactive Global Assemblies" & Chr(1) & "select ct.Description Category,c.area AreaID,c.description " & FMain.CD_Community & ",a.Model modelid,a.OptionID,a.OptionID [Option],a.Assembly,a.Description,a.assemblytype,0 Source from tbldbassemblymaster a left outer join tbllocality c on(a.community=c.area) left outer join tblcategories ct on(a.category=ct.category) where a.DivisionID = " & HFApp.DivisionID & " and a.inactive=1 and a.assemblytype=3 order by 1,2,3,4,5,6,7"
                Else
                    s = ""
                    s = s & "Models" & Chr(1) & "select c.area AreaID,c.description " & FMain.CD_Community & ",a.Model,a.Model modelid,a.OptionID,a.Assembly,a.Description,a.assemblytype,0 Source,a.Series,a.Style,a.Elevation from tbldbassemblymaster a left outer join tbllocality c on(a.community=c.area) where a.DivisionID = " & HFApp.DivisionID & " and isnull(a.inactive,0)=0 and a.assemblytype=0 order by 1,2,3,4,5,6,7" & Chr(0)
                    'old model specific options -- too slow when client has tens of thousands of options
                    's = s & "Model Specific Options" & Chr(1) & "select ct.Description Category,c.area AreaID,c.description " & FMain.CD_Community & ",a.Model,dm.description ModelDescription,a.Series,a.Model modelid,a.OptionID,a.OptionID [Option],a.Assembly,a.Description,a.assemblytype,0 Source from tbldbassemblymaster a left outer join tbllocality c on(a.community=c.area) left outer join tblcategories ct on(a.category=ct.category) left outer join distinctmodelsbydivision dm on(a.model=dm.model and a.divisionid=dm.divisionid) where a.DivisionID = " & HFApp.DivisionID & " and isnull(a.inactive,0)=0 and a.assemblytype=2 order by 1,2,3,4,5,6,7" & Chr(0)
                    s = s & "Model Specific Options" & Chr(1)
                        s = s & "select distinct m.Model,m.Description,0 Source,m.Model ModelID" & vbCrLf
                        s = s & "from tbldbassemblymaster o" & vbCrLf
                        s = s & "join DistinctModelsByDivision m on o.divisionid=m.divisionid and m.model=o.model" & vbCrLf
                        s = s & "where o.assemblytype=2" & vbCrLf
                        s = s & "and isnull(o.inactive,0)=0 and o.divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf & Chr(0)

                    
                    s = s & "Design Center Options" & Chr(1) & "select ct.Description Category,c.area AreaID,c.description " & FMain.CD_Community & ",a.Model modelid,a.OptionID,a.OptionID [Option],a.Assembly,a.Description,a.assemblytype,0 Source from tbldbassemblymaster a left outer join tbllocality c on(a.community=c.area) left outer join tblcategories ct on(a.category=ct.category) where a.DivisionID = " & HFApp.DivisionID & " and isnull(a.inactive,0)=0 and a.assemblytype=4  order by 1,2,3,4,5,6,7" & Chr(0)
                    s = s & "Global Options" & Chr(1) & "select ct.Description Category,c.area AreaID,c.description " & FMain.CD_Community & ",a.Model modelid,a.OptionID,a.OptionID [Option],a.Assembly,a.Description,a.assemblytype,0 Source from tbldbassemblymaster a left outer join tbllocality c on(a.community=c.area) left outer join tblcategories ct on(a.category=ct.category) where a.DivisionID = " & HFApp.DivisionID & " and isnull(a.inactive,0)=0 and a.assemblytype=3 order by 1,2,3,4,5,6,7" & Chr(0)
                    s = s & "-------------------------" & Chr(1) & "select '' where 1=0" & Chr(0)
                    s = s & "Sales Models" & Chr(1) & "select community areaID,CommunityDesc " & FMain.CD_Community & ",Model,Series,Model modelid,OptionID,Assembly,Description,assemblytype,1 Source from salesassemblylist where DivisionID = " & HFApp.DivisionID & " and assemblytype=0 order by 1,2,3,4,5,6,7" & Chr(0)
                    s = s & "Sales Model Specific Options" & Chr(1) & "select CategoryDesc Category,community areaID,CommunityDesc " & FMain.CD_Community & ",a.Model,dm.description ModelDescription,Series,a.Model modelid,OptionID,OptionID [Option],Assembly,a.Description,assemblytype,1 Source from salesassemblylist a left outer join distinctmodelsbydivision dm on(a.model=dm.model and a.divisionid=dm.divisionid) where a.DivisionID = " & HFApp.DivisionID & " and assemblytype=2 order by 1,2,3,4,5,6,7" & Chr(0)
                    s = s & "Sales Design Center Options" & Chr(1) & "select CategoryDesc Category,community areaID,CommunityDesc " & FMain.CD_Community & ",Model modelid,OptionID,OptionID [Option],Assembly,Description,assemblytype,1 Source from salesassemblylist where DivisionID = " & HFApp.DivisionID & " and assemblytype=4 order by 1,2,3,4,5,6,7" & Chr(0)
                    s = s & "Sales Global Options" & Chr(1) & "select CategoryDesc Category,community areaID,CommunityDesc " & FMain.CD_Community & ",Model modelid,OptionID,OptionID [Option],Assembly,Description,assemblytype,1 Source from salesassemblylist where DivisionID = " & HFApp.DivisionID & " and assemblytype=3 order by 1,2,3,4,5,6,7" & Chr(0)
                    s = s & "-------------------------" & Chr(1) & "select '' where 1=0" & Chr(0)
                    s = s & "Inactive Models" & Chr(1) & "select c.area AreaID,c.description " & FMain.CD_Community & ",a.Model,a.Model modelid,a.OptionID,a.Assembly,a.Description,a.assemblytype,0 Source,a.Series,a.Style,a.Elevation from tbldbassemblymaster a left outer join tbllocality c on(a.community=c.area) where a.DivisionID = " & HFApp.DivisionID & " and a.inactive=1 and a.assemblytype=0 order by 1,2,3,4,5,6,7" & Chr(0)
                    s = s & "Inactive Model Specific Options" & Chr(1) & "select ct.Description Category,c.area AreaID,c.description " & FMain.CD_Community & ",a.Model,dm.description ModelDescription,a.Series,a.Model modelid,a.OptionID,a.OptionID [Option],a.Assembly,a.Description,a.assemblytype,0 Source from tbldbassemblymaster a left outer join tbllocality c on(a.community=c.area) left outer join tblcategories ct on(a.category=ct.category) left outer join distinctmodelsbydivision dm on(a.model=dm.model and a.divisionid=dm.divisionid) where a.DivisionID = " & HFApp.DivisionID & " and a.inactive=1 and a.assemblytype=2 order by 1,2,3,4,5,6,7" & Chr(0)
                    s = s & "Inactive Design Center Options" & Chr(1) & "select ct.Description Category,c.area AreaID,c.description " & FMain.CD_Community & ",a.Model modelid,a.OptionID,a.OptionID [Option],a.Assembly,a.Description,a.assemblytype,0 Source from tbldbassemblymaster a left outer join tbllocality c on(a.community=c.area) left outer join tblcategories ct on(a.category=ct.category) where a.DivisionID = " & HFApp.DivisionID & " and a.inactive=1 and a.assemblytype=4  order by 1,2,3,4,5,6,7" & Chr(0)
                    s = s & "Inactive Global Options" & Chr(1) & "select ct.Description Category,c.area AreaID,c.description " & FMain.CD_Community & ",a.Model modelid,a.OptionID,a.OptionID [Option],a.Assembly,a.Description,a.assemblytype,0 Source from tbldbassemblymaster a left outer join tbllocality c on(a.community=c.area) left outer join tblcategories ct on(a.category=ct.category) where a.DivisionID = " & HFApp.DivisionID & " and a.inactive=1 and a.assemblytype=3 order by 1,2,3,4,5,6,7"
                End If
                
                Select Case mAssemblyType
                Case atModel:        view = 1
                Case atoption:       view = 2
                Case atGlobal:       view = 4
                Case atDesignCenter: view = 3
                Case atCustomOption: view = 2
                End Select
                                
                                
                                
        
                                
                If FPickList.Choose(HFApp.Databases(dbHomefront), IIf(HFApp.Options.ValueByName("BuilderType") = "Commercial", "Assemblies", "Model and Option"), s, , , , , "areaid,modelid,optionid,assemblytype,Source," & IIf(lblElevation.Visible, "", "Elevation"), , view) Then
                
                
                    mCopiedCommunity = ""
                    mCopiedAssembly = ""
                    mCopiedAssemblyID = 0
                    mCopiedModel = ""
                    mCommunity = FPickList.SelectedItem("areaID")
                    mAssembly = FPickList.SelectedItem("assembly")
                    mModel = FPickList.SelectedItem("modelid")
                    mOptionID = FPickList.SelectedItem("optionid")
                    mDataSource = FPickList.SelectedItem("Source")
                    mAssemblyType = Val(FPickList.SelectedItem("assemblytype"))
                    
                    
                    'if model option
                    If FPickList.SelectedView = 2 Then
                        
                        'now pick option
                        s = ""
                        s = s & "select ct.Description Category,c.area AreaID,c.description " & FMain.CD_Community & ",a.Model,dm.description ModelDescription,a.Series,a.Model modelid,a.OptionID,a.OptionID [Option],a.Assembly,a.Description,a.assemblytype,0 Source" & vbCrLf
                        s = s & "from tbldbassemblymaster a left outer join tbllocality c on(a.community=c.area) left outer join tblcategories ct on(a.category=ct.category) left outer join distinctmodelsbydivision dm on(a.model=dm.model and a.divisionid=dm.divisionid)" & vbCrLf
                        s = s & "where a.DivisionID = " & HFApp.DivisionID & vbCrLf
                        s = s & " and isnull(a.inactive,0)=0" & vbCrLf
                        s = s & " and a.assemblytype=2" & vbCrLf
                        s = s & " and a.model=" & DbQuote(Str, mModel) & vbCrLf
                        s = s & " order by 1,2,3,4,5,6,7"
                        If Not FPickList.Choose(HFApp.Databases(dbHomefront), IIf(HFApp.Options.ValueByName("BuilderType") = "Commercial", "Assemblies", "Model and Option"), s, , , , , "areaid,modelid,optionid,assemblytype,Source", , view) Then Exit Sub
                        mCommunity = FPickList.SelectedItem("areaID")
                        mAssembly = FPickList.SelectedItem("assembly")
                        mModel = FPickList.SelectedItem("modelid")
                        mOptionID = FPickList.SelectedItem("optionid")
                        mDataSource = FPickList.SelectedItem("Source")
                        mAssemblyType = Val(FPickList.SelectedItem("assemblytype"))
                    End If
                                        
                    Me.cboPriceCommunity.ListIndex = -1
                    Call cboPriceCommunity_Click
                End If
            
            
            End If
            
        Case "TakeoffOneTime", "TakeoffItem", "TakeoffAssembly", "TakeoffCustom", "TakeoffPlanSwift"
            Call FTakeoff.Takeoff(Me, False, Mid(Button.key, 8), 0, "", mAssemblyType, txtDescription.Text, PricingCommunity, PricingCommunityPhase, cboModel.Text, "", txtAssembly.Text, "")
             
        Case "TakeoffItemChart"
            If Not FPickList.Choose(HFApp.Databases(dbHomefront), "Item Chart", "select name,formula from itemcharts", , , , , "formula", True) Then Exit Sub
            ReadOnly = False
            mDirty = True
            With gItems
            
            For i = 1 To FPickList.SelectedItems
            
                .AddItem ""
                .TextMatrix(.Rows - 1, .ColIndex("ItemChart")) = FPickList.SelectedItem("name", i)
                .TextMatrix(.Rows - 1, .ColIndex("formula")) = FPickList.SelectedItem("formula", i)
                .TextMatrix(.Rows - 1, .ColIndex("ItemConversionFactor")) = 1
                .TextMatrix(.Rows - 1, .ColIndex("TakeoffQty")) = 1
                .TextMatrix(.Rows - 1, .ColIndex("OrderQty")) = 1
                        
                .RowData(.Rows - 1) = "new"
            Next
            End With
        

        
    End Select
Exit Sub
eh:
    If InStr(1, Err.Description, "needs to be upgraded") Then
        MsgBox Err.Description, vbExclamation, App.ProductName
        Screen.MousePointer = vbDefault
    Else
        Call errHandler(SRCFILE & "Toolbar_ButtonClick")
        Unload FProgress
    End If
End Sub

Private Sub Toolbar_ButtonDropDown(ByVal Button As MSComctlLib.Button)
On Error GoTo eh
    Select Case Button.key
        Case "New"
            PopupMenu FMain.mnuAssemblyNew, , Button.Left, Button.Top + Button.Height
    End Select
    Exit Sub
eh: Call errHandler(SRCFILE & "Toolbar_ButtonDropDown")
End Sub

Public Property Get PricingCommunity() As String
    PricingCommunity = Parse(GetComboBoxListKey(cboPriceCommunity), 1, Chr(2))
End Property
Private Property Get PricingCommunityPhase() As String
    PricingCommunityPhase = Parse(GetComboBoxListKey(cboPriceCommunity), 2, Chr(2))
End Property



Public Sub LoadData()
On Error GoTo eh

    Dim s As String
    Dim rs As Recordset
    Dim r As Long, ra As Long
    Dim i As Long
    Dim w As Long
    
    Dim AssemblyCost As Double
    Dim ComponentCost As Double
    
    
    
    ReadOnly = False
    Screen.MousePointer = vbHourglass
    
    
    lblAssemblyCost(0) = ""
    lblAssemblyCost(1) = ""
    lblComponentCost = ""
    lblTotalCost = ""
    
    
    If mAssemblyType = atoption Then
        cboModel.Style = vbComboDropdownList
    Else
        cboModel.Style = vbComboDropdown
    End If
    
    If mAssemblyType = atoption Then
        Call LoadComboBox(cboModel, HFApp.Databases(dbHomefront), "select distinct model,'',0 from tbldbassemblymaster where isnull(isbaseassembly,0)=0 and assemblytype=0 and divisionid=" & DbQuote(Num, HFApp.DivisionID) & " order by 1")
    Else
        Call LoadComboBox(cboModel, HFApp.Databases(dbHomefront), "select distinct model,'',0 from tbldbassemblymaster where assemblytype=0 and divisionid=" & DbQuote(Num, HFApp.DivisionID) & " order by 1")
    End If
    
    s = ""
    s = s & "SELECT AssemblyId,BillingFactor,BillingCode,BillingCategory,GraphicPath,SpecDocument,ScheduleTemplate,MaxWidth,MaxLength,ConstCutoff,Community,AssemblyType,Model,Elevation,OptionID,Assembly,Inactive,Description,Comments,Notes,Series,Style,AssemblyUOM,Bedrooms,Bathrooms,FloorArea,Category,JCExtra,Qty,Location,IncludedOption,ColorListID,StyleListID,FinishListID,OtherListID,Color,StyleValue,FinishValue,OtherValue,TakeoffRequired,UseNormalSalesQtyFactors,IsBaseAssembly,DesignCenterSalesOnly,SelectByRoom,DisplayTotalOnly" & vbCrLf
    s = s & "  FROM " & IIf(mDataSource = 0, "tblDBAssemblyMaster m", "SalesAssemblyList m") & vbCrLf
    s = s & " WHERE m.DivisionID = " & HFApp.DivisionID & vbCrLf
    s = s & "   and ISNULL(m.Community,'')=" & DbQuote(Str, mCommunity) & vbCrLf
    s = s & "   AND ISNULL(m.Model,'')=" & DbQuote(Str, mModel) & vbCrLf
    s = s & "   AND ISNULL(m.Optionid,'')=" & DbQuote(Str, mOptionID) & vbCrLf
    s = s & "   AND ISNULL(m.Assembly,'')=" & DbQuote(Str, mAssembly) & vbCrLf
    Set rs = HFApp.SqlExec(s, dbHomefront, ra)
    If rs.EOF Then
        mAssemblyID = 0
        cboCommunity.ListIndex = -1
        Call SetComboBoxListIndex(cboCommunity, , mCommunity)
        Call SetComboBoxListIndex(cboModel, mModel)
        
        txtOption.Text = mOptionID
        txtElevation.Text = ""
        txtAssembly.Text = mAssembly
        txtDescription.Text = mAssemblyDesc
        txtGraphicPath.Text = ""
        txtSpecDocument.Text = ""
        txtMaxWidth.Text = ""
        txtMaxLength.Text = ""
        txtConstCutoff.Text = ""
        txtComments.Text = ""
        txtNotes.Text = ""
        chkActive.value = vbChecked
        cboSeries.ListIndex = -1
        cboUOM.Text = ""
        txtStyle.Text = ""
        cboStyle.ListIndex = -1
        txtBedrooms.Text = ""
        txtBathrooms.Text = ""
        txtFloorArea.Text = ""
        cboCategory.ListIndex = -1
        txtJCExtra.Text = ""
        chkIsBaseAssembly.value = vbUnchecked
        txtBillingFactor.Text = "1"
        cboBillingCode.ListIndex = -1
        cboBillingCategory.ListIndex = -1
        
        txtQty.Text = "1"
        cboScheduleTemplate.ListIndex = -1
        cboLocation.ListIndex = -1
        chkIncludedOption.value = vbUnchecked
        chkDCSalesOnly.value = vbUnchecked
        chkSelectByRoom.value = vbUnchecked
        chkDisplayTotalOnly.value = vbUnchecked
        
        chkTakeoffRequired.value = vbUnchecked
        chkUseNormalSalesQtyFactors.value = vbChecked
        
        
        
        txtLColor.Text = ""
        txtLStyle.Text = ""
        txtLFinish.Text = ""
        txtLOther.Text = ""
        txtAColor.Text = ""
        txtAStyle.Text = ""
        txtAFinish.Text = ""
        txtAOther.Text = ""
        
        
        CtrlEnabled(cboCommunity) = True
        CtrlEnabled(cboModel) = mAssemblyType = atModel Or mAssemblyType = atoption
        CtrlEnabled(txtOption) = mAssemblyType <> atModel
        CtrlEnabled(txtAssembly) = True
    Else
        mAssemblyID = Val("" & rs("AssemblyID"))
        mCommunity = "" & rs("Community")
        mAssemblyType = Val("" & rs("AssemblyType"))
        
        txtConstCutoff.Visible = mAssemblyType <> atDesignCenter
        lblConstCutoff.Visible = mAssemblyType <> atDesignCenter
        
        
        
        Call SetComboBoxListIndex(cboModel, "" & rs("Model"))
        If cboModel.ListIndex = -1 Then
            cboModel.AddItem ("" & rs("Model"))
            Call SetComboBoxListIndex(cboModel, "" & rs("Model"))
        End If
        
        If "" & rs("ScheduleTemplate") = "" Then
            cboScheduleTemplate.ListIndex = 0
        Else
            Call SetComboBoxListIndex(cboScheduleTemplate, "" & rs("ScheduleTemplate"))
            If cboScheduleTemplate.ListIndex = -1 Then
                Call cboScheduleTemplate.AddItem("" & rs("ScheduleTemplate"))
                Call SetComboBoxListIndex(cboScheduleTemplate, "" & rs("ScheduleTemplate"))
            End If
        End If
        
        txtOption.Text = "" & mOptionID
        txtElevation.Text = "" & rs("Elevation")
        txtAssembly.Text = "" & mAssembly
        chkActive.value = IIf("" & rs("InActive") = "True", vbUnchecked, vbChecked)
        txtDescription.Text = "" & rs("Description")
        txtGraphicPath.Text = "" & rs("GraphicPath")
        txtSpecDocument.Text = "" & rs("SpecDocument")
        txtMaxWidth.Text = Val("" & rs("MaxWidth"))
        txtMaxLength.Text = Val("" & rs("MaxLength"))
        txtConstCutoff.Text = Val("" & rs("ConstCutoff"))
        txtComments.Text = "" & rs("Comments")
        txtNotes.Text = "" & rs("Notes")
        Call SetComboBoxListIndex(cboSeries, , "" & rs("Series"))
        txtStyle.Text = "" & rs("Style")
        Call SetComboBoxListIndex(cboCommunity, , mCommunity)
        If "" & rs("Style") = "" Then
            cboStyle.ListIndex = -1
        Else
            Call SetComboBoxListIndex(cboStyle, "" & rs("Style"))
        End If
        cboUOM.Text = "" & rs("AssemblyUOM")
        txtBedrooms.Text = Val("" & rs("Bedrooms"))
        txtBathrooms.Text = Val("" & rs("Bathrooms"))
        txtFloorArea.Text = Val("" & rs("FloorArea"))
        Call SetComboBoxListIndex(cboCategory, , "" & rs("Category"))
        txtJCExtra.Text = "" & rs("JCExtra")
        
        txtQty.Text = Val("" & rs("Qty"))
        Call SetComboBoxListIndex(cboLocation, , "" & rs("Location"))
        chkIncludedOption.value = IIf("" & rs("IncludedOption") = "True", vbChecked, vbUnchecked)
        chkDCSalesOnly.value = IIf("" & rs("DesignCenterSalesOnly") = "True", vbChecked, vbUnchecked)
        chkSelectByRoom.value = IIf("" & rs("SelectByRoom") = "True", vbChecked, vbUnchecked)
        chkDisplayTotalOnly.value = IIf("" & rs("DisplayTotalOnly") = "True", vbChecked, vbUnchecked)
        
        txtBillingFactor.Text = Val("" & rs("BillingFactor"))
        Call SetComboBoxListIndex(cboBillingCode, , "" & rs("BillingCode"))
        Call SetComboBoxListIndex(cboBillingCategory, , "" & rs("BillingCategory"))
        
        txtLColor.Text = Val("" & rs("ColorListID"))
        txtLStyle.Text = Val("" & rs("StyleListID"))
        txtLFinish.Text = Val("" & rs("FinishListID"))
        txtLOther.Text = Val("" & rs("otherListID"))
        If txtLColor.Text = "0" Then txtLColor.Text = ""
        If txtLStyle.Text = "0" Then txtLStyle.Text = ""
        If txtLFinish.Text = "0" Then txtLFinish.Text = ""
        If txtLOther.Text = "0" Then txtLOther.Text = ""
        
        txtAColor.Text = "" & rs("Color")
        txtAStyle.Text = "" & rs("StyleValue")
        txtAFinish.Text = "" & rs("FinishValue")
        txtAOther.Text = "" & rs("OtherValue")
        
        chkTakeoffRequired.value = IIf("" & rs("TakeoffRequired") = "True", vbChecked, vbUnchecked)
        chkUseNormalSalesQtyFactors.value = IIf("" & rs("UseNormalSalesQtyFactors") = "True", vbChecked, vbUnchecked)
        chkIsBaseAssembly.value = IIf("" & rs("IsBaseAssembly") = "True", vbChecked, vbUnchecked)
        CtrlEnabled(cboCommunity) = mDataSource = 1
        CtrlEnabled(cboModel) = mDataSource = 1
        CtrlEnabled(txtOption) = mDataSource = 1
        CtrlEnabled(txtAssembly) = mDataSource = 1
    End If
    
    Select Case mAssemblyType
        Case atModel
            txtOption.Visible = False
            cboModel.Visible = True
            Label144.Caption = "Model"
        
        Case atoption
            txtOption.Visible = True
            cboModel.Visible = True
            txtOption.Left = cboModel.Left + cboModel.Width + 60
            Label144.Caption = "Model/Option"
        
        Case atGlobal
            txtOption.Visible = True
            cboModel.Visible = False
            txtOption.Left = cboModel.Left
            Label144.Caption = "Option"
        
        Case atDesignCenter
            txtOption.Visible = True
            cboModel.Visible = False
            txtOption.Left = cboModel.Left
            Label144.Caption = "Option"
            
        Case atCustomOption
    End Select
    
    CtrlEnabled(cboUOM) = mAssemblyType <> atModel
    
    Call Form_Resize
    
    'cboModel.AutoCompleteListItemsOnly = mAssemblyType <> atModel
    frmModelFields.Visible = mAssemblyType = atModel
    frmOptionFields.Visible = mAssemblyType <> atModel
    frmEstimatorNotes.Visible = frmOptionFields.Visible
 
 
    If PricingCommunity = "" Then
        s = ""
        s = s & "SELECT d.sequence,d.Invertable,case when v.phase is null then 0 else 1 end hasInverseItem,d.UseModelCost,i.Description,i.OrderUOM,isnull(nullif(d.POIndex,''),isnull(i.POIndex,''))POIndex,p.FullDescription POIndexDescription, cc.CostCode JCCostCode,cc.Description JCCostCodeDesc, ct.Category JCCategory,ct.Description JCCategoryDesc,i.TakeoffUOM,i.rounddir,i.roundto" & vbCrLf
        s = s & ",d.Phase,d.Item,d.itemchart,i.itemnumber,d.Formula,d.TakeoffQty,d.OrderQty,d.notes,i.IsQuote,i.WastePercent" & vbCrLf
        s = s & ",d.Location,d.WBS01,d.WBS02,d.WBS03,d.WBS04,d.WBS05,d.WBS06,d.WBS07,d.WBS08,d.WBS09,d.WBS10,d.WBS11" & vbCrLf
        s = s & ",d.WBS12,d.WBS13,d.WBS14,d.WBS15,d.WBS16,d.WBS17,d.WBS18,d.WBS19,d.WBS20,d.WBS21,d.WBS22,d.WBS23" & vbCrLf
        s = s & ",d.WBS24,d.WBS25,d.WBS26,d.WBS27,d.WBS28,d.WBS29,d.WBS30,d.WBS31,d.WBS32,d.WBS33,d.WBS34,d.WBS35" & vbCrLf
        s = s & ",d.WBS36,d.WBS37,d.WBS38,d.WBS39,d.WBS40,d.ConversionFactor AssemblyConversionFactor,i.ConversionFactor ItemConversionFactor" & vbCrLf
        s = s & "FROM tblDBAssemblyDetails d" & vbCrLf
        s = s & "LEFT OUTER JOIN tblPhaseItem i ON(d.DivisionID = i.DivisionID and d.Phase=i.Phase AND d.Item=i.Item)" & vbCrLf
        s = s & "LEFT OUTER JOIN tblPhaseItem v ON(i.DivisionID = v.DivisionID and i.inversePhase=v.Phase AND i.inverseItem=v.Item)" & vbCrLf
        s = s & "LEFT OUTER JOIN tblPOIndex p ON(p.DivisionID = d.DivisionID and isnull(nullif(d.POIndex,''),isnull(i.POIndex,''))=p.POIndex)" & vbCrLf
        s = s & "LEFT OUTER JOIN appoptions o on(o.optionname='UseAltCostCodesForCO' and o.DivisionID = d.DivisionID)" & vbCrLf
        s = s & "LEFT OUTER JOIN StandardCostCodes  cc ON(cc.DivisionID = i.DivisionID and cc.costcode= isnull(case when isnull(d.optionid,'')='' or o.optionvalue='true' then i.jccostcode else isnull(nullif(i.altjccostcode,''),i.jccostcode) end,p.jccostcode))" & vbCrLf
        s = s & "LEFT OUTER JOIN StandardCategories ct ON(ct.DivisionID = i.DivisionID and ct.category= isnull(case when isnull(d.optionid,'')='' or o.optionvalue='true' then i.jccategory else isnull(nullif(i.altjccategory,''),i.jccategory) end,p.jccategory))" & vbCrLf
        s = s & "WHERE d.DivisionID = " & HFApp.DivisionID & " and isnull(d.Community,'')=" & DbQuote(Str, mCommunity) & vbCrLf
        s = s & "AND isnull(d.Assembly,'')=" & DbQuote(Str, mAssembly) & vbCrLf
        s = s & "AND isnull(d.Model,'')=" & DbQuote(Str, mModel) & vbCrLf
        s = s & "AND isnull(d.Optionid,'')=" & DbQuote(Str, mOptionID) & vbCrLf
        s = s & "ORDER BY d.Phase,d.Item,d.itemchart,d.Sequence"
    Else
        s = ""
        s = s & "SELECT d.sequence,d.Invertable,case when iv.phase is null then 0 else 1 end hasInverseItem,d.UseModelCost,i.Description,i.OrderUOM,isnull(nullif(d.POIndex,''),isnull(i.POIndex,'')) POIndex,p.FullDescription POIndexDescription,cc.CostCode JCCostCode,cc.Description JCCostCodeDesc, ct.Category JCCategory,ct.Description JCCategoryDesc,i.TakeoffUOM,i.rounddir,i.roundto" & vbCrLf
        s = s & ",d.Phase,d.Item,d.itemchart,i.itemnumber,d.Formula,d.TakeoffQty,d.OrderQty,d.notes,i.IsQuote,i.WastePercent" & vbCrLf
        s = s & ",v.vendor_id Vendor,v.vendor_name VendorDesc" & vbCrLf
        s = s & ",dbo.Purch_GetItemRate(0,0," & DbQuote(Str, PricingCommunity) & ", " & DbQuote(Str, PricingCommunityPhase) & ", d.Assembly,d.Model,d.OptionID,d.Phase,d.Item,d.Sequence,v.Vendor_id,getdate()," & HFApp.DivisionID & ")        Rate" & vbCrLf
        s = s & ",dbo.Purch_GetItemRateQuality(0,0," & DbQuote(Str, PricingCommunity) & ", " & DbQuote(Str, PricingCommunityPhase) & ", d.Assembly,d.Model,d.OptionID,d.Phase,d.Item,d.Sequence,v.Vendor_id,getdate()," & HFApp.DivisionID & ") RateQuality" & vbCrLf
        s = s & ",t.JCRate JCTaxRate,d.Location,d.WBS01,d.WBS02,d.WBS03,d.WBS04,d.WBS05,d.WBS06,d.WBS07,d.WBS08,d.WBS09,d.WBS10,d.WBS11" & vbCrLf
        s = s & ",d.WBS12,d.WBS13,d.WBS14,d.WBS15,d.WBS16,d.WBS17,d.WBS18,d.WBS19,d.WBS20,d.WBS21,d.WBS22,d.WBS23" & vbCrLf
        s = s & ",d.WBS24,d.WBS25,d.WBS26,d.WBS27,d.WBS28,d.WBS29,d.WBS30,d.WBS31,d.WBS32,d.WBS33,d.WBS34,d.WBS35" & vbCrLf
        s = s & ",d.WBS36,d.WBS37,d.WBS38,d.WBS39,d.WBS40,d.ConversionFactor AssemblyConversionFactor,i.ConversionFactor ItemConversionFactor" & vbCrLf
        s = s & "FROM tblDBAssemblyDetails d" & vbCrLf
        s = s & "LEFT OUTER JOIN tblPhaseItem i ON(d.DivisionID = i.DivisionID and d.Phase=i.Phase AND d.Item=i.Item)" & vbCrLf
        s = s & "LEFT OUTER JOIN tblPhaseItem iv ON(i.DivisionID = iv.DivisionID and i.inversePhase=iv.Phase AND i.inverseItem=iv.Item)" & vbCrLf
        s = s & "LEFT OUTER JOIN tblPOIndex p ON(p.DivisionID = " & HFApp.DivisionID & " and isnull(nullif(d.POIndex,''),isnull(i.POIndex,''))=p.POIndex)" & vbCrLf
        s = s & "LEFT OUTER JOIN appoptions o on(o.optionname='UseAltCostCodesForCO' and o.DivisionID = d.DivisionID)" & vbCrLf
        s = s & "LEFT OUTER JOIN StandardCostCodes  cc ON(cc.DivisionID = i.DivisionID and cc.costcode= isnull(case when isnull(d.optionid,'')='' or o.optionvalue='true' then i.jccostcode else isnull(nullif(i.altjccostcode,''),i.jccostcode) end,p.jccostcode))" & vbCrLf
        s = s & "LEFT OUTER JOIN StandardCategories ct ON(ct.DivisionID = i.DivisionID and ct.category= isnull(case when isnull(d.optionid,'')='' or o.optionvalue='true' then i.jccategory else isnull(nullif(i.altjccategory,''),i.jccategory) end,p.jccategory))" & vbCrLf
        s = s & "LEFT OUTER JOIN tblvendors v ON(v.DivisionID = d.DivisionID and v.vendor_id=dbo.Purch_GetCommunityVendor(" & DbQuote(Str, PricingCommunity) & ", p.POIndex," & HFApp.DivisionID & "))" & vbCrLf
        s = s & "left outer join taxGroups t on (t.DivisionID = " & HFApp.DivisionID & " and t.TaxGroup=dbo.Purch_GetDefaultTaxGroup('' , " & DbQuote(Str, PricingCommunity) & " ,'' , d.Model , d.Assembly ,i.phase ,i.item , v.Vendor_ID , i.JCCategory," & HFApp.DivisionID & "))"
        s = s & "WHERE d.DivisionID = " & HFApp.DivisionID & " and isnull(d.Community,'')=" & DbQuote(Str, mCommunity) & vbCrLf
        s = s & "AND isnull(d.Assembly,'')=" & DbQuote(Str, mAssembly) & vbCrLf
        s = s & "AND isnull(d.Model,'')=" & DbQuote(Str, mModel) & vbCrLf
        s = s & "AND isnull(d.Optionid,'')=" & DbQuote(Str, mOptionID) & vbCrLf
        s = s & "ORDER BY d.Phase,d.Item,d.itemchart,d.Sequence" & vbCrLf
    End If
    Set rs = HFApp.SqlExec(s)
    AssemblyCost = 0
    With gItems
        .Rows = 1
        While Not rs.EOF
            r = .Rows
            .AddItem ""
            .TextMatrix(r, .ColIndex("ItemType")) = IIf("" & rs("IsQuote") = "True", "Quote", "Unit Price")
            
            If PricingCommunity <> "" Then
                i = Val("" & rs("RateQuality"))
                Select Case i
                    'assembly specific
                    Case 4:  .TextMatrix(r, .ColIndex("PriceLevel")) = "Phase Specific"
                             .TextMatrix(r, .ColIndex("ItemType")) = "Quote"
                    Case 5:  .TextMatrix(r, .ColIndex("PriceLevel")) = FMain.CD_Community & " Specific"
                             .TextMatrix(r, .ColIndex("ItemType")) = "Quote"
                    Case 6:  .TextMatrix(r, .ColIndex("PriceLevel")) = "Global (any " & FMain.CD_Community & ")"
                             .TextMatrix(r, .ColIndex("ItemType")) = "Quote"
                    
                    'not assembly specific
                    Case 9:  .TextMatrix(r, .ColIndex("PriceLevel")) = "Phase Specific"
                             .TextMatrix(r, .ColIndex("ItemType")) = "Unit Price"
                    Case 10: .TextMatrix(r, .ColIndex("PriceLevel")) = FMain.CD_Community & " Specific"
                             .TextMatrix(r, .ColIndex("ItemType")) = "Unit Price"
                    Case 11: .TextMatrix(r, .ColIndex("PriceLevel")) = "Global (any " & FMain.CD_Community & ")"
                             .TextMatrix(r, .ColIndex("ItemType")) = "Unit Price"
                    Case 12: .TextMatrix(r, .ColIndex("PriceLevel")) = "Corporate (any Division)"
                             .TextMatrix(r, .ColIndex("ItemType")) = "Unit Price"
                    
                    'no vendor rate found
                    Case 13: .TextMatrix(r, .ColIndex("PriceLevel")) = "Item DB"
                             .TextMatrix(r, .ColIndex("ItemType")) = "Unit Price"
                
                End Select
                .TextMatrix(r, .ColIndex("Price")) = "" & rs("Rate")
                .TextMatrix(r, .ColIndex("TaxRate")) = "" & rs("JCTaxRate") / 100
                .TextMatrix(r, .ColIndex("TaxAmount")) = (Val("" & rs("Rate")) * (Val("" & rs("JCTaxRate")) / 100)) * Val("" & rs("OrderQty"))
                .TextMatrix(r, .ColIndex("PretaxAmount")) = Val("" & rs("Rate")) * Val("" & rs("OrderQty"))
                .TextMatrix(r, .ColIndex("ExtendedAmount")) = (Val("" & rs("Rate")) * (1 + Val("" & rs("JCTaxRate")) / 100)) * Val("" & rs("OrderQty"))
                AssemblyCost = AssemblyCost + (Val("" & rs("Rate")) * (1 + Val("" & rs("JCTaxRate")) / 100)) * Val("" & rs("OrderQty")) 'Val("" & rs("Rate")) * Val("" & rs("OrderQty"))
                
                
                
                
                .TextMatrix(r, .ColIndex("Vendor")) = "" & rs("Vendor")
                .TextMatrix(r, .ColIndex("VendorDesc")) = "" & rs("VendorDesc")
                'save original value so we can remove old cost records
                .TextMatrix(r, .ColIndex("OldPriceLevel")) = .TextMatrix(r, .ColIndex("PriceLevel"))
                .TextMatrix(r, .ColIndex("OldItemType")) = .TextMatrix(r, .ColIndex("ItemType"))
            End If
            
            If PricingCommunity <> "" Then
                lblAssemblyCost(0).Caption = format(AssemblyCost, "$#,##0.00")
                lblAssemblyCost(1).Caption = format(AssemblyCost, "$#,##0.00")
            End If
            
            .TextMatrix(r, .ColIndex("hasInverseItem")) = "" & rs("hasInverseItem")
            .Cell(flexcpChecked, r, .ColIndex("Invertable")) = IIf("" & rs("Invertable") = "True" And "" & rs("hasInverseItem") = "1", flexChecked, flexUnchecked)
            .Cell(flexcpChecked, r, .ColIndex("UseModelCost")) = IIf("" & rs("UseModelCost") = "True", flexChecked, flexUnchecked)
            
            .TextMatrix(r, .ColIndex("Sequence")) = "" & rs("Sequence")
            .TextMatrix(r, .ColIndex("Phase")) = "" & rs("Phase")
            .TextMatrix(r, .ColIndex("Item")) = "" & rs("Item")
            .TextMatrix(r, .ColIndex("ItemNumber")) = "" & rs("ItemNumber")
            .TextMatrix(r, .ColIndex("ItemChart")) = "" & rs("ItemChart")
            .TextMatrix(r, .ColIndex("WastePercent")) = "" & rs("WastePercent")
            .TextMatrix(r, .ColIndex("Description")) = "" & rs("Description")
            .TextMatrix(r, .ColIndex("TakeoffUOM")) = "" & rs("TakeoffUOM")
            .TextMatrix(r, .ColIndex("AssemblyConversionFactor")) = Val("" & rs("AssemblyConversionFactor"))
            If .ValueMatrix(r, .ColIndex("AssemblyConversionFactor")) = 0 Then .TextMatrix(r, .ColIndex("AssemblyConversionFactor")) = ""

            .TextMatrix(r, .ColIndex("ItemConversionFactor")) = Val("" & rs("ItemConversionFactor"))
            
            
            .TextMatrix(r, .ColIndex("Formula")) = "" & rs("Formula")
            .TextMatrix(r, .ColIndex("TakeoffQty")) = Val("" & rs("TakeoffQty"))
            .TextMatrix(r, .ColIndex("OrderQty")) = Val("" & rs("OrderQty"))
            .TextMatrix(r, .ColIndex("RoundTo")) = Val("" & rs("RoundTo"))
            .TextMatrix(r, .ColIndex("RoundDir")) = Val("" & rs("RoundDir"))
            .TextMatrix(r, .ColIndex("OrderUOM")) = "" & rs("OrderUOM")
            .TextMatrix(r, .ColIndex("Notes")) = "" & rs("Notes")
            .Cell(flexcpData, r, .ColIndex("POIndex")) = "" & rs("POIndex")
            .Cell(flexcpText, r, .ColIndex("POIndex")) = "" & rs("POIndexDescription")
            .TextMatrix(r, .ColIndex("JCCostCode")) = "" & rs("JCCostCode")
            .TextMatrix(r, .ColIndex("JCCostCodeDesc")) = "" & rs("JCCostCodeDesc")
            .TextMatrix(r, .ColIndex("JCCategory")) = "" & rs("JCCategory")
            .TextMatrix(r, .ColIndex("JCCategoryDesc")) = "" & rs("JCCategoryDesc")
            
            .TextMatrix(r, .ColIndex("Location")) = "" & rs("Location")
            For w = 1 To 40
            .TextMatrix(r, .ColIndex("WBS" & format(w, "00"))) = "" & rs("WBS" & format(w, "00"))
            Next
            
            rs.MoveNext
        Wend
    
        Call GroupGrid
    End With
    
    'check if this assembly is a child component of another assembly.
    'cant have components if I am a component
    s = "select parentassemblyid from tbldbassemblycomponents where ComponentAssemblyID=" & DbQuote(Str, mAssemblyID)
    Set rs = HFApp.SqlExec(s)
    mComponentsEnabled = rs.EOF
    
    If Not mUseComponents Then mComponentsEnabled = False
    
    gComponents.Rows = 1
    gComponents.Enabled = mComponentsEnabled And Not mReadOnly
    If mComponentsEnabled Then
    
'        s = ""
'        s = s & "select m.assemblyid,m.community,m.assembly,m.model,m.optionid,m.description,c.qty" & vbCrLf
'        s = s & "from tbldbassemblycomponents c" & vbCrLf
'        s = s & "join tbldbassemblymaster m on c.componentassemblyid=m.assemblyid" & vbCrLf
'        s = s & "where c.ParentAssemblyid=" & DbQuote(Num, mAssemblyID) & vbCrLf
        
        s = ""
        s = s & "select" & vbCrLf
        s = s & "  m.assemblyid,m.community,m.assembly,m.model,m.optionid,m.description,c.qty" & vbCrLf
        s = s & " ,round(sum(" & vbCrLf
        s = s & "     isnuLL(dbo.Purch_GetItemRate(0,0," & DbQuote(Str, PricingCommunity) & "," & DbQuote(Str, PricingCommunityPhase) & ", d.Assembly,d.Model,d.OptionID,d.Phase,d.Item,d.Sequence,v.Vendor_id,getdate(),d.DivisionID),0)" & vbCrLf
        s = s & "   *(100 + isnull(t.JCRate,0))/100" & vbCrLf
        s = s & "   *isnull(d.OrderQty,0)" & vbCrLf
        s = s & " ),2) Cost" & vbCrLf
        s = s & "from tbldbassemblycomponents c" & vbCrLf
        s = s & "join tbldbassemblymaster m on c.componentassemblyid=m.assemblyid" & vbCrLf
        s = s & "left outer join tblDBAssemblyDetails d on m.assemblyid=d.assemblyid" & vbCrLf
        s = s & "left outer join tblPhaseItem i ON(d.DivisionID = i.DivisionID and d.Phase=i.Phase AND d.Item=i.Item)" & vbCrLf
        s = s & "left outer join tblPOIndex p ON(p.DivisionID = " & HFApp.DivisionID & " and isnull(nullif(d.POIndex,''),isnull(i.POIndex,''))=p.POIndex)" & vbCrLf
        s = s & "left outer join StandardCostCodes cc ON(cc.DivisionID = " & HFApp.DivisionID & " and isnull(nullif(i.JCCostCode,''),isnull(p.JCCostCode,''))=cc.CostCode)" & vbCrLf
        s = s & "left outer join StandardCategories ct ON(i.DivisionID = ct.DivisionID and i.JCCategory=ct.Category)" & vbCrLf
        s = s & "left outer join tblvendors v ON(v.DivisionID = d.DivisionID and v.vendor_id=dbo.Purch_GetCommunityVendor(" & DbQuote(Str, PricingCommunity) & ", p.POIndex," & HFApp.DivisionID & "))" & vbCrLf
        s = s & "left outer join taxGroups t on (t.DivisionID = " & HFApp.DivisionID & " and t.TaxGroup=dbo.Purch_GetDefaultTaxGroup('' , " & DbQuote(Str, PricingCommunity) & " ,'' , d.Model , d.Assembly ,i.phase ,i.item , v.Vendor_ID , i.JCCategory," & HFApp.DivisionID & "))" & vbCrLf
        s = s & "where c.ParentAssemblyid=" & DbQuote(Num, mAssemblyID) & vbCrLf
        s = s & "group by m.assemblyid,m.community,m.assembly,m.model,m.optionid,m.description,c.qty" & vbCrLf
        Set rs = HFApp.SqlExec(s)
        With gComponents
            .Rows = 1
            ComponentCost = 0
            While Not rs.EOF
                r = .Rows
                .AddItem ""
                .TextMatrix(r, .ColIndex("assemblyid")) = "" & rs("assemblyid")
                .TextMatrix(r, .ColIndex("community")) = "" & rs("community")
                .TextMatrix(r, .ColIndex("Model")) = "" & rs("model")
                .TextMatrix(r, .ColIndex("option")) = "" & rs("optionid")
                .TextMatrix(r, .ColIndex("assembly")) = "" & rs("assembly")
                .TextMatrix(r, .ColIndex("description")) = "" & rs("description")
                .TextMatrix(r, .ColIndex("qty")) = "" & rs("qty")
                
                If PricingCommunity <> "" Then
                    .TextMatrix(r, .ColIndex("cost")) = "" & rs("Cost")
                End If
                
                ComponentCost = ComponentCost + Val("" & rs("cost")) * Val("" & rs("qty"))
                
                rs.MoveNext
            Wend
        End With
    End If
    
    If PricingCommunity <> "" Then
        lblComponentCost.Caption = format(ComponentCost, "$#,##0.00")
        lblTotalCost.Caption = format(Round(ComponentCost, 2) + Round(AssemblyCost, 2), "$#,##0.00")
    End If
    
    
    Call Form_Resize   'to show/hide components
    
    Screen.MousePointer = vbDefault
    mDirty = False
    
Exit Sub
eh: Call errHandler(SRCFILE & "LoadData", s)
End Sub

Public Function SaveData(prompt As Boolean) As Boolean
On Error GoTo eh
    Dim i As Long
    Dim s As String
    Dim d As String
    Dim a As String
    Dim w As Long
    Dim step As String
    Dim rs As Recordset
    Dim c As New Connection
    Dim b As Boolean
    
'weird scenario... if you close the screen while the model dropdown is open, cboModel_Validate doesnt fire so the description is not removed. so call it explicitly
    
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
    
    If Not ValidateData Then Exit Function
    
    Screen.MousePointer = vbHourglass
    
    
    
    
    
    
step = "ASSEMBLY"
    If mCopiedAssembly <> "" Then
        'check that something changed
        
        If mCopiedKey = GetComboBoxListKey(cboCommunity) & Chr(1) & txtAssembly.Text & Chr(1) & cboModel.Text & Chr(1) & txtOption.Text Then
            Screen.MousePointer = vbDefault
            MsgBox "When copying an assembly you must change one of Community, Assembly, Model, or Option", vbExclamation, App.ProductName
            SaveData = False
            Exit Function
        End If
   
        'check that it doesnt already exist
        s = ""
        s = s & "select count(*)" & vbCrLf
        s = s & "from tblDBAssemblyMaster" & vbCrLf
        s = s & "where DivisionID=" & HFApp.DivisionID & vbCrLf
        s = s & "and community=" & DbQuote(Str, GetComboBoxListKey(cboCommunity)) & vbCrLf
        s = s & "and model=" & DbQuote(Str, cboModel.Text) & vbCrLf
        s = s & "and optionid=" & DbQuote(Str, txtOption.Text) & vbCrLf
        s = s & "and assembly=" & DbQuote(Str, txtAssembly.Text) & vbCrLf
        If HFApp.SqlExec(s)(0) <> 0 Then
            Screen.MousePointer = vbDefault
            MsgBox "This assembly already exists. You must specify a unique set of Community, Assembly, Model and/or Option", vbExclamation, App.ProductName
            SaveData = False
            Exit Function
        End If
   




        
        If vbYes = MsgBox("Do you want to copy the vendor pricelists also?", vbQuestion + vbYesNo, App.ProductName) Then
            s = ""
            s = s & "INSERT INTO tblVendorCost(DivisionID,Community,CommunityPhase,Assembly,Model,Phase,Item,Vendor,Current_Cost,Next_Cost1,Next_Cost2,Next_Effective1,Next_Effective2,Last_Cost1,Last_Cost2,Last_Cost3,Last1_Expiry,Last2_Expiry,Last3_Expiry,Forecast1,Forecast2,Forecast3,Forecast4,Forecast5,Forecast6,Forecast7,Forecast8,Forecast9,Forecast10,Forecast11,Forecast12,TaxGroup,PartNumber,PriceLink)" & vbCrLf
            s = s & "SELECT old.DivisionID" & vbCrLf
            's = s & "      ,old.CommunityPhase" & vbCrLf
            's = s & "      ,old.Community" & vbCrLf
            s = s & "      ," & DbQuote(Str, GetComboBoxListKey(cboCommunity)) & vbCrLf
            s = s & "      ,''" & vbCrLf
            s = s & "      ," & DbQuote(Str, txtAssembly.Text) & vbCrLf
            s = s & "      ," & DbQuote(Str, cboModel.Text) & vbCrLf
            s = s & "      ,old.Phase" & vbCrLf
            s = s & "      ,old.Item" & vbCrLf
            s = s & "      ,old.Vendor" & vbCrLf
            s = s & "      ,old.Current_Cost,old.Next_Cost1,old.Next_Cost2,old.Next_Effective1,old.Next_Effective2,old.Last_Cost1,old.Last_Cost2,old.Last_Cost3,old.Last1_Expiry,old.Last2_Expiry,old.Last3_Expiry" & vbCrLf
            s = s & "      ,old.Forecast1,old.Forecast2,old.Forecast3,old.Forecast4,old.Forecast5,old.Forecast6,old.Forecast7,old.Forecast8,old.Forecast9,old.Forecast10,old.Forecast11,old.Forecast12" & vbCrLf
            s = s & "      ,old.TaxGroup,old.PartNumber,old.PriceLink" & vbCrLf
            s = s & "  FROM tblVendorCost old" & vbCrLf
            s = s & "       LEFT OUTER JOIN tblVendorCost new ON(new.DivisionID = old.DivisionID" & vbCrLf
'            s = s & "                                        and new.Community=old.Community" & vbCrLf
'            s = s & "                                        AND new.CommunityPhase=old.CommunityPhase" & vbCrLf
            s = s & "                                        AND new.Community=" & DbQuote(Str, GetComboBoxListKey(cboCommunity)) & vbCrLf
            s = s & "                                        AND new.communityphase=''" & vbCrLf
            s = s & "                                        AND new.Assembly=" & DbQuote(Str, txtAssembly.Text) & vbCrLf
            s = s & "                                        AND new.Model=" & DbQuote(Str, cboModel.Text) & vbCrLf
            s = s & "                                        AND new.Phase=old.Phase" & vbCrLf
            s = s & "                                        AND new.Item=old.Item" & vbCrLf
            s = s & "                                        AND new.Vendor=old.Vendor)" & vbCrLf
            s = s & " WHERE new.Vendor IS NULL" & vbCrLf
            s = s & "   AND old.Assembly=" & DbQuote(Str, mCopiedAssembly) & vbCrLf
            s = s & "   AND old.Model=" & DbQuote(Str, mCopiedModel) & vbCrLf
            s = s & "   AND old.Community=" & DbQuote(Str, mCopiedCommunity) & vbCrLf
            s = s & "   AND isnull(old.CommunityPhase,'')=''" & vbCrLf
            s = s & "   AND old.DivisionID = " & HFApp.DivisionID
            Call HFApp.SqlExec(s)
        End If
    End If
    
    On Error Resume Next
    s = ""
    s = s & "INSERT INTO tblDBAssemblyMaster(DivisionID,UStmp,TStmp,AssemblyType,InActive,Community,Model,SCheduleTemplate,BillingFactor,BillingCode,BillingCategory,OptionID,Assembly,AssemblyUOM,Description,GraphicPath,MaxWidth,MaxLength,ConstCutoff,SpecDocument,Comments,Notes,Series,Elevation,Style,Bedrooms,Bathrooms,FloorArea,Category,JCExtra,TakeoffRequired,UseNormalSalesQtyFactors,IsBaseAssembly,Qty,Color,Location,IncludedOption,DesignCenterSalesOnly,SelectByRoom,DisplayTotalOnly)" & vbCrLf
    s = s & "VALUES(" & HFApp.DivisionID & "," & DbQuote(Str, HFApp.LoginID) & vbCrLf
    s = s & "      ,GETDATE()" & vbCrLf
    s = s & "      ," & DbQuote(Num, mAssemblyType) & vbCrLf
    s = s & "      ," & DbQuote(Bit, chkActive.value = vbUnchecked) & vbCrLf
    s = s & "      ," & DbQuote(Str, GetComboBoxListKey(cboCommunity)) & vbCrLf
    s = s & "      ," & DbQuote(Str, cboModel.Text) & vbCrLf
    s = s & "      ," & DbQuote(Str, cboScheduleTemplate.Text) & vbCrLf
    
    s = s & "      ," & DbQuote(Num, txtBillingFactor.Text) & vbCrLf
    s = s & "      ," & DbQuote(Str, GetComboBoxListKey(cboBillingCode)) & vbCrLf
    s = s & "      ," & DbQuote(Str, GetComboBoxListKey(cboBillingCategory)) & vbCrLf
    
    s = s & "      ," & DbQuote(Str, txtOption.Text) & vbCrLf
    s = s & "      ," & DbQuote(Str, txtAssembly.Text) & vbCrLf
    s = s & "      ," & DbQuote(Str, cboUOM.Text) & vbCrLf
    s = s & "      ," & DbQuote(Str, txtDescription.Text) & vbCrLf
    s = s & "      ," & DbQuote(Str, txtGraphicPath.Text) & vbCrLf
    s = s & "      ," & DbQuote(Num, txtMaxWidth.Text) & vbCrLf
    s = s & "      ," & DbQuote(Num, txtMaxLength.Text) & vbCrLf
    s = s & "      ," & DbQuote(Num, txtConstCutoff.Text) & vbCrLf
    s = s & "      ," & DbQuote(Str, txtSpecDocument.Text) & vbCrLf
    s = s & "      ," & DbQuote(Str, txtComments.Text) & vbCrLf
    s = s & "      ," & DbQuote(Str, txtNotes.Text) & vbCrLf
    s = s & "      ," & DbQuote(Str, GetComboBoxListKey(cboSeries)) & vbCrLf
    s = s & "      ," & DbQuote(Str, txtElevation.Text) & vbCrLf
    s = s & "      ," & DbQuote(Str, IIf(cboStyle.Visible, cboStyle.Text, txtStyle.Text)) & vbCrLf
    s = s & "      ," & DbQuote(Num, txtBedrooms.Text) & vbCrLf
    s = s & "      ," & DbQuote(Num, txtBathrooms.Text) & vbCrLf
    s = s & "      ," & DbQuote(Num, txtFloorArea.Text) & vbCrLf
    s = s & "      ," & DbQuote(Str, GetComboBoxListKey(cboCategory)) & vbCrLf
    s = s & "      ," & DbQuote(Str, txtJCExtra.Text) & vbCrLf
    s = s & "      ," & DbQuote(Bit, chkTakeoffRequired.value = vbChecked) & vbCrLf
    s = s & "      ," & DbQuote(Bit, chkUseNormalSalesQtyFactors.value = vbChecked) & vbCrLf
    s = s & "      ," & DbQuote(Bit, chkIsBaseAssembly.value = vbChecked) & vbCrLf
    s = s & "      ," & DbQuote(Num, txtQty.Text) & vbCrLf
    s = s & "      ," & DbQuote(Str, txtAColor.Text) & vbCrLf
    s = s & "      ," & DbQuote(Str, GetComboBoxListKey(cboLocation)) & vbCrLf
    s = s & "      ," & DbQuote(Bit, chkIncludedOption.value = vbChecked) & vbCrLf
    s = s & "      ," & DbQuote(Bit, Me.chkDCSalesOnly.value = vbChecked) & vbCrLf
    s = s & "      ," & DbQuote(Bit, Me.chkSelectByRoom.value = vbChecked) & vbCrLf
    s = s & "      ," & DbQuote(Bit, Me.chkDisplayTotalOnly.value = vbChecked) & ")" & vbCrLf
    Call HFApp.SqlExec(s)
    If Err.Number = 0 Then
        mAssemblyID = HFApp.SqlIdentity("tbldbassemblymaster")
    End If
    
    If mCopiedAssemblyID <> 0 Then
        s = ""
        s = s & "insert tbldbassemblycomponents(parentassemblyid,componentassemblyid,qty)" & vbCrLf
        s = s & "select " & DbQuote(Num, mAssemblyID) & ",componentassemblyid,qty " & vbCrLf
        s = s & "from tbldbassemblycomponents" & vbCrLf
        s = s & "where parentassemblyid=" & DbQuote(Num, mCopiedAssemblyID) & vbCrLf
        Call HFApp.SqlExec(s)
        
        
        s = ""
        s = s & "insert RoomQtyByModel(Model, RoomID, SubCategory, UOM, Qty, Inactive)" & vbCrLf
        s = s & "select " & DbQuote(Str, cboModel.Text) & " Model, a.RoomID, a.SubCategory, a.UOM, a.Qty, a.Inactive" & vbCrLf
        s = s & "from RoomMaster r" & vbCrLf
        s = s & "join RoomQtyByModel a on r.roomid=a.roomid" & vbCrLf
        s = s & "left join RoomQtyByModel b on b.model=" & DbQuote(Str, cboModel.Text) & " and a.roomid=b.roomid and a.subcategory=b.subcategory and a.uom=b.uom" & vbCrLf
        s = s & "where r.divisionid=1" & vbCrLf
        s = s & "and a.Model=" & DbQuote(Str, mCopiedModel) & vbCrLf
        s = s & "and b.model is null" & vbCrLf
        Call HFApp.SqlExec(s)
    
    
    End If
    
    CtrlEnabled(cboCommunity) = False
    CtrlEnabled(cboModel) = False
    
    'save these temporarily
    s = cboModel.Text
    d = txtDescription.Text
    a = txtAssembly.Text
    'restore selected values
    Call SetComboBoxListIndex(cboModel, s)
    txtDescription.Text = d
    txtAssembly.Text = a
    
    CtrlEnabled(txtOption) = False
    CtrlEnabled(txtAssembly) = False
    On Error GoTo eh
    
    mCommunity = GetComboBoxListKey(cboCommunity)
    mAssembly = txtAssembly.Text
    mModel = cboModel.Text
    mOptionID = txtOption.Text
    
    
    
    
    
    s = ""
    s = s & "UPDATE tblDBAssemblyMaster" & vbCrLf
    s = s & "SET UStmp=" & DbQuote(Str, HFApp.LoginID) & vbCrLf
    s = s & "   ,TStmp=GETDATE()" & vbCrLf
    s = s & "   ,Community=" & DbQuote(Str, GetComboBoxListKey(cboCommunity)) & vbCrLf
    
    s = s & "   ,BillingFactor=" & DbQuote(Num, txtBillingFactor.Text) & vbCrLf
    s = s & "   ,BillingCode=" & DbQuote(Str, GetComboBoxListKey(cboBillingCode)) & vbCrLf
    s = s & "   ,BillingCategory=" & DbQuote(Str, GetComboBoxListKey(cboBillingCategory)) & vbCrLf
    
    s = s & "   ,Model=" & DbQuote(Str, cboModel.Text) & vbCrLf
    s = s & "   ,ScheduleTemplate=" & DbQuote(Str, cboScheduleTemplate.Text) & vbCrLf
    s = s & "   ,OptionID=" & DbQuote(Str, txtOption.Text) & vbCrLf
    s = s & "   ,Assembly=" & DbQuote(Str, txtAssembly.Text) & vbCrLf
    s = s & "   ,AssemblyUOM=" & DbQuote(Str, cboUOM.Text) & vbCrLf
    s = s & "   ,Description=" & DbQuote(Str, txtDescription.Text) & vbCrLf
    s = s & "   ,GraphicPath=" & DbQuote(Str, txtGraphicPath.Text) & vbCrLf
    s = s & "   ,SpecDocument=" & DbQuote(Str, txtSpecDocument.Text) & vbCrLf
    s = s & "   ,MaxWidth=" & DbQuote(Str, txtMaxWidth.Text) & vbCrLf
    s = s & "   ,MaxLength=" & DbQuote(Str, txtMaxLength.Text) & vbCrLf
    s = s & "   ,ConstCutoff=" & DbQuote(Str, txtConstCutoff.Text) & vbCrLf
    s = s & "   ,Comments=" & DbQuote(Str, txtComments.Text) & vbCrLf
    s = s & "   ,Notes=" & DbQuote(Str, txtNotes.Text) & vbCrLf
    s = s & "   ,InActive=" & DbQuote(Bit, chkActive.value = vbUnchecked) & vbCrLf
    s = s & "   ,Series=" & DbQuote(Str, GetComboBoxListKey(cboSeries)) & vbCrLf
    s = s & "   ,Style=" & DbQuote(Str, IIf(cboStyle.Visible, cboStyle.Text, txtStyle.Text)) & vbCrLf
    s = s & "   ,Elevation=" & DbQuote(Str, txtElevation.Text) & vbCrLf
    s = s & "   ,Bedrooms=" & DbQuote(Num, txtBedrooms.Text) & vbCrLf
    s = s & "   ,Bathrooms=" & DbQuote(Num, txtBathrooms.Text) & vbCrLf
    s = s & "   ,FloorArea=" & DbQuote(Num, txtFloorArea.Text) & vbCrLf
    s = s & "   ,Category=" & DbQuote(Str, GetComboBoxListKey(cboCategory)) & vbCrLf
    s = s & "   ,JCExtra=" & DbQuote(Str, txtJCExtra.Text) & vbCrLf
    s = s & "   ,TakeoffRequired=" & DbQuote(Bit, chkTakeoffRequired.value = vbChecked) & vbCrLf
    s = s & "   ,UseNormalSalesQtyFactors=" & DbQuote(Bit, chkUseNormalSalesQtyFactors.value = vbChecked) & vbCrLf
    s = s & "   ,IsBaseAssembly=" & DbQuote(Bit, chkIsBaseAssembly.value = vbChecked) & vbCrLf
    s = s & "   ,qty=" & DbQuote(Num, txtQty.Text) & vbCrLf
    s = s & "   ,color=" & DbQuote(Str, txtAColor.Text) & vbCrLf
    s = s & "   ,stylevalue=" & DbQuote(Str, txtAStyle.Text) & vbCrLf
    s = s & "   ,finishvalue=" & DbQuote(Str, txtAFinish.Text) & vbCrLf
    s = s & "   ,othervalue=" & DbQuote(Str, txtAOther.Text) & vbCrLf
    s = s & "   ,colorlistid=" & DbQuote(Num, txtLColor.Text) & vbCrLf
    s = s & "   ,stylelistid=" & DbQuote(Num, txtLStyle.Text) & vbCrLf
    s = s & "   ,finishlistid=" & DbQuote(Num, txtLFinish.Text) & vbCrLf
    s = s & "   ,otherlistid=" & DbQuote(Num, txtLOther.Text) & vbCrLf
    s = s & "   ,Location=" & DbQuote(Str, GetComboBoxListKey(cboLocation)) & vbCrLf
    s = s & "   ,IncludedOption=" & DbQuote(Bit, chkIncludedOption.value = vbChecked) & vbCrLf
    
    s = s & "   ,DesignCenterSalesOnly=" & DbQuote(Bit, Me.chkDCSalesOnly.value = vbChecked) & vbCrLf
    s = s & "   ,SelectByRoom=" & DbQuote(Bit, Me.chkSelectByRoom.value = vbChecked) & vbCrLf
    s = s & "   ,DisplayTotalOnly=" & DbQuote(Bit, Me.chkDisplayTotalOnly.value = vbChecked) & vbCrLf
    
    s = s & "WHERE Community=" & DbQuote(Str, GetComboBoxListKey(cboCommunity)) & vbCrLf
    s = s & "  AND Model=" & DbQuote(Str, cboModel.Text) & vbCrLf
    s = s & "  AND OptionID=" & DbQuote(Str, txtOption.Text) & vbCrLf
    s = s & "  AND Assembly=" & DbQuote(Str, txtAssembly.Text) & vbCrLf
    s = s & "  AND DivisionID = " & HFApp.DivisionID
    
    Call HFApp.SqlExec(s)
    
step = "ITEMS"
    With gItems
        
        'deletes only do deletes if you are not copying
        If mCopiedAssembly = "" Then
            For i = .Rows - 1 To 1 Step -1
            If .RowData(i) = "delete" And .ValueMatrix(i, .ColIndex("Sequence")) <> 0 Then
                s = "delete tbldbassemblydetails where sequence=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Sequence")))
                HFApp.SqlExec s
                .RemoveItem i
            End If
            Next
        End If
        
        
        'inserts
        For i = 1 To .Rows - 1
        If Not .IsSubtotal(i) And .RowData(i) = "new" Then
            s = ""
            s = s & "INSERT INTO tblDBAssemblyDetails(AssemblyID,DivisionID,Community,Assembly,Model,OptionID,Phase,Item,ItemChart,Notes,UStmp,TStmp,Formula,TakeoffQty,OrderQty,POIndex,Location,UseModelCost,Invertable" & vbCrLf
            For w = 1 To 40
                s = s & ",WBS" & format(w, "00")
            Next
            s = s & ",ConversionFactor)" & vbCrLf
            s = s & "VALUES(" & DbQuote(Num, mAssemblyID)
            s = s & "," & DbQuote(Num, HFApp.DivisionID)
            s = s & "," & DbQuote(Str, GetComboBoxListKey(cboCommunity))
            s = s & "," & DbQuote(Str, txtAssembly.Text)
            s = s & "," & DbQuote(Str, cboModel.Text)
            s = s & "," & DbQuote(Str, txtOption.Text)
            s = s & "," & DbQuote(Str, .TextMatrix(i, .ColIndex("Phase")))
            s = s & "," & DbQuote(Str, .TextMatrix(i, .ColIndex("Item")))
            s = s & "," & DbQuote(Str, .TextMatrix(i, .ColIndex("ItemChart")))
            s = s & "," & DbQuote(Str, .TextMatrix(i, .ColIndex("Notes")))
            s = s & "," & DbQuote(Str, HFApp.LoginID)
            s = s & ",GETDATE()"
            s = s & "," & DbQuote(Str, .TextMatrix(i, .ColIndex("Formula")))
            s = s & "," & DbQuote(Num, .TextMatrix(i, .ColIndex("TakeoffQty")))
            s = s & "," & DbQuote(Num, .TextMatrix(i, .ColIndex("OrderQty")))
            s = s & "," & DbQuote(Str, .Cell(flexcpData, i, .ColIndex("POIndex")))
            s = s & "," & DbQuote(Str, .TextMatrix(i, .ColIndex("Location")))
            s = s & "," & DbQuote(Bit, mAssemblyType = atoption And .Cell(flexcpChecked, i, .ColIndex("UseModelCost")) = flexChecked) & vbCrLf
            s = s & "," & DbQuote(Bit, .Cell(flexcpChecked, i, .ColIndex("Invertable")) = flexChecked) & vbCrLf
            For w = 1 To 40
                s = s & "," & DbQuote(Str, .TextMatrix(i, .ColIndex("WBS" & format(w, "00"))))
            Next
            s = s & "," & DbQuote(Num, .TextMatrix(i, .ColIndex("AssemblyConversionFactor")))
            s = s & ")" & vbCrLf & vbCrLf
            HFApp.SqlExec s
            .TextMatrix(i, .ColIndex("Sequence")) = HFApp.SqlIdentity("tbldbassemblydetails")
                        
            'if itemtype changed then update item db
            If .Cell(flexcpData, i, .ColIndex("ItemType")) = "Dirty" And .TextMatrix(i, .ColIndex("ItemType")) = "Quote" Then
                s = s & "UPDATE tblPhaseItem"
                s = s & " SET IsQuote=1"
                s = s & " WHERE DivisionID=" & HFApp.DivisionID
                s = s & " AND Phase=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Phase")))
                s = s & " AND Item=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Item"))) & vbCrLf & vbCrLf
            End If
                
            'if price changed then update vendor pricelists
            If .Cell(flexcpData, i, .ColIndex("Price")) = "Dirty" Then
                Call UpdateItemPrice(i)
            End If
            
            'reset statuses
            .Cell(flexcpData, i, .ColIndex("ItemType")) = ""
            .Cell(flexcpData, i, .ColIndex("Price")) = ""
            .RowData(i) = ""
            
        End If
        Next

        'updates
        For i = 1 To .Rows - 1
        If Not .IsSubtotal(i) And .RowData(i) = "update" Then
            s = ""
            s = s & "UPDATE tblDBAssemblyDetails SET"
            s = s & " Assemblyid=" & DbQuote(Num, mAssemblyID)
            s = s & ",DivisionID=" & DbQuote(Num, HFApp.DivisionID)
            s = s & ",Community=" & DbQuote(Str, GetComboBoxListKey(cboCommunity))
            s = s & ",Assembly=" & DbQuote(Str, txtAssembly.Text)
            s = s & ",Model=" & DbQuote(Str, cboModel.Text)
            s = s & ",OptionID=" & DbQuote(Str, txtOption.Text)
            s = s & ",Phase=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Phase")))
            s = s & ",Item=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Item")))
            s = s & ",ItemChart=" & DbQuote(Str, .TextMatrix(i, .ColIndex("ItemChart")))
            s = s & ",Notes=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Notes")))
            s = s & ",ustmp=" & DbQuote(Str, HFApp.LoginID)
            s = s & ",tstmp=GETDATE()"
            s = s & ",formula=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Formula")))
            s = s & ",takeoffqty=" & DbQuote(Num, .TextMatrix(i, .ColIndex("TakeoffQty")))
            s = s & ",orderqty=" & DbQuote(Num, .TextMatrix(i, .ColIndex("OrderQty")))
            s = s & ",poindex=" & DbQuote(Str, .Cell(flexcpData, i, .ColIndex("POIndex")))
            s = s & ",location=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Location")))
            s = s & ",usemodelcost=" & DbQuote(Bit, mAssemblyType = atoption And .Cell(flexcpChecked, i, .ColIndex("UseModelCost")) = flexChecked) & vbCrLf
            s = s & ",invertable=" & DbQuote(Bit, .Cell(flexcpChecked, i, .ColIndex("Invertable")) = flexChecked) & vbCrLf
            For w = 1 To 40
                s = s & ",WBS" & format(w, "00") & "=" & DbQuote(Str, .TextMatrix(i, .ColIndex("WBS" & format(w, "00"))))
            Next
            s = s & ",ConversionFactor=" & DbQuote(Num, .TextMatrix(i, .ColIndex("AssemblyConversionFactor"))) & vbCrLf
            s = s & "where sequence=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Sequence"))) & vbCrLf
            HFApp.SqlExec s
            
            'if itemtype changed then update item db
            If .Cell(flexcpData, i, .ColIndex("ItemType")) = "Dirty" And .TextMatrix(i, .ColIndex("ItemType")) = "Quote" Then
                s = s & "UPDATE tblPhaseItem"
                s = s & " SET IsQuote=1"
                s = s & " WHERE DivisionID=" & HFApp.DivisionID
                s = s & " AND Phase=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Phase")))
                s = s & " AND Item=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Item"))) & vbCrLf & vbCrLf
            End If
                
            'if price changed then update vendor pricelists
            If .Cell(flexcpData, i, .ColIndex("Price")) = "Dirty" Then
                Call UpdateItemPrice(i)
            End If
            
            'reset statuses
            .Cell(flexcpData, i, .ColIndex("ItemType")) = ""
            .Cell(flexcpData, i, .ColIndex("Price")) = ""
            .RowData(i) = ""
            
        End If
        Next
    
    
    End With
 
    With gComponents
    For i = .Rows - 1 To 1 Step -1
        Select Case .RowData(i)
        Case "delete"
            s = ""
            s = s & "delete tbldbassemblycomponents" & vbCrLf
            s = s & "where parentassemblyid=" & DbQuote(Num, mAssemblyID) & vbCrLf
            s = s & "and componentassemblyid=" & DbQuote(Num, .TextMatrix(i, .ColIndex("assemblyid")))
            Call HFApp.SqlExec(s)
            Call .RemoveItem(i)
            
        Case "new"
            s = ""
            s = s & "insert tbldbassemblycomponents(ParentAssemblyID, ComponentAssemblyID, Qty) values" & vbCrLf
            s = s & "(" & DbQuote(Num, mAssemblyID) & vbCrLf
            s = s & "," & DbQuote(Str, .TextMatrix(i, .ColIndex("assemblyid"))) & vbCrLf
            s = s & "," & DbQuote(Num, .TextMatrix(i, .ColIndex("Qty"))) & vbCrLf
            s = s & ")"
            Call HFApp.SqlExec(s)
            .RowData(i) = ""
        
        Case "dirty"
            s = ""
            s = s & "update tbldbassemblycomponents" & vbCrLf
            s = s & "set qty=" & DbQuote(Num, .TextMatrix(i, .ColIndex("qty"))) & vbCrLf
            s = s & "where parentassemblyid=" & DbQuote(Num, mAssemblyID) & vbCrLf
            s = s & "and componentassemblyid=" & DbQuote(Num, .TextMatrix(i, .ColIndex("assemblyid")))
            Call HFApp.SqlExec(s)
            .RowData(i) = ""
            
        End Select
    Next
    End With
    
    If mDataSource = 1 Then
        'created an assembly for an unestimated sales item so assign assembly back to sales
        s = ""
        Select Case mAssemblyType
            Case atModel
                s = s & "UPDATE tblModels" & vbCrLf
                s = s & "SET Assembly=" & DbQuote(Str, txtAssembly.Text) & vbCrLf
                s = s & "   ,Description=" & DbQuote(Str, txtDescription.Text) & vbCrLf
                If Trim(txtComments.Text) <> "" Then s = s & "   ,Comments=" & DbQuote(Str, txtComments.Text) & vbCrLf
                s = s & "   ,Series=" & DbQuote(Str, GetComboBoxListKey(cboSeries)) & vbCrLf
                s = s & "   ,Style=" & DbQuote(Str, IIf(cboStyle.Visible, cboStyle.Text, txtStyle.Text)) & vbCrLf
                s = s & "   ,NoOfBedrooms=" & DbQuote(Num, txtBedrooms.Text) & vbCrLf
                s = s & "   ,NoOfBathrooms=" & DbQuote(Num, txtBathrooms.Text) & vbCrLf
                s = s & "   ,modelsize=" & DbQuote(Num, txtFloorArea.Text) & vbCrLf
                s = s & "WHERE ISNULL(Area,'')=" & DbQuote(Str, GetComboBoxListKey(cboCommunity)) & vbCrLf
                s = s & "  AND ISNULL(Assembly,'')=''" & vbCrLf
                s = s & "  AND ISNULL(Model,'')=" & DbQuote(Str, cboModel.Text) & vbCrLf
                s = s & "  AND ISNULL(Series,'')=" & DbQuote(Str, GetComboBoxListKey(cboSeries)) & vbCrLf
                s = s & "  AND DivisionID = " & HFApp.DivisionID & vbCrLf
            Case atoption
                s = s & "UPDATE tblOptions" & vbCrLf
                s = s & "SET Assembly=" & DbQuote(Str, txtAssembly.Text) & vbCrLf
                s = s & "   ,Description=" & DbQuote(Str, txtDescription.Text) & vbCrLf
                s = s & "   ,UOM=" & DbQuote(Str, cboUOM.Text) & vbCrLf
                If Trim(txtComments.Text) <> "" Then s = s & "   ,Comments=" & DbQuote(Str, txtComments.Text) & vbCrLf
                s = s & "   ,Category=" & DbQuote(Str, GetComboBoxListKey(cboCategory)) & vbCrLf
                s = s & "   ,TL_Extra=" & DbQuote(Str, txtJCExtra.Text) & vbCrLf
                s = s & "WHERE ISNULL(Area,'')=" & DbQuote(Str, GetComboBoxListKey(cboCommunity)) & vbCrLf
                s = s & "  AND ISNULL(Assembly,'')=''" & vbCrLf
                s = s & "  AND Model=" & DbQuote(Str, cboModel.Text) & vbCrLf
                s = s & "  AND Series=" & DbQuote(Str, GetComboBoxListKey(cboSeries)) & vbCrLf
                s = s & "  AND Opt=" & DbQuote(Str, txtOption.Text) & vbCrLf
                s = s & "  AND DivisionID = " & HFApp.DivisionID & vbCrLf
            Case atDesignCenter
                s = s & "UPDATE tblDCOptions" & vbCrLf
                s = s & "SET Assembly=" & DbQuote(Str, txtAssembly.Text) & vbCrLf
                s = s & "   ,Description=" & DbQuote(Str, txtDescription.Text) & vbCrLf
                s = s & "   ,UOM=" & DbQuote(Str, cboUOM.Text) & vbCrLf
                If Trim(txtComments.Text) <> "" Then s = s & "   ,Comments=" & DbQuote(Str, txtComments.Text) & vbCrLf
                s = s & "   ,Category=" & DbQuote(Str, GetComboBoxListKey(cboCategory)) & vbCrLf
                s = s & "   ,TL_Extra=" & DbQuote(Str, txtJCExtra.Text) & vbCrLf
                s = s & "WHERE ISNULL(Community,'')=" & DbQuote(Str, GetComboBoxListKey(cboCommunity)) & vbCrLf
                s = s & "  AND ISNULL(Assembly,'')=''" & vbCrLf
                s = s & "  AND Opt=" & DbQuote(Str, txtOption.Text) & vbCrLf
                s = s & "  AND DivisionID = " & HFApp.DivisionID & vbCrLf
            Case atGlobal
                s = s & "UPDATE tblGlobalOptions" & vbCrLf
                s = s & "SET Assembly=" & DbQuote(Str, txtAssembly.Text) & vbCrLf
                s = s & "   ,Description=" & DbQuote(Str, txtDescription.Text) & vbCrLf
                s = s & "   ,UOM=" & DbQuote(Str, cboUOM.Text) & vbCrLf
                If Trim(txtComments.Text) <> "" Then s = s & "   ,Comments=" & DbQuote(Str, txtComments.Text) & vbCrLf
                s = s & "   ,Category=" & DbQuote(Str, GetComboBoxListKey(cboCategory)) & vbCrLf
                s = s & "   ,TL_Extra=" & DbQuote(Str, txtJCExtra.Text) & vbCrLf
                If GetComboBoxListKey(cboCommunity) <> "" Then
                    s = s & "WHERE ISNULL(Community,'')=" & DbQuote(Str, GetComboBoxListKey(cboCommunity)) & vbCrLf
                    s = s & "  AND ISNULL(Assembly,'')=''" & vbCrLf
                    s = s & "  AND Opt=" & DbQuote(Str, txtOption.Text) & vbCrLf
                Else
                    s = s & "WHERE ISNULL(Assembly,'')=''" & vbCrLf
                    s = s & "  AND Opt=" & DbQuote(Str, txtOption.Text) & vbCrLf
                End If
                s = s & "  AND DivisionID = " & HFApp.DivisionID & vbCrLf
        End Select
        Call HFApp.SqlExec(s)
        mDataSource = 0
    End If
    
    'now clear these
    mCopiedCommunity = ""
    mCopiedAssembly = ""
    mCopiedAssemblyID = 0
    mCopiedModel = ""
    mCopiedKey = ""
    
    SaveData = True
    
    mDirty = False
    Screen.MousePointer = vbDefault
Exit Function
eh: If step = "ASSEMBLY" And InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        Screen.MousePointer = vbDefault
        If vbYes = MsgBox("This assembly already exists. Do you want to replace it with this one?", vbExclamation + vbYesNo, App.ProductName) Then
            Screen.MousePointer = vbHourglass
            Resume Next
        End If
    ElseIf step = "ITEMS" Then
        Call errHandler(SRCFILE & "SaveData")
    Else
        Call errHandler(SRCFILE & "SaveData", s)
    End If
End Function




Private Property Get ReadOnly() As Boolean
    ReadOnly = mReadOnly
End Property

Private Property Let ReadOnly(RHS As Boolean)
    mReadOnly = RHS
    
    cboPriceCommunity.Enabled = Not RHS
    chkUseNormalSalesQtyFactors.Enabled = Not RHS
    chkIsBaseAssembly.Enabled = Not RHS
    txtElevation.Enabled = Not RHS
    gComponents.Enabled = Not RHS
    gItems.Enabled = Not RHS
    cboCommunity.Enabled = Not RHS
    cboModel.Enabled = Not RHS
    txtAssembly.Enabled = Not RHS
    chkActive.Enabled = Not RHS
        
    txtDescription.Enabled = Not RHS
    txtGraphicPath.Enabled = Not RHS
    txtMaxWidth.Enabled = Not RHS
    txtMaxLength.Enabled = Not RHS
    txtSpecDocument.Enabled = Not RHS
    txtConstCutoff.Enabled = Not RHS
    cmdBrowse(0).Enabled = Not RHS
    cmdBrowse(1).Enabled = Not RHS
    txtComments.Enabled = Not RHS
    txtNotes.Enabled = Not RHS
    txtQty.Enabled = Not RHS
    cboLocation.Enabled = Not RHS
    txtAColor.Enabled = Not RHS
    chkIncludedOption.Enabled = Not RHS
    
    Me.chkDCSalesOnly.Enabled = Not RHS
    Me.chkSelectByRoom.Enabled = Not RHS
    Me.chkDisplayTotalOnly.Enabled = Not RHS
    
    cboSeries.Enabled = Not RHS
    txtStyle.Enabled = Not RHS
    txtBedrooms.Enabled = Not RHS
    txtBathrooms.Enabled = Not RHS
    txtFloorArea.Enabled = Not RHS
    cboCategory.Enabled = Not RHS
    txtJCExtra.Enabled = Not RHS
    chkTakeoffRequired.Enabled = Not RHS
    
    txtBillingFactor.Enabled = Not RHS
    cboBillingCode.Enabled = Not RHS
    cboBillingCategory.Enabled = Not RHS
    
    
    Toolbar.Buttons("Save").Enabled = Not RHS
    'Toolbar.Buttons("Delete").Enabled = Not RHS
    Toolbar.Buttons("TakeoffOneTime").Enabled = False
    Toolbar.Buttons("TakeoffItem").Enabled = Not RHS
    Toolbar.Buttons("TakeoffAssembly").Enabled = Not RHS
    Toolbar.Buttons("TakeoffPlanSwift").Enabled = Not RHS
    Toolbar.Buttons("TakeoffCustom").Enabled = Not RHS
    Toolbar.Buttons("TakeoffItemChart").Enabled = Not RHS
    
    Toolbar.Buttons("PricelistExport").Enabled = Not RHS
    
    
End Property



Private Sub Toolbar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.key
        Case "ImportModels"
            mImportGlobals = False
            If Not SaveData(True) Then Exit Sub
            Call ImportPipelineModelsAndOptions
        
        Case "ImportGlobals":
            mImportGlobals = True
            If Not SaveData(True) Then Exit Sub
            Call ImportPipelineGlobalOptions
        
        Case Else
            Call FItemChart.Edit("")
    End Select

End Sub

Private Sub txtBathrooms_Validate(Cancel As Boolean)
    txtBathrooms.Text = Val(txtBathrooms.Text)
End Sub

Private Sub txtBedrooms_Validate(Cancel As Boolean)
    txtBedrooms.Text = Val(txtBedrooms.Text)
End Sub

Private Sub txtBillingFactor_Change()
    txtBillingFactor.Text = Val(txtBillingFactor.Text)
    mDirty = True
End Sub

Private Sub txtConstCutoff_GotFocus()
    SelectAll txtConstCutoff
End Sub

Private Sub txtElevation_Change()
    mDirty = True
End Sub

Private Sub txtElevation_GotFocus()
    SelectAll txtElevation
End Sub

Private Sub txtFloorArea_Change()
    mDirty = True
End Sub

Private Sub txtFloorArea_GotFocus()
    SelectAll txtFloorArea
End Sub

Private Sub txtMaxLength_GotFocus()
    SelectAll txtMaxLength
End Sub
Private Sub txtMaxWidth_GotFocus()
    SelectAll txtMaxWidth
End Sub

Private Sub txtMaxWidth_Change()
    txtMaxWidth.Text = Val(txtMaxWidth.Text)
    mDirty = True
End Sub
Private Sub txtMaxLength_Change()
    txtMaxLength.Text = Val(txtMaxLength.Text)
    mDirty = True
End Sub
Private Sub txtConstCutoff_Change()
    txtConstCutoff.Text = Val(txtConstCutoff.Text)
    mDirty = True
End Sub

Private Sub txtNotes_Change()
    mDirty = True
End Sub

Private Sub txtOption_GotFocus()
    SelectAll txtOption
End Sub
Private Sub txtAssembly_GotFocus()
    SelectAll txtAssembly
End Sub
Private Sub txtDescription_GotFocus()
    SelectAll txtDescription
End Sub
Private Sub txtGraphicPath_GotFocus()
    SelectAll txtGraphicPath
End Sub

Private Sub txtQty_Change()
    mDirty = True
End Sub

Private Sub txtQty_GotFocus()
    SelectAll txtQty
End Sub



Private Sub txtQty_Validate(Cancel As Boolean)
    txtQty.Text = Val(txtQty.Text)
End Sub

Private Sub txtSpecDocument_GotFocus()
    SelectAll txtSpecDocument
End Sub

Private Sub txtStyle_GotFocus()
    SelectAll txtStyle
End Sub
Private Sub txtBedrooms_GotFocus()
    SelectAll txtBedrooms
End Sub
Private Sub txtBathrooms_GotFocus()
    SelectAll txtBathrooms
End Sub
Private Sub txtJCExtra_GotFocus()
    SelectAll txtJCExtra
End Sub




Private Sub txtOption_Change()
    mDirty = True
    If mAssemblyType <> atModel Then txtAssembly.Text = txtOption.Text
End Sub
Private Sub txtAssembly_Change()
    mDirty = True
End Sub
Private Sub txtDescription_Change()
    mDirty = True
End Sub
Private Sub txtSpecDocument_Change()
    mDirty = True
End Sub
Private Sub txtGraphicPath_Change()
    mDirty = True
End Sub
Private Sub txtComments_Change()
    mDirty = True
End Sub
Private Sub txtStyle_Change()
    mDirty = True
End Sub
Private Sub txtBedrooms_Change()
    mDirty = True
End Sub
Private Sub txtBathrooms_Change()
    mDirty = True
End Sub
Private Sub txtJCExtra_Change()
    mDirty = True
End Sub
Private Sub cboCommunity_Click()
    mDirty = True
    Call LoadPricingCommunities(GetComboBoxListKey(cboCommunity))
End Sub
Private Sub cboSeries_Click()
    mDirty = True
End Sub
Private Sub cboCategory_Click()
    mDirty = True
End Sub

Private Sub gItems_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
    Dim grpd As Boolean
    
    With gItems
        If Not .Enabled Then Exit Sub
        If Button = vbRightButton Then
            If .MouseRow < 1 Then
                Cancel = True
                
                
                FMain.mnuGrid2Sub(mcGRID_GROUP).checked = .ColData(.MouseCol) = "GROUPED"
                grpd = FMain.mnuGrid2Sub(mcGRID_GROUP).checked
                Call FMain.ShowColumnMenu(gItems, Not grpd, , , True)
                
            Else
            
            
                FMain.mnuAssemblyItemsSub(mcITEM_EDITITEMCHART).Enabled = .TextMatrix(.Row, .ColIndex("ItemChart")) <> ""
            
                Call PopupMenu(FMain.mnuAssemblyItems)
            End If
        End If
    End With
End Sub


Public Sub AddItem(Assembly As String, _
                   AssemblyDescription As String, _
                   Model As String, _
                   Phase As String, _
                   Item As String, Sequence As Long, _
                   Description As String, _
                   OrderQty As Double, _
                   OrderUOM As String, _
                   TakeoffQty As Double, _
                   TakeoffUOM As String, _
                   ConversionFactor As Double, RoundTo As Double, RoundDir As Long, WastePercent As Long, _
                   JCExtra As String, _
                   JCCostCode As String, _
                   JCCostCodeDesc As String, JCCategory As String, _
                   JCCategoryDesc As String, Vendor As String, _
                   VendorName As String, price As Double, _
                   TaxGroup As String, TaxGroupName As String, _
                   JCTaxRate As Double, NJCTaxRate As Double, _
                   POIndex As String, _
                   Comments As String, Formula As String, SalesQty As Double, Location As String, _
                   WBS, Optional Job As String)
On Error GoTo eh
    
    
'WBS is a 40 element array
    
    Dim i As Long
    Dim w As Long
    Dim s As String
    Dim rs As Recordset
    
    ReadOnly = False
    mDirty = True
    
    If ConversionFactor = 0 Then ConversionFactor = 1
    
    
    With gItems
        .AddItem ""
        i = .Rows - 1
        .TextMatrix(i, .ColIndex("Phase")) = Phase
        .TextMatrix(i, .ColIndex("Item")) = Item
        
        On Error Resume Next
        .TextMatrix(i, .ColIndex("ItemNumber")) = HFApp.SqlExec("select itemnumber from tblphaseitem where divisionid=" & HFApp.DivisionID & " and phase=" & DbQuote(Str, Phase) & " and item=" & DbQuote(Str, Item))(0)
        .TextMatrix(i, .ColIndex("ItemConversionFactor")) = HFApp.SqlExec("select conversionfactor from tblphaseitem where divisionid=" & HFApp.DivisionID & " and phase=" & DbQuote(Str, Phase) & " and item=" & DbQuote(Str, Item))(0)
        On Error GoTo eh
        
        .TextMatrix(i, .ColIndex("Description")) = Description
        .TextMatrix(i, .ColIndex("JCCostCode")) = JCCostCode
        .TextMatrix(i, .ColIndex("JCCostCodeDesc")) = JCCostCodeDesc
        .TextMatrix(i, .ColIndex("JCCategory")) = JCCategory
        .TextMatrix(i, .ColIndex("JCCategoryDesc")) = JCCategoryDesc
        .Cell(flexcpData, i, .ColIndex("POIndex")) = POIndex
        .TextMatrix(i, .ColIndex("AssemblyConversionFactor")) = ConversionFactor
        .TextMatrix(i, .ColIndex("RoundTo")) = RoundTo
        .TextMatrix(i, .ColIndex("RoundDir")) = RoundDir
        .TextMatrix(i, .ColIndex("WastePercent")) = WastePercent
        .TextMatrix(i, .ColIndex("TakeoffUOM")) = TakeoffUOM
        .TextMatrix(i, .ColIndex("TakeoffQty")) = TakeoffQty
        .TextMatrix(i, .ColIndex("Formula")) = Formula
        .TextMatrix(i, .ColIndex("OrderUOM")) = OrderUOM
        .TextMatrix(i, .ColIndex("OrderQty")) = OrderQty
        .TextMatrix(i, .ColIndex("Location")) = Location
        .TextMatrix(i, .ColIndex("Notes")) = Comments
        .RowData(i) = "new"
        
        For w = 1 To 40
            .TextMatrix(i, .ColIndex("WBS" & format(w, "00"))) = WBS(w)
        Next
        
        Set rs = HFApp.SqlExec("select fulldescription from tblpoindex where divisionid=" & DbQuote(Num, HFApp.DivisionID) & " and poindex=" & DbQuote(Str, POIndex))
        If Not rs.EOF Then
            .Cell(flexcpText, i, .ColIndex("POIndex")) = "" & rs(0)
        End If
        
    End With
    
    
    Call GroupGrid
Exit Sub
eh: Call errHandler(SRCFILE & "AddItem")
End Sub



Private Function ValidateData() As Boolean
    Dim s As String
    Dim r As Long
    Dim at As Long
    
    If Trim(txtAssembly.Text) = "" Then s = s & " " & vbBullet & " Assembly is required" & vbCrLf
    If Trim(txtDescription.Text) = "" Then s = s & " " & vbBullet & " Description is required" & vbCrLf
    
    at = mAssemblyType
    If at = atCustomOption Then
        If Trim(cboModel.Text) = "" Then
            at = atGlobal
        Else
            at = atoption
        End If
    End If
    
    
    Select Case at
        Case atModel
            If Trim(cboModel.Text) = "" Then s = s & " " & vbBullet & " Model is required" & vbCrLf
        Case atoption
            If Trim(cboModel.Text) = "" Then s = s & " " & vbBullet & " Model is required" & vbCrLf
            If Trim(txtOption.Text) = "" Then s = s & " " & vbBullet & " Option is required" & vbCrLf
            If cboCategory.ListIndex < 0 And HFApp.Options(ReqCategory) Then s = s & " " & vbBullet & " Category is required" & vbCrLf
        Case atGlobal, atDesignCenter
            If Trim(txtOption.Text) = "" Then s = s & " " & vbBullet & " Option is required" & vbCrLf
            If cboCategory.ListIndex < 0 And HFApp.Options(ReqCategory) Then s = s & " " & vbBullet & " Category is required" & vbCrLf
    End Select
    
    With gItems
        For r = 1 To .Rows - 1
            If (.TextMatrix(r, .ColIndex("Phase")) = "" Or .TextMatrix(r, .ColIndex("Item")) = "") And .TextMatrix(r, .ColIndex("ItemChart")) = "" And Not .RowHidden(r) And Not .IsSubtotal(r) Then
                s = s & " " & vbBullet & " Row " & r & " """ & .TextMatrix(r, .ColIndex("Description")) & """ is a one time item. It must be replaced with a database item." & vbCrLf
            End If
        Next
    End With
    
    
    If s = "" Then
        ValidateData = True
        mAssemblyType = at
    Else
        MsgBox "Unable to save assembly" & vbCrLf & vbCrLf & s, vbExclamation, App.ProductName
    End If
    
    
    
End Function


Public Sub mnuAssemblyItemsSub_Click(Index As Integer)
    Dim s      As String
    Dim phaseitem As String
    Dim c      As Long
    Dim r      As Long
    Dim NewRow As Long
    Dim ObjectID As String
    
    With gItems
    Select Case Index
        
        Case mcITEM_EDITITEMCHART
            Call FItemChart.Edit(.TextMatrix(.Row, .ColIndex("ItemChart")))
    
        Case mcITEM_VIEWFILES
            ObjectID = "ASM~" & mCommunity & "~" & mModel & "~" & mOptionID & "~" & mAssembly & "~" & .TextMatrix(.Row, .ColIndex("Phase")) & "~" & .TextMatrix(.Row, .ColIndex("Item"))
            s = "ITM~" & .TextMatrix(.Row, .ColIndex("Phase")) & "~" & .TextMatrix(.Row, .ColIndex("Item")) & "|Item Database"
            
            mTimerTask = "EditAttachments|" & ObjectID & "|Assembly Items|" & s & "|Item Database"
            Me.Timer1.Enabled = True
            Me.Timer1.Interval = 10
            
            
        
        Case mcITEM_COPY
            If .Row <= 0 Then Exit Sub
            For r = Max(.Row, .RowSel) To Min(.Row, .RowSel) Step -1
                mDirty = True
                NewRow = Max(.Row, .RowSel) + 1
                .AddItem "", NewRow
                For c = 0 To .Cols - 1
                    .TextMatrix(NewRow, c) = .TextMatrix(r, c)
                Next
                .RowData(NewRow) = "new"
                .TextMatrix(NewRow, .ColIndex("Sequence")) = ""
                
            Next
                   
        Case mcITEM_SUBSTITUE
            If .Row <= 0 Then Exit Sub
            s = ""
            s = s & "SELECT isnull(phase,'')+char(1)+isnull(item,'') phaseitem " & vbCrLf
            s = s & "      ,i.Phase" & vbCrLf
            s = s & "      ,i.Item" & vbCrLf
            s = s & "      ,i.Description" & vbCrLf
            s = s & "      ,i.TakeoffUOM" & vbCrLf
            s = s & "      ,i.ConversionFactor" & vbCrLf
            s = s & "      ,i.OrderUOM" & vbCrLf
            s = s & "      ,p.POIndex,i.PartNumber" & vbCrLf
            s = s & "      ,p.FullDescription POIndexDescription" & vbCrLf
            s = s & "      ,i.Notes" & vbCrLf
            s = s & "      ,i.JCCostCode CostCode" & vbCrLf
            s = s & "      ,cod.description CostCodeDesc" & vbCrLf
            s = s & "      ,i.JCCategory Category" & vbCrLf
            s = s & "      ,cat.Description CategoryDesc" & vbCrLf
            s = s & "  FROM tblPhaseItem i" & vbCrLf
            s = s & "LEFT OUTER JOIN StandardCostCodes cod ON(i.DivisionID = cod.DivisionID and i.jccostcode=cod.costcode)" & vbCrLf
            s = s & "LEFT OUTER JOIN StandardCategories cat ON(i.DivisionID = cat.DivisionID and i.jccategory=cat.category)" & vbCrLf
            s = s & "LEFT OUTER JOIN tblPOIndex p ON(i.DivisionID = p.DivisionID and i.poindex=p.poindex)" & vbCrLf
            s = s & "WHERE i.DivisionID = " & HFApp.DivisionID
            .HighLight = flexHighlightAlways
            
            
            phaseitem = .TextMatrix(.Row, .ColIndex("Phase")) & Chr(1) & .TextMatrix(.Row, .ColIndex("Item"))
            If FPickList.Choose(HFApp.Databases(dbHomefront), "Item", s, phaseitem, , , , "phaseitem,TakeoffUOM,EffectiveConversionFactor,OrderUOM,POIndex,Notes,CategoryDesc") Then
                mDirty = True
                For r = Max(.Row, .RowSel) To Min(.Row, .RowSel) Step -1
                    .TextMatrix(r, .ColIndex("Phase")) = FPickList.SelectedItem("Phase")
                    .TextMatrix(r, .ColIndex("Item")) = FPickList.SelectedItem("Item")
                    .TextMatrix(r, .ColIndex("Description")) = FPickList.SelectedItem("Description")
                    .TextMatrix(r, .ColIndex("TakeoffUOM")) = FPickList.SelectedItem("TakeoffUOM")
                    .TextMatrix(r, .ColIndex("ItemConversionFactor")) = Val("" & FPickList.SelectedItem("ConversionFactor"))
                    .TextMatrix(r, .ColIndex("OrderUOM")) = FPickList.SelectedItem("OrderUOM")
                    .Cell(flexcpData, r, .ColIndex("POIndex")) = FPickList.SelectedItem("POIndex")
                    .Cell(flexcpText, r, .ColIndex("POIndex")) = FPickList.SelectedItem("POIndexDescription")
                    .TextMatrix(r, .ColIndex("Notes")) = FPickList.SelectedItem("Notes")
                    .TextMatrix(r, .ColIndex("JCCostCode")) = FPickList.SelectedItem("CostCode")
                    .TextMatrix(r, .ColIndex("JCCostCodeDesc")) = FPickList.SelectedItem("CostCodeDesc")
                    .TextMatrix(r, .ColIndex("JCCategory")) = FPickList.SelectedItem("Category")
                    .TextMatrix(r, .ColIndex("JCCategoryDesc")) = FPickList.SelectedItem("CategoryDesc")
                Next
                Call gItems_AfterEdit(.Row, .ColIndex("TakeoffQty"))
            End If
            .HighLight = flexHighlightWithFocus
                
                    
        Case mcITEM_REMOVE
            Call gItems_KeyDown(vbKeyDelete, vbCtrlMask)
            
             
            
    End Select
    End With
End Sub

Public Sub mnuAssemblyNewSub_Click(Index As Integer)
    Dim i As Long
    If Not SaveData(True) Then Exit Sub
    
    Select Case Index
        Case mcNEW_COPYOF
            CtrlEnabled(cboCommunity) = True
            CtrlEnabled(txtAssembly) = True
            CtrlEnabled(cboModel) = True
            CtrlEnabled(txtOption) = True
            Select Case mAssemblyType
                Case atModel:         CtrlEnabled(txtOption) = False
                Case atGlobal:        CtrlEnabled(cboModel) = False
                Case atDesignCenter:  CtrlEnabled(cboModel) = False
                Case atoption:
            End Select
            
            'save these so we can duplicate the vendor pricing when we save the new one
            mCopiedCommunity = mCommunity
            mCopiedAssembly = mAssembly
            mCopiedAssemblyID = mAssemblyID
            mCopiedModel = mModel
            mCopiedOption = mOptionID
            mCopiedKey = mCopiedCommunity & Chr(1) & mCopiedAssembly & Chr(1) & mCopiedModel & Chr(1) & mCopiedOption
            
            mCommunity = ""
            mAssembly = ""
            mModel = ""
            mOptionID = ""
            chkIncludedOption.value = 0
            Me.chkDCSalesOnly.value = 0
            Me.chkSelectByRoom.value = 0
            Me.chkDisplayTotalOnly.value = 0
            mDirty = True
            
            With gItems
            For i = 1 To .Rows - 1
                If Not .IsSubtotal(i) And .RowData(i) <> "delete" Then .RowData(i) = "new"
            Next
            End With
            
        Case Else
            Select Case Index
                Case mcNEW_MODEL:    mAssemblyType = atModel
                Case mcNEW_OPTION:   mAssemblyType = atoption
                Case mcNEW_GLOBAL:   mAssemblyType = atGlobal
                Case mcNEW_DCOPTION: mAssemblyType = atDesignCenter
            End Select
            mCopiedCommunity = ""
            mCopiedAssembly = ""
            mCopiedAssemblyID = 0
            mCopiedModel = ""
            mCommunity = ""
            mAssembly = ""
            mModel = ""
            mOptionID = ""
            Call LoadData
            ReadOnly = False
    End Select
    
End Sub


Private Sub DeleteAssembly()
On Error GoTo eh
    Dim s As String
    
    If mAssembly = "" Then Exit Sub
    If MsgBox("Are you sure you want to delete this assembly?", vbCritical + vbDefaultButton2 + vbOKCancel, App.ProductName) <> vbOK Then Exit Sub
    If MsgBox("STOP!" & vbCrLf & vbCrLf & "You are about to delete this entire assembly." & vbCrLf & "If you click YES the assembly and all it's items will be gone forever." & vbCrLf & vbCrLf & "Are you sure this is what you want to do?", vbCritical + vbDefaultButton2 + vbYesNo, App.ProductName) <> vbYes Then Exit Sub
    
            
    s = ""
    s = s & "exec Purch_DeleteAssembly "
    s = s & "       " & DbQuote(Str, GetComboBoxListKey(cboCommunity)) & vbCrLf
    s = s & "      ," & DbQuote(Str, txtAssembly.Text) & vbCrLf
    s = s & "      ," & DbQuote(Str, cboModel.Text) & vbCrLf
    s = s & "      ," & DbQuote(Str, txtOption.Text) & vbCrLf
    s = s & "      ," & DbQuote(Str, MachineName) & vbCrLf
    s = s & "      ," & DbQuote(Str, HFApp.LoginID)
    s = s & "      ," & HFApp.DivisionID
    Call HFApp.SqlExec(s)
    mDirty = False
    mCopiedCommunity = ""
    mCopiedAssembly = ""
    mCopiedAssemblyID = 0
    mCopiedModel = ""
    mAssembly = ""
    mModel = ""
    mOptionID = ""
    mAssemblyDesc = ""
    Call LoadData
    CtrlEnabled(cboCommunity) = False
    CtrlEnabled(cboModel) = False
    CtrlEnabled(txtOption) = False
    CtrlEnabled(txtAssembly) = False
    ReadOnly = True

Exit Sub
eh: Select Case True
        Case Parse(Err.Description, Parse(Err.Description, , "]"), "]") = "Unable to delete model assembly. Model specific options assemblies exist."
            MsgBox "Unable to delete model assembly. Model specific options assemblies exist.", vbExclamation, App.ProductName
        Case Else: Call errHandler(SRCFILE & "DeleteAssembly")
    End Select
End Sub

Private Sub cboModel_Click()
    mDirty = True
End Sub

Private Sub cboModel_Change()
    mDirty = True
    If mAssemblyType = atModel Then txtAssembly.Text = cboModel.Text
End Sub










Private Sub ExportPriceList()
On Error GoTo eh
    Dim FilterIndex As Long
    Dim rs As Recordset
    Dim s As String
    Dim Vendor As String
    Dim FileName As String
    
    If Not SaveData(False) Then Exit Sub
    If Not VBGetSaveFileName(FileName, , , "Quotes Only (*.xls)|*.xls|All Items (*.xls)|*.xls", FilterIndex, , , "xls", Me.hwnd) Then Exit Sub
 
'filterindex=1  quotes only
'filterindex=2  all items
    
    s = ""
    s = s & "select " & vbCrLf
    s = s & " v.vendor_id Vendor" & vbCrLf
    s = s & ",v.vendor_name VendorDesc" & vbCrLf
    s = s & ",m.Community" & vbCrLf
    s = s & ",'' CommunityPhase" & vbCrLf
    s = s & ",l.Description CommunityDesc" & vbCrLf
    s = s & ",m.Model" & vbCrLf
    s = s & ",m.Assembly" & vbCrLf
    s = s & ",m.Description AssemblyDesc" & vbCrLf
    s = s & ",d.Phase" & vbCrLf
    s = s & ",d.Item" & vbCrLf
    s = s & ",i.POIndex" & vbCrLf
    s = s & ",i.Description ItemDesc" & vbCrLf
    s = s & ",'' PartNumber" & vbCrLf
    s = s & ",i.OrderUOM" & vbCrLf
    s = s & ",0 Price" & vbCrLf
    s = s & "from tbldbassemblymaster m" & vbCrLf
    s = s & "left outer join tbldbassemblydetails d on (m.Divisionid = d.Divisionid and m.assembly=d.assembly and m.community=d.community and m.model=d.model and m.optionid=d.optionid)" & vbCrLf
    s = s & "left outer join tblphaseitem i on (d.DivisionID = i.DivisionID and d.phase=i.phase and d.item=i.item)" & vbCrLf
    If Replace(GetComboBoxListKey(cboPriceCommunity), Chr(2), "") <> "" Then
        s = s & "left outer join tblvendors v on(v.Divisionid = " & HFApp.DivisionID & " and vendor_id=dbo.purch_getcommunityvendor(" & DbQuote(Str, Replace(GetComboBoxListKey(cboPriceCommunity), Chr(2), "")) & ",i.poindex," & HFApp.DivisionID & "))" & vbCrLf
    Else
        s = s & "left outer join tblvendors v on(v.Divisionid = " & HFApp.DivisionID & " and vendor_id=dbo.purch_getcommunityvendor(" & DbQuote(Str, Replace(GetComboBoxListKey(cboCommunity), Chr(2), "")) & ",i.poindex," & HFApp.DivisionID & "))" & vbCrLf
    End If
    s = s & "left outer join tbllocality l on(m.community=l.area)" & vbCrLf
    s = s & "where m.Community=" & DbQuote(Str, GetComboBoxListKey(cboCommunity)) & vbCrLf
    s = s & "  and m.Assembly=" & DbQuote(Str, txtAssembly.Text) & vbCrLf
    s = s & "  and m.model=" & DbQuote(Str, cboModel.Text) & vbCrLf
    s = s & "  and m.OptionID=" & DbQuote(Str, txtOption.Text) & vbCrLf
    s = s & "  and m.DivisionID = " & HFApp.DivisionID & vbCrLf
    If FilterIndex = 1 Then
        s = s & "  and i.isquote=1" & vbCrLf
    End If
    s = s & "order by v.vendor_id,d.phase,d.item" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    
    With gExportData
        .Rows = 0
        Vendor = Chr(1)
        While Not rs.EOF
            
            If Vendor <> rs("vendor") Then
            
                Vendor = rs("Vendor")
                
                'write vendor heading
                .AddItem ""
                .AddItem ""
                .AddItem "Vendor" & vbTab & Vendor
                .AddItem "Company Name" & vbTab & rs("VendorDesc") & vbTab & rs("POIndex")
                .Cell(flexcpFontBold, .Rows - 1, 0, .Rows - 2, 0) = True
                
                'write column headings
                .AddItem "Community" & vbTab & "Phase" & vbTab & "Community Name" & vbTab & "Model" & vbTab & "Specification" & vbTab & "Assembly Description" & vbTab & "Group" & vbTab & "Item" & vbTab & "Item Description" & vbTab & "Part Number" & vbTab & "UOM" & vbTab & "Price"
                .Cell(flexcpFontBold, .Rows - 1, 0, .Rows - 1, 11) = True
                
            End If
            
            .AddItem "" & rs("Community") & vbTab & _
                          rs("CommunityPhase") & vbTab & _
                          rs("CommunityDesc") & vbTab & _
                          rs("Model") & vbTab & _
                          rs("Assembly") & vbTab & _
                          rs("AssemblyDesc") & vbTab & _
                          rs("Phase") & vbTab & _
                          rs("Item") & vbTab & _
                          rs("ItemDesc") & vbTab & _
                          rs("PartNumber") & vbTab & _
                          rs("OrderUOM") & vbTab & _
                          rs("Price")
                          
            rs.MoveNext
            
        Wend
    
        'save file
        If .Rows > 0 Then
            Call .AutoSize(0, .Cols - 1)
            .ColWidth(1) = 615
            .Cell(flexcpBackColor, 0, 9, .Rows - 1, 9) = RGB(255, 255, 193)
            .Cell(flexcpBackColor, 0, 11, .Rows - 1, 11) = RGB(255, 255, 193)
            Call .SaveGrid(FileName, flexFileExcel)
            .Rows = 0
        End If
    
        
        
    End With
    Exit Sub
eh: Select Case Err.Number
        Case 70, 75
            Call MsgBox("Cannot create the " & s & " file." & vbCrLf & "You may not have access to the folder or the file may be in use.", vbExclamation, App.ProductName)
        Case Else
            errHandler (SRCFILE & "ExportPriceList")
    End Select
End Sub
 

Private Function PriceLevelIndex(Name As String) As Long
    Select Case Name
        Case "Phase Specific":                                  PriceLevelIndex = 1
        Case FMain.CD_Community & " Specific":                  PriceLevelIndex = 2
        Case "Global (any " & FMain.CD_Community & ")":         PriceLevelIndex = 3
        Case "Corporate (any Division)":                        PriceLevelIndex = 4
        Case "Item DB":                                         PriceLevelIndex = 5
    End Select
End Function


Private Sub UpdateItemPrice(i As Long)
On Error GoTo eh
    Dim j As Long
    Dim s As String
    Dim oldLevel As Long
    Dim newLevel As Long
    Dim oldType  As String
    Dim newType  As String

    If PricingCommunity = "" Then Exit Sub
    
    With gItems
    
        oldType = .TextMatrix(i, .ColIndex("OldItemType"))
        newType = .TextMatrix(i, .ColIndex("ItemType"))
        oldLevel = PriceLevelIndex(.TextMatrix(i, .ColIndex("OldPriceLevel")))
        newLevel = PriceLevelIndex(.TextMatrix(i, .ColIndex("PriceLevel")))
        ' pricelevels 1 = phase specific
        '             2 = community specific
        '             3 = global
        '             4 = corporate
        '             5 = item db
    
        '----------------------------
        'if new price level is less specific then we have to remove the old one
        '----------------------------
        If (newLevel > oldLevel Or oldType = "Quote" And newType <> "Quote") And oldLevel <> 0 And oldLevel <> 4 Then
            s = ""
            s = s & "delete tblvendorcost" & vbCrLf
            s = s & "from tblphaseitem i" & vbCrLf
            s = s & "join tblphaseitem x on i.divisionid=x.divisionid and ((i.pricelink<>0 and i.pricelink=x.pricelink) or (i.phase=x.phase and i.item=x.item))" & vbCrLf
            s = s & "join tblvendorcost c on c.divisionid=x.divisionid and c.phase=x.phase and c.item=x.item" & vbCrLf
            s = s & "where i.divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
            s = s & "and i.phase=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Phase"))) & vbCrLf
            s = s & "and i.item=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Item"))) & vbCrLf
            Select Case oldLevel
                Case 1 'phase
                    s = s & "  and isnull(c.community,'')=" & DbQuote(Str, PricingCommunity) & vbCrLf
                    s = s & "  and isnull(c.communityphase,'')=" & DbQuote(Str, PricingCommunityPhase) & vbCrLf
                Case 2 'community
                    s = s & "  and isnull(c.community,'')=" & DbQuote(Str, PricingCommunity) & vbCrLf
                    s = s & "  and isnull(c.communityphase,'')=''" & vbCrLf
                Case Else 'global
                    s = s & "  and isnull(c.community,'')=''" & vbCrLf
                    s = s & "  and isnull(c.communityphase,'')=''" & vbCrLf
            End Select
            If oldType = "Quote" Then
                s = s & "  and isnull(c.assembly,'')=" & DbQuote(Str, txtAssembly.Text) & vbCrLf
                s = s & "  and isnull(c.model,'')=" & DbQuote(Str, cboModel.Text) & vbCrLf
            Else
                s = s & "  and isnull(c.assembly,'')=''" & vbCrLf
                s = s & "  and isnull(c.model,'')=''" & vbCrLf
            End If
            Call HFApp.SqlExec(s, dbHomefront)
        End If
            
        
        '----------------------------
        ' now do insert/update
        '----------------------------
        If newLevel = 5 Then
        ElseIf .TextMatrix(i, .ColIndex("Vendor")) <> "" Then
            '-----------------------
            ' insert any that dont exist
            '-----------------------
            s = ""
            s = s & "insert into tblvendorcost(phase,item,DivisionID,vendor,community,communityphase,assembly,model,Current_Cost,Forecast1,Forecast2,Forecast3,Forecast4,Forecast5,Forecast6,Forecast7,Forecast8,Forecast9,Forecast10,Forecast11,Forecast12)" & vbCrLf
            s = s & "select x.phase,x.item" & vbCrLf
            s = s & " ," & IIf(newLevel = 4, 0, HFApp.DivisionID) & vbCrLf
            s = s & " ," & DbQuote(Str, .TextMatrix(i, .ColIndex("Vendor"))) & vbCrLf
            Select Case newLevel
                Case 1 'phase
                    s = s & " ," & DbQuote(Str, PricingCommunity) & vbCrLf
                    s = s & " ," & DbQuote(Str, PricingCommunityPhase) & vbCrLf
                Case 2 'community
                    s = s & " ," & DbQuote(Str, PricingCommunity) & vbCrLf
                    s = s & " ,''" & vbCrLf
                Case Else 'global, corporate
                    s = s & " ,'',''" & vbCrLf
            End Select
            If newType = "Quote" Then
                s = s & " ," & DbQuote(Str, txtAssembly.Text) & vbCrLf
                s = s & " ," & DbQuote(Str, cboModel.Text) & vbCrLf
            Else
                s = s & " ,'',''" & vbCrLf
            End If
            For j = 1 To 13
                s = s & " ," & DbQuote(Num, .TextMatrix(i, .ColIndex("Price")))
            Next
            s = s & vbCrLf
            s = s & "from tblphaseitem i" & vbCrLf
            s = s & "join tblphaseitem x on i.divisionid=x.divisionid and ((i.pricelink<>0 and i.pricelink=x.pricelink) or (i.phase=x.phase and i.item=x.item))" & vbCrLf
            s = s & "left outer join tblVendorCost c on c.divisionid=" & IIf(newLevel = 4, 0, HFApp.DivisionID) & " and x.phase=c.phase and x.item=c.item " & vbCrLf
            s = s & " and c.vendor=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Vendor"))) & vbCrLf
            Select Case newLevel
                Case 1 'phase
                    s = s & " and c.community=" & DbQuote(Str, PricingCommunity) & vbCrLf
                    s = s & " and c.communityphase=" & DbQuote(Str, PricingCommunityPhase) & vbCrLf
                Case 2 'community
                    s = s & " and c.community=" & DbQuote(Str, PricingCommunity) & vbCrLf
                    s = s & " and c.communityphase=''" & vbCrLf
                Case Else 'global,corporate
                    s = s & " and c.community=''" & vbCrLf
                    s = s & " and c.communityphase=''" & vbCrLf
            End Select
            If newType = "Quote" Then
                s = s & " and c.assembly=" & DbQuote(Str, txtAssembly.Text) & vbCrLf
                s = s & " and c.model=" & DbQuote(Str, cboModel.Text) & vbCrLf
            Else
                s = s & " and c.assembly=''" & vbCrLf
                s = s & " and c.model=''" & vbCrLf
            End If
            s = s & "where c.vendor is null" & vbCrLf
            s = s & "and i.divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
            s = s & "and i.phase=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Phase"))) & vbCrLf
            s = s & "and i.item=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Item"))) & vbCrLf
            Call HFApp.SqlExec(s, dbHomefront)



            '-----------------------
            ' update
            '-----------------------
            s = ""
            s = s & "update c set current_cost=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Price"))) & vbCrLf
            s = s & "from tblphaseitem i" & vbCrLf
            s = s & "join tblphaseitem x on i.divisionid=x.divisionid and ((i.pricelink<>0 and i.pricelink=x.pricelink) or (i.phase=x.phase and i.item=x.item))" & vbCrLf
            s = s & "join tblVendorCost c on c.divisionid=" & IIf(newLevel = 4, 0, HFApp.DivisionID) & " and x.phase=c.phase and x.item=c.item " & vbCrLf
            s = s & " and c.vendor=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Vendor"))) & vbCrLf
            Select Case newLevel
                Case 1 'phase
                    s = s & " and c.community=" & DbQuote(Str, PricingCommunity) & vbCrLf
                    s = s & " and c.communityphase=" & DbQuote(Str, PricingCommunityPhase) & vbCrLf
                Case 2 'community
                    s = s & " and c.community=" & DbQuote(Str, PricingCommunity) & vbCrLf
                    s = s & " and c.communityphase=''" & vbCrLf
                Case Else 'global, corporate
                    s = s & " and c.community=''" & vbCrLf
                    s = s & " and c.communityphase=''" & vbCrLf
            End Select
            If newType = "Quote" Then
                s = s & " and c.assembly=" & DbQuote(Str, txtAssembly.Text) & vbCrLf
                s = s & " and c.model=" & DbQuote(Str, cboModel.Text) & vbCrLf
            Else
                s = s & " and c.assembly=''" & vbCrLf
                s = s & " and c.model=''" & vbCrLf
            End If
            s = s & "where i.divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
            s = s & "and i.phase=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Phase"))) & vbCrLf
            s = s & "and i.item=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Item"))) & vbCrLf
            Call HFApp.SqlExec(s, dbHomefront)
            
            
        End If
    
        'reset these
        .TextMatrix(i, .ColIndex("OldPriceLevel")) = .TextMatrix(i, .ColIndex("PriceLevel"))
        .TextMatrix(i, .ColIndex("OldItemType")) = .TextMatrix(i, .ColIndex("ItemType"))
        .RowData(i) = ""
    End With
    
Exit Sub
eh: If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        Resume Next
    Else
        Call errHandler(SRCFILE & "UpdateItemPrice", s)
    End If
End Sub

Private Sub ShowTip(Row As Long, Col As Long, X As Long, Y As Long, Optional tip As String)
'
'   THIS ALMOST WORKS. NEED TO RETRIEVE TAXGROUP AND RATE WITH THE PRICING COMMUNITY
'
    
    
    With gItems
        If Row < 0 Or Col < 0 Or (.RowSel = .Row And tip <> "") Then
            picWarningMessages.Visible = False
        Else
        
            lblWarningMessages = IIf(tip <> "", tip, gItems.Cell(flexcpData, Row, Col))
            If lblWarningMessages.Text = "Dirty" Then lblWarningMessages.Text = ""
            
            
            imgTipIcon.Picture = IIf(tip <> "", imgInfo.Picture, imgWarning.Picture)


            picWarningMessages.Visible = lblWarningMessages <> ""

            Set Me.Font = lblWarningMessages.Font

            lblWarningMessages.Alignment = IIf(tip <> "", 1, 0)
            lblWarningMessages.Width = Me.TextWidth(lblWarningMessages) + lblWarningMessages.Left
            lblWarningMessages.Height = Me.TextHeight(lblWarningMessages) + lblWarningMessages.Top

            picWarningMessages.Width = lblWarningMessages.Width + lblWarningMessages.Left + lblWarningMessages.Top
            picWarningMessages.Height = lblWarningMessages.Height + 2 * lblWarningMessages.Top
            If .Height - Y - picWarningMessages.Height - 365 < 0 Then Y = .Height - picWarningMessages.Height - 365
            If .Width - X - picWarningMessages.Width - 365 < 0 Then X = .Width - picWarningMessages.Width - 365
            picWarningMessages.Top = Y
            picWarningMessages.Left = X

        End If
    End With
End Sub



Private Sub LoadWBSDescriptions()
    Dim i As Long
    Dim c As Long
    Dim s As String
    Dim rs As Recordset
    
    With gItems
    
        'add columns
        c = .Cols - 1
        .Cols = .Cols + 40
        For i = 1 To 40
            .ColKey(c + i) = "WBS" & format(i, "00")
            .ColHidden(c + i) = True
        Next
        
        'read column names
        s = "select Item,custom_description Description from customDescriptions where isnull(custom_description,'')<>'' and item like 'WBS__' order by item"
        Set rs = HFApp.SqlExec(s, dbHomefront)
        While Not rs.EOF
            c = .ColIndex("" & rs(0))
            If c <> -1 Then
                .TextMatrix(0, c) = "" & rs(1)
                .ColHidden(c) = False
            End If
            rs.MoveNext
        Wend
    
    
    End With

End Sub


Public Sub GroupGrid()
On Error GoTo eh
    
    Dim i As Long
    Dim Row As Long
    Dim Col As Long
    
    
    
    With gItems
        Row = .Row
        Col = .Col
        
        GroupedColumns = 0
        For i = 0 To .Cols - 1
            If .ColData(i) = "GROUPED" Then
                .ColPosition(i) = GroupedColumns
                GroupedColumns = GroupedColumns + 1
            End If
        Next
        
        If GroupedColumns = 0 Then
            Call .SubTotal(flexSTClear)
        Else
            .SubTotal flexSTClear
            mFunnyFlag = True
            .Col = 0
            .ColSel = GroupedColumns - 1
            .Sort = flexSortGenericAscending
            mFunnyFlag = False
            
            
            .OutlineCol = 0
            .SubtotalPosition = flexSTAbove
            For i = 0 To GroupedColumns - 1
                .SubTotal flexSTSum, i, .ColIndex("ExtendedAmount"), "$(#,###.00)", , vbHighlight, True, "%s", 0
            Next
        End If
    
        On Error Resume Next
        .Row = Row
        .Col = Col
    End With
Exit Sub
eh: Call errHandler(SRCFILE & "GroupGrid")
End Sub

Private Sub ClearGroups()
    Dim i As Long
    With gItems
    For i = 0 To .Cols - 1
        .ColData(i) = ""
    Next
    End With
End Sub



Private Sub LoadCustomDescriptions()
    
    lblSeries.Caption = FMain.CD_Series
    lblCommunity.Caption = FMain.CD_Community
    
End Sub

