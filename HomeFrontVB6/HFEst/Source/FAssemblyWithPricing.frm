VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{6C83CF2C-BE8D-4EE2-9B07-7FF5E27AB2FD}#1.0#0"; "zybCombo.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form FAssembly 
   Caption         =   "Model & Option Library"
   ClientHeight    =   6090
   ClientLeft      =   150
   ClientTop       =   1830
   ClientWidth     =   15015
   Icon            =   "FAssemblyWithPricing.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6090
   ScaleWidth      =   15015
   Begin VB.Timer Timer1 
      Left            =   9360
      Top             =   60
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   1005
      ButtonWidth     =   1376
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
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
            Caption         =   "Delete"
            Key             =   "Delete"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "One Time"
            Key             =   "TakeoffOneTime"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Item"
            Key             =   "TakeoffItem"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Assembly"
            Key             =   "TakeoffAssembly"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Custom"
            Key             =   "TakeoffCustom"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Change"
            Key             =   "MassChange"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pricelist"
            Key             =   "PricelistExport"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Export"
            Key             =   "EEExport"
            Object.ToolTipText     =   "Push this assembly into Timberline Estimating"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame HeaderFrame 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   -30
      TabIndex        =   20
      Top             =   810
      Width           =   15315
      Begin VB.CheckBox chkActive 
         Caption         =   "Active"
         Height          =   195
         Left            =   3660
         TabIndex        =   4
         Top             =   540
         Width           =   1935
      End
      Begin VB.CheckBox chkUseNormalSalesQtyFactors 
         Caption         =   "Use normal sales quantity factors."
         Height          =   195
         Left            =   6300
         TabIndex        =   16
         Top             =   1470
         Width           =   3075
      End
      Begin VB.Frame frmSeriesFields 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   255
         Left            =   6000
         TabIndex        =   32
         Top             =   120
         Visible         =   0   'False
         Width           =   3855
         Begin zybCombo.zybCombobox cboSeries 
            Height          =   240
            Left            =   930
            TabIndex        =   7
            Top             =   0
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   423
            Style           =   2
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ExtendedUI      =   0   'False
            DropDownWidth   =   0
            AutoCompleteListItemsOnly=   -1  'True
            AutoCompleteItemsAreSorted=   -1  'True
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
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   405
            TabIndex        =   33
            Top             =   30
            Width           =   435
         End
      End
      Begin VB.Frame frmOptionFields 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   750
         Left            =   6000
         TabIndex        =   28
         Top             =   360
         Visible         =   0   'False
         Width           =   3855
         Begin VB.TextBox txtJCExtra 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   230
            Left            =   930
            MaxLength       =   10
            TabIndex        =   13
            Top             =   255
            Width           =   1155
         End
         Begin zybCombo.zybCombobox cboCategory 
            Height          =   240
            Left            =   930
            TabIndex        =   12
            Top             =   0
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   423
            Style           =   2
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ExtendedUI      =   0   'False
            DropDownWidth   =   0
            AutoCompleteListItemsOnly=   -1  'True
            AutoCompleteItemsAreSorted=   -1  'True
         End
         Begin zybCombo.zybCombobox cboUOM 
            Height          =   240
            Left            =   930
            TabIndex        =   14
            Top             =   495
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   423
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ExtendedUI      =   0   'False
            DropDownWidth   =   0
            AutoCompleteListItemsOnly=   -1  'True
            AutoCompleteItemsAreSorted=   -1  'True
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "/"
            Height          =   195
            Index           =   2
            Left            =   480
            TabIndex        =   37
            Top             =   0
            Width           =   75
         End
         Begin VB.Label lblGroup 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Grp"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   210
            TabIndex        =   36
            Top             =   0
            Width           =   255
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "UOM"
            Height          =   195
            Index           =   11
            Left            =   495
            TabIndex        =   31
            Top             =   510
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "JC Extra"
            Height          =   195
            Index           =   9
            Left            =   270
            TabIndex        =   30
            Top             =   240
            Width           =   585
         End
         Begin VB.Label lblCategory 
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
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   555
            TabIndex        =   29
            Top             =   0
            Width           =   240
         End
      End
      Begin VB.CheckBox chkTakeoffRequired 
         Caption         =   "Takeoff is required on every sale."
         Height          =   195
         Left            =   6300
         TabIndex        =   15
         Top             =   1260
         Width           =   3075
      End
      Begin VB.Frame frmModelFields 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   750
         Left            =   6000
         TabIndex        =   25
         Top             =   360
         Visible         =   0   'False
         Width           =   3855
         Begin VB.TextBox txtFloorArea 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   230
            Left            =   930
            TabIndex        =   11
            Top             =   510
            Width           =   930
         End
         Begin zybCombo.zybCombobox cboStyle 
            Height          =   240
            Left            =   930
            TabIndex        =   8
            Top             =   15
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   423
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ExtendedUI      =   0   'False
            DropDownWidth   =   0
            AutoCompleteListItemsOnly=   -1  'True
            DoAutoComplete  =   -1  'True
         End
         Begin VB.TextBox txtBathrooms 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   230
            Left            =   1575
            TabIndex        =   10
            Top             =   270
            Width           =   630
         End
         Begin VB.TextBox txtBedrooms 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   230
            Left            =   930
            TabIndex        =   9
            Top             =   270
            Width           =   630
         End
         Begin VB.TextBox txtStyle 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   240
            Left            =   930
            MaxLength       =   30
            TabIndex        =   17
            Top             =   15
            Width           =   2895
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Sqr Foot"
            Height          =   195
            Index           =   5
            Left            =   240
            TabIndex        =   38
            Top             =   510
            Width           =   600
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Beds/Baths"
            Height          =   195
            Index           =   7
            Left            =   0
            TabIndex        =   27
            Top             =   270
            Width           =   840
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Style"
            Height          =   195
            Index           =   6
            Left            =   495
            TabIndex        =   26
            Top             =   30
            Width           =   345
         End
      End
      Begin VB.TextBox txtComments 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   705
         Left            =   1440
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   1005
         Width           =   4125
      End
      Begin VB.TextBox txtAssembly 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000011&
         Height          =   230
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   3
         Top             =   525
         Width           =   2055
      End
      Begin VB.TextBox txtOption 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000011&
         Height          =   240
         Left            =   3540
         MaxLength       =   20
         TabIndex        =   2
         Top             =   270
         Width           =   2025
      End
      Begin VB.TextBox txtDescription 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   230
         Left            =   1440
         MaxLength       =   200
         TabIndex        =   5
         Top             =   765
         Width           =   4125
      End
      Begin zybCombo.zybCombobox cboModel 
         Height          =   240
         Left            =   1440
         TabIndex        =   1
         Top             =   270
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   423
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ExtendedUI      =   0   'False
         DropDownWidth   =   320
         AutoCompleteItemsAreSorted=   -1  'True
         DoAutoComplete  =   -1  'True
      End
      Begin zybCombo.zybCombobox cboCommunity 
         Height          =   240
         Left            =   1440
         TabIndex        =   0
         Top             =   15
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   423
         Style           =   2
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ExtendedUI      =   0   'False
         DropDownWidth   =   0
         AutoCompleteListItemsOnly=   -1  'True
         AutoCompleteItemsAreSorted=   -1  'True
      End
      Begin zybCombo.zybCombobox cboPriceCommunity 
         Height          =   240
         Left            =   10140
         TabIndex        =   39
         Top             =   330
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   423
         Style           =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ExtendedUI      =   0   'False
         DropDownWidth   =   0
         AutoCompleteListItemsOnly=   -1  'True
         AutoCompleteItemsAreSorted=   -1  'True
      End
      Begin VB.Label lblTotalCost 
         AutoSize        =   -1  'True
         Caption         =   "$248,900.00"
         Height          =   195
         Left            =   10950
         TabIndex        =   42
         Top             =   720
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total Cost"
         Height          =   195
         Index           =   12
         Left            =   10140
         TabIndex        =   41
         Top             =   720
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Show vendors and costs for:"
         Height          =   195
         Index           =   8
         Left            =   10140
         TabIndex        =   40
         Top             =   150
         Width           =   2025
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Community"
         Height          =   195
         Index           =   3
         Left            =   600
         TabIndex        =   34
         Top             =   30
         Width           =   765
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Estimator Notes"
         Height          =   195
         Index           =   10
         Left            =   255
         TabIndex        =   24
         Top             =   1005
         Width           =   1110
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Description"
         Height          =   195
         Index           =   4
         Left            =   570
         TabIndex        =   23
         Top             =   765
         Width           =   795
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Model/Option"
         Height          =   195
         Index           =   1
         Left            =   390
         TabIndex        =   22
         Top             =   270
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Assembly"
         Height          =   195
         Index           =   0
         Left            =   705
         TabIndex        =   21
         Top             =   525
         Width           =   660
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid gItems 
      Height          =   2655
      Left            =   30
      TabIndex        =   18
      Top             =   2790
      Width           =   13515
      _cx             =   23839
      _cy             =   4683
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
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   25
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FAssemblyWithPricing.frx":000C
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
   Begin VSFlex8Ctl.VSFlexGrid gExportData 
      Height          =   3285
      Left            =   9000
      TabIndex        =   35
      Top             =   4050
      Visible         =   0   'False
      Width           =   3045
      _cx             =   5371
      _cy             =   5794
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
      FormatString    =   $"FAssemblyWithPricing.frx":0447
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

Private mCommunity     As String
Private mModel         As String
Private mOptionID      As String
Private mAssembly      As String
Private mAssemblyDesc  As String
Private mAssemblyType  As AssemblyTypes
Private mDataSource    As Long '0=estimating, 1 = salesdata

Private mCopiedCommunity     As String
Private mCopiedModel         As String
Private mCopiedAssembly      As String

'new menu constants
Private Const mcNEW_MODEL = 0
Private Const mcNEW_OPTION = 1
Private Const mcNEW_GLOBAL = 2
Private Const mcNEW_DCOPTION = 3
Private Const mcNEW_COPYOF = 5

'items menu constants
Private Const mcITEM_COPY = 0
Private Const mcITEM_SUBSTITUE = 1
Private Const mcITEM_REMOVE = 2
Private Const mcITEM_VIEWFILES = 4
Private Const mcITEM_INSERTFILE = 5
Private Const mcITEM_LINKTOFILE = 6





Private Sub cboCommunity_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = vbKeyDelete Then
        cboCommunity.ListIndex = -1
        mDirty = True
    End If
End Sub

Private Sub cboPriceCommunity_Click()
    Dim s As String
    Dim community As String
    Dim phase As String
    If SaveData(True) Then
        s = GetComboBoxListKey(cboPriceCommunity)
        community = Parse(s, 1, Chr(3))
        phase = Parse(s, 2, Chr(3))
        Select Case True
            Case phase = "":          s = "#2;Community|#3;Global|#4;Item DB"
            Case Else:                s = "#1;Phase|#2;Community|#3;Global|#4;Item DB"
        End Select
        gItems.ColComboList(gItems.ColIndex("PriceLevel")) = s
        Call LoadData
    End If
End Sub

Private Sub cboStyle_Click()
    mDirty = True
End Sub

Private Sub cboUOM_Click()
    mDirty = True
End Sub

Private Sub cboUOM_Validate(Cancel As Boolean)
    mDirty = True
    cboUOM.Text = left(cboUOM.Text, 10)
End Sub


Private Sub chkActive_Click()
    mDirty = True
End Sub

Private Sub chkTakeoffRequired_Click()
    mDirty = True
End Sub


Private Sub chkUseNormalSalesQtyFactors_Click()
    mDirty = True
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
    
    HeaderFrame.BackColor = vbButtonFace

    Call SetToolbarIcons(Toolbar, FMain.LargeIcons)
    Call IniGetGrid(Me, gItems)
    Call IniGetForm(Me)
    gItems.Rows = 1
    ReadOnly = True
    Me.Show
    
    
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
    
    
    Call LoadComboBox(cboUOM, HFApp.Databases(dbHomefront), "SELECT DISTINCT ISNULL(AssemblyUOM,''),'',0 FROM tblDBAssemblyMaster ORDER BY 1")
    Call LoadComboBox(cboModel, HFApp.Databases(dbHomefront), "SELECT distinct isnull(model,'') + ' - ' + isnull(description,''),model,0 FROM DistinctModels order by 1")
    Call LoadComboBox(cboCommunity, HFApp.Databases(dbHomefront), "select '','',0 union all SELECT isnull(area,'') + ' - ' + isnull(description,''),area,0 FROM tblLocality order by 1")
    
    s = ""
    s = s & "select '',char(3),0" & vbCrLf
    s = s & "union" & vbCrLf
    s = s & "select c.area + ' -- ' + c.Description,c.area + char(3),0" & vbCrLf
    s = s & "from tbllocality c " & vbCrLf
    s = s & "union" & vbCrLf
    s = s & "select c.area + '.' + p.communityphase + ' -- ' + c.description + isnull(' ' + p.description,' phase ' + p.communityphase),c.area + char(3) + isnull(p.communityphase,''),0" & vbCrLf
    s = s & "from tbllocality c " & vbCrLf
    s = s & "join communityphase p on(c.area=p.community)" & vbCrLf
    s = s & "order by 1" & vbCrLf
    Call LoadComboBox(cboPriceCommunity, HFApp.Databases(dbHomefront), s)
    
    Call LoadSeries
    Call LoadCategories

    
End Sub
Private Sub LoadSeries()
    
    lblSeries.FontUnderline = HFApp.Options(SalesSystem) <> asHomeFront
    lblSeries.ForeColor = IIf(HFApp.Options(SalesSystem) <> asHomeFront, &HFF0000, vbButtonText)
    
    Call LoadComboBox(cboSeries, HFApp.Databases(dbHomefront), "SELECT isnull(description,series),series,0 FROM tblSeries order by 1")
End Sub

Private Sub LoadCategories()
    Call LoadComboBox(cboCategory, HFApp.Databases(dbHomefront), "SELECT isnull(description,category),category,0 FROM tblcategories order by 1")
End Sub
Public Function ShowForm(SaveAs As Boolean, Optional community As String, Optional Model As String, Optional OptionID As String, Optional Assembly As String, Optional AssemblyDesc As String, Optional AssemblyType As AssemblyTypes) As Boolean
    
    mDirty = False
    ShowForm = True
    
    If community & Model & OptionID & Assembly = "" And Not SaveAs Then
        Call Toolbar_ButtonClick(Toolbar.Buttons("Open"))
    Else
    
        mCopiedCommunity = ""
        mCopiedAssembly = ""
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

    frmSeriesFields.Visible = False
    frmModelFields.Visible = True
    frmOptionFields.Visible = True
    
    Select Case mAssemblyType
        Case atModel
            frmSeriesFields.Visible = True
            frmModelFields.Visible = True
            frmOptionFields.Visible = False
            
        Case atOption
            frmSeriesFields.Visible = True
            frmModelFields.Visible = True
            frmOptionFields.Visible = False
            
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
    gItems.Move 0, HeaderFrame.Top + HeaderFrame.Height, Me.ScaleWidth, Me.ScaleHeight - Toolbar.Height - HeaderFrame.Height - 120
    
    With frmModelFields
        frmOptionFields.Move .left, .Top, .Width, .Height
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not SaveData(True) Then
        Cancel = True
        Exit Sub
    End If
    Call IniPutGrid(Me, gItems)
    Call IniPutForm(Me)
End Sub


Private Sub gItems_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    Dim r As Long
    Dim TakeoffQty As Double
    Dim OrderQty As Double
    Dim Conversion As Double
    Dim RoundDir As Long
    Dim RoundUnit As Double
    Dim WastePercent As Long
    With gItems
        Select Case .ColKey(Col)
            Case "Price"
                For r = Min(Row, .RowSel) To Max(Row, .RowSel)
                    .TextMatrix(r, .ColIndex("Price")) = Val(.TextMatrix(r, .ColIndex("Price")))
                Next
                Call CalcTotal
                
            Case "TakeoffQty"
                For r = Min(Row, .RowSel) To Max(Row, .RowSel)
                    TakeoffQty = Val(.TextMatrix(r, .ColIndex("TakeoffQty")))
                    Conversion = Val(.TextMatrix(r, .ColIndex("ConversionFactor")))
                    If Conversion = 0 Then Conversion = 1
                    RoundDir = Val(.TextMatrix(r, .ColIndex("RoundDir")))
                    RoundUnit = Val(.TextMatrix(r, .ColIndex("RoundTo")))
                    WastePercent = Val(.TextMatrix(r, .ColIndex("WastePercent")))
                    OrderQty = RoundTo(TakeoffQty * (100 + WastePercent) / 100 * Conversion, RoundUnit, RoundDir)
                    .TextMatrix(r, .ColIndex("OrderQty")) = OrderQty
                Next
                Call CalcTotal
                
            Case "OrderQty"
                For r = Min(Row, .RowSel) To Max(Row, .RowSel)
                    OrderQty = Val(.TextMatrix(r, .ColIndex("OrderQty")))
                    Conversion = Val(.TextMatrix(r, .ColIndex("ConversionFactor")))
                    If Conversion = 0 Then Conversion = 1
                    TakeoffQty = OrderQty / Conversion
                    .TextMatrix(r, .ColIndex("TakeoffQty")) = TakeoffQty
                Next
                Call CalcTotal
                
            Case "PriceLevel"
                For r = Min(Row, .RowSel) To Max(Row, .RowSel)
                    .Cell(flexcpForeColor, r, .ColIndex("PriceLevel")) = IIf(4 = Val(.TextMatrix(r, Col)), vbHighlight, vbWindowText)
                    .Cell(flexcpForeColor, r, .ColIndex("Price")) = IIf(4 = Val(.TextMatrix(r, Col)), vbHighlight, vbWindowText)
                    .Cell(flexcpForeColor, r, .ColIndex("Vendor")) = IIf(4 = Val(.TextMatrix(r, Col)), vbHighlight, vbWindowText)
                    .Cell(flexcpForeColor, r, .ColIndex("VendorDesc")) = IIf(4 = Val(.TextMatrix(r, Col)), vbHighlight, vbWindowText)
                Next
                
            Case "ItemType"
                For r = Min(Row, .RowSel) To Max(Row, .RowSel)
                    For i = 1 To .Rows - 1
                        If .TextMatrix(r, .ColIndex("Phase")) = .TextMatrix(i, .ColIndex("Phase")) And .TextMatrix(r, .ColIndex("Item")) = .TextMatrix(i, .ColIndex("Item")) Then
                            .TextMatrix(i, Col) = .TextMatrix(r, Col)
                        End If
                    Next
                Next
        End Select
    End With
End Sub

Private Sub gItems_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gItems
        .ComboList = ""
        .AutoSearch = flexSearchNone
        Select Case .ColKey(Col)
            Case "OrderQty", "TakeoffQty", "ItemType"
            
            Case "Price", "PriceLevel":
                Cancel = cboPriceCommunity.ListIndex = 0
                
            Case "Notes":         .ComboList = "|..."
            Case Else
                Cancel = True
                .AutoSearch = flexSearchFromCursor
        End Select
    End With
End Sub

Private Sub gItems_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim s As String
    With gItems
        Select Case .ColKey(Col)
            Case "Notes"
                s = .Text
                If FComments.Edit(s, gItems, , "Notes", 4000) Then
                    .Text = s
                End If
        End Select
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
                Call gItems.RemoveItem(r)
                mDirty = True
            Next
            .Row = .Row
            End If
            
        Case Shift = vbCtrlMask And KeyCode = vbKeyC
            s = .Clip
            s = Replace(s, Chr(10), "")
            s = Replace(s, Chr(13), vbCrLf)
            Call Clipboard.SetText(s)
                        
        Case Shift = vbCtrlMask And KeyCode = vbKeyV And Not ReadOnly And .Row > 0 And .RowSel > 0
            mDirty = True
            .Clip = Clipboard.GetText
            
    End Select
    End With
Exit Sub
eh: Call errHandler(SRCFILE & "gItems_KeyDown")
End Sub



Private Sub gItems_SelChange()
'On Error Resume Next
'change selection to only single column
    Static bInHere As Boolean
    If bInHere Then Exit Sub
    bInHere = True
    gItems.ColSel = gItems.Col
    bInHere = False
End Sub

Private Sub gItems_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    mDirty = True
End Sub





Private Sub lblCategory_Click()
    Dim i As Long
    Call FDBGrid.ShowForm("select group_code [Group], Category,Description from tblcategories order by 1,2", "Option Categories", False, "")
    i = cboCategory.ListIndex
    Call LoadCategories
On Error Resume Next
    cboCategory.ListIndex = i
End Sub

Private Sub lblGroup_Click()
    Call FDBGrid.ShowForm("select major_group [Group] ,Description from tblmajorgroups order by 1", "Option Groups", False, "")
End Sub

Private Sub lblSeries_Click()
    Dim i As Long
    If HFApp.Options(SalesSystem) = asHomeFront Then Exit Sub
    
    Call FDBGrid.ShowForm("select Series,Description from tblseries", "Series", False, "")
    i = cboSeries.ListIndex
    Call LoadSeries
On Error Resume Next
    cboSeries.ListIndex = i
End Sub


Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo eh
    Dim rs As Recordset
    Dim s As String
    
    Dim sCommunity As String
    Dim sModel     As String
    Dim sOptionID  As String
    Dim sAssembly  As String
    
    Select Case Button.Key
    
        Case "MassChange"
            Call FMassChange.ShowForm
    
        Case "PricelistExport"
            Call ExportPriceList
        
        Case "New"
            Call Toolbar_ButtonDropDown(Button)
            
        Case "Save"
            Call SaveData(False)
            
        Case "Delete"
            Call DeleteAssembly
            
        Case "Open"
            If SaveData(True) Then
                s = ""
                s = s & "Models" & Chr(1) & "select c.area,c.description Community,a.Model,a.Model modelid,a.OptionID,a.Assembly,a.Description,a.assemblytype,0 Source,a.Series,a.Style,a.Elevation from tbldbassemblymaster a left outer join tbllocality c on(a.community=c.area) where a.assemblytype=0" & Chr(0)
                s = s & "Model Specific Options" & Chr(1) & "select ct.Description Category,c.area,c.description Community,a.Model,a.Model modelid,a.OptionID,a.OptionID [Option],a.Assembly,a.Description,a.assemblytype,0 Source from tbldbassemblymaster a left outer join tbllocality c on(a.community=c.area) left outer join tblcategories ct on(a.category=ct.category) where a.assemblytype=2" & Chr(0)
                s = s & "Design Center Options" & Chr(1) & "select ct.Description Category,c.area,c.description Community,a.Model modelid,a.OptionID,a.OptionID [Option],a.Assembly,a.Description,a.assemblytype,0 Source from tbldbassemblymaster a left outer join tbllocality c on(a.community=c.area) left outer join tblcategories ct on(a.category=ct.category) where a.assemblytype=4" & Chr(0)
                s = s & "Global Options" & Chr(1) & "select ct.Description Category,c.area,c.description Community,a.Model modelid,a.OptionID,a.OptionID [Option],a.Assembly,a.Description,a.assemblytype,0 Source from tbldbassemblymaster a left outer join tbllocality c on(a.community=c.area) left outer join tblcategories ct on(a.category=ct.category) where a.assemblytype=3" & Chr(0)
                s = s & "-------------------------" & Chr(1) & "select '' where 1=0" & Chr(0)
                s = s & "Sales Models" & Chr(1) & "select community area,CommunityDesc Community,Model,Model modelid,OptionID,Assembly,Description,assemblytype,1 Source from salesassemblylist where assemblytype=0" & Chr(0)
                s = s & "Sales Model Specific Options" & Chr(1) & "select CategoryDesc Category,community area,CommunityDesc Community,Model,Model modelid,OptionID,OptionID [Option],Assembly,Description,assemblytype,1 Source from salesassemblylist where assemblytype=2" & Chr(0)
                s = s & "Sales Design Center Options" & Chr(1) & "select CategoryDesc Category,community area,CommunityDesc Community,Model modelid,OptionID,OptionID [Option],Assembly,Description,assemblytype,1 Source from salesassemblylist where assemblytype=4" & Chr(0)
                s = s & "Sales Global Options" & Chr(1) & "select CategoryDesc Category,community area,CommunityDesc Community,Model modelid,OptionID,OptionID [Option],Assembly,Description,assemblytype,1 Source from salesassemblylist where assemblytype=3"
                
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Model and Option", s, , , , , "area,modelid,optionid,assemblytype,Source") Then
                
                    mCopiedCommunity = ""
                    mCopiedAssembly = ""
                    mCopiedModel = ""
                    mCommunity = FPickList.SelectedItem("area")
                    mAssembly = FPickList.SelectedItem("assembly")
                    mModel = FPickList.SelectedItem("modelid")
                    mOptionID = FPickList.SelectedItem("optionid")
                    mDataSource = FPickList.SelectedItem("Source")
                    mAssemblyType = Val(FPickList.SelectedItem("assemblytype"))
                    Call LoadData
                End If
            
            
            End If
            
        Case "TakeoffOneTime", "TakeoffItem", "TakeoffAssembly", "TakeoffCustom"
            Call FTakeoff.Takeoff(Me, Mid(Button.Key, 8), mAssemblyType, txtDescription.Text, mCommunity, "", cboModel.Text, txtAssembly.Text, "")
             
        Case "EEExport"
            If Not HFApp.Databases(dbEstimating).State = adStateOpen Then Exit Sub
            If MsgBox("This will write this assembly into Timberline Estimating." & vbCrLf & vbCrLf & "Are you sure this is what you want to do?", vbYesNo + vbExclamation, App.ProductName) = vbNo Then Exit Sub
            If Not ValidateData Then Exit Sub
            If Not SaveData(False) Then Exit Sub
            
            Screen.MousePointer = vbHourglass
            'get key
            sCommunity = GetComboBoxListKey(cboCommunity)
            sModel = cboModel.Text
            sOptionID = txtOption.Text
            sAssembly = txtAssembly.Text
            'update stuff
            Call ExportAssembly(sCommunity, sModel, sOptionID, sAssembly)
            
            Unload FProgress
            Screen.MousePointer = vbDefault
        
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
    Select Case Button.Key
        Case "New"
            PopupMenu FMain.mnuAssemblyNew, , Button.left, Button.Top + Button.Height
    End Select
    Exit Sub
eh: Call errHandler(SRCFILE & "Toolbar_ButtonDropDown")
End Sub



Public Sub LoadData()
    Dim i As Long
    Dim s As String
    Dim rs As Recordset
    Dim r As Long
    
    'for pricing
    Dim community As String
    Dim phase As String
    Dim GetRates As Boolean
    GetRates = Me.cboPriceCommunity.ListIndex > 0
    
Screen.MousePointer = vbHourglass
    
    ReadOnly = False
    
    
    s = ""
    s = s & "SELECT *" & vbCrLf
    s = s & "  FROM " & IIf(mDataSource = 0, "tblDBAssemblyMaster m", "SalesAssemblyList m") & vbCrLf
    s = s & " WHERE ISNULL(m.Community,'')=" & DbQuote(Str, mCommunity) & vbCrLf
    s = s & "   AND ISNULL(m.Model,'')=" & DbQuote(Str, mModel) & vbCrLf
    s = s & "   AND ISNULL(m.Optionid,'')=" & DbQuote(Str, mOptionID) & vbCrLf
    s = s & "   AND ISNULL(m.Assembly,'')=" & DbQuote(Str, mAssembly) & vbCrLf
    Set rs = HFApp.SqlExec(s)
    If rs.EOF Then
        cboCommunity.ListIndex = -1
        Call SetComboBoxListIndex(cboCommunity, , mCommunity)
        cboModel.Text = mModel
        txtOption.Text = mOptionID
        txtAssembly.Text = mAssembly
        txtDescription.Text = mAssemblyDesc
        txtComments.Text = ""
        chkActive.Value = vbChecked
        cboSeries.ListIndex = -1
        cboUOM.Text = ""
        txtStyle.Text = ""
        cboStyle.ListIndex = -1
        txtBedrooms.Text = ""
        txtBathrooms.Text = ""
        txtFloorArea.Text = ""
        cboCategory.ListIndex = -1
        txtJCExtra.Text = ""
        chkTakeoffRequired.Value = vbUnchecked
        chkUseNormalSalesQtyFactors.Value = vbChecked
        
        CtrlEnabled(cboCommunity) = True
        CtrlEnabled(cboModel) = mAssemblyType = atModel Or mAssemblyType = atOption
        CtrlEnabled(txtOption) = mAssemblyType <> atModel
        CtrlEnabled(txtAssembly) = True
        
    Else
        mCommunity = "" & rs("Community")
        
        mAssemblyType = Val("" & rs("AssemblyType"))
        cboCommunity.ListIndex = 0
        Call SetComboBoxListIndex(cboCommunity, , mCommunity)
        cboModel.Text = "" & rs("Model")
        txtOption.Text = "" & rs("OptionID")
        txtAssembly.Text = "" & rs("Assembly")
        
        chkActive.Value = IIf("" & rs("InActive") = "True", vbUnchecked, vbChecked)
        
        txtDescription.Text = "" & rs("Description")
        txtComments.Text = "" & rs("Notes")
        Call SetComboBoxListIndex(cboSeries, , "" & rs("Series"))
        txtStyle.Text = "" & rs("Style")
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
        chkTakeoffRequired.Value = IIf("" & rs("TakeoffRequired") = "True", vbChecked, vbUnchecked)
        chkUseNormalSalesQtyFactors.Value = IIf("" & rs("UseNormalSalesQtyFactors") = "True", vbChecked, vbUnchecked)
        
        CtrlEnabled(cboCommunity) = False
        CtrlEnabled(cboModel) = False
        CtrlEnabled(txtOption) = False
        CtrlEnabled(txtAssembly) = mDataSource = 1
    
    End If
    CtrlEnabled(cboUOM) = mAssemblyType <> atModel
    
    Call Form_Resize
    
    cboModel.AutoCompleteListItemsOnly = mAssemblyType <> atModel
    frmModelFields.Visible = mAssemblyType = atModel
    frmOptionFields.Visible = mAssemblyType <> atModel
 
 
    'for pricing
    s = GetComboBoxListKey(cboPriceCommunity)
    community = Parse(s, 1, Chr(3))
    phase = Parse(s, 2, Chr(3))
 
 
    s = ""
    s = s & "SELECT i.Description,i.OrderUOM,p.POIndex,p.Description POIndexDescription,i.JCCostCode,cc.Description JCCostCodeDesc,i.JCCategory,ct.Description JCCategoryDesc,i.TakeoffUOM,i.ConversionFactor,i.rounddir,i.roundto" & vbCrLf
    s = s & "      ,d.Phase,d.Item,d.TakeoffQty,d.OrderQty,d.notes" & vbCrLf
    s = s & "      ,i.isquote" & vbCrLf
    If GetRates Then
        'get vendor rates
        s = s & "      ,v.vendor_id Vendor" & vbCrLf
        s = s & "      ,v.vendor_name VendorDesc" & vbCrLf
        s = s & "      ,dbo.Purch_GetItemRate(0,0," & DbQuote(Str, community) & ", " & DbQuote(Str, phase) & ", d.Assembly,d.Model,d.OptionID,d.Phase,d.Item,d.Sequence,v.Vendor_id,getdate()) Price" & vbCrLf
        s = s & "      ,dbo.Purch_GetItemRateQuality(0,0," & DbQuote(Str, community) & ", " & DbQuote(Str, phase) & ", d.Assembly,d.Model,d.OptionID,d.Phase,d.Item,d.Sequence,v.Vendor_id,getdate()) PriceLevel" & vbCrLf
    Else
        'no community selected...
        s = s & "      ,'' Vendor" & vbCrLf
        s = s & "      ,'' VendorDesc" & vbCrLf
        s = s & "      ,0 Price" & vbCrLf
        s = s & "      ,4 PriceLevel" & vbCrLf
    End If
    s = s & "  FROM tblDBAssemblyDetails d" & vbCrLf
    s = s & "       LEFT OUTER JOIN tblPhaseItem i ON(d.Phase=i.Phase AND d.Item=i.Item)" & vbCrLf
    s = s & "       LEFT OUTER JOIN tblPOIndex p ON(i.POIndex=p.POIndex)" & vbCrLf
    s = s & "       LEFT OUTER JOIN StandardCostCodes cc ON(i.JCCostCode=cc.CostCode)" & vbCrLf
    s = s & "       LEFT OUTER JOIN StandardCategories ct ON(i.JCCategory=ct.Category)" & vbCrLf
    s = s & "       LEFT OUTER JOIN tblvendors v ON(v.vendor_id=dbo.Purch_GetCommunityVendor(" & DbQuote(Str, community) & ", p.POIndex))" & vbCrLf
    s = s & " WHERE isnull(d.Community,'')=" & DbQuote(Str, mCommunity) & vbCrLf
    s = s & "   AND isnull(d.Model,'')=" & DbQuote(Str, mModel) & vbCrLf
    s = s & "   AND isnull(d.Optionid,'')=" & DbQuote(Str, mOptionID) & vbCrLf
    s = s & "   AND isnull(d.Assembly,'')=" & DbQuote(Str, mAssembly) & vbCrLf
    s = s & "ORDER BY d.Phase,d.Item,d.Sequence" & vbCrLf
    
    Set rs = HFApp.SqlExec(s)
    With gItems
        .Rows = 1
        While Not rs.EOF
            r = .Rows
            .AddItem ""
            
            
            If GetRates Then
                Select Case Val("" & rs("PriceLevel"))
                    Case 4:  i = 1
                    Case 5:  i = 2
                    Case 6:  i = 3
                    Case 9:  i = 1
                    Case 10: i = 2
                    Case 11: i = 3
                    Case 12: i = 4
                End Select
                If i = 4 Then
                    .Cell(flexcpForeColor, r, .ColIndex("PriceLevel")) = vbHighlight
                    .Cell(flexcpForeColor, r, .ColIndex("Price")) = vbHighlight
                    .Cell(flexcpForeColor, r, .ColIndex("Vendor")) = vbHighlight
                    .Cell(flexcpForeColor, r, .ColIndex("VendorDesc")) = vbHighlight
                End If
                .TextMatrix(r, .ColIndex("PriceLevel")) = i
                .TextMatrix(r, .ColIndex("Price")) = "" & rs("Price")
                .TextMatrix(r, .ColIndex("Vendor")) = "" & rs("Vendor")
                .TextMatrix(r, .ColIndex("VendorDesc")) = "" & rs("VendorDesc")
            Else
                .TextMatrix(r, .ColIndex("PriceLevel")) = ""
                .TextMatrix(r, .ColIndex("Price")) = ""
                .TextMatrix(r, .ColIndex("Vendor")) = ""
                .TextMatrix(r, .ColIndex("VendorDesc")) = ""
                .Cell(flexcpBackColor, r, .ColIndex("PriceLevel")) = vbButtonFace
                .Cell(flexcpBackColor, r, .ColIndex("Price")) = vbButtonFace
                .Cell(flexcpBackColor, r, .ColIndex("Vendor")) = vbButtonFace
                .Cell(flexcpBackColor, r, .ColIndex("VendorDesc")) = vbButtonFace
            End If
            
            
            .TextMatrix(r, .ColIndex("ItemType")) = IIf("" & rs("IsQuote") = "True", 1, 0)
            'save original values so we can remove old cost records
            .TextMatrix(r, .ColIndex("OldPriceLevel")) = .TextMatrix(r, .ColIndex("PriceLevel"))
            .TextMatrix(r, .ColIndex("OldItemType")) = .TextMatrix(r, .ColIndex("ItemType"))
            
            
            .TextMatrix(r, .ColIndex("Phase")) = "" & rs("Phase")
            .TextMatrix(r, .ColIndex("Item")) = "" & rs("Item")
            .TextMatrix(r, .ColIndex("Description")) = "" & rs("Description")
            .TextMatrix(r, .ColIndex("TakeoffQty")) = Val("" & rs("TakeoffQty"))
            .TextMatrix(r, .ColIndex("TakeoffUOM")) = "" & rs("TakeoffUOM")
            .TextMatrix(r, .ColIndex("ConversionFactor")) = Val("" & rs("ConversionFactor"))
            .TextMatrix(r, .ColIndex("OrderQty")) = Val("" & rs("OrderQty"))
            .TextMatrix(r, .ColIndex("RoundTo")) = Val("" & rs("RoundTo"))
            .TextMatrix(r, .ColIndex("RoundDir")) = Val("" & rs("RoundDir"))
            .TextMatrix(r, .ColIndex("OrderUOM")) = "" & rs("OrderUOM")
            .TextMatrix(r, .ColIndex("Notes")) = "" & rs("Notes")
            .TextMatrix(r, .ColIndex("POIndex")) = "" & rs("POIndex")
            .TextMatrix(r, .ColIndex("POIndexDescription")) = "" & rs("POIndexDescription")
            .TextMatrix(r, .ColIndex("JCCostCode")) = "" & rs("JCCostCode")
            .TextMatrix(r, .ColIndex("JCCostCodeDesc")) = "" & rs("JCCostCodeDesc")
            .TextMatrix(r, .ColIndex("JCCategory")) = "" & rs("JCCategory")
            .TextMatrix(r, .ColIndex("JCCategoryDesc")) = "" & rs("JCCategoryDesc")
            
            rs.MoveNext
        Wend
    End With
    Call CalcTotal
    Screen.MousePointer = vbDefault
    mDirty = False
 
End Sub

Public Function SaveData(prompt As Boolean) As Boolean
On Error GoTo eh
    Dim i As Long
    Dim s As String
    Dim step As String
    Dim CommunityPhase As String
    Dim rs As Recordset
    
    
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
    'insert if is new or is from sales...
    If mAssembly = "" Or mDataSource = 1 Then
        
        If mCopiedAssembly <> "" Then
            If vbYes = MsgBox("Do you want to copy the vendor pricelists also?", vbQuestion + vbYesNo, App.ProductName) Then
                s = ""
                s = s & "INSERT INTO tblVendorCost(Community,CommunityPhase,Assembly,Model,Phase,Item,Vendor,Current_Cost,Next_Cost1,Next_Cost2,Next_Effective1,Next_Effective2,Last_Cost1,Last_Cost2,Last_Cost3,Last1_Expiry,Last2_Expiry,Last3_Expiry,Forecast1,Forecast2,Forecast3,Forecast4,Forecast5,Forecast6,Forecast7,Forecast8,Forecast9,Forecast10,Forecast11,Forecast12,TaxGroup,PartNumber,PriceLink)" & vbCrLf
                s = s & "SELECT " & DbQuote(Str, GetComboBoxListKey(cboCommunity)) & vbCrLf
                s = s & "      ,old.CommunityPhase" & vbCrLf
                s = s & "      ," & DbQuote(Str, txtAssembly.Text) & vbCrLf
                s = s & "      ," & DbQuote(Str, cboModel.Text) & vbCrLf
                s = s & "      ,old.Phase" & vbCrLf
                s = s & "      ,old.Item" & vbCrLf
                s = s & "      ,old.Vendor" & vbCrLf
                s = s & "      ,old.Current_Cost,old.Next_Cost1,old.Next_Cost2,old.Next_Effective1,old.Next_Effective2,old.Last_Cost1,old.Last_Cost2,old.Last_Cost3,old.Last1_Expiry,old.Last2_Expiry,old.Last3_Expiry" & vbCrLf
                s = s & "      ,old.Forecast1,old.Forecast2,old.Forecast3,old.Forecast4,old.Forecast5,old.Forecast6,old.Forecast7,old.Forecast8,old.Forecast9,old.Forecast10,old.Forecast11,old.Forecast12" & vbCrLf
                s = s & "      ,old.TaxGroup,old.PartNumber,old.PriceLink" & vbCrLf
                s = s & "  FROM tblVendorCost old" & vbCrLf
                s = s & "       LEFT OUTER JOIN tblVendorCost new ON(new.Community=" & DbQuote(Str, GetComboBoxListKey(cboCommunity)) & vbCrLf
                s = s & "                                        AND new.CommunityPhase=''" & vbCrLf
                s = s & "                                        AND new.Assembly=" & DbQuote(Str, txtAssembly.Text) & vbCrLf
                s = s & "                                        AND new.Model=" & DbQuote(Str, cboModel.Text) & vbCrLf
                s = s & "                                        AND new.Phase=old.Phase" & vbCrLf
                s = s & "                                        AND new.Item=old.Item" & vbCrLf
                s = s & "                                        AND new.Vendor=old.Vendor)" & vbCrLf
                s = s & " WHERE new.Vendor IS NULL" & vbCrLf
                s = s & "   AND old.Community=" & DbQuote(Str, mCopiedCommunity) & vbCrLf
                s = s & "   AND old.Assembly=" & DbQuote(Str, mCopiedAssembly) & vbCrLf
                s = s & "   AND old.Model=" & DbQuote(Str, mCopiedModel) & vbCrLf
                Call HFApp.SqlExec(s)

            End If
        End If
    
    
        s = ""
        s = s & "INSERT INTO tblDBAssemblyMaster(UStmp,TStmp,AssemblyType,InActive,Community,Model,OptionID,Assembly,AssemblyUOM,Description,Notes,Series,Style,Bedrooms,Bathrooms,FloorArea,Category,JCExtra,TakeoffRequired,UseNormalSalesQtyFactors)" & vbCrLf
        s = s & "VALUES(" & DbQuote(Str, HFApp.LoginID) & vbCrLf
        s = s & "      ,GETDATE()" & vbCrLf
        s = s & "      ," & DbQuote(Num, mAssemblyType) & vbCrLf
        s = s & "      ," & DbQuote(Bit, chkActive.Value = vbUnchecked) & vbCrLf
        s = s & "      ," & DbQuote(Str, GetComboBoxListKey(cboCommunity)) & vbCrLf
        s = s & "      ," & DbQuote(Str, cboModel.Text) & vbCrLf
        s = s & "      ," & DbQuote(Str, txtOption.Text) & vbCrLf
        s = s & "      ," & DbQuote(Str, txtAssembly.Text) & vbCrLf
        s = s & "      ," & DbQuote(Str, cboUOM.Text) & vbCrLf
        s = s & "      ," & DbQuote(Str, txtDescription.Text) & vbCrLf
        s = s & "      ," & DbQuote(Str, txtComments.Text) & vbCrLf
        s = s & "      ," & DbQuote(Str, GetComboBoxListKey(cboSeries)) & vbCrLf
        s = s & "      ," & DbQuote(Str, IIf(cboStyle.Visible, cboStyle.Text, txtStyle.Text)) & vbCrLf
        s = s & "      ," & DbQuote(Num, txtBedrooms.Text) & vbCrLf
        s = s & "      ," & DbQuote(Num, txtBathrooms.Text) & vbCrLf
        s = s & "      ," & DbQuote(Num, txtFloorArea.Text) & vbCrLf
        s = s & "      ," & DbQuote(Str, GetComboBoxListKey(cboCategory)) & vbCrLf
        s = s & "      ," & DbQuote(Str, txtJCExtra.Text) & vbCrLf
        s = s & "      ," & DbQuote(Bit, chkTakeoffRequired.Value = vbChecked) & vbCrLf
        s = s & "      ," & DbQuote(Bit, chkUseNormalSalesQtyFactors.Value = vbChecked) & ")" & vbCrLf
        Call HFApp.SqlExec(s)
        CtrlEnabled(cboCommunity) = False
        CtrlEnabled(cboModel) = False
        CtrlEnabled(txtOption) = False
        CtrlEnabled(txtAssembly) = False
        
        
        'now clear these
        mCopiedCommunity = ""
        mCopiedAssembly = ""
        mCopiedModel = ""
        
        mCommunity = GetComboBoxListKey(cboCommunity)
        mAssembly = txtAssembly.Text
        mModel = cboModel.Text
        mOptionID = txtOption.Text
    End If
    
    s = ""
    s = s & "UPDATE tblDBAssemblyMaster" & vbCrLf
    s = s & "SET UStmp=" & DbQuote(Str, HFApp.LoginID) & vbCrLf
    s = s & "   ,TStmp=GETDATE()" & vbCrLf
    s = s & "   ,Community=" & DbQuote(Str, GetComboBoxListKey(cboCommunity)) & vbCrLf
    s = s & "   ,Model=" & DbQuote(Str, cboModel.Text) & vbCrLf
    s = s & "   ,OptionID=" & DbQuote(Str, txtOption.Text) & vbCrLf
    s = s & "   ,Assembly=" & DbQuote(Str, txtAssembly.Text) & vbCrLf
    s = s & "   ,AssemblyUOM=" & DbQuote(Str, cboUOM.Text) & vbCrLf
    s = s & "   ,Description=" & DbQuote(Str, txtDescription.Text) & vbCrLf
    s = s & "   ,Notes=" & DbQuote(Str, txtComments.Text) & vbCrLf
    s = s & "   ,InActive=" & DbQuote(Bit, chkActive.Value = vbUnchecked) & vbCrLf
    s = s & "   ,Series=" & DbQuote(Str, GetComboBoxListKey(cboSeries)) & vbCrLf
    s = s & "   ,Style=" & DbQuote(Str, IIf(cboStyle.Visible, cboStyle.Text, txtStyle.Text)) & vbCrLf
    s = s & "   ,Bedrooms=" & DbQuote(Num, txtBedrooms.Text) & vbCrLf
    s = s & "   ,Bathrooms=" & DbQuote(Num, txtBathrooms.Text) & vbCrLf
    s = s & "   ,FloorArea=" & DbQuote(Num, txtFloorArea.Text) & vbCrLf
    s = s & "   ,Category=" & DbQuote(Str, GetComboBoxListKey(cboCategory)) & vbCrLf
    s = s & "   ,JCExtra=" & DbQuote(Str, txtJCExtra.Text) & vbCrLf
    s = s & "   ,TakeoffRequired=" & DbQuote(Bit, chkTakeoffRequired.Value = vbChecked) & vbCrLf
    s = s & "   ,UseNormalSalesQtyFactors=" & DbQuote(Bit, chkUseNormalSalesQtyFactors.Value = vbChecked) & vbCrLf
    s = s & "WHERE Community=" & DbQuote(Str, mCommunity) & vbCrLf
    s = s & "  AND Model=" & DbQuote(Str, mModel) & vbCrLf
    s = s & "  AND OptionID=" & DbQuote(Str, mOptionID) & vbCrLf
    s = s & "  AND Assembly=" & DbQuote(Str, mAssembly) & vbCrLf
    Call HFApp.SqlExec(s)
    
step = "ITEMS"
    With gItems
        
        s = ""
        s = s & "DELETE FROM tblDBAssemblyDetails" & vbCrLf
        s = s & "WHERE Community=" & DbQuote(Str, GetComboBoxListKey(cboCommunity)) & vbCrLf
        s = s & "  AND Model=" & DbQuote(Str, mModel) & vbCrLf
        s = s & "  AND OptionID=" & DbQuote(Str, mOptionID) & vbCrLf
        s = s & "  AND Assembly=" & DbQuote(Str, mAssembly) & vbCrLf
        Call HFApp.SqlExec(s)
        
        For i = 1 To .Rows - 1
        
            s = ""
            s = s & "INSERT INTO tblDBAssemblyDetails(Community,Assembly,Model,OptionID,Phase,Item,Notes,UStmp,TStmp,TakeoffQty,OrderQty)" & vbCrLf
            s = s & "VALUES(" & DbQuote(Str, GetComboBoxListKey(cboCommunity)) & vbCrLf
            s = s & "      ," & DbQuote(Str, txtAssembly.Text) & vbCrLf
            s = s & "      ," & DbQuote(Str, cboModel.Text) & vbCrLf
            s = s & "      ," & DbQuote(Str, txtOption.Text) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("Phase"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("Item"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("Notes"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, HFApp.LoginID) & vbCrLf
            s = s & "      ,GETDATE()" & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("TakeoffQty"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("OrderQty"))) & ")"
            Call HFApp.SqlExec(s)
            
            s = ""
            s = s & "UPDATE tblPhaseItem" & vbCrLf
            s = s & "SET IsQuote=" & DbQuote(Bit, .TextMatrix(i, .ColIndex("ItemType"))) & vbCrLf
            s = s & "WHERE Phase=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Phase"))) & vbCrLf
            s = s & "  AND Item=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Item"))) & vbCrLf
            Call HFApp.SqlExec(s)
            
            Call UpdatePrices(i)
            
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
            Case atOption
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
            Case atGlobal
                s = s & "UPDATE tblGlobalOptions" & vbCrLf
                s = s & "SET Assembly=" & DbQuote(Str, txtAssembly.Text) & vbCrLf
                s = s & "   ,Description=" & DbQuote(Str, txtDescription.Text) & vbCrLf
                s = s & "   ,UOM=" & DbQuote(Str, cboUOM.Text) & vbCrLf
                If Trim(txtComments.Text) <> "" Then s = s & "   ,Comments=" & DbQuote(Str, txtComments.Text) & vbCrLf
                s = s & "   ,Category=" & DbQuote(Str, GetComboBoxListKey(cboCategory)) & vbCrLf
                s = s & "   ,TL_Extra=" & DbQuote(Str, txtJCExtra.Text) & vbCrLf
                s = s & "WHERE ISNULL(Community,'')=" & DbQuote(Str, GetComboBoxListKey(cboCommunity)) & vbCrLf
                s = s & "  AND ISNULL(Assembly,'')=''" & vbCrLf
                s = s & "  AND Opt=" & DbQuote(Str, txtOption.Text) & vbCrLf
        End Select
        Call HFApp.SqlExec(s)
        mDataSource = 0
    End If
    
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
    Else
        Call errHandler(SRCFILE & "SaveData", s)
    End If
End Function




Private Property Get ReadOnly() As Boolean
    ReadOnly = mReadOnly
End Property

Private Property Let ReadOnly(RHS As Boolean)
    mReadOnly = RHS
    
    txtDescription.Enabled = Not RHS
    txtComments.Enabled = Not RHS
    cboSeries.Enabled = Not RHS
    txtStyle.Enabled = Not RHS
    txtBedrooms.Enabled = Not RHS
    txtBathrooms.Enabled = Not RHS
    txtFloorArea.Enabled = Not RHS
    cboCategory.Enabled = Not RHS
    txtJCExtra.Enabled = Not RHS
    chkTakeoffRequired.Enabled = Not RHS
    
    Toolbar.Buttons("Save").Enabled = Not RHS
    Toolbar.Buttons("Delete").Enabled = Not RHS
    Toolbar.Buttons("TakeoffOneTime").Enabled = False
    Toolbar.Buttons("TakeoffItem").Enabled = Not RHS
    Toolbar.Buttons("TakeoffAssembly").Enabled = Not RHS
    Toolbar.Buttons("TakeoffCustom").Enabled = Not RHS
    Toolbar.Buttons("EEExport").Enabled = Not RHS
    
End Property



Private Sub txtBathrooms_Validate(Cancel As Boolean)
    txtBathrooms.Text = Val(txtBathrooms.Text)
End Sub

Private Sub txtBedrooms_Validate(Cancel As Boolean)
    txtBedrooms.Text = Val(txtBedrooms.Text)
End Sub

Private Sub txtFloorArea_Change()
    mDirty = True
End Sub

Private Sub txtFloorArea_GotFocus()
    SelectAll txtFloorArea
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
Private Sub txtComments_GotFocus()
    SelectAll txtComments
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
End Sub
Private Sub cboSeries_Click()
    mDirty = True
End Sub
Private Sub cboCategory_Click()
    mDirty = True
End Sub

Private Sub gItems_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
    With gItems
        If Button = vbRightButton Then
            If .MouseRow < 1 Then
                Cancel = True
                Call FMain.ShowColumnMenu(gItems, False)
            Else
                Call PopupMenu(FMain.mnuAssemblyItems)
            End If
        End If
    End With
End Sub


Public Sub AddItem(Assembly As String, _
                   AssemblyDescription As String, _
                   Model As String, _
                   phase As String, _
                   item As String, _
                   Description As String, _
                   OrderQty As Double, _
                   OrderUOM As String, _
                   TakeoffQty As Double, _
                   TakeoffUOM As String, _
                   ConversionFactor As Double, RoundTo As Double, RoundDir As Long, WastePercent As Long, _
                   JCExtra As String, _
                   JCCostCode As String, _
                   JCCostCodeDesc As String, _
                   JCCategory As String, _
                   JCCategoryDesc As String, _
                   Vendor As String, _
                   vendorName As String, _
                   price As Double, _
                   TaxGroup As String, _
                   TaxGroupName As String, _
                   JCTaxRate As Double, _
                   NJCTaxRate As Double, _
                   POIndex As String, _
                   Comments As String)
On Error GoTo eh
    
    Dim i As Long
    Dim r As Long
    Dim s As String
    Dim rs As Recordset
    Dim sComm  As String
    Dim sPhase As String
    
    ReadOnly = False
    mDirty = True
    
    
    With gItems
        .AddItem ""
        r = .Rows - 1
        .TextMatrix(r, .ColIndex("Phase")) = phase
        .TextMatrix(r, .ColIndex("Item")) = item
        .TextMatrix(r, .ColIndex("Description")) = Description
        .TextMatrix(r, .ColIndex("JCCostCode")) = JCCostCode
        .TextMatrix(r, .ColIndex("JCCostCodeDesc")) = JCCostCodeDesc
        .TextMatrix(r, .ColIndex("JCCategory")) = JCCategory
        .TextMatrix(r, .ColIndex("JCCategoryDesc")) = JCCategoryDesc
        .TextMatrix(r, .ColIndex("POIndex")) = POIndex
        .TextMatrix(r, .ColIndex("ConversionFactor")) = ConversionFactor
        .TextMatrix(r, .ColIndex("RoundTo")) = RoundTo
        .TextMatrix(r, .ColIndex("RoundDir")) = RoundDir
        .TextMatrix(r, .ColIndex("WastePercent")) = WastePercent
        .TextMatrix(r, .ColIndex("Notes")) = Comments
        .TextMatrix(r, .ColIndex("TakeoffUOM")) = TakeoffUOM
        .TextMatrix(r, .ColIndex("TakeoffQty")) = TakeoffQty
        .TextMatrix(r, .ColIndex("OrderUOM")) = OrderUOM
        .TextMatrix(r, .ColIndex("OrderQty")) = OrderQty
    
        '--------------------------------------------------
        'get pricing for selected community/phase
        '--------------------------------------------------
        s = GetComboBoxListKey(cboPriceCommunity)
        sComm = Parse(s, 1, Chr(3))
        sPhase = Parse(s, 2, Chr(3))
        If sComm = "" Then
            .TextMatrix(r, .ColIndex("PriceLevel")) = ""
            .TextMatrix(r, .ColIndex("Price")) = ""
            .TextMatrix(r, .ColIndex("Vendor")) = ""
            .TextMatrix(r, .ColIndex("VendorDesc")) = ""
            .Cell(flexcpBackColor, r, .ColIndex("PriceLevel")) = vbButtonFace
            .Cell(flexcpBackColor, r, .ColIndex("Price")) = vbButtonFace
            .Cell(flexcpBackColor, r, .ColIndex("Vendor")) = vbButtonFace
            .Cell(flexcpBackColor, r, .ColIndex("VendorDesc")) = vbButtonFace
            s = ""
            s = s & "SELECT IsQuote FROM tblPhaseItem" & vbCrLf
            s = s & " WHERE isnull(phase,'')=" & DbQuote(Str, phase) & vbCrLf
            s = s & "   AND isnull(item,'')=" & DbQuote(Str, item) & vbCrLf
            Set rs = HFApp.SqlExec(s, dbHomefront)
            If Not rs.EOF Then
                .TextMatrix(r, .ColIndex("ItemType")) = IIf("" & rs("IsQuote") = "True", 1, 0)
            End If
            Exit Sub
        End If
            
        s = ""
        s = s & "SELECT i.IsQuote,v.vendor_id Vendor" & vbCrLf
        s = s & "      ,v.vendor_name VendorDesc" & vbCrLf
        s = s & "      ,dbo.Purch_GetItemRate(0,0," & DbQuote(Str, sComm) & ", " & DbQuote(Str, sPhase) & "," & DbQuote(Str, txtAssembly.Text) & "," & DbQuote(Str, cboModel.Text) & "," & DbQuote(Str, txtOption.Text) & ",i.Phase,i.Item,0,v.Vendor_id,getdate()) Price" & vbCrLf
        s = s & "      ,dbo.Purch_GetItemRateQuality(0,0," & DbQuote(Str, sComm) & ", " & DbQuote(Str, sPhase) & "," & DbQuote(Str, txtAssembly.Text) & "," & DbQuote(Str, cboModel.Text) & "," & DbQuote(Str, txtOption.Text) & ",i.Phase,i.Item,0,v.Vendor_id,getdate()) PriceLevel" & vbCrLf
        s = s & "  FROM tblPhaseItem i " & vbCrLf
        s = s & "       LEFT OUTER JOIN tblvendors v ON(v.vendor_id=dbo.Purch_GetCommunityVendor(" & DbQuote(Str, sComm) & "," & DbQuote(Str, POIndex) & "))" & vbCrLf
        s = s & " WHERE isnull(i.phase,'')=" & DbQuote(Str, phase) & vbCrLf
        s = s & "   AND isnull(i.item,'')=" & DbQuote(Str, item) & vbCrLf
        Set rs = HFApp.SqlExec(s, dbHomefront)
        If Not rs.EOF Then
            Select Case Val("" & rs("PriceLevel"))
                Case 4:  i = 1
                Case 5:  i = 2
                Case 6:  i = 3
                Case 9:  i = 1
                Case 10: i = 2
                Case 11: i = 3
                Case 12: i = 4
            End Select
            If i = 4 Then
                .Cell(flexcpForeColor, r, .ColIndex("PriceLevel")) = vbHighlight
                .Cell(flexcpForeColor, r, .ColIndex("Price")) = vbHighlight
                .Cell(flexcpForeColor, r, .ColIndex("Vendor")) = vbHighlight
                .Cell(flexcpForeColor, r, .ColIndex("VendorDesc")) = vbHighlight
            End If
            .TextMatrix(r, .ColIndex("ItemType")) = IIf("" & rs("IsQuote") = "True", 1, 0)
            .TextMatrix(r, .ColIndex("PriceLevel")) = i
            .TextMatrix(r, .ColIndex("Price")) = "" & rs("Price")
            .TextMatrix(r, .ColIndex("Vendor")) = "" & rs("Vendor")
            .TextMatrix(r, .ColIndex("VendorDesc")) = "" & rs("VendorDesc")
        End If
    End With
    
    
    
    
    
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
            at = atOption
        End If
    End If
    
    
    Select Case at
        Case atModel
            If Trim(cboModel.Text) = "" Then s = s & " " & vbBullet & " Model is required" & vbCrLf
            If cboSeries.ListIndex < 0 Then s = s & " " & vbBullet & " Series is required" & vbCrLf
        Case atOption
            If Trim(cboModel.Text) = "" Then s = s & " " & vbBullet & " Model is required" & vbCrLf
            If cboSeries.ListIndex < 0 Then s = s & " " & vbBullet & " Series is required" & vbCrLf
            If Trim(txtOption.Text) = "" Then s = s & " " & vbBullet & " Option is required" & vbCrLf
            If cboCategory.ListIndex < 0 And HFApp.Options(ReqCategory) Then s = s & " " & vbBullet & " Category is required" & vbCrLf
        Case atGlobal, atDesignCenter
            If Trim(txtOption.Text) = "" Then s = s & " " & vbBullet & " Option is required" & vbCrLf
            If cboCategory.ListIndex < 0 And HFApp.Options(ReqCategory) Then s = s & " " & vbBullet & " Category is required" & vbCrLf
    End Select
    
    With gItems
        For r = 1 To .Rows - 1
            If .TextMatrix(r, .ColIndex("Phase")) = "" Or .TextMatrix(r, .ColIndex("Item")) = "" And Not .RowHidden(r) Then
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

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    With gItems
    
        Call FAttachments.ShowForm("File Attachments - " & .TextMatrix(.Row, .ColIndex("Description")), _
                                   "ASM~" & mCommunity & "~" & mModel & "~" & mOptionID & "~" & mAssembly & "~" & .TextMatrix(.Row, .ColIndex("Phase")) & "~" & .TextMatrix(.Row, .ColIndex("Item")), "Assembly Master", _
                                   "ITM~" & .TextMatrix(.Row, .ColIndex("Phase")) & "~" & .TextMatrix(.Row, .ColIndex("Item")), "Item Database")
    End With
End Sub

Public Sub mnuAssemblyItemsSub_Click(Index As Integer)
    Dim s      As String
    Dim phaseitem As String
    Dim c      As Long
    Dim r      As Long
    Dim newrow As Long
    Dim ObjectID As String
    
    With gItems
    Select Case Index
    
        Case mcITEM_VIEWFILES
            ObjectID = "ASM~" & mCommunity & "~" & mModel & "~" & mOptionID & "~" & mAssembly & "~" & .TextMatrix(.Row, .ColIndex("Phase")) & "~" & .TextMatrix(.Row, .ColIndex("Item"))
            'Call FAttachments.Showform(ObjectID, .TextMatrix(.Row, .ColIndex("Description")), "")
            'this doesn't work. the popup menus on FAttachments don't work if the screen is displayed from inside this procedure
            'to get around it we have to do something weird.
            ' 1.) turn on a timer then exit sub
            ' 2.) when timer goes disable it and launch the attachments window
            Timer1.Interval = 10
            Timer1.Enabled = True
            
            
        Case mcITEM_INSERTFILE, mcITEM_LINKTOFILE
            ObjectID = "ASM~" & mCommunity & "~" & mModel & "~" & mOptionID & "~" & mAssembly & "~" & .TextMatrix(.Row, .ColIndex("Phase")) & "~" & .TextMatrix(.Row, .ColIndex("Item"))
            Call FAttachments.AddFile(ObjectID, Index = mcITEM_INSERTFILE)
        
        Case mcITEM_COPY
            For r = Max(.Row, .RowSel) To Min(.Row, .RowSel) Step -1
                mDirty = True
                newrow = Max(.Row, .RowSel) + 1
                .AddItem "", newrow
                For c = 0 To .cols - 1
                    .TextMatrix(newrow, c) = .TextMatrix(r, c)
                Next
            Next
                   
        Case mcITEM_SUBSTITUE
            s = ""
            

            s = s & "SELECT i.Phase" & vbCrLf
            s = s & "      ,i.Item" & vbCrLf
            s = s & "      ,isnull(phase,'')+char(1)+isnull(item,'') phaseitem " & vbCrLf
            s = s & "      ,i.Description" & vbCrLf
            s = s & "      ,i.TakeoffUOM" & vbCrLf
            s = s & "      ,i.ConversionFactor" & vbCrLf
            s = s & "      ,i.OrderUOM" & vbCrLf
            s = s & "      ,i.POIndex" & vbCrLf
            s = s & "      ,i.Notes" & vbCrLf
            s = s & "      ,i.JCCostCode CostCode" & vbCrLf
            s = s & "      ,cod.description CostCodeDesc" & vbCrLf
            s = s & "      ,i.JCCategory Category" & vbCrLf
            s = s & "      ,cat.Description CategoryDesc" & vbCrLf
            s = s & "  FROM tblPhaseItem i" & vbCrLf
            s = s & "LEFT OUTER JOIN StandardCostCodes cod ON(i.jccostcode=cod.costcode)" & vbCrLf
            s = s & "LEFT OUTER JOIN StandardCategories cat ON(i.jccategory=cat.category)" & vbCrLf
            .HighLight = flexHighlightAlways
            
            phaseitem = .TextMatrix(r, .ColIndex("Phase")) & Chr(1) & .TextMatrix(r, .ColIndex("Item"))
            If FPickList.Choose(HFApp.Databases(dbHomefront), "Item", s, phaseitem, , , , "phaseitem,TakeoffUOM,ConversionFactor,OrderUOM,POIndex,Notes,CostCode,CostCodeDesc,Category,CategoryDesc") Then
                mDirty = True
                For r = Max(.Row, .RowSel) To Min(.Row, .RowSel) Step -1
                    .TextMatrix(r, .ColIndex("Phase")) = FPickList.SelectedItem("Phase")
                    .TextMatrix(r, .ColIndex("Item")) = FPickList.SelectedItem("Item")
                    .TextMatrix(r, .ColIndex("Description")) = FPickList.SelectedItem("Description")
                    .TextMatrix(r, .ColIndex("TakeoffUOM")) = FPickList.SelectedItem("TakeoffUOM")
                    .TextMatrix(r, .ColIndex("ConversionFactor")) = Val("" & FPickList.SelectedItem("ConversionFactor"))
                    .TextMatrix(r, .ColIndex("OrderUOM")) = FPickList.SelectedItem("OrderUOM")
                    .TextMatrix(r, .ColIndex("POIndex")) = FPickList.SelectedItem("POIndex")
                    .TextMatrix(r, .ColIndex("Notes")) = FPickList.SelectedItem("Notes")
                    .TextMatrix(r, .ColIndex("JCCostCode")) = FPickList.SelectedItem("CostCode")
                    .TextMatrix(r, .ColIndex("JCCostCodeDesc")) = FPickList.SelectedItem("CostCodeDesc")
                    .TextMatrix(r, .ColIndex("JCCategory")) = FPickList.SelectedItem("Category")
                    .TextMatrix(r, .ColIndex("JCCategoryDesc")) = FPickList.SelectedItem("CategoryDesc")
                    .TextMatrix(r, .ColIndex("OrderQty")) = .ValueMatrix(r, .ColIndex("TakeoffQty")) * .ValueMatrix(r, .ColIndex("ConversionFactor"))
                Next
            End If
            .HighLight = flexHighlightWithFocus
                
                    
        Case mcITEM_REMOVE
            Call gItems_KeyDown(vbKeyDelete, vbCtrlMask)
            
             
            
    End Select
    End With
End Sub

Public Sub mnuAssemblyNewSub_Click(Index As Integer)
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
                Case atOption:
            End Select
            
            'save these so we can duplicate the vendor pricing when we save the new one
            mCopiedCommunity = mCommunity
            mCopiedAssembly = mAssembly
            mCopiedModel = mModel
            mCommunity = ""
            mAssembly = ""
            mModel = ""
            mOptionID = ""
            mDirty = True
            
        Case Else
            Select Case Index
                Case mcNEW_MODEL:    mAssemblyType = atModel
                Case mcNEW_OPTION:   mAssemblyType = atOption
                Case mcNEW_GLOBAL:   mAssemblyType = atGlobal
                Case mcNEW_DCOPTION: mAssemblyType = atDesignCenter
            End Select
            mCopiedCommunity = ""
            mCopiedAssembly = ""
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
    If MsgBox("Are you sure you want to delete this assembly?", vbExclamation + vbDefaultButton2 + vbOKCancel, App.ProductName) = vbOK Then
        
        s = ""
        s = s & "exec Purch_DeleteAssembly "
        s = s & "       " & DbQuote(Str, GetComboBoxListKey(cboCommunity)) & vbCrLf
        s = s & "      ," & DbQuote(Str, txtAssembly.Text) & vbCrLf
        s = s & "      ," & DbQuote(Str, cboModel.Text) & vbCrLf
        s = s & "      ," & DbQuote(Str, txtOption.Text) & vbCrLf
        s = s & "      ," & DbQuote(Str, MachineName) & vbCrLf
        s = s & "      ," & DbQuote(Str, HFApp.LoginID)
        
        Call HFApp.SqlExec(s)
        mDirty = False
        mCopiedCommunity = ""
        mCopiedAssembly = ""
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
    End If
Exit Sub
eh: Select Case True
        Case Parse(Err.Description, Parse(Err.Description, , "]"), "]") = "Unable to delete model assembly. Model specific options assemblies exist."
            MsgBox "Unable to delete model assembly. Model specific options assemblies exist.", vbExclamation, App.ProductName
        Case Else: Call errHandler(SRCFILE & "DeleteAssembly")
    End Select
End Sub









Private Sub cboModel_Click()
    mDirty = True
    If mAssemblyType = atModel Then
        txtAssembly.Text = Parse(cboModel.Text, 1, " - ")
        txtDescription.Text = Parse(cboModel.Text, 2, " - ")
    End If
    cboModel.Text = Parse(cboModel.Text, 1, " - ")
End Sub

Private Sub cboModel_Change()
    mDirty = True
    If mAssemblyType = atModel Then txtAssembly.Text = cboModel.Text
End Sub
Private Sub cboModel_Validate(Cancel As Boolean)
    cboModel.Text = Parse(cboModel.Text, 1, " - ")
End Sub




Private Sub ExportAssembly(community As String, Model As String, OptionID As String, Assembly As String)
On Error GoTo eh
    Dim TLAssembly As String
    Dim s As String
    Dim Row As Long
    Dim rs As Recordset
    Dim Count As Long
    Dim header As Long


    TLAssembly = FormatTSField(20, Trim(Assembly), False, False)
    
    Call FProgress.Progress("Synchronizing assembly...", "querying...", 1, 2)
    
    ' get row count
    Row = 0
    s = ""
    s = s & "SELECT COUNT(*)" & vbCrLf
    s = s & "  FROM tblDBAssemblyMaster m" & vbCrLf
    s = s & "       JOIN tblDBAssemblyDetails d ON(m.Community=d.Community and m.Assembly=d.Assembly and m.Model=d.Model and m.OptionID=d.OptionID)" & vbCrLf
    s = s & " WHERE m.Community=" & DbQuote(Str, community) & vbCrLf
    s = s & "   AND m.Assembly=" & DbQuote(Str, Assembly) & vbCrLf
    s = s & "   AND m.Model=" & DbQuote(Str, Model) & vbCrLf
    s = s & "   AND m.OptionID=" & DbQuote(Str, OptionID) & vbCrLf
    Count = Val("" & HFApp.SqlExec(s)(0))
    
    ' get items
    s = ""
    s = s & "SELECT DISTINCT m.Assembly,m.Description,m.Notes,i.Phase,i.ItemNumber,d.TakeoffQty,d.Rate" & vbCrLf
    s = s & "  FROM tblDBAssemblyMaster m" & vbCrLf
    s = s & "       JOIN tblDBAssemblyDetails d ON(m.Community=d.Community and m.Assembly=d.Assembly and m.Model=d.Model and m.OptionID=d.OptionID)" & vbCrLf
    s = s & "       JOIN tblPhaseItem i ON(d.phase=i.phase and d.item=i.item)" & vbCrLf
    s = s & " WHERE m.Community=" & DbQuote(Str, community) & vbCrLf
    s = s & "   AND m.Assembly=" & DbQuote(Str, Assembly) & vbCrLf
    s = s & "   AND m.Model=" & DbQuote(Str, Model) & vbCrLf
    s = s & "   AND m.OptionID=" & DbQuote(Str, OptionID) & vbCrLf
    s = s & "ORDER BY m.Assembly,m.Description,m.Notes,i.Phase,i.ItemNumber,d.TakeoffQty,d.Rate" & vbCrLf
    Set rs = HFApp.SqlExec(s)

    If Not rs.EOF Then
        
        'get header id
        On Error Resume Next
        header = Val("" & HFApp.SqlExec("select header from dat_pwa__db_assembly_header where assembly_id=" & DbQuote(Str, TLAssembly), dbEstimating)(0))
        If header = 0 Then
            header = Val("" & HFApp.SqlExec("SELECT MAX(Header) FROM DAT_PWA__DB_ASSEMBLY_HEADER", dbEstimating)(0))
            header = header + 1
        End If
        On Error GoTo eh

        'now insert/update master
        s = ""
        s = s & "INSERT INTO DAT_PWA__DB_ASSEMBLY_HEADER(Header,Assembly_ID,Assembly_Description,Assembly_Notes,Assembly_Level,Calculation,Date_Stamp,Time_Stamp)" & vbCrLf
        s = s & "VALUES(" & header & vbCrLf
        s = s & "      ," & DbQuote(Str, TLAssembly) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("Description"), , , HFApp.EstFieldSize("AssemblyDesc")) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("Notes"), , , HFApp.EstFieldSize("AssemblyNotes")) & vbCrLf
        s = s & "      ,0" & vbCrLf
        s = s & "      ,'1'" & vbCrLf
        s = s & "      ," & DbQuote(Date, Now) & vbCrLf
        s = s & "      ," & DbQuote(Time, Now) & ")"
        Call HFApp.SqlExec(s, dbEstimating)

        s = ""
        s = s & "UPDATE DAT_PWA__DB_ASSEMBLY_HEADER" & vbCrLf
        s = s & "SET Assembly_Description=" & DbQuote(Str, "" & rs("Description"), , , HFApp.EstFieldSize("AssemblyDesc")) & vbCrLf
        s = s & "   ,Assembly_Notes=" & DbQuote(Str, "" & rs("Notes"), , , HFApp.EstFieldSize("AssemblyNotes")) & vbCrLf
        s = s & "WHERE header=" & DbQuote(Num, header) & vbCrLf
        Call HFApp.SqlExec(s, dbEstimating)
        
        'remove all items
        Call HFApp.SqlExec("delete from dat_pwa__db_assembly_detail where asb_header_number=" & DbQuote(Num, header), dbEstimating)
      
      
        'now insert new items
        While Not rs.EOF
          
            Row = Row + 1
          
            Call FProgress.Progress(, "updating database...", Row, Count)
            
            s = ""
            s = s & "INSERT INTO DAT_PWA__DB_ASSEMBLY_DETAIL(Asb_header_number,Asb_detail_number,Ordinality,Phase_Code,Item_code,Factor)" & vbCrLf
            s = s & "VALUES(" & header & vbCrLf
            s = s & "      ," & Row & vbCrLf
            s = s & "      ," & Row & vbCrLf
            s = s & "      ," & DbQuote(Str, HFApp.FormatPhase("" & rs("Phase"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, HFApp.FormatItem("" & rs("ItemNumber"))) & vbCrLf
            If Val("" & rs("Rate")) < 0 Then
                s = s & "      ," & DbQuote(Num, -1 * Val("" & rs("TakeoffQty"))) & ")" & vbCrLf
            Else
                s = s & "      ," & DbQuote(Num, "" & rs("TakeoffQty")) & ")" & vbCrLf
            End If
            Call HFApp.SqlExec(s, dbEstimating)
            
            rs.MoveNext
        Wend
            
    End If

Exit Sub
eh: If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        Resume Next
    Else
        Call errHandler(SRCFILE & "ExportAssembly", s)
    End If
End Sub







Private Sub ExportPriceList()
On Error GoTo eh
    Dim rs As Recordset
    Dim s As String
    Dim Vendor As String
    Dim FileName As String
    
    If Not SaveData(False) Then Exit Sub
    If Not VBGetSaveFileName(FileName, , , "Excel Files (*.xls)|*.xls", , , , "xls", Me.hwnd) Then Exit Sub
 
 
        
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
    s = s & "left outer join tbldbassemblydetails d on (m.assembly=d.assembly and m.community=d.community and m.model=d.model and m.optionid=d.optionid)" & vbCrLf
    s = s & "left outer join tblphaseitem i on (d.phase=i.phase and d.item=i.item)" & vbCrLf
    s = s & "left outer join tblvendors v on(vendor_id=dbo.purch_getcommunityvendor(m.community,i.poindex))" & vbCrLf
    s = s & "left outer join tbllocality l on(m.community=l.area)" & vbCrLf
    s = s & "where m.Community=" & DbQuote(Str, GetComboBoxListKey(cboCommunity)) & vbCrLf
    s = s & "  and m.Assembly=" & DbQuote(Str, txtAssembly.Text) & vbCrLf
    s = s & "  and m.model=" & DbQuote(Str, cboModel.Text) & vbCrLf
    s = s & "  and m.OptionID=" & DbQuote(Str, txtOption.Text) & vbCrLf
    s = s & "  and i.isquote=1" & vbCrLf
    s = s & "order by v.vendor_id,d.phase,d.item" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    
    With gExportData
        .Rows = 0
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
            Call .AutoSize(0, .cols - 1)
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
            Call MsgBox("Cannot create the " & s & " file." & vbCrLf & "You may not have access to the folder or the file may be in use.", vbExclamation)
        Case Else
            errHandler (SRCFILE & "ExportPriceList")
    End Select
End Sub
 

Private Sub CalcTotal()
    Dim r As Long
    Dim t As Double
    
    'no community selected so do nothing
    If Me.cboPriceCommunity.ListIndex < 1 Then Exit Sub
    
    
    With Me.gItems
        t = 0
        For r = 1 To .Rows - 1
            t = t + Round(Val(.TextMatrix(r, .ColIndex("OrderQty"))) * Val(.TextMatrix(r, .ColIndex("Price"))), 2)
        Next
        Me.lblTotalCost.Caption = format(t, "currency")
    End With
    
End Sub



Private Sub UpdatePrices(Row As Long)
    Dim s As String
    
    Dim Vendor As String
    Dim Model As String
    Dim OptionID As String
    Dim Assembly As String
    Dim community As String
    Dim CommunityPhase As String
    Dim phase As String
    Dim item As String
    Dim itemtype As Long
    Dim pricelevel As Long
    Dim price As Double
    Dim olditemtype As Long
    Dim oldpricelevel As Long
    
    
    
    
    With Me.gItems
                
        oldpricelevel = Val(.TextMatrix(Row, .ColIndex("OldPriceLevel")))
        olditemtype = Val(.TextMatrix(Row, .ColIndex("OldItemType")))
        
        Vendor = .TextMatrix(Row, .ColIndex("Vendor"))
        Model = mModel
        OptionID = mOptionID
        Assembly = mAssembly
        
        s = GetComboBoxListKey(cboPriceCommunity)
        community = Parse(s, 1, Chr(3))
        CommunityPhase = Parse(s, 2, Chr(3))
        price = Val(.TextMatrix(Row, .ColIndex("Price")))
        phase = .TextMatrix(Row, .ColIndex("phase"))
        item = .TextMatrix(Row, .ColIndex("item"))
        itemtype = Val(.TextMatrix(Row, .ColIndex("itemtype")))
        pricelevel = Val(.TextMatrix(Row, .ColIndex("PriceLevel")))
        
    End With
    
    
    'remove old entries that are at a higher level
    'pricelevels
    '  1 = phase
    '  2 = community
    '  3 = global
    '  4 = item db
    If oldpricelevel < pricelevel Then
        s = ""
        s = s & "delete from tblvendorcost" & vbCrLf
        s = s & "where vendor=" & DbQuote(Str, Vendor) & vbCrLf
        s = s & "  and phase=" & DbQuote(Str, phase) & vbCrLf
        s = s & "  and item=" & DbQuote(Str, item) & vbCrLf
        If olditemtype = 1 Then
            s = s & "  and assembly=" & DbQuote(Str, Assembly) & vbCrLf
        Else
            s = s & "  and isnull(assembly,'')=''" & vbCrLf
        End If
        Select Case oldpricelevel
        Case 1
            s = s & "  and community=" & DbQuote(Str, community) & vbCrLf
            s = s & "  and communityphase=" & DbQuote(Str, CommunityPhase) & vbCrLf
        Case 2
            s = s & "  and community=" & DbQuote(Str, community) & vbCrLf
            s = s & "  and isnull(communityphase,'')=''" & vbCrLf
        Case Else
            s = s & "  and isnull(community,'')=''" & vbCrLf
            s = s & "  and isnull(communityphase,'')=''" & vbCrLf
        End Select
        HFApp.SqlExec s, dbHomefront
    End If
    
    
    'now write new price
    If pricelevel = 4 Then
        s = ""
        s = s & "update tblphaseitem" & vbCrLf
        s = s & "   set price=" & DbQuote(Num, price) & vbCrLf
        s = s & " where phase=" & DbQuote(Str, phase) & vbCrLf
        s = s & "   and item=" & DbQuote(Str, item) & vbCrLf
        HFApp.SqlExec s, dbHomefront
    Else
        s = ""
        s = s & "insert into tblvendorcost(Current_cost,vendor,phase,item,assembly,community,communityphase)" & vbCrLf
        s = s & "values(" & DbQuote(Num, price) & vbCrLf
        s = s & "      ," & DbQuote(Str, Vendor) & vbCrLf
        s = s & "      ," & DbQuote(Str, phase) & vbCrLf
        s = s & "      ," & DbQuote(Str, item) & vbCrLf
        s = s & "      ," & IIf(itemtype = 1, DbQuote(Str, Assembly), "''") & vbCrLf
        s = s & "      ," & IIf(pricelevel < 3, DbQuote(Str, community), "''") & vbCrLf
        s = s & "      ," & IIf(pricelevel < 2, DbQuote(Str, CommunityPhase), "''") & ")"
        On Error Resume Next
        HFApp.SqlExec s, dbHomefront
        On Error GoTo 0
        
        s = ""
        s = s & "update tblvendorcost" & vbCrLf
        s = s & "   set current_cost=" & DbQuote(Num, price) & vbCrLf
        s = s & " where vendor=" & DbQuote(Str, Vendor) & vbCrLf
        s = s & "   and phase=" & DbQuote(Str, phase) & vbCrLf
        s = s & "   and item=" & DbQuote(Str, item) & vbCrLf
        s = s & "   and assembly=" & IIf(itemtype = 1, DbQuote(Str, Assembly), "''") & vbCrLf
        s = s & "   and community=" & IIf(pricelevel < 3, DbQuote(Str, community), "''") & vbCrLf
        s = s & "   and communityphase=" & IIf(pricelevel < 2, DbQuote(Str, CommunityPhase), "''")
        HFApp.SqlExec s, dbHomefront
    
    End If
    
    
End Sub
