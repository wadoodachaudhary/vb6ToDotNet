VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Begin VB.Form FGenerateInvoices 
   Caption         =   "Select purchase orders for payment"
   ClientHeight    =   8175
   ClientLeft      =   1155
   ClientTop       =   3150
   ClientWidth     =   15090
   Icon            =   "FGenerateInvoices.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   15090
   Begin HFPayables.Slider Slider 
      Height          =   4335
      Left            =   6945
      Top             =   2055
      Width           =   60
      _ExtentX        =   106
      _ExtentY        =   7646
   End
   Begin VB.Frame FrameSearch 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "4"
      Height          =   1635
      Left            =   -15
      TabIndex        =   26
      Top             =   0
      Width           =   16845
      Begin VB.TextBox txtVendorDesc 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2655
         TabIndex        =   52
         ToolTipText     =   "Use the % wildcard in search criteria"
         Top             =   780
         Width           =   2310
      End
      Begin VB.TextBox txtJobDesc 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2655
         TabIndex        =   51
         ToolTipText     =   "Use the % wildcard in search criteria"
         Top             =   1020
         Width           =   2310
      End
      Begin VB.TextBox txtPODesc 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2655
         TabIndex        =   50
         ToolTipText     =   "Use the % wildcard in search criteria"
         Top             =   1260
         Width           =   2310
      End
      Begin VB.TextBox txtPaymentDate 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   7815
         TabIndex        =   5
         Top             =   1020
         Width           =   1875
      End
      Begin VB.TextBox txtInvoiceDate 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   7815
         TabIndex        =   4
         Top             =   780
         Width           =   1875
      End
      Begin VB.TextBox txtAccountingDate 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   7815
         TabIndex        =   6
         Top             =   1260
         Width           =   1875
      End
      Begin VB.TextBox txtPO 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   1500
         TabIndex        =   2
         ToolTipText     =   "Use the % wildcard in search criteria"
         Top             =   1260
         Width           =   1140
      End
      Begin VB.TextBox txtJob 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   1500
         TabIndex        =   1
         ToolTipText     =   "Use the % wildcard in search criteria"
         Top             =   1020
         Width           =   1140
      End
      Begin VB.TextBox txtVendor 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   1500
         TabIndex        =   0
         ToolTipText     =   "Use the % wildcard in search criteria"
         Top             =   780
         Width           =   1140
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "S&earch"
         Height          =   315
         Left            =   5220
         TabIndex        =   3
         Top             =   1170
         Width           =   1155
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6705
         TabIndex        =   47
         Top             =   1035
         Width           =   1005
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6795
         TabIndex        =   46
         Top             =   795
         Width           =   915
      End
      Begin VB.Label lblAccountingDate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Accounting Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6510
         TabIndex        =   28
         Top             =   1275
         Width           =   1200
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   210
         Picture         =   "FGenerateInvoices.frx":000C
         Top             =   210
         Width           =   480
      End
      Begin VB.Label lblJob 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Job/Description"
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
         Left            =   270
         TabIndex        =   24
         Top             =   1020
         Width           =   1125
      End
      Begin VB.Label lblPO 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PO/Description"
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
         Left            =   300
         TabIndex        =   25
         Top             =   1260
         Width           =   1095
      End
      Begin VB.Label lblVendor 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor/Name"
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
         Left            =   390
         TabIndex        =   23
         Top             =   780
         Width           =   1005
      End
      Begin VB.Label lblCaption 
         BackStyle       =   0  'Transparent
         Caption         =   $"FGenerateInvoices.frx":08D6
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   900
         TabIndex        =   27
         Top             =   60
         Width           =   9075
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000F&
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   2.30016e6
      End
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&Save"
      Height          =   375
      Index           =   0
      Left            =   5640
      TabIndex        =   21
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Index           =   1
      Left            =   6960
      TabIndex        =   22
      Top             =   7680
      Width           =   1215
   End
   Begin VSFlex8Ctl.VSFlexGrid gPOs 
      Height          =   5595
      Left            =   0
      TabIndex        =   7
      Top             =   1650
      Width           =   6720
      _cx             =   11853
      _cy             =   9869
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
      GridColorFixed  =   -2147483633
      TreeColor       =   -2147483632
      FloodColor      =   -2147483635
      SheetBorder     =   -2147483643
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   63
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FGenerateInvoices.frx":09CE
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   2
      AutoSearchDelay =   3
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   5
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   2
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
   Begin VB.Frame frmInvoice 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   5655
      Left            =   7830
      TabIndex        =   29
      Top             =   1800
      Width           =   12105
      Begin VSFlex8Ctl.VSFlexGrid gItems 
         Height          =   3015
         Left            =   30
         TabIndex        =   20
         Top             =   2250
         Width           =   8745
         _cx             =   15425
         _cy             =   5318
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
         GridColorFixed  =   -2147483633
         TreeColor       =   -2147483632
         FloodColor      =   -2147483635
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   0
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   35
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FGenerateInvoices.frx":12F6
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   2
         AutoSearchDelay =   3
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   5
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   3
         PicturesOver    =   0   'False
         FillStyle       =   1
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   1
         OwnerDraw       =   0
         Editable        =   1
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
      Begin TDBNumber6Ctl.TDBNumber numAmount 
         Height          =   225
         Left            =   3660
         TabIndex        =   30
         Top             =   105
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   397
         Calculator      =   "FGenerateInvoices.frx":186A
         Caption         =   "FGenerateInvoices.frx":188A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FGenerateInvoices.frx":18EF
         Keys            =   "FGenerateInvoices.frx":190D
         Spin            =   "FGenerateInvoices.frx":1965
         AlignHorizontal =   1
         AlignVertical   =   2
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   0
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,###,##0.00;-###,###,##0.00;0.00;0.00"
         EditMode        =   0
         Enabled         =   0
         ErrorBeep       =   -1
         ForeColor       =   -2147483640
         Format          =   "###,###,##0.00;-###,###,##0.00"
         HighlightText   =   -1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999.99
         MinValue        =   -999999999.99
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   62390273
         Value           =   0
         MaxValueVT      =   1952776197
         MinValueVT      =   1146880005
      End
      Begin TDBDate6Ctl.TDBDate dteReceived 
         Height          =   225
         Left            =   5610
         TabIndex        =   12
         Top             =   105
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   397
         Calendar        =   "FGenerateInvoices.frx":198D
         Caption         =   "FGenerateInvoices.frx":1A82
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FGenerateInvoices.frx":1AE7
         Keys            =   "FGenerateInvoices.frx":1B05
         Spin            =   "FGenerateInvoices.frx":1B71
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   0
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "dd-mmm-yy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   -1
         FirstMonth      =   4
         ForeColor       =   -2147483640
         Format          =   "dd-mmm-yy"
         HighlightText   =   2
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxDate         =   2958465
         MinDate         =   -657434
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__-___-__"
         ValidateMode    =   0
         ValueVT         =   1
         Value           =   3.66819414244724E-316
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate dteInvoice 
         Height          =   225
         Left            =   840
         TabIndex        =   10
         Top             =   585
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   397
         Calendar        =   "FGenerateInvoices.frx":1B99
         Caption         =   "FGenerateInvoices.frx":1C8E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FGenerateInvoices.frx":1CF3
         Keys            =   "FGenerateInvoices.frx":1D11
         Spin            =   "FGenerateInvoices.frx":1D75
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   0
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "dd-mmm-yy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   -1
         FirstMonth      =   4
         ForeColor       =   -2147483640
         Format          =   "dd-mmm-yy"
         HighlightText   =   2
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxDate         =   2958465
         MinDate         =   -657434
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__-___-__"
         ValidateMode    =   0
         ValueVT         =   1
         Value           =   1.10537306944062E-317
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate dtePayment 
         Height          =   225
         Left            =   5610
         TabIndex        =   13
         Top             =   345
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   397
         Calendar        =   "FGenerateInvoices.frx":1D9D
         Caption         =   "FGenerateInvoices.frx":1E92
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FGenerateInvoices.frx":1EF7
         Keys            =   "FGenerateInvoices.frx":1F15
         Spin            =   "FGenerateInvoices.frx":1F81
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   0
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "dd-mmm-yy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   -1
         FirstMonth      =   4
         ForeColor       =   -2147483640
         Format          =   "dd-mmm-yy"
         HighlightText   =   2
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxDate         =   2958465
         MinDate         =   -657434
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__-___-__"
         ValidateMode    =   0
         ValueVT         =   1
         Value           =   2.69366467562102E-316
         CenturyMode     =   0
      End
      Begin TDBNumber6Ctl.TDBNumber numTax 
         Height          =   225
         Left            =   3660
         TabIndex        =   31
         Top             =   345
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   397
         Calculator      =   "FGenerateInvoices.frx":1FA9
         Caption         =   "FGenerateInvoices.frx":1FC9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FGenerateInvoices.frx":202E
         Keys            =   "FGenerateInvoices.frx":204C
         Spin            =   "FGenerateInvoices.frx":20A4
         AlignHorizontal =   1
         AlignVertical   =   2
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   0
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,###,##0.00;-###,###,##0.00;0.00;0.00"
         EditMode        =   0
         Enabled         =   0
         ErrorBeep       =   -1
         ForeColor       =   -2147483640
         Format          =   "###,###,##0.00;-###,###,##0.00"
         HighlightText   =   -1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999.99
         MinValue        =   -999999999.99
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   56623105
         Value           =   0
         MaxValueVT      =   1952776197
         MinValueVT      =   1146880005
      End
      Begin TDBNumber6Ctl.TDBNumber numDiscount 
         Height          =   225
         Left            =   3660
         TabIndex        =   11
         Top             =   585
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   397
         Calculator      =   "FGenerateInvoices.frx":20CC
         Caption         =   "FGenerateInvoices.frx":20EC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FGenerateInvoices.frx":2151
         Keys            =   "FGenerateInvoices.frx":216F
         Spin            =   "FGenerateInvoices.frx":21A9
         AlignHorizontal =   1
         AlignVertical   =   2
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   0
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,###,##0.00;-###,###,##0.00;0.00;0.00"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   -1
         ForeColor       =   -2147483640
         Format          =   "###,###,##0.00;-###,###,##0.00"
         HighlightText   =   -1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999.99
         MinValue        =   -999999999.99
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   1
         Value           =   0
         MaxValueVT      =   1952776197
         MinValueVT      =   1146880005
      End
      Begin VB.TextBox txtInvoice 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   840
         MaxLength       =   15
         TabIndex        =   9
         Top             =   345
         Width           =   1695
      End
      Begin VB.TextBox txtDescription 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5130
         MaxLength       =   30
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   1170
         Width           =   3285
      End
      Begin TDBDate6Ctl.TDBDate dteDiscount 
         Height          =   225
         Left            =   5610
         TabIndex        =   14
         Top             =   585
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   397
         Calendar        =   "FGenerateInvoices.frx":21D1
         Caption         =   "FGenerateInvoices.frx":22C6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FGenerateInvoices.frx":232B
         Keys            =   "FGenerateInvoices.frx":2349
         Spin            =   "FGenerateInvoices.frx":23B5
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   0
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "dd-mmm-yy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   -1
         FirstMonth      =   4
         ForeColor       =   -2147483640
         Format          =   "dd-mmm-yy"
         HighlightText   =   2
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxDate         =   2958465
         MinDate         =   -657434
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__-___-__"
         ValidateMode    =   0
         ValueVT         =   1
         Value           =   9.4965543544696E-318
         CenturyMode     =   0
      End
      Begin VB.TextBox txtInvoiceCode1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5130
         MaxLength       =   10
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   930
         Width           =   1635
      End
      Begin VB.TextBox txtInvoiceCode2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6780
         MaxLength       =   10
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   930
         Width           =   1635
      End
      Begin VB.Label lblPayPoints 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pay Points: 10% - 20% - 20% - 20% - 30%"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   30
         TabIndex        =   53
         Top             =   2025
         Width           =   2880
      End
      Begin VB.Label lblPOComments 
         BackColor       =   &H80000005&
         Height          =   465
         Left            =   5130
         TabIndex        =   49
         Top             =   1410
         UseMnemonic     =   0   'False
         Width           =   3300
      End
      Begin VB.Label lblPOComm 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PO Comments"
         Height          =   195
         Left            =   4050
         TabIndex        =   48
         Top             =   1410
         Width           =   1005
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice"
         Height          =   195
         Index           =   15
         Left            =   225
         TabIndex        =   45
         Top             =   360
         Width           =   525
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor"
         Height          =   195
         Index           =   14
         Left            =   240
         TabIndex        =   44
         Top             =   120
         Width           =   510
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inv Date"
         Height          =   195
         Index           =   8
         Left            =   165
         TabIndex        =   40
         Top             =   585
         Width           =   630
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Received"
         Height          =   195
         Index           =   7
         Left            =   4905
         TabIndex        =   38
         Top             =   105
         Width           =   660
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   32
         Top             =   915
         Width           =   585
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Terms"
         Height          =   195
         Index           =   4
         Left            =   360
         TabIndex        =   43
         Top             =   1500
         Width           =   435
      End
      Begin VB.Label lblTerms 
         BackColor       =   &H80000005&
         Height          =   465
         Left            =   840
         TabIndex        =   16
         Top             =   1530
         UseMnemonic     =   0   'False
         Width           =   2955
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount"
         Height          =   195
         Index           =   5
         Left            =   3000
         TabIndex        =   42
         Top             =   585
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tax"
         Height          =   195
         Index           =   3
         Left            =   3345
         TabIndex        =   41
         Top             =   345
         Width           =   270
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount"
         Height          =   195
         Index           =   6
         Left            =   4950
         TabIndex        =   39
         Top             =   585
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Payment"
         Height          =   195
         Index           =   9
         Left            =   4935
         TabIndex        =   37
         Top             =   345
         Width           =   630
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   195
         Index           =   2
         Left            =   3060
         TabIndex        =   36
         Top             =   105
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   11
         Left            =   4260
         TabIndex        =   35
         Top             =   1170
         Width           =   795
      End
      Begin VB.Label lblInvoiceCodes 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code1 / Code2"
         Height          =   195
         Left            =   3990
         TabIndex        =   34
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblJobDesc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   2910
         TabIndex        =   33
         Top             =   825
         Width           =   45
      End
      Begin VB.Label lblAddress 
         BackColor       =   &H80000005&
         Height          =   615
         Left            =   840
         TabIndex        =   15
         Top             =   915
         UseMnemonic     =   0   'False
         Width           =   2955
      End
      Begin VB.Label lblCompany 
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   840
         TabIndex        =   8
         Top             =   105
         UseMnemonic     =   0   'False
         Width           =   2145
      End
   End
   Begin VB.Menu mnuGrid 
      Caption         =   "<mnuGrid>"
      Visible         =   0   'False
      Begin VB.Menu mnuGridSub 
         Caption         =   "Select All"
         Index           =   0
      End
      Begin VB.Menu mnuGridSub 
         Caption         =   "Un-Select All"
         Index           =   1
      End
      Begin VB.Menu mnuGridSub 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuGridSub 
         Caption         =   "Sort Ascending"
         Index           =   3
      End
      Begin VB.Menu mnuGridSub 
         Caption         =   "Sort Descending"
         Index           =   4
      End
      Begin VB.Menu mnuGridSub 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuGridSub 
         Caption         =   "Remove this column"
         Index           =   6
      End
      Begin VB.Menu mnuColumns 
         Caption         =   "Insert a column"
         Index           =   7
         Begin VB.Menu mnuColumnsSub 
            Caption         =   "(none available)"
            Enabled         =   0   'False
            Index           =   0
         End
      End
   End
End
Attribute VB_Name = "FGenerateInvoices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private Const SRCFILE = "FGenerateInvoices::"

'these are set in LoadItems and used by the frmInvoice fields
Private mParentRow As Long
Private mDiscPercent  As Double
Private mDiscDays As Long
Private mTermsDays As Long
Private mTermsType As String


'grid menu
Private MouseGrid As VSFlexGrid
Private MouseCol  As Long

Private Const mcGRID_SELECTALL = 0
Private Const mcGRID_DESELECTALL = 1
Private Const mcGRID_ASC = 3
Private Const mcGRID_DESC = 4
Private Const mcGRID_HIDE = 6


Private Sub cmdNav_Click(Index As Integer)
On Error GoTo eh
    
    Select Case Index
        Case 0 'generate
            Call SaveData(Index = 1)
            
        Case 1 'close
            If SaveData(True) Then Unload Me

    End Select

Exit Sub
eh: Call ErrHandler(SRCFILE & "cmdNav_Click")
End Sub

Public Function ValidateData() As Boolean
    Dim s As String
    Dim i As Long
    Dim v As Double
    Dim xxx As String
        
    s = ""
    If txtAccountingDate.Visible And txtAccountingDate.Text = "" Then
        s = s & "Accounting date is required." & vbCrLf
    End If
    
    With gPOs
    For i = 1 To .Rows - 1
        If .RowOutlineLevel(i) = 0 And .Cell(flexcpChecked, i, .ColIndex("Add")) <> flexUnchecked Then
        
            If App.Options.Value(PaymentDateRequired) And .TextMatrix(i, .ColIndex("PaymentDate")) = "" Then
                s = s & "PO " & .TextMatrix(i, .ColIndex("PO")) & " payment date is required." & vbCrLf
            End If
            If App.Options.Value(ReceivedDateRequired) And .TextMatrix(i, .ColIndex("ReceivedDate")) = "" Then
                s = s & "PO " & .TextMatrix(i, .ColIndex("PO")) & " received date is required." & vbCrLf
            End If
            
            
            
            If .ValueMatrix(i, .ColIndex("InvoicePretax")) + .ValueMatrix(i, .ColIndex("InvoiceTax")) > .ValueMatrix(i, .ColIndex("RemainingPretax")) + .ValueMatrix(i, .ColIndex("RemainingTax")) Then
                xxx = ""
                xxx = xxx & "Unable to create the invoice as entered for po number " & .TextMatrix(i, .ColIndex("PO")) & "." & vbCrLf & vbCrLf
                xxx = xxx & "amount entered:   " & Format(.ValueMatrix(i, .ColIndex("InvoicePretax")) + .ValueMatrix(i, .ColIndex("InvoiceTax")), "#,##0.00") & vbCrLf
                xxx = xxx & "amount remaining:   " & Format(.ValueMatrix(i, .ColIndex("RemainingPretax")) + .ValueMatrix(i, .ColIndex("RemainingTax")), "#,##0.00") & vbCrLf
                xxx = xxx & "" & vbCrLf
                xxx = xxx & "Change the invoice to the remaining amount?" & vbCrLf
                If MsgBox(xxx, vbYesNo + vbQuestion, App.ProductName) = vbYes Then
                    Call SetInvoiceToRemaining(i)
                Else
                    s = s & "PO " & .TextMatrix(i, .ColIndex("PO")) & " is over-invoiced." & vbCrLf
                End If
                
            End If
        End If
    Next
    End With
        
    If s <> "" Then
        MsgBox "Unable to save" & vbCrLf & vbCrLf & s, vbExclamation, App.ProductName
    Else
        ValidateData = True
    End If
    
End Function

Private Sub SetInvoiceToRemaining(Row As Long)
On Error GoTo eh
    Dim i As Long
    Dim done As Boolean
    
    With gPOs
        done = False
        i = Row
        While .TextMatrix(Row, .ColIndex("PO")) = .Cell(flexcpData, i, .ColIndex("PO"))
            .Cell(flexcpText, i, .ColIndex("InvoicePretax")) = .Cell(flexcpText, i, .ColIndex("RemainingPretax"))
            .Cell(flexcpText, i, .ColIndex("InvoiceTax")) = .Cell(flexcpText, i, .ColIndex("RemainingTax"))
            i = i + 1
            If i = .Rows Then Exit Sub
        Wend
    End With
Exit Sub
eh: Call ErrHandler(SRCFILE & "SetInvoiceToRemaining")
End Sub


Public Function SaveData(prompt As Boolean) As Boolean
On Error GoTo eh
    
    Dim paypointsused As Boolean
    Dim p As Long
    Dim i As Long
    Dim s As String
    Dim Invoice As String
    Dim InvoiceID As Long
    Dim CreatedInvoices As String
    Dim d As String
    Dim AccountingDate As String
    Dim InvoiceDate As String
    Dim ReceivedDate As String
    Dim PODesc As String
    Dim RetainageRate As Double
    Dim DeptID As Long
    Dim Approver As String
    Dim rs As Recordset
    Dim RequireLienRelease As Boolean
    Dim WrapInsuranceExempt As Boolean
    
    Call UnloadItems
    gItems.Rows = 1
    
    If Not SomethingSelected Then
        SaveData = True
        Exit Function
    End If
    If prompt Then
        Select Case MsgBox("You have PO's selected for payment but have not yet generated the invoices." & vbCrLf & vbCrLf & "Do you want to generate them?" & vbCrLf, vbExclamation + vbYesNoCancel, App.ProductName)
            Case vbNo
                SaveData = True
                Exit Function
            Case vbCancel
                SaveData = False
                Exit Function
        End Select
    End If
    
    
    If Not ValidateData Then
        SaveData = False
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    With gPOs
    
    For i = 1 To .Rows - 1
        If Round(.ValueMatrix(i, .ColIndex("InvoicePreTax")), 2) <> 0 Then
            If .RowOutlineLevel(i) = 0 Then
            
                InvoiceDate = .Cell(flexcpText, i, .ColIndex("InvoiceDate"))
                If InvoiceDate = "" Then InvoiceDate = Now()
                
                AccountingDate = txtAccountingDate.Text
                If AccountingDate = "" Then AccountingDate = InvoiceDate
                
                ReceivedDate = .Cell(flexcpText, i, .ColIndex("ReceivedDate"))
                If ReceivedDate = "" Then ReceivedDate = Now()
                
                Invoice = GetInvoiceNumber(.Cell(flexcpData, i, .ColIndex("vendor")), .Cell(flexcpData, i, .ColIndex("po")), .Cell(flexcpText, i, .ColIndex("invoice")))
                
                'get department and approver
                s = ""
                s = s & "select top 1 d.deptid,isnull(nullif(j.pm,''),l.project_manager) pm,isnull(x.RequireLienRelease,0) RequireLienRelease" & vbCrLf
                s = s & "from pomaster p" & vbCrLf
                s = s & "join tbljobs j on p.job=j.job_no" & vbCrLf
                s = s & "left outer join tblpoindex x on p.poindex=x.poindex and p.divisionid=x.divisionid" & vbCrLf
                s = s & "left outer join tbllocality l on j.community=l.area" & vbCrLf
                s = s & "left outer join departmentapprovers da on da.pm=isnull(nullif(j.pm,''),l.project_manager)" & vbCrLf
                s = s & "left outer join departments d on da.deptid=d.deptid and d.divisionid=j.divisionid" & vbCrLf
                s = s & "where p.ponumber=" & DbQuote(Str, .Cell(flexcpData, i, .ColIndex("po"))) & vbCrLf
                s = s & "order by da.invoicelimit desc" & vbCrLf
                Set rs = HFApp.SqlExec(s)
                If rs.EOF Then
                    DeptID = 0
                    Approver = ""
                    RequireLienRelease = False
                Else
                    DeptID = Val("" & rs(0))
                    Approver = "" & rs(1)
                    RequireLienRelease = "" & rs(2) = "True"
                End If
                
                'insert invoice
                WrapInsuranceExempt = .Cell(flexcpChecked, i, .ColIndex("WrapInsuranceExempt")) = flexChecked
                RetainageRate = .ValueMatrix(i, .ColIndex("RetainageRate"))
                InvoiceID = WriteInvoice(0, RequireLienRelease, _
                                         .Cell(flexcpData, i, .ColIndex("vendor")), "", .Cell(flexcpText, i, .ColIndex("vendor")), "", "", "", "", "", _
                                         Invoice, "", "", "Pending", Approver, .Cell(flexcpText, i, .ColIndex("POComments")), DeptID, .Cell(flexcpValue, i, .ColIndex("InvoicePretax")), .Cell(flexcpValue, i, .ColIndex("InvoiceTax")), .Cell(flexcpText, i, .ColIndex("DiscountPretax")), _
                                         .Cell(flexcpValue, i, .ColIndex("InvoicePretax")) * .Cell(flexcpValue, i, .ColIndex("WrapInsuranceRate")), .Cell(flexcpValue, i, .ColIndex("WrapInsuranceRate")), .Cell(flexcpText, i, .ColIndex("DiscountDate")), ReceivedDate, InvoiceDate, .Cell(flexcpText, i, .ColIndex("PaymentDate")), AccountingDate, .Cell(flexcpText, i, .ColIndex("InvoiceDesc")), _
                                         .Cell(flexcpText, i, .ColIndex("MiscCode1")), .Cell(flexcpText, i, .ColIndex("MiscCode2")), "", False, 0)
                CreatedInvoices = CreatedInvoices & "," & InvoiceID
                
                
                'now create invoicepaypoints
                paypointsused = False
                s = ""
                s = s & "insert into invoicepaymentpoints(invoiceid,ponumber,paypoint1,paypoint2,paypoint3,paypoint4,paypoint5)" & vbCrLf
                s = s & "values(" & DbQuote(num, InvoiceID)
                s = s & "," & DbQuote(Str, .Cell(flexcpData, i, .ColIndex("po")))
                For p = 1 To 5
                    If .Cell(flexcpChecked, i, .ColIndex("PayPoint" & p)) = flexChecked Then
                        paypointsused = True
                        s = s & ",1"
                    Else
                        s = s & ",0"
                    End If
                Next
                s = s & ")"
                If paypointsused Then Call HFApp.SqlExec(s)
            Else
                Call WriteInvoiceItem(InvoiceID, 0, .Cell(flexcpData, i, .ColIndex("vendor")), Invoice, _
                                      .Cell(flexcpData, i, .ColIndex("vendor")), .Cell(flexcpData, i, .ColIndex("PO")), PODesc, .TextMatrix(i, .ColIndex("Line")), _
                                      .Cell(flexcpData, i, .ColIndex("Job")), "", .Cell(flexcpData, i, .ColIndex("Extra")), "", .Cell(flexcpData, i, .ColIndex("CostCode")), .Cell(flexcpText, i, .ColIndex("CostCodeDesc")), .Cell(flexcpData, i, .ColIndex("Category")), .Cell(flexcpText, i, .ColIndex("CategoryDesc")), .Cell(flexcpData, i, .ColIndex("DebitAccount")), "", _
                                      "", "", "", "", _
                                      .TextMatrix(i, .ColIndex("TaxGroup")), "", .ValueMatrix(i, .ColIndex("TaxRate")), _
                                      .TextMatrix(i, .ColIndex("CommittedQty")), .TextMatrix(i, .ColIndex("CommittedRate")), _
                                      .TextMatrix(i, .ColIndex("Qty")), .TextMatrix(i, .ColIndex("Rate")), _
                                      Round(.ValueMatrix(i, .ColIndex("InvoicePreTax")), 2), Round(.ValueMatrix(i, .ColIndex("InvoiceTax")), 2), _
                                      Round(.ValueMatrix(i, .ColIndex("InvoicePreTax")) * RetainageRate / 100, 2), RetainageRate, _
                                      .TextMatrix(i, .ColIndex("Description")), False, WrapInsuranceExempt, .TextMatrix(i, .ColIndex("JointPayee")))
            End If
            .RowHidden(i) = True
        End If
    Next
    For i = .Rows - 1 To 1 Step -1
        If .Cell(flexcpChecked, i, .ColIndex("Add")) <> flexUnchecked Then
            Call .RemoveItem(i)
        End If
    Next
        
    
    'finalize invoices, lookup descriptions, create hb invoices, approve
    CreatedInvoices = Mid(CreatedInvoices, 2)
    
    s = ""
    s = s & "update invoiceitems" & vbCrLf
    s = s & "set jobdesc = j.description" & vbCrLf
    s = s & "   ,phasedesc=cc.description" & vbCrLf
    s = s & "   ,categorydesc=c.description" & vbCrLf
    s = s & "   ,debitaccountdesc=gl.description" & vbCrLf
    s = s & "   ,taxgroupdesc=t.description   " & vbCrLf
    s = s & "from invoiceitems i" & vbCrLf
    s = s & "left outer join tbljobs             j on i.DivisionID = j.DivisionID and i.job=j.job_no" & vbCrLf
    s = s & "left outer join standardcostcodes  cc on i.DivisionID = cc.DivisionID and i.phase=cc.costcode" & vbCrLf
    s = s & "left outer join glaccounts         gl on i.DivisionID = gl.DivisionID and i.debitaccount=gl.account" & vbCrLf
    s = s & "left outer join standardcategories  c on i.DivisionID = c.DivisionID and i.category=c.category" & vbCrLf
    s = s & "left outer join taxgroups           t on i.DivisionID = t.DivisionID and i.taxgroup=t.taxgroup" & vbCrLf
    s = s & "where i.DivisionID =" & DbQuote(num, HFApp.DivisionID) & vbCrLf
    s = s & "  and i.invoiceid in(" & CreatedInvoices & ")" & vbCrLf
    Call HFApp.SqlExec(s, dbHomeFront)
    
    For i = 1 To Parse(CreatedInvoices)
        Call HFApp.SqlExec("exec dbo.Payables_CreateHBInvoice " & DbQuote(num, Parse(CreatedInvoices, i)))
    Next

    Call ApproveInvoices(CreatedInvoices)
    
    End With
    Screen.MousePointer = vbDefault

Exit Function
eh: Call ErrHandler(SRCFILE & "SaveData", s)
End Function



Public Sub Form_Load()
On Error GoTo eh
    
    Dim i As Long
    
    Call IniGetForm(Me)
    Call IniGetGrid(Me, gPOs, , , "new3")
    Call IniGetGrid(Me, gItems, , , "new3")
    gPOs.Rows = 1
    gItems.Rows = 1
    
    
    gPOs.ColPosition(gPOs.ColIndex("Add")) = 0
    lblPayPoints.Caption = ""
    gItems.ColHidden(gItems.ColIndex("PayPoint1")) = True
    gItems.ColHidden(gItems.ColIndex("PayPoint2")) = True
    gItems.ColHidden(gItems.ColIndex("PayPoint3")) = True
    gItems.ColHidden(gItems.ColIndex("PayPoint4")) = True
    gItems.ColHidden(gItems.ColIndex("PayPoint5")) = True
    
    
    
    txtInvoiceDate.Text = Format(VBA.Date, App.Options(TimberlineDateFormat))
    txtPaymentDate.Text = Format(VBA.Date, App.Options(TimberlineDateFormat))
    
    If Not IsIn(HFApp.Options(AccountingSystem), asIntacct, asTimberline) Then
        txtAccountingDate.Visible = False
        dteDiscount.Visible = False
        numDiscount.Visible = False
        Label1(5).Visible = False
        Label1(6).Visible = False
        lblAccountingDate.Visible = False
    
        gItems.ColHidden(gItems.ColIndex("JointPayee")) = True
        gItems.TextMatrix(0, gItems.ColIndex("JointPayee")) = ""
    Else
        If gItems.TextMatrix(0, gItems.ColIndex("JointPayee")) = "" Then gItems.TextMatrix(0, gItems.ColIndex("JointPayee")) = "Joint Payee"
    End If
    
    
    
    
    If IsIn(HFApp.Options(AccountingSystem), asQuickBooks, asQuickBooksOnline) Then
        txtVendor.Visible = False
        lblVendor.Caption = "Vendor"
        txtVendorDesc.left = txtVendor.left
        txtVendorDesc.Width = 2 * txtVendor.Width
    End If
    
    
    
    
    dteInvoice.Format = App.Options.Value(TimberlineDateFormat)
    dteInvoice.DisplayFormat = App.Options.Value(TimberlineDateFormat)
    dteReceived.Format = App.Options.Value(TimberlineDateFormat)
    dteReceived.DisplayFormat = App.Options.Value(TimberlineDateFormat)
    dtePayment.Format = App.Options.Value(TimberlineDateFormat)
    dtePayment.DisplayFormat = App.Options.Value(TimberlineDateFormat)
    dteDiscount.Format = App.Options.Value(TimberlineDateFormat)
    dteDiscount.DisplayFormat = App.Options.Value(TimberlineDateFormat)
    
    txtInvoiceCode1.Visible = App.Options.Value(InvoiceCode1Usage) <> "Hidden"
    txtInvoiceCode2.Visible = App.Options.Value(InvoiceCode2Usage) <> "Hidden"
    lblInvoiceCodes.Caption = App.Options.Value(InvoiceCode1Label) & IIf(App.Options.Value(InvoiceCode2Usage) <> "Hidden", "/" & App.Options.Value(InvoiceCode2Label), "")
    lblInvoiceCodes.Visible = App.Options.Value(InvoiceCode1Usage) <> "Hidden" Or App.Options.Value(InvoiceCode2Usage) <> "Hidden"
    
    lblPOComments.Visible = Not TimberlineAccounting
    lblPOComm.Visible = Not TimberlineAccounting
    
    
Exit Sub
eh: Call ErrHandler(SRCFILE & "Form_Load")
End Sub

Private Sub cmdSearch_Click()
    If SaveData(True) Then Call LoadData
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call IniPutForm(Me)
    Call IniPutGrid(Me, gPOs, , "new3")
    Call IniPutGrid(Me, gItems, , "new3")
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Const margin = 60


    Slider.Min = 960
    Slider.Max = Me.ScaleWidth - 960
    If Slider.left > Slider.Max Then Slider.left = Slider.Max
    
    FrameSearch.Width = Me.ScaleWidth
    
    
    
    Slider.Move Slider.left, FrameSearch.Height, Slider.Width, Me.ScaleHeight - FrameSearch.Height - cmdNav(0).Height - 2 * margin
    gPOs.Move 0, FrameSearch.Height, Slider.left, Slider.Height
    frmInvoice.Move Slider.left + Slider.Width, Slider.Top, Me.ScaleWidth - Slider.left - Slider.Width, Slider.Height
    gItems.Move 0, gItems.Top, frmInvoice.Width, frmInvoice.Height - gItems.Top
    
    
    cmdNav(0).Move Me.ScaleWidth - 2 * cmdNav(0).Width - 2 * margin, Me.ScaleHeight - cmdNav(0).Height - margin
    cmdNav(1).Move Me.ScaleWidth - cmdNav(0).Width - margin, Me.ScaleHeight - cmdNav(0).Height - margin
End Sub


Private Sub LoadData()
On Error GoTo eh
    Dim s  As String
    Dim rs As Recordset
    Dim r  As Long
    Dim i As Long
Dim t As Single
    Dim prevPO As String
    
    Dim terms As String
    Dim termdays As Long
    Dim d As Date
    
    Screen.MousePointer = vbHourglass
    
    With gPOs
        gItems.Rows = 1
        .Rows = 1
        .Redraw = flexRDNone
        
        If TimberlineAccounting Then
            s = ""
            s = s & "SELECT " & vbCrLf
            s = s & "   v.vendor      Vendor" & vbCrLf
            s = s & "  ,v.vname       VendorDesc" & vbCrLf
            s = s & "  ,v.vtype       VendorType" & vbCrLf
            s = s & "  ,v.vdscpct     DiscPercent" & vbCrLf
            s = s & "  ,v.vdscday     DiscDays" & vbCrLf
            s = s & "  ,v.vpmtday     NetDays" & vbCrLf
            s = s & "  ,v.vdaytyp     NetType" & vbCrLf
            s = s & "  ,v.vaddr1+'\n'+v.vaddr2+'\n'+v.vcity+', '+v.vstate   Address" & vbCrLf
            s = s & "  ,p.sub         PO" & vbCrLf
            s = s & "  ,p.sdesc       PODesc" & vbCrLf
            s = s & "  ,''            POComments" & vbCrLf
            s = s & "  ,p.sactsd      POStarted" & vbCrLf
            s = s & "  ,p.sactcd      POCompleted" & vbCrLf
            s = s & "  ,p.sjob        POJob" & vbCrLf
            s = s & "  ,pj.jdesc    POJobDesc" & vbCrLf
            s = s & "  ,l.item        Line" & vbCrLf
            s = s & "  ,l.idesc       LineDesc" & vbCrLf
            s = s & "  ,l.iamt-l.itxamt LinePretax" & vbCrLf
            s = s & "  ,l.itxamt      LineTax" & vbCrLf
            s = s & "  ,l.iappcoa-l.iaptcoa   LineChangesPretax" & vbCrLf
            s = s & "  ,l.iaptcoa           LineChangesTax" & vbCrLf
            s = s & "  ,round(l.iamtinv/(100+t.grate)*100,2)                      LineInvoicedPretax" & vbCrLf
            s = s & "  ,round(l.iamtinv - (round(l.iamtinv/(100+t.grate)*100,2)),2) LineInvoicedTax" & vbCrLf
            s = s & "  ,(l.iamt-l.itxamt) + (l.iappcoa-l.iaptcoa) - round(l.iamtinv/(100+t.grate)*100,2) LineRemainingPretax" & vbCrLf
            s = s & "  ,l.itxamt+l.iaptcoa-(l.iamtinv - (round(l.iamtinv/(100+t.grate)*100,2))) LineRemainingTax" & vbCrLf
            s = s & "  ,l.itxgrp   LineTaxGroup" & vbCrLf
            s = s & "  ,t.grate   LineTaxRate" & vbCrLf
            s = s & "  ,l.ijob      LineJob" & vbCrLf
            s = s & "  ,lj.jdesc  LineJobDesc" & vbCrLf
            s = s & "  ,l.iextra    LineExtra" & vbCrLf
            s = s & "  ,l.iphase    LineCostCode" & vbCrLf
            s = s & "  ,lc.spdesc    LineCostCodeDesc" & vbCrLf
            s = s & "  ,l.icat      LineCategory" & vbCrLf
            s = s & "  ,lcc.scdesc    LineCategoryDesc" & vbCrLf
            s = s & "  ,0     RetainageRate" & vbCrLf
            s = s & "  ,''    LineAccount" & vbCrLf
            s = s & "  ,''    LineAccountDesc" & vbCrLf
            
            s = s & "  ,ifnull(nullif(l.iunits,0),1) LineCommittedQuantity" & vbCrLf
            s = s & "  ,ifnull(nullif(l.iuntcst,0),l.iamt) LineCommittedUnitPrice" & vbCrLf
            s = s & "  ,ifnull(nullif(l.iunits,0),1)-l.iuntinv LineRemainingQuantity" & vbCrLf
            
            s = s & "  ,0     PayPoint1Percent" & vbCrLf
            s = s & "  ,0     PayPoint2Percent" & vbCrLf
            s = s & "  ,0     PayPoint3Percent" & vbCrLf
            s = s & "  ,0     PayPoint4Percent" & vbCrLf
            s = s & "  ,0     PayPoint5Percent" & vbCrLf
            s = s & "  ,0     PayPoint1" & vbCrLf
            s = s & "  ,0     PayPoint2" & vbCrLf
            s = s & "  ,0     PayPoint3" & vbCrLf
            s = s & "  ,0     PayPoint4" & vbCrLf
            s = s & "  ,0     PayPoint5" & vbCrLf
            s = s & "  ,v.vmdrate     WrapInsuranceRate" & vbCrLf
            s = s & "  ,0 WrapInsuranceExempt" & vbCrLf
            s = s & "FROM" & vbCrLf
            s = s & " master_jcm_record_12 p" & vbCrLf
            s = s & " INNER JOIN master_apm_record_9 v ON(p.svendor=v.vendor)" & vbCrLf
            s = s & " INNER JOIN master_jcm_record_13 l ON(p.sub=l.isub)" & vbCrLf
            s = s & " LEFT OUTER JOIN master_jcm_record_1_1 pj ON(pj.job=p.sjob)" & vbCrLf
            s = s & " INNER JOIN master_jcm_record_1_1 lj ON(lj.job=l.ijob)" & vbCrLf
            s = s & " INNER JOIN master_jcm_record_16 lc ON(l.iphase=lc.sphase)" & vbCrLf
            s = s & " INNER JOIN master_jcm_record_17 lcc ON(l.icat=lcc.scat)" & vbCrLf
            s = s & " LEFT OUTER JOIN master_txm_record_2 t ON(l.itxgrp=t.""group"")" & vbCrLf
            s = s & "WHERE p.sclosed=0 AND p.samt+p.sapprco-p.samtinv<>0" & vbCrLf
            If AccountingDB = dbTimberlinePVdata Then
                If txtVendor.Text <> "" Then s = s & " AND v.vendor = " & DbQuote(Str, FormatTSField(10, txtVendor.Text)) & vbCrLf
                If txtJob.Text <> "" Then s = s & "AND pj.job = " & DbQuote(Str, FormatTSField(10, txtJob.Text, True)) & vbCrLf
                If txtPO.Text <> "" Then s = s & "AND p.sub = " & DbQuote(Str, FormatTSField(12, txtPO.Text)) & vbCrLf
            Else
                If txtVendor.Text <> "" Then s = s & "AND v.vendor" & IIf(txtVendor.Text Like "*%*", " like ", " = ") & DbQuote(Str, txtVendor.Text) & vbCrLf
                If txtVendorDesc.Text <> "" Then s = s & "AND v.vname" & IIf(txtVendorDesc.Text Like "*%*", " like ", " = ") & DbQuote(Str, txtVendorDesc.Text) & vbCrLf
                If txtJob.Text <> "" Then s = s & "AND pj.job" & IIf(txtJob.Text Like "*%*", " like ", " = ") & DbQuote(Str, txtJob.Text) & vbCrLf
                If txtJobDesc.Text <> "" Then s = s & "AND pj.jdesc" & IIf(txtJobDesc.Text Like "*%*", " like ", " = ") & DbQuote(Str, txtJobDesc.Text) & vbCrLf
                If txtPO.Text <> "" Then s = s & "AND p.sub" & IIf(txtPO.Text Like "*%*", " like ", " = ") & DbQuote(Str, txtPO.Text) & vbCrLf
                If txtPODesc.Text <> "" Then s = s & "AND p.sdesc" & IIf(txtPODesc.Text Like "*%*", " like ", " = ") & DbQuote(Str, txtPODesc.Text) & vbCrLf
            End If
            s = s & "UNION ALL" & vbCrLf
            s = s & "SELECT" & vbCrLf
            s = s & "   xv.vendor      Vendor" & vbCrLf
            s = s & "  ,xv.vname       VendorDesc" & vbCrLf
            s = s & "  ,xv.vtype       VendorType" & vbCrLf
            s = s & "  ,xv.vdscpct     DiscPercent" & vbCrLf
            s = s & "  ,xv.vdscday     DiscDays" & vbCrLf
            s = s & "  ,xv.vpmtday     NetDays" & vbCrLf
            s = s & "  ,xv.vdaytyp     NetType" & vbCrLf
            s = s & "  ,xv.vaddr1+'\n'+xv.vaddr2+'\n'+xv.vcity+', '+xv.vstate   Address" & vbCrLf
            s = s & "  ,xp.sub         PO" & vbCrLf
            s = s & "  ,xp.sdesc       PODesc" & vbCrLf
            s = s & "  ,''             POComments" & vbCrLf
            s = s & "  ,xp.sactsd      POStarted" & vbCrLf
            s = s & "  ,xp.sactcd      POCompleted" & vbCrLf
            s = s & "  ,xp.sjob        POJob" & vbCrLf
            s = s & "  ,xpj.jdesc    POJobDesc" & vbCrLf
            s = s & "  ,xl.item        Line" & vbCrLf
            s = s & "  ,xl.idesc       LineDesc" & vbCrLf
            s = s & "  ,xl.iamt-xl.itxamt LinePretax" & vbCrLf
            s = s & "  ,xl.itxamt      LineTax" & vbCrLf
            s = s & "  ,xl.iappcoa-xl.iaptcoa   LineChangesPretax" & vbCrLf
            s = s & "  ,xl.iaptcoa           LineChangesTax" & vbCrLf
            s = s & "  ,xl.iamtinv           LineInvoicedPretax" & vbCrLf
            s = s & "  ,0.0               LineInvoicedTax" & vbCrLf
            s = s & "  ,(xl.iamt-xl.itxamt) + (xl.iappcoa-xl.iaptcoa) - xl.iamtinv LineRemainingPretax" & vbCrLf
            s = s & "  ,0.0           LineRemainingTax" & vbCrLf
            s = s & "  ,xl.itxgrp  LineTaxGroup" & vbCrLf
            s = s & "  ,xt.grate   LineTaxRate" & vbCrLf
            s = s & "  ,xl.ijob      LineJob" & vbCrLf
            s = s & "  ,xlj.jdesc  LineJobDesc" & vbCrLf
            s = s & "  ,xl.iextra    LineExtra" & vbCrLf
            s = s & "  ,xl.iphase    LineCostCode" & vbCrLf
            s = s & "  ,xlc.spdesc    LineCostCodeDesc" & vbCrLf
            s = s & "  ,xl.icat      LineCategory" & vbCrLf
            s = s & "  ,xlcc.scdesc    LineCategoryDesc" & vbCrLf
            s = s & "  ,0     RetainageRate" & vbCrLf
            s = s & "  ,''    LineAccount" & vbCrLf
            s = s & "  ,''    LineAccountDesc" & vbCrLf
            
            s = s & "  ,ifnull(nullif(l.iunits,0),1) LineCommittedQty" & vbCrLf
            s = s & "  ,ifnull(nullif(l.iuntcst,0),l.iamt) LineCommittedUnitPrice" & vbCrLf
            s = s & "  ,ifnull(nullif(l.iunits,0),1)-l.iuntinv LineRemainingQuantity" & vbCrLf
            
            s = s & "  ,0     PayPoint1Percent" & vbCrLf
            s = s & "  ,0     PayPoint2Percent" & vbCrLf
            s = s & "  ,0     PayPoint3Percent" & vbCrLf
            s = s & "  ,0     PayPoint4Percent" & vbCrLf
            s = s & "  ,0     PayPoint5Percent" & vbCrLf
            s = s & "  ,0     PayPoint1" & vbCrLf
            s = s & "  ,0     PayPoint2" & vbCrLf
            s = s & "  ,0     PayPoint3" & vbCrLf
            s = s & "  ,0     PayPoint4" & vbCrLf
            s = s & "  ,0     PayPoint5" & vbCrLf
            s = s & "  ,xv.vmdrate     WrapInsuranceRate" & vbCrLf
            s = s & "  ,0 WrapInsuranceExempt" & vbCrLf
            s = s & "FROM" & vbCrLf
            s = s & " master_jcm_record_12 xp" & vbCrLf
            s = s & " INNER JOIN master_apm_record_9 xv ON(xp.svendor=xv.vendor)" & vbCrLf
            s = s & " INNER JOIN master_jcm_record_13 xl ON(xp.sub=xl.isub)" & vbCrLf
            s = s & " LEFT OUTER JOIN master_jcm_record_1_1 xpj ON(xpj.job=xp.sjob)" & vbCrLf
            s = s & " INNER JOIN master_jcm_record_1_1 xlj ON(xlj.job=xl.ijob)" & vbCrLf
            s = s & " INNER JOIN master_jcm_record_16 xlc ON(xl.iphase=xlc.sphase)" & vbCrLf
            s = s & " INNER JOIN master_jcm_record_17 xlcc ON(xl.icat=xlcc.scat)" & vbCrLf
            s = s & " LEFT OUTER JOIN master_txm_record_2 xt ON(xl.itxgrp=xt.""group"")" & vbCrLf
            s = s & "WHERE xt.""group"" is null and xp.sclosed=0 AND xp.samt+xp.sapprco-xp.samtinv<>0" & vbCrLf
            If AccountingDB = dbTimberlinePVdata Then
                If txtVendor.Text <> "" Then s = s & "AND xv.vendor = " & DbQuote(Str, FormatTSField(10, txtVendor.Text)) & vbCrLf
                If txtJob.Text <> "" Then s = s & "AND xpj.job = " & DbQuote(Str, FormatTSField(10, txtJob.Text, True)) & vbCrLf
                If txtPO.Text <> "" Then s = s & "AND xp.sub = " & DbQuote(Str, FormatTSField(12, txtPO.Text)) & vbCrLf
            Else
                If txtVendor.Text <> "" Then s = s & "AND xv.vendor" & IIf(txtVendor.Text Like "*%*", " like ", " = ") & DbQuote(Str, txtVendor.Text) & vbCrLf
                If txtVendorDesc.Text <> "" Then s = s & "AND xv.vname" & IIf(txtVendorDesc.Text Like "*%*", " like ", " = ") & DbQuote(Str, txtVendorDesc.Text) & vbCrLf
                If txtJob.Text <> "" Then s = s & "AND xpj.job" & IIf(txtJob.Text Like "*%*", " like ", " = ") & DbQuote(Str, txtJob.Text) & vbCrLf
                If txtJobDesc.Text <> "" Then s = s & "AND xpj.jdesc" & IIf(txtJobDesc.Text Like "*%*", " like ", " = ") & DbQuote(Str, txtJobDesc.Text) & vbCrLf
                If txtPO.Text <> "" Then s = s & "AND xp.sub" & IIf(txtPO.Text Like "*%*", " like ", " = ") & DbQuote(Str, txtPO.Text) & vbCrLf
                If txtPODesc.Text <> "" Then s = s & "AND xp.sdesc" & IIf(txtPODesc.Text Like "*%*", " like ", " = ") & DbQuote(Str, txtPODesc.Text) & vbCrLf
            End If
            s = s & "ORDER BY 9,1,3,6,14" & vbCrLf
        Else
                              
            s = ""
            s = s & "select p.*" & vbCrLf
            s = s & " ,isnull(v.miscDeductionRate,0) WrapInsuranceRate" & vbCrLf
            s = s & " ,isnull(c.WrapInsuranceExempt,0) WrapInsuranceExempt" & vbCrLf
            s = s & "from jcpodetails p" & vbCrLf
            s = s & "left join tblvendors v on p.divisionid=v.divisionid and p.vendor=v.vendor_id" & vbCrLf
            s = s & "left join tblJobs j on p.divisionid=j.divisionid and p.pojob=j.job_no" & vbCrLf
            s = s & "left join tbllocality c on j.community=c.area" & vbCrLf
            s = s & "where isnull(p.TBDVendor,0)=0 and p.pocancelled<>1 and p.DivisionID=" & HFApp.DivisionID & vbCrLf
            s = s & "and p.LineRemainingPretax<>0" & vbCrLf
            
            If HFApp.Options(AccountingSystem) = asTimberline Then
                s = s & "and p.postingbatch<>0" & vbCrLf
            End If
            
            If txtVendor.Visible And txtVendor.Text <> "" Then s = s & "and p.vendor" & IIf(txtVendor.Text Like "*%*", " like ", " = ") & DbQuote(Str, txtVendor.Text) & vbCrLf
            
            If txtVendorDesc.Text <> "" Then s = s & "and p.vendordesc" & IIf(txtVendorDesc.Text Like "*%*", " like ", " = ") & DbQuote(Str, txtVendorDesc.Text) & vbCrLf
            If txtJob.Text <> "" Then s = s & "and p.pojob" & IIf(txtJob.Text Like "*%*", " like ", " = ") & DbQuote(Str, txtJob.Text) & vbCrLf
            If txtJobDesc.Text <> "" Then s = s & "and p.linejobdesc" & IIf(txtJobDesc.Text Like "*%*", " like ", " = ") & DbQuote(Str, txtJobDesc.Text) & vbCrLf
            If txtPO.Text <> "" Then s = s & "and p.po" & IIf(txtPO.Text Like "*%*", " like ", " = ") & DbQuote(Str, txtPO.Text) & vbCrLf
            If txtPODesc.Text <> "" Then s = s & "and p.podesc" & IIf(txtPODesc.Text Like "*%*", " like ", " = ") & DbQuote(Str, txtPODesc.Text) & vbCrLf
            
            If HFApp.Options.ValueByName("SortInvoiceByJobExtra") = "true" Then
                s = s & "order by p.po,p.linejob,p.lineextra,p.linecostcode,p.linecategory" & vbCrLf
            Else
                s = s & "order by p.po,p.line" & vbCrLf
            End If
            
        End If
        Set rs = HFApp.SqlExec(s, AccountingDB)
        
        
        r = 0
        prevPO = ""
        
        While Not rs.EOF
            If prevPO <> "" & rs("po") Then
                prevPO = "" & rs("po")
                r = r + 1
                .AddItem ""
                .Cell(flexcpData, r, .ColIndex("Vendor")) = "" & rs("Vendor")
                .Cell(flexcpData, r, .ColIndex("PO")) = "" & rs("PO")
                .Cell(flexcpData, r, .ColIndex("Job")) = "" & rs("POJob")
                
                .Cell(flexcpText, r, .ColIndex("Vendor")) = "" & rs("VendorDesc")
                .Cell(flexcpText, r, .ColIndex("VendorType")) = "" & rs("VendorType")
                .Cell(flexcpText, r, .ColIndex("Address")) = Replace(Replace("" & rs("Address"), "\n\n", "\n"), "\n", vbCrLf)
                
                .Cell(flexcpText, r, .ColIndex("WrapInsuranceRate")) = Val("" & rs("WrapInsuranceRate")) / 100
                .Cell(flexcpChecked, r, .ColIndex("WrapInsuranceExempt")) = IIf("" & rs("WrapInsuranceExempt") = "True", flexChecked, flexUnchecked)
                
                .Cell(flexcpText, r, .ColIndex("DiscPercent")) = "" & rs("DiscPercent")
                
                .Cell(flexcpText, r, .ColIndex("DiscDays")) = "" & rs("DiscDays")
                .Cell(flexcpText, r, .ColIndex("TermsDays")) = "" & rs("NetDays")
                .Cell(flexcpText, r, .ColIndex("TermsType")) = "" & rs("NetType")
            
            
                terms = ""
                If Val("" & rs("DiscPercent")) <> 0 Then
                    terms = terms & Val("" & rs("DiscPercent")) & "% discount"
                    If Val("" & rs("DiscDays")) <> 0 Then
                        terms = terms & " if paid in " & Val("" & rs("DiscDays")) & Choose(Min(Val("" & rs("DiscDays")), 2), " day", " days")
                    End If
                    terms = terms & vbCrLf
                End If
                If Val("" & rs("NetDays")) <> 0 Then
                    If "" & rs("NetType") = "Number of days" Then
                        terms = terms & "Net due in " & Val("" & rs("NetDays")) & Choose(Min(Val("" & rs("NetDays")), 2), " day", " days") & vbCrLf
                    Else
                        terms = terms & "Net due on the " & Val("" & rs("NetDays")) & Choose(Min(Val("" & rs("NetDays")), 4), "st", "nd", "rd", "th") & vbCrLf
                    End If
                End If
                If Right(terms, 2) = vbCrLf Then terms = Mid(terms, 1, Len(terms) - 2)
                .Cell(flexcpText, r, .ColIndex("PaymentTerms")) = terms
                
                
                .Cell(flexcpText, r, .ColIndex("PO")) = "" & rs("po")
                
                
                
                
                If App.Options(ReceivedDateRequired) Or Not App.Options(CalcDatesFromInvoiceDate) Then
                    .Cell(flexcpText, r, .ColIndex("ReceivedDate")) = txtInvoiceDate.Text
                End If
                
                
                termdays = Val("" & rs("NetDays"))
                If termdays <> 0 And Not App.Options.Value(DontSetPaymentDate) Then
                    d = VBA.Date
                    Select Case "" & rs("NetType")
                        Case "Number of days"
                            .Cell(flexcpText, r, .ColIndex("PaymentDate")) = DateAdd("d", termdays, d)
                        Case "Day of month"
                            If d >= DateSerial(Year(d), Month(d), termdays) Then
                                d = DateSerial(Year(d), Month(d) + 1, termdays)
                            Else
                                d = DateSerial(Year(d), Month(d), termdays)
                            End If
                            .Cell(flexcpText, r, .ColIndex("PaymentDate")) = d
                        Case "Day of next month"
                            .Cell(flexcpText, r, .ColIndex("PaymentDate")) = DateSerial(Year(d), Month(d) + 1, termdays)
                    End Select
                End If
                
                
                If Val("" & rs("DiscDays")) <> 0 Then
                    .Cell(flexcpText, r, .ColIndex("DiscountDate")) = VBA.Date + Val("" & rs("DiscDays"))
                End If
                
                
                .Cell(flexcpText, r, .ColIndex("RetainageRate")) = Val("" & rs("RetainageRate"))
                .Cell(flexcpText, r, .ColIndex("POStarted")) = "" & rs("POStarted")
                .Cell(flexcpText, r, .ColIndex("POCompleted")) = "" & rs("POCompleted")
                .Cell(flexcpText, r, .ColIndex("Description")) = "" & rs("PODesc")
                .Cell(flexcpText, r, .ColIndex("POComments")) = "" & rs("POComments")
                
                If App.Options.Value(InvDescDefaultsToJobAddr) Then
                    .Cell(flexcpText, r, .ColIndex("InvoiceDesc")) = "" & rs("POJobDesc")
                Else
                    .Cell(flexcpText, r, .ColIndex("InvoiceDesc")) = "" & rs("PODesc")
                End If
                
                .Cell(flexcpText, r, .ColIndex("Job")) = "" & rs("POJobDesc")
                
                .Cell(flexcpText, r, .ColIndex("InvoicePretax")) = 0
                .Cell(flexcpText, r, .ColIndex("InvoiceTax")) = 0
                
                
                
                .Cell(flexcpChecked, r, 0) = flexUnchecked
                
                Call FormatAmounts(gPOs, r)
                .IsSubtotal(r) = True
                .RowOutlineLevel(r) = 0
            End If
            
            r = r + 1
            .AddItem ""
            .Cell(flexcpData, r, .ColIndex("Vendor")) = "" & rs("Vendor")
            .Cell(flexcpData, r, .ColIndex("PO")) = "" & rs("PO")
            .Cell(flexcpData, r, .ColIndex("Job")) = "" & rs("LineJob")
            .Cell(flexcpData, r, .ColIndex("Extra")) = "" & rs("LineExtra")
            .Cell(flexcpData, r, .ColIndex("CostCode")) = "" & rs("LineCostCode")
            .Cell(flexcpData, r, .ColIndex("Category")) = "" & rs("LineCategory")
            .Cell(flexcpData, r, .ColIndex("DebitAccount")) = "" & rs("LineAccount")
            
            .Cell(flexcpText, r, .ColIndex("Rate")) = "" & rs("LineCommittedUnitPrice")
            .Cell(flexcpText, r, .ColIndex("CommittedRate")) = "" & rs("LineCommittedUnitPrice")
            .Cell(flexcpText, r, .ColIndex("Qty")) = "" & rs("LineRemainingQuantity")
            .Cell(flexcpText, r, .ColIndex("CommittedQty")) = "" & rs("LineCommittedQuantity")
            
            .Cell(flexcpText, r, .ColIndex("CostCodeDesc")) = "" & rs("LineCostCodeDesc")
            .Cell(flexcpText, r, .ColIndex("CategoryDesc")) = "" & rs("LineCategoryDesc")
            .Cell(flexcpText, r, .ColIndex("Line")) = "" & rs("Line")
            .Cell(flexcpText, r, .ColIndex("Description")) = "" & rs("LineDesc")
            
            .Cell(flexcpText, r, .ColIndex("OriginalPretax")) = Val("" & rs("LinePretax"))
            .Cell(flexcpText, r, .ColIndex("OriginalTax")) = Val("" & rs("LineTax"))
            .Cell(flexcpText, r, .ColIndex("ChangesPretax")) = Val("" & rs("LineChangesPretax"))
            .Cell(flexcpText, r, .ColIndex("ChangesTax")) = Val("" & rs("LineChangesTax"))
            .Cell(flexcpText, r, .ColIndex("Pretax")) = Val("" & rs("LinePretax")) + Val("" & rs("LineChangesPretax"))
            .Cell(flexcpText, r, .ColIndex("Tax")) = Val("" & rs("LineTax")) + Val("" & rs("LineChangesTax"))
            .Cell(flexcpText, r, .ColIndex("TaxGroup")) = "" & rs("LineTaxGroup")
            .Cell(flexcpText, r, .ColIndex("TaxRate")) = "" & rs("LineTaxRate")
            
            .Cell(flexcpText, r, .ColIndex("InvoicedPretax")) = Val("" & rs("LineInvoicedPretax"))
            .Cell(flexcpText, r, .ColIndex("InvoicedTax")) = Val("" & rs("LineInvoicedTax"))
            .Cell(flexcpText, r, .ColIndex("RemainingPretax")) = Val("" & rs("LineRemainingPretax"))
            .Cell(flexcpText, r, .ColIndex("RemainingTax")) = Val("" & rs("LineRemainingTax"))
            
            .Cell(flexcpText, r, .ColIndex("PayPoint1Percent")) = "" & rs("PayPoint1Percent")
            .Cell(flexcpText, r, .ColIndex("PayPoint2Percent")) = "" & rs("PayPoint2Percent")
            .Cell(flexcpText, r, .ColIndex("PayPoint3Percent")) = "" & rs("PayPoint3Percent")
            .Cell(flexcpText, r, .ColIndex("PayPoint4Percent")) = "" & rs("PayPoint4Percent")
            .Cell(flexcpText, r, .ColIndex("PayPoint5Percent")) = "" & rs("PayPoint5Percent")
            .Cell(flexcpText, r, .ColIndex("PayPoint1")) = "" & rs("PayPoint1")
            .Cell(flexcpText, r, .ColIndex("PayPoint2")) = "" & rs("PayPoint2")
            .Cell(flexcpText, r, .ColIndex("PayPoint3")) = "" & rs("PayPoint3")
            .Cell(flexcpText, r, .ColIndex("PayPoint4")) = "" & rs("PayPoint4")
            .Cell(flexcpText, r, .ColIndex("PayPoint5")) = "" & rs("PayPoint5")
            
            
            
            
            .Cell(flexcpChecked, r, 0) = flexUnchecked
            Call FormatAmounts(gPOs, r)
            .IsSubtotal(r) = True
            .RowOutlineLevel(r) = 1
            rs.MoveNext
        Wend
        
        
        Call SetInitialTotals

        
        .Outline 0
        .OutlineBar = flexOutlineBarSimple
        .OutlineCol = -1
        
        .ColWidth(0) = 555
        
        .Redraw = flexRDBuffered
        Screen.MousePointer = vbDefault
        
        
        mParentRow = 0
        Call .SetFocus
        
        
    End With
Exit Sub
eh:
    Call ErrHandler(SRCFILE & "LoadData", s)
    gPOs.Redraw = flexRDBuffered
    Screen.MousePointer = vbDefault
End Sub


Private Sub CalcInvoiceTotal(CalcDiscount As Boolean)
    Dim Pretax As Double
    Dim Tax As Double
    Dim Discount As Double
    Dim i As Long
    
    With gItems
        Pretax = 0
        Tax = 0
        For i = 1 To .Rows - 1
            Tax = Tax + .Cell(flexcpValue, i, .ColIndex("InvoiceTax"))
            Pretax = Pretax + .Cell(flexcpValue, i, .ColIndex("InvoicePreTax"))
        Next
        numAmount = Pretax + Tax
        numTax = Tax
    
        If CalcDiscount And mDiscPercent <> 0 Then
             Discount = Round(Pretax * mDiscPercent / 100, 2)
             numDiscount = Discount
        End If
        
    End With
    
    With gPOs
        .TextMatrix(mParentRow, .ColIndex("InvoicePretax")) = Pretax
        .TextMatrix(mParentRow, .ColIndex("InvoiceTax")) = Tax
        .TextMatrix(mParentRow, .ColIndex("DiscountPretax")) = Discount
        
        .Cell(flexcpChecked, mParentRow, .ColIndex("Add")) = IIf(Pretax <> 0, flexChecked, flexUnchecked)
    End With
    
End Sub






Private Sub gPOs_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call UnloadItems
    Call LoadItems(NewRow)
End Sub

Private Sub gPOs_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gPOs
    .AutoSearch = flexSearchNone
    .EditMaxLength = 0
    
    Cancel = .RowOutlineLevel(Row) <> 0
    If Cancel Then Exit Sub
    
    Select Case .ColKey(Col)
        Case "Invoice"
             .EditMaxLength = 15
       
        Case "InvoiceDate", "Add"
            
        Case Else
            .AutoSearch = flexSearchFromCursor
            Cancel = True
    
    End Select
    End With
End Sub



Private Sub gPOs_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then Call ShowColumnMenu(gPOs, False)
End Sub


Private Sub ShowColumnMenu(Grid As VSFlexGrid, Optional Sortable As Boolean = True)
On Error GoTo eh
    Dim i As Long
    Dim j As Long

    'save this stuff for menu click
    Set MouseGrid = Grid
    MouseCol = Grid.MouseCol

    'set these
    mnuGridSub(mcGRID_ASC).Enabled = Sortable
    mnuGridSub(mcGRID_DESC).Enabled = Sortable
    mnuGridSub(mcGRID_HIDE).Enabled = MouseCol >= 0

    'load Grid column names
    mnuColumnsSub(0).Visible = True
    For i = mnuColumnsSub.UBound To 1 Step -1
        Unload mnuColumnsSub(i)
    Next
    For i = 0 To MouseGrid.Cols - 1
        If MouseGrid.ColHidden(i) And MouseGrid.TextMatrix(0, i) <> "" Then
            j = j + 1
            Load mnuColumnsSub(j)
            mnuColumnsSub(j).tag = MouseGrid.ColKey(i)
            mnuColumnsSub(j).Caption = MouseGrid.TextMatrix(0, i)
            mnuColumnsSub(j).Visible = True
            mnuColumnsSub(j).Enabled = True
        End If
    Next
    If j = 0 Then mnuColumnsSub(0).Caption = "(none available)"
    mnuColumnsSub(0).Visible = j = 0
    mnuColumnsSub(0).Enabled = False

    'show menu
    PopupMenu mnuGrid

    Exit Sub
eh: Call ErrHandler(SRCFILE & "ShowColumnMenu")
End Sub





Private Sub mnuColumnsSub_Click(Index As Integer)
    MouseGrid.ColHidden(MouseGrid.ColIndex(mnuColumnsSub(Index).tag)) = False
    MouseGrid.ColPosition(MouseGrid.ColIndex(mnuColumnsSub(Index).tag)) = IIf(MouseCol < 0, MouseGrid.Cols - 1, MouseCol)

    gPOs.ColPosition(gPOs.ColIndex("Add")) = 0
    MouseGrid.ColHidden(0) = False

End Sub
Private Sub mnuGridSub_Click(Index As Integer)
    Dim i As Long
    Select Case Index
    
        Case mcGRID_SELECTALL
            For i = 1 To MouseGrid.Rows - 1
                gPOs.Cell(flexcpChecked, i, gPOs.ColIndex("Add")) = flexChecked
                Call gPOs_AfterEdit(i, gPOs.ColIndex("Add"))
            Next
        
        Case mcGRID_DESELECTALL
            For i = 1 To MouseGrid.Rows - 1
                gPOs.Cell(flexcpChecked, i, gPOs.ColIndex("Add")) = flexUnchecked
                Call gPOs_AfterEdit(i, gPOs.ColIndex("Add"))
            Next
    
        Case mcGRID_ASC
            MouseGrid.Col = MouseCol
            MouseGrid.Sort = flexSortGenericAscending

        Case mcGRID_DESC
            MouseGrid.Col = MouseCol
            MouseGrid.Sort = flexSortGenericDescending

        Case mcGRID_HIDE
            For i = 0 To MouseGrid.Cols - 1
                If MouseGrid.ColHidden(i) = False Then
                    MouseGrid.ColHidden(MouseCol) = True
                    Exit Sub
                End If
            Next
    End Select
End Sub

Private Sub lblJob_Click()
On Error Resume Next
    If FPickList.Choose(HFApp.Databases(AccountingDB), App.Options(Caption_Job), SelectJob, txtJob.Text) Then
        txtJob.Text = FPickList.SelectedItem(1)
    End If
End Sub

Private Sub lblPO_Click()
On Error Resume Next
    If FPickList.Choose(HFApp.Databases(AccountingDB), App.Options(Caption_Commitment), SelectPO, txtPO.Text) Then
        txtPO.Text = FPickList.SelectedItem(1)
    End If
End Sub



Private Sub lblVendor_Click()
On Error Resume Next
    If FPickList.Choose(HFApp.Databases(dbHomeFront), App.Options(Caption_Vendor), SelectVendor) Then
        txtVendor.Text = FPickList.SelectedItem("vendor")
        txtVendorDesc.Text = FPickList.SelectedItem("name")
    End If
End Sub


Private Sub FormatAmounts(g, r As Long)
On Error Resume Next
    With g
        .Cell(flexcpForeColor, r, .ColIndex("OriginalPretax")) = IIf(.ValueMatrix(r, .ColIndex("OriginalPretax")) < 0, vbRed, vbWindowText)
        .Cell(flexcpForeColor, r, .ColIndex("ChangesPretax")) = IIf(.ValueMatrix(r, .ColIndex("ChangesPretax")) < 0, vbRed, vbWindowText)
        .Cell(flexcpForeColor, r, .ColIndex("Pretax")) = IIf(.ValueMatrix(r, .ColIndex("Pretax")) < 0, vbRed, vbWindowText)
        .Cell(flexcpForeColor, r, .ColIndex("InvoicedPretax")) = IIf(.ValueMatrix(r, .ColIndex("InvoicedPretax")) < 0, vbRed, vbWindowText)
        .Cell(flexcpForeColor, r, .ColIndex("RemainingPretax")) = IIf(.ValueMatrix(r, .ColIndex("RemainingPretax")) < 0, vbRed, vbWindowText)
        .Cell(flexcpForeColor, r, .ColIndex("InvoicePretax")) = IIf(.ValueMatrix(r, .ColIndex("InvoicePretax")) < 0, vbRed, vbWindowText)
        .Cell(flexcpForeColor, r, .ColIndex("DiscountPretax")) = IIf(.ValueMatrix(r, .ColIndex("DiscountPretax")) < 0, vbRed, vbWindowText)
    
        .Cell(flexcpForeColor, r, .ColIndex("OriginalTax")) = IIf(.ValueMatrix(r, .ColIndex("OriginalTax")) < 0, vbRed, vbWindowText)
        .Cell(flexcpForeColor, r, .ColIndex("ChangesTax")) = IIf(.ValueMatrix(r, .ColIndex("ChangesTax")) < 0, vbRed, vbWindowText)
        .Cell(flexcpForeColor, r, .ColIndex("Tax")) = IIf(.ValueMatrix(r, .ColIndex("Tax")) < 0, vbRed, vbWindowText)
        .Cell(flexcpForeColor, r, .ColIndex("InvoicedTax")) = IIf(.ValueMatrix(r, .ColIndex("InvoicedTax")) < 0, vbRed, vbWindowText)
        .Cell(flexcpForeColor, r, .ColIndex("RemainingTax")) = IIf(.ValueMatrix(r, .ColIndex("RemainingTax")) < 0, vbRed, vbWindowText)
        .Cell(flexcpForeColor, r, .ColIndex("InvoiceTax")) = IIf(.ValueMatrix(r, .ColIndex("InvoiceTax")) < 0, vbRed, vbWindowText)
    End With
End Sub






Private Sub SetInitialTotals()
    Dim i As Long
    Dim r As Long
    
    Dim cOrgP As Long
    Dim cOrgT As Long
    Dim cChgP As Long
    Dim cChgT As Long
    Dim cAmtP As Long
    Dim cAmtT As Long
    Dim cInvP As Long
    Dim cInvT As Long
    Dim cRemP As Long
    Dim cRemT As Long
    
    With gPOs
        cOrgP = .ColIndex("OriginalPretax")
        cChgP = .ColIndex("ChangesPretax")
        cAmtP = .ColIndex("Pretax")
        cInvP = .ColIndex("InvoicedPretax")
        cRemP = .ColIndex("RemainingPretax")
        
        cOrgT = .ColIndex("OriginalTax")
        cChgT = .ColIndex("ChangesTax")
        cAmtT = .ColIndex("Tax")
        cInvT = .ColIndex("InvoicedTax")
        cRemT = .ColIndex("RemainingTax")
        
        For i = 1 To .Rows - 1
            If .RowOutlineLevel(i) = 0 Then
                .Cell(flexcpText, i, cOrgP) = 0
                .Cell(flexcpText, i, cChgP) = 0
                .Cell(flexcpText, i, cAmtP) = 0
                .Cell(flexcpText, i, cInvP) = 0
                .Cell(flexcpText, i, cRemP) = 0
                .Cell(flexcpText, i, cOrgT) = 0
                .Cell(flexcpText, i, cChgT) = 0
                .Cell(flexcpText, i, cAmtT) = 0
                .Cell(flexcpText, i, cInvT) = 0
                .Cell(flexcpText, i, cRemT) = 0
            Else
                r = .GetNodeRow(i, flexNTParent)
                .Cell(flexcpText, r, cOrgP) = .Cell(flexcpValue, i, cOrgP) + .Cell(flexcpValue, r, cOrgP)
                .Cell(flexcpText, r, cChgP) = .Cell(flexcpValue, i, cChgP) + .Cell(flexcpValue, r, cChgP)
                .Cell(flexcpText, r, cAmtP) = .Cell(flexcpValue, i, cAmtP) + .Cell(flexcpValue, r, cAmtP)
                .Cell(flexcpText, r, cInvP) = .Cell(flexcpValue, i, cInvP) + .Cell(flexcpValue, r, cInvP)
                .Cell(flexcpText, r, cRemP) = .Cell(flexcpValue, i, cRemP) + .Cell(flexcpValue, r, cRemP)
            
                .Cell(flexcpText, r, cOrgT) = .Cell(flexcpValue, i, cOrgT) + .Cell(flexcpValue, r, cOrgT)
                .Cell(flexcpText, r, cChgT) = .Cell(flexcpValue, i, cChgT) + .Cell(flexcpValue, r, cChgT)
                .Cell(flexcpText, r, cAmtT) = .Cell(flexcpValue, i, cAmtT) + .Cell(flexcpValue, r, cAmtT)
                .Cell(flexcpText, r, cInvT) = .Cell(flexcpValue, i, cInvT) + .Cell(flexcpValue, r, cInvT)
                .Cell(flexcpText, r, cRemT) = .Cell(flexcpValue, i, cRemT) + .Cell(flexcpValue, r, cRemT)
            
            
                
            End If
        Next
    
    
    
    End With



End Sub









Private Function GetInvoiceNumber(Vendor As String, po As String, Invoice As String) As String
    Dim s  As String
    Dim r As Long
    Dim Number As String
    
    
    Number = ""
    
    'check local
    s = ""
    s = s & "select invoice " & vbCrLf
    s = s & "  from invoices" & vbCrLf
    s = s & " where DivisionID =" & HFApp.DivisionID & " and vendor=" & DbQuote(Str, Vendor) & vbCrLf
    s = s & "   and invoice like " & DbQuote(Str, IIf(Invoice = "", po, Invoice) & "%") & vbCrLf
    s = s & "order by 1 desc"
    On Error Resume Next
    Number = Max(Number, HFApp.SqlExec(s)(0))
    On Error GoTo 0
    
    
    If HFApp.Options(AccountingSystem) = asTimberline Then
        s = ""
        s = s & "select oiinv" & vbCrLf
        s = s & "  from new_api_record_1" & vbCrLf
        s = s & " where oivnd=" & DbQuote(Str, Vendor) & vbCrLf
        s = s & "   and oiinv like " & DbQuote(Str, IIf(Invoice = "", po, Invoice) & "%") & vbCrLf
        s = s & "order by 1 desc"
        On Error Resume Next
        Number = Max(Number, HFApp.SqlExec(s, dbAccountingDictionary)(0))
        On Error GoTo 0
            
        s = ""
        s = s & "select oiinv" & vbCrLf
        s = s & "  from master_apm_record_1" & vbCrLf
        s = s & " where oivnd=" & DbQuote(Str, Vendor) & vbCrLf
        s = s & "   and oiinv like " & DbQuote(Str, IIf(Invoice = "", po, Invoice) & "%") & vbCrLf
        s = s & "order by 1 desc"
        On Error Resume Next
        Number = Max(Number, HFApp.SqlExec(s, dbAccountingDictionary)(0))
        On Error GoTo 0
    End If
    
    If Number = "" Then
        Number = IIf(Invoice = "", po, Invoice)
    Else
        If Parse(Number, , ".") = 1 Then
            Number = Number & ".1"
        Else
            Number = IIf(Invoice = "", po, Invoice) & "." & Val(Parse(Number, Parse(Number, , "."), ".")) + 1
        End If
    End If
    
    
    
    GetInvoiceNumber = left(Number, MaxInvoiceLength())
    
        
End Function



Private Sub Slider_Move()
    Call Form_Resize
End Sub

Private Sub txtAccountingDate_GotFocus()
    SelectAll txtAccountingDate
End Sub

Private Sub txtAccountingDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 And Shift = 0 Then
        DCalendar.Popup txtAccountingDate
    End If
End Sub

Private Sub txtAccountingDate_Validate(Cancel As Boolean)
    If txtAccountingDate.Text <> "" Then
        If IsDate(txtAccountingDate.Text) Then
            txtAccountingDate.Text = Format(txtAccountingDate.Text, App.Options(TimberlineDateFormat))
        Else
            Cancel = True
        End If
    End If
End Sub



Private Sub gPOs_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim n As VSFlexNode

    Dim parentRow As Long
    Dim childRow As Long
    Dim p As Double
    Dim i As Long
    
    With gPOs
        .Redraw = flexRDNone
    
        'one of Add, invoice,invoicedate have changed.
        If .ColKey(Col) = "Add" Then
            
            parentRow = Row
            childRow = .GetNodeRow(parentRow, flexNTFirstChild)
            While childRow > 0
            
                'set checks on line items
                .Cell(flexcpChecked, childRow, .ColIndex("PayPoint1")) = .Cell(flexcpChecked, parentRow, .ColIndex("Add"))
                .Cell(flexcpChecked, childRow, .ColIndex("PayPoint2")) = .Cell(flexcpChecked, parentRow, .ColIndex("Add"))
                .Cell(flexcpChecked, childRow, .ColIndex("PayPoint3")) = .Cell(flexcpChecked, parentRow, .ColIndex("Add"))
                .Cell(flexcpChecked, childRow, .ColIndex("PayPoint4")) = .Cell(flexcpChecked, parentRow, .ColIndex("Add"))
                .Cell(flexcpChecked, childRow, .ColIndex("PayPoint5")) = .Cell(flexcpChecked, parentRow, .ColIndex("Add"))
                
                'set/remove values
                If .Cell(flexcpChecked, parentRow, .ColIndex("Add")) = flexChecked Then
                    .Cell(flexcpText, childRow, .ColIndex("InvoicePretax")) = .Cell(flexcpText, childRow, .ColIndex("remainingPretax"))
                    .Cell(flexcpText, childRow, .ColIndex("InvoiceTax")) = .Cell(flexcpText, childRow, .ColIndex("remainingtax"))
                Else
                    .Cell(flexcpText, childRow, .ColIndex("InvoicePretax")) = 0
                    .Cell(flexcpText, childRow, .ColIndex("InvoiceTax")) = 0
                End If
                Call FormatAmounts(gPOs, childRow)
                
                
                childRow = .GetNodeRow(childRow, flexNTNextSibling)
            Wend
        End If
           
        .Redraw = flexRDBuffered
    End With

    Call LoadItems(Row)

End Sub



Private Sub gPOs_AfterMoveColumn(ByVal Col As Long, Position As Long)
    gPOs.ColPosition(gPOs.ColIndex("Add")) = 0
End Sub

Private Sub gPOs_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = Col = gPOs.ColIndex("Add")
End Sub


Private Sub gPOs_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim parentRow As Long
    Dim childRow As Long
    
    Dim i As Long
    Dim p As Double
    
    With gPOs
    Select Case .ColKey(Col)
        
        Case "Invoice"
            .Cell(flexcpChecked, Row, .ColIndex("Add")) = flexChecked
        
        
        Case "InvoiceDate"
            If .EditText <> "" Then
                If IsDate(.EditText) Then
                    .EditText = Format(.EditText, App.Options(TimberlineDateFormat))
                    .Cell(flexcpChecked, Row, .ColIndex("Add")) = flexChecked
                Else
                    Cancel = True
                End If
            End If
         
            
            
    End Select
    End With
End Sub



Private Sub SetPayPointCheckboxes(Row As Long, ColKey As String)
    Dim PercentPaid As Double
    Dim p1 As Double
    Dim p2 As Double
    Dim p3 As Double
    Dim p4 As Double
    Dim p5 As Double
    Dim s As String
    Dim i As Long
    Dim c As Long
    
    If Row < 1 Then Exit Sub
    
    With gItems
        
        'get paypoints percents
        p1 = .ValueMatrix(Row, .ColIndex("PayPoint1Percent"))
        p2 = .ValueMatrix(Row, .ColIndex("PayPoint2Percent"))
        p3 = .ValueMatrix(Row, .ColIndex("PayPoint3Percent"))
        p4 = .ValueMatrix(Row, .ColIndex("PayPoint4Percent"))
        p5 = .ValueMatrix(Row, .ColIndex("PayPoint5Percent"))

        s = p1
        If p2 <> 0 Then s = s & " " & p2
        If p3 <> 0 Then s = s & " " & p3
        If p4 <> 0 Then s = s & " " & p4
        If p5 <> 0 Then s = s & " " & p5
        s = s & "%"
        If p2 = 0 Then s = ""
        lblPayPoints.Caption = s
        
        
        
        'show/hide paypoint checkbox columns
        '.ColHidden(.ColIndex("PayPoint1")) = p2 = 0 'hide 1 if 2 is zero. Yes this is on purpose.
        .ColHidden(.ColIndex("PayPoint1")) = False 'changed to expose the checkbox for 100% paypoints too
        
        .ColHidden(.ColIndex("PayPoint2")) = p2 = 0 'also hide 2 if 2 is zero
        .ColHidden(.ColIndex("PayPoint3")) = p3 = 0
        .ColHidden(.ColIndex("PayPoint4")) = p4 = 0
        .ColHidden(.ColIndex("PayPoint5")) = p5 = 0

        'which p did the user change?
        If ColKey Like "PayPoint*" Then
            i = Right(ColKey, 1)
            'if updated col is checked ensure everything less than me is also checked
            If .Cell(flexcpChecked, Row, .ColIndex(ColKey)) = flexChecked Then
            For c = i To 1 Step -1
                .Cell(flexcpChecked, Row, .ColIndex("PayPoint" & c)) = flexChecked
            Next
            End If
            
            'if updated col is UNchecked ensure everything greater than me is also unchecked
            If .Cell(flexcpChecked, Row, .ColIndex(ColKey)) = flexUnchecked Then
            For c = i To 5
                .Cell(flexcpChecked, Row, .ColIndex("PayPoint" & c)) = flexUnchecked
            Next
            End If
            
        End If
        
    
        'get percent paid for this row
        If .ValueMatrix(Row, .ColIndex("Pretax")) = 0 Then
            PercentPaid = 100
        Else
            PercentPaid = .ValueMatrix(Row, .ColIndex("InvoicedPretax")) / .ValueMatrix(Row, .ColIndex("Pretax")) * 100
        End If
    
        'set checkboxes on this row
        If PercentPaid >= p1 Then .Cell(flexcpChecked, Row, .ColIndex("PayPoint1")) = flexTSGrayed
        If PercentPaid >= p1 + p2 Then .Cell(flexcpChecked, Row, .ColIndex("PayPoint2")) = flexTSGrayed
        If PercentPaid >= p1 + p2 + p3 Then .Cell(flexcpChecked, Row, .ColIndex("PayPoint3")) = flexTSGrayed
        If PercentPaid >= p1 + p2 + p3 + p4 Then .Cell(flexcpChecked, Row, .ColIndex("PayPoint4")) = flexTSGrayed
        
        'spacebar cycles tristate checkboxes thru checked-graychecked-unchecked potentially leaving in a gray state when it should be unchecked
        'check each box, if is gray but percent paid is not enough then change to unchecked.
        If .Cell(flexcpChecked, Row, .ColIndex("PayPoint1")) = flexTSGrayed And PercentPaid < p1 Then .Cell(flexcpChecked, Row, .ColIndex("PayPoint1")) = flexUnchecked
        If .Cell(flexcpChecked, Row, .ColIndex("PayPoint2")) = flexTSGrayed And PercentPaid < p1 + p2 Then .Cell(flexcpChecked, Row, .ColIndex("PayPoint2")) = flexUnchecked
        If .Cell(flexcpChecked, Row, .ColIndex("PayPoint3")) = flexTSGrayed And PercentPaid < p1 + p2 + p3 Then .Cell(flexcpChecked, Row, .ColIndex("PayPoint3")) = flexUnchecked
        If .Cell(flexcpChecked, Row, .ColIndex("PayPoint4")) = flexTSGrayed And PercentPaid < p1 + p2 + p3 + p4 Then .Cell(flexcpChecked, Row, .ColIndex("PayPoint4")) = flexUnchecked
        If .Cell(flexcpChecked, Row, .ColIndex("PayPoint5")) = flexTSGrayed And PercentPaid < p1 + p2 + p3 + p4 + p5 Then .Cell(flexcpChecked, Row, .ColIndex("PayPoint5")) = flexUnchecked
    
        'ensure that if p1 is UNchecked then so is p2 and p3 etc.
        If .Cell(flexcpChecked, Row, .ColIndex("PayPoint1")) = flexUnchecked Then .Cell(flexcpChecked, Row, .ColIndex("PayPoint2")) = flexUnchecked
        If .Cell(flexcpChecked, Row, .ColIndex("PayPoint2")) = flexUnchecked Then .Cell(flexcpChecked, Row, .ColIndex("PayPoint3")) = flexUnchecked
        If .Cell(flexcpChecked, Row, .ColIndex("PayPoint3")) = flexUnchecked Then .Cell(flexcpChecked, Row, .ColIndex("PayPoint4")) = flexUnchecked
        If .Cell(flexcpChecked, Row, .ColIndex("PayPoint4")) = flexUnchecked Then .Cell(flexcpChecked, Row, .ColIndex("PayPoint5")) = flexUnchecked
        
        'ensure that if p4 is checked then so is p3 and p2 etc.
        If .Cell(flexcpChecked, Row, .ColIndex("PayPoint5")) = flexChecked Then If .Cell(flexcpChecked, Row, .ColIndex("PayPoint4")) <> flexTSGrayed Then .Cell(flexcpChecked, Row, .ColIndex("PayPoint4")) = flexChecked
        If .Cell(flexcpChecked, Row, .ColIndex("PayPoint4")) = flexChecked Then If .Cell(flexcpChecked, Row, .ColIndex("PayPoint3")) <> flexTSGrayed Then .Cell(flexcpChecked, Row, .ColIndex("PayPoint3")) = flexChecked
        If .Cell(flexcpChecked, Row, .ColIndex("PayPoint3")) = flexChecked Then If .Cell(flexcpChecked, Row, .ColIndex("PayPoint2")) <> flexTSGrayed Then .Cell(flexcpChecked, Row, .ColIndex("PayPoint2")) = flexChecked
        If .Cell(flexcpChecked, Row, .ColIndex("PayPoint2")) = flexChecked Then If .Cell(flexcpChecked, Row, .ColIndex("PayPoint1")) <> flexTSGrayed Then .Cell(flexcpChecked, Row, .ColIndex("PayPoint1")) = flexChecked
    
        'if all rows are checked or unchecked then header should be same
        For c = 1 To 5
            .Cell(flexcpChecked, 0, .ColIndex("PayPoint" & c)) = flexChecked
            For i = 1 To .Rows - 1
                If .Cell(flexcpChecked, i, .ColIndex("PayPoint" & c)) = flexUnchecked Then
                    .Cell(flexcpChecked, 0, .ColIndex("PayPoint" & c)) = flexUnchecked
                    Exit For
                End If
            Next
                        
        Next
    
    
    End With
    
End Sub



Private Sub LoadItems(Row As Long)
    'selected PO from gPOs to gItems
    Dim childRow As Long
    Dim itemRow  As Long
    Dim c As Long
    Dim s As String
    
    With gPOs
            
        mParentRow = Row
        
        
        frmInvoice.Enabled = Row > 0
        
        If Row < 1 Then
            lblCompany.tag = ""
            lblCompany.Caption = ""
            txtInvoice.Text = ""
            dteInvoice = ""
            numDiscount.Text = ""
            lblAddress.Caption = ""
            lblTerms.Caption = ""
            dteReceived = ""
            dtePayment = ""
            dteDiscount = ""
            txtInvoiceCode1.Text = ""
            txtInvoiceCode2.Text = ""
            txtDescription.Text = ""
            lblPOComments.Caption = ""
            gItems.Rows = 1
            Exit Sub
        End If
        
        If .Cell(flexcpText, Row, .ColIndex("InvoiceDate")) = "" Then .Cell(flexcpText, Row, .ColIndex("InvoiceDate")) = txtInvoiceDate.Text
        If .Cell(flexcpText, Row, .ColIndex("PaymentDate")) = "" Then .Cell(flexcpText, Row, .ColIndex("PaymentDate")) = txtPaymentDate.Text
                
        
        mDiscPercent = .ValueMatrix(Row, .ColIndex("DiscPercent"))
        mDiscDays = .ValueMatrix(Row, .ColIndex("DiscDays"))
        mTermsDays = .ValueMatrix(Row, .ColIndex("TermsDays"))
        mTermsType = .TextMatrix(Row, .ColIndex("TermsType"))
        
        lblCompany.tag = .Cell(flexcpData, Row, .ColIndex("Vendor"))
        lblCompany.Caption = .TextMatrix(Row, .ColIndex("Vendor"))
        txtInvoice.Text = .TextMatrix(Row, .ColIndex("Invoice"))
        dteInvoice = .TextMatrix(Row, .ColIndex("InvoiceDate"))
        numDiscount = .TextMatrix(Row, .ColIndex("DiscountPretax"))
        lblAddress.Caption = .TextMatrix(Row, .ColIndex("Address"))
        lblTerms.Caption = .TextMatrix(Row, .ColIndex("PaymentTerms"))
        dteReceived = .TextMatrix(Row, .ColIndex("ReceivedDate"))
        dtePayment = .TextMatrix(Row, .ColIndex("PaymentDate"))
        dteDiscount = .TextMatrix(Row, .ColIndex("DiscountDate"))
        txtInvoiceCode1.Text = .TextMatrix(Row, .ColIndex("MiscCode1"))
        txtInvoiceCode2.Text = .TextMatrix(Row, .ColIndex("MiscCode2"))
        txtDescription.Text = .TextMatrix(Row, .ColIndex("InvoiceDesc"))
        lblPOComments.Caption = .TextMatrix(Row, .ColIndex("POComments"))
            
        gItems.Redraw = flexRDNone
        gItems.Rows = 1
        
        
        itemRow = 0
        childRow = .GetNodeRow(Row, flexNTFirstChild)
        
        'set paypoint checkboxes in the header
        For c = 1 To 5
            gItems.Cell(flexcpChecked, 0, gItems.ColIndex("PayPoint" & c)) = flexUnchecked
        Next
        
        
        While childRow > 0
        
            gItems.AddItem ""
            itemRow = itemRow + 1
             
            gItems.RowData(itemRow) = childRow
             
            
            gItems.Cell(flexcpText, itemRow, gItems.ColIndex("Line")) = .Cell(flexcpText, childRow, .ColIndex("Line"))
            gItems.Cell(flexcpText, itemRow, gItems.ColIndex("Description")) = .Cell(flexcpText, childRow, .ColIndex("Description"))
            gItems.Cell(flexcpText, itemRow, gItems.ColIndex("OriginalPretax")) = .Cell(flexcpText, childRow, .ColIndex("OriginalPretax"))
            gItems.Cell(flexcpText, itemRow, gItems.ColIndex("OriginalTax")) = .Cell(flexcpText, childRow, .ColIndex("OriginalTax"))
            gItems.Cell(flexcpText, itemRow, gItems.ColIndex("ChangesPretax")) = .Cell(flexcpText, childRow, .ColIndex("ChangesPretax"))
            gItems.Cell(flexcpText, itemRow, gItems.ColIndex("ChangesTax")) = .Cell(flexcpText, childRow, .ColIndex("ChangesTax"))
            gItems.Cell(flexcpText, itemRow, gItems.ColIndex("Pretax")) = .Cell(flexcpText, childRow, .ColIndex("Pretax"))
            gItems.Cell(flexcpText, itemRow, gItems.ColIndex("Tax")) = .Cell(flexcpText, childRow, .ColIndex("Tax"))
            gItems.Cell(flexcpText, itemRow, gItems.ColIndex("InvoicedPretax")) = .Cell(flexcpText, childRow, .ColIndex("InvoicedPretax"))
            gItems.Cell(flexcpText, itemRow, gItems.ColIndex("InvoicedTax")) = .Cell(flexcpText, childRow, .ColIndex("InvoicedTax"))
            gItems.Cell(flexcpText, itemRow, gItems.ColIndex("RemainingPretax")) = .Cell(flexcpText, childRow, .ColIndex("RemainingPretax"))
            gItems.Cell(flexcpText, itemRow, gItems.ColIndex("RemainingTax")) = .Cell(flexcpText, childRow, .ColIndex("RemainingTax"))
            gItems.Cell(flexcpText, itemRow, gItems.ColIndex("InvoicePretax")) = .Cell(flexcpText, childRow, .ColIndex("InvoicePretax"))
            gItems.Cell(flexcpText, itemRow, gItems.ColIndex("InvoiceTax")) = .Cell(flexcpText, childRow, .ColIndex("InvoiceTax"))
            gItems.Cell(flexcpText, itemRow, gItems.ColIndex("TaxGroup")) = .Cell(flexcpText, childRow, .ColIndex("TaxGroup"))
            gItems.Cell(flexcpText, itemRow, gItems.ColIndex("TaxRate")) = .Cell(flexcpText, childRow, .ColIndex("TaxRate"))
            
            gItems.Cell(flexcpText, itemRow, gItems.ColIndex("Job")) = .Cell(flexcpData, childRow, .ColIndex("Job"))
            gItems.Cell(flexcpText, itemRow, gItems.ColIndex("Extra")) = .Cell(flexcpData, childRow, .ColIndex("Extra"))
            gItems.Cell(flexcpText, itemRow, gItems.ColIndex("CostCode")) = .Cell(flexcpData, childRow, .ColIndex("CostCode"))
            gItems.Cell(flexcpText, itemRow, gItems.ColIndex("Category")) = .Cell(flexcpData, childRow, .ColIndex("Category"))
            gItems.Cell(flexcpText, itemRow, gItems.ColIndex("CostCodeDesc")) = .Cell(flexcpText, childRow, .ColIndex("CostCodeDesc"))
            gItems.Cell(flexcpText, itemRow, gItems.ColIndex("CategoryDesc")) = .Cell(flexcpText, childRow, .ColIndex("CategoryDesc"))
            
            gItems.Cell(flexcpText, itemRow, gItems.ColIndex("Rate")) = .Cell(flexcpText, childRow, .ColIndex("Rate"))
            gItems.Cell(flexcpText, itemRow, gItems.ColIndex("Qty")) = .Cell(flexcpText, childRow, .ColIndex("Qty"))
            gItems.Cell(flexcpData, itemRow, gItems.ColIndex("Rate")) = .Cell(flexcpText, childRow, .ColIndex("CommittedRate"))
            gItems.Cell(flexcpData, itemRow, gItems.ColIndex("Qty")) = .Cell(flexcpText, childRow, .ColIndex("CommittedQty"))
            gItems.Cell(flexcpText, itemRow, gItems.ColIndex("JointPayee")) = .Cell(flexcpText, childRow, .ColIndex("JointPayee"))
            
            
            gItems.Cell(flexcpData, itemRow, gItems.ColIndex("Job")) = .Cell(flexcpData, childRow, .ColIndex("Job"))
            gItems.Cell(flexcpData, itemRow, gItems.ColIndex("Extra")) = .Cell(flexcpData, childRow, .ColIndex("Extra"))
            gItems.Cell(flexcpData, itemRow, gItems.ColIndex("CostCode")) = .Cell(flexcpData, childRow, .ColIndex("CostCode"))
            gItems.Cell(flexcpData, itemRow, gItems.ColIndex("Category")) = .Cell(flexcpData, childRow, .ColIndex("Category"))
            
            
            For c = 1 To 5
                gItems.Cell(flexcpChecked, itemRow, gItems.ColIndex("PayPoint" & c)) = .Cell(flexcpChecked, childRow, .ColIndex("PayPoint" & c))
                gItems.Cell(flexcpText, itemRow, gItems.ColIndex("PayPoint" & c)) = .Cell(flexcpText, childRow, .ColIndex("PayPoint" & c))
                gItems.Cell(flexcpText, itemRow, gItems.ColIndex("PayPoint" & c & "Percent")) = .Cell(flexcpText, childRow, .ColIndex("PayPoint" & c & "Percent"))
            Next
            
            Call SetPayPointCheckboxes(itemRow, "")
            Call FormatAmounts(gItems, itemRow)
            
            childRow = .GetNodeRow(childRow, flexNTNextSibling)
        Wend
    End With
    
    If gItems.Rows > 1 Then gItems.Select 1, 0
    gItems.Redraw = flexRDBuffered
    
    Call CalcInvoiceTotal(False)
End Sub
Private Sub UnloadItems()
    'move editted PO from gItems back to gPOs
    
    Dim i As Long
    Dim c As Long
    
    With gPOs
        If mParentRow < 1 Then Exit Sub
        If mParentRow > .Rows - 1 Then Exit Sub
            
        
        .TextMatrix(mParentRow, .ColIndex("Invoice")) = txtInvoice.Text
        .TextMatrix(mParentRow, .ColIndex("DiscountPretax")) = numDiscount.Text
        
        .TextMatrix(mParentRow, .ColIndex("InvoiceDate")) = "" & dteInvoice
        .TextMatrix(mParentRow, .ColIndex("ReceivedDate")) = "" & dteReceived
        .TextMatrix(mParentRow, .ColIndex("PaymentDate")) = "" & dtePayment
        .TextMatrix(mParentRow, .ColIndex("DiscountDate")) = "" & dteDiscount
        
        .TextMatrix(mParentRow, .ColIndex("MiscCode1")) = txtInvoiceCode1.Text
        .TextMatrix(mParentRow, .ColIndex("MiscCode2")) = txtInvoiceCode2.Text
        .TextMatrix(mParentRow, .ColIndex("InvoiceDesc")) = txtDescription.Text
            
        For i = 1 To gItems.Rows - 1
            
            On Error Resume Next
            c = Val(gItems.RowData(i))
            If c = 0 Then Exit Sub
            On Error GoTo 0
             
            .Cell(flexcpChecked, c, .ColIndex("PayPoint1")) = gItems.Cell(flexcpChecked, i, gItems.ColIndex("PayPoint1"))
            .Cell(flexcpChecked, c, .ColIndex("PayPoint2")) = gItems.Cell(flexcpChecked, i, gItems.ColIndex("PayPoint2"))
            .Cell(flexcpChecked, c, .ColIndex("PayPoint3")) = gItems.Cell(flexcpChecked, i, gItems.ColIndex("PayPoint3"))
            .Cell(flexcpChecked, c, .ColIndex("PayPoint4")) = gItems.Cell(flexcpChecked, i, gItems.ColIndex("PayPoint4"))
            .Cell(flexcpChecked, c, .ColIndex("PayPoint5")) = gItems.Cell(flexcpChecked, i, gItems.ColIndex("PayPoint5"))
            
            .Cell(flexcpText, c, .ColIndex("InvoicePretax")) = gItems.Cell(flexcpText, i, gItems.ColIndex("InvoicePretax"))
            .Cell(flexcpText, c, .ColIndex("InvoiceTax")) = gItems.Cell(flexcpText, i, gItems.ColIndex("InvoiceTax"))
            
            .Cell(flexcpText, c, .ColIndex("Rate")) = gItems.Cell(flexcpText, i, gItems.ColIndex("Rate"))
            .Cell(flexcpText, c, .ColIndex("Qty")) = gItems.Cell(flexcpText, i, gItems.ColIndex("Qty"))
            .Cell(flexcpText, c, .ColIndex("JointPayee")) = gItems.Cell(flexcpText, i, gItems.ColIndex("JointPayee"))
            
            
        Next
    End With
End Sub

Private Sub txtInvoiceCode1_Validate(Cancel As Boolean)
On Error GoTo eh

    If App.Options(InvoiceCode1Usage) = "Required" And txtInvoiceCode1.Text = "" Then
        MsgBox App.Options(InvoiceCode1Label) & " is required", vbExclamation, App.ProductName
        Cancel = True
        Exit Sub
    End If
    
    Select Case left(App.Options(InvoiceCode1Validate), 4)
        Case "None"
        Case "Stri"
            If InvoiceCodeExists(lblCompany.tag, txtInvoice.Text, 1, txtInvoiceCode1.Text) Then
                MsgBox App.Options(InvoiceCode1Label) & " """ & txtInvoiceCode1.Text & """ has already been entered.", vbExclamation, App.ProductName
                Cancel = True
                Exit Sub
            End If
        Case "Warn"
            If InvoiceCodeExists(lblCompany.tag, txtInvoice.Text, 1, txtInvoiceCode1.Text) Then
                If vbCancel = MsgBox(App.Options(InvoiceCode1Label) & " """ & txtInvoiceCode1.Text & """ has already been entered.", vbOKCancel + vbExclamation, App.ProductName) Then
                    Cancel = True
                    Exit Sub
                End If
            End If
    End Select
    
Exit Sub
eh: Call ErrHandler(SRCFILE & "txtInvoiceCode1_Validate")
End Sub

Private Sub txtInvoiceCode2_Validate(Cancel As Boolean)
On Error GoTo eh

    If App.Options(InvoiceCode2Usage) = "Required" And txtInvoiceCode2.Text = "" Then
        MsgBox App.Options(InvoiceCode2Label) & " is required", vbExclamation, App.ProductName
        Cancel = True
        Exit Sub
    End If
    
    Select Case left(App.Options(InvoiceCode2Validate), 4)
        Case "None"
        Case "Stri"
            If InvoiceCodeExists(lblCompany.tag, txtInvoice.Text, 2, txtInvoiceCode2.Text) Then
                MsgBox App.Options(InvoiceCode2Label) & " """ & txtInvoiceCode2.Text & """ has already been entered.", vbExclamation, App.ProductName
                Cancel = True
                Exit Sub
            End If
        Case "Warn"
            If InvoiceCodeExists(lblCompany.tag, txtInvoice.Text, 2, txtInvoiceCode2.Text) Then
                If vbCancel = MsgBox(App.Options(InvoiceCode2Label) & " """ & txtInvoiceCode2.Text & """ has already been entered.", vbOKCancel + vbInformation, App.ProductName) Then
                    Cancel = True
                    Exit Sub
                End If
            End If
    End Select
    
Exit Sub
eh: Call ErrHandler(SRCFILE & "txtInvoiceCode2_Validate")
End Sub



Private Sub dteInvoice_Change()
On Error Resume Next
    Dim d As Date
    
    If Not IsDate(dteInvoice) Then Exit Sub
    If Not App.Options(CalcDatesFromInvoiceDate) Then Exit Sub
    
    'discount date
    If mDiscPercent > 0 Then
        If mDiscDays > 0 Then
            d = dteInvoice
            d = d + mDiscDays
            dteDiscount = d
        Else
            dteDiscount = ""
        End If
    Else
        dteDiscount = ""
    End If
    
    'payment date
    If mTermsDays > 0 And Not App.Options(DontSetPaymentDate) Then
        d = dteInvoice
        Select Case mTermsType
        
            Case "Number of days"
                dtePayment = DateAdd("d", mTermsDays, d)
                
            Case "Day of month"
                If d >= DateSerial(Year(d), Month(d), mTermsDays) Then
                    d = DateSerial(Year(d), Month(d) + 1, mTermsDays)
                Else
                    d = DateSerial(Year(d), Month(d), mTermsDays)
                End If
                dtePayment = d
                
            Case "Day of next month"
                dtePayment = DateSerial(Year(d), Month(d) + 1, mTermsDays)
        
        End Select
    End If

End Sub

Private Sub dteReceived_Change()
On Error Resume Next
    Dim d As Date
    
    If Not IsDate(dteReceived) Then Exit Sub
    If App.Options(CalcDatesFromInvoiceDate) Then Exit Sub
    
    'discount date
    If mDiscPercent > 0 Then
        If mDiscDays > 0 Then
            d = dteReceived
            d = d + mDiscDays
            dteDiscount = d
        Else
            dteDiscount = ""
        End If
    Else
        dteDiscount = ""
    End If
    
    'payment date
    If mTermsDays > 0 And Not App.Options(DontSetPaymentDate) Then
        d = dteReceived
        Select Case mTermsType
        
            Case "Number of days"
                dtePayment = DateAdd("d", mTermsDays, d)
                
            Case "Day of month"
                If d >= DateSerial(Year(d), Month(d), mTermsDays) Then
                    d = DateSerial(Year(d), Month(d) + 1, mTermsDays)
                Else
                    d = DateSerial(Year(d), Month(d), mTermsDays)
                End If
                dtePayment = d
                
            Case "Day of next month"
                dtePayment = DateSerial(Year(d), Month(d) + 1, mTermsDays)
                
        End Select
    End If


End Sub




Private Sub txtInvoiceDate_Validate(Cancel As Boolean)
    If txtInvoiceDate.Text <> "" Then
        If IsDate(txtInvoiceDate.Text) Then
            txtInvoiceDate.Text = Format(txtInvoiceDate.Text, App.Options(TimberlineDateFormat))
        Else
            Cancel = True
        End If
    End If
End Sub


Private Sub txtPaymentDate_Validate(Cancel As Boolean)
    If txtPaymentDate.Text <> "" Then
        If IsDate(txtPaymentDate.Text) Then
            txtPaymentDate.Text = Format(txtPaymentDate.Text, App.Options(TimberlineDateFormat))
        Else
            Cancel = True
        End If
    End If
End Sub


Private Sub ApproveInvoices(InvoiceIDs As String)
    Dim s As String
    
    Dim total As Long
    Dim approvable As Long
    
    If Not App.ApproveInvoices Then Exit Sub
    
    'count approvable invoices
    s = ""
    s = s & "select count(*) from invoices where invoiceid in(" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "select" & vbCrLf
    s = s & "   i.invoiceid" & vbCrLf
    s = s & "from invoices i" & vbCrLf
    s = s & "join invoiceitems d on i.invoiceid=d.invoiceid" & vbCrLf
    s = s & "where i.status<>'Lien Hold' and i.invoiceid in (" & InvoiceIDs & ")" & vbCrLf
    s = s & "group by i.invoiceid,i.pretax,i.tax" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "having " & vbCrLf
    s = s & "(" & vbCrLf
    s = s & "   --ExceedsInvMax = InvoiceTotal > InvoiceMax" & vbCrLf
    s = s & "   case when i.pretax+i.tax > " & DbQuote(num, App.InvoiceMax) & " then 1 else 0 end =0" & vbCrLf
    s = s & "AND" & vbCrLf
    s = s & "   --ExceedsOvrMax = PO>0 and Variance > OverrideMax" & vbCrLf
    s = s & "   case when i.pretax+i.tax - sum(case when isnull(d.commitment,'')='' then d.pretax+d.tax else 0 end) > 0 " & vbCrLf
    s = s & "        and sum(case when isnull(d.commitment,'')='' then d.pretax+d.tax else 0 end) > " & DbQuote(num, App.InvOverrideMax) & vbCrLf
    s = s & "        then 1 else 0 end =0" & vbCrLf
    s = s & "AND" & vbCrLf
    s = s & "   --ExceedsOvrPct = PO>0 and Variance > PO * OverridePcnt" & vbCrLf
    s = s & "  case when i.pretax+i.tax - sum(case when isnull(d.commitment,'')='' then d.pretax+d.tax else 0 end) > 0 " & vbCrLf
    s = s & "       and sum(case when isnull(d.commitment,'')='' then d.pretax+d.tax else 0 end) > (i.pretax+i.tax - sum(case when isnull(d.commitment,'')='' then d.pretax+d.tax else 0 end)) * " & DbQuote(num, App.InvOverridePcnt) & "/100" & vbCrLf
    s = s & "       then 1 else 0 end =0" & vbCrLf
    s = s & ")" & vbCrLf
    s = s & ")" & vbCrLf
    approvable = Val(HFApp.SqlExec(s)(0))
    
    'approve them
    If approvable > 0 Then
        total = Parse(InvoiceIDs)
        If vbYes = MsgBox(total & " invoices have been created. " & approvable & " are within your approval limits." & vbCrLf & vbCrLf & "Do you want to approve them now?", vbQuestion + vbYesNo, App.ProductName) Then
            s = ""
            s = s & "update invoices" & vbCrLf
            s = s & "set status='Approved'" & vbCrLf
            s = s & "   ,approver=" & DbQuote(Str, HFApp.LoginID) & vbCrLf
            s = s & "   ,tstmp=getdate(),ustmp=" & DbQuote(Str, HFApp.LoginID) & vbCrLf
            s = s & "where invoiceid in(" & vbCrLf
            s = s & "" & vbCrLf
            s = s & "select" & vbCrLf
            s = s & "   i.invoiceid" & vbCrLf
            s = s & "from invoices i" & vbCrLf
            s = s & "join invoiceitems d on i.invoiceid=d.invoiceid" & vbCrLf
            s = s & "where i.status<>'Lien Hold' and i.invoiceid in (" & InvoiceIDs & ")" & vbCrLf
            s = s & "group by i.invoiceid,i.pretax,i.tax" & vbCrLf
            s = s & "" & vbCrLf
            s = s & "having " & vbCrLf
            s = s & "(" & vbCrLf
            s = s & "   --ExceedsInvMax = InvoiceTotal > InvoiceMax" & vbCrLf
            s = s & "   case when i.pretax+i.tax > " & DbQuote(num, App.InvoiceMax) & " then 1 else 0 end =0" & vbCrLf
            s = s & "AND" & vbCrLf
            s = s & "   --ExceedsOvrMax = PO>0 and Variance > OverrideMax" & vbCrLf
            s = s & "   case when i.pretax+i.tax - sum(case when isnull(d.commitment,'')='' then d.pretax+d.tax else 0 end) > 0 " & vbCrLf
            s = s & "        and sum(case when isnull(d.commitment,'')='' then d.pretax+d.tax else 0 end) > " & DbQuote(num, App.InvOverrideMax) & vbCrLf
            s = s & "        then 1 else 0 end =0" & vbCrLf
            s = s & "AND" & vbCrLf
            s = s & "   --ExceedsOvrPct = PO>0 and Variance > PO * OverridePcnt" & vbCrLf
            s = s & "  case when i.pretax+i.tax - sum(case when isnull(d.commitment,'')='' then d.pretax+d.tax else 0 end) > 0 " & vbCrLf
            s = s & "       and sum(case when isnull(d.commitment,'')='' then d.pretax+d.tax else 0 end) > (i.pretax+i.tax - sum(case when isnull(d.commitment,'')='' then d.pretax+d.tax else 0 end)) * " & DbQuote(num, App.InvOverridePcnt) & "/100" & vbCrLf
            s = s & "       then 1 else 0 end =0" & vbCrLf
            s = s & ")" & vbCrLf
            s = s & ")" & vbCrLf
            Call HFApp.SqlExec(s)
        End If
    End If
    
    'log status
    s = ""
    s = s & "insert approvallog(docid,ustmp,dstmp,action,assigneddept,assignedappr)" & vbCrLf
    s = s & "select docid,ustmp,tstmp,status,deptid,approver" & vbCrLf
    s = s & "from invoices" & vbCrLf
    s = s & "where invoiceid in (" & InvoiceIDs & ")" & vbCrLf
    Call HFApp.SqlExec(s)


End Sub
Private Function SomethingSelected() As Boolean
    Dim i As Long
    With gPOs
    For i = 1 To .Rows - 1
        If .Cell(flexcpChecked, i, .ColIndex("Add")) = flexChecked Then
            SomethingSelected = True
            Exit Function
        End If
    Next
    End With
End Function


Private Sub gItems_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim n As VSFlexNode
    Dim DesiredPcnt As Double
    Dim NewInvoiceAmt As Double
    Dim NothingSelected As Boolean


    With gItems
        .Redraw = flexRDNone

        'when using the mouse to check/uncheck a checkbox column, the afteredit event is fired before the selection is updated.
        If Row <> .Row Or Col <> .Col Then .Select Row, Col


        'did you click the column heading?
        If Row = 0 Then
            'set checks on all rows, iif is required!
            '.Cell(flexcpChecked, 1, Col, .Rows - 1, Col) = IIf(.Cell(flexcpChecked, Row, Col) = flexChecked, flexUnchecked, flexChecked)
            .Cell(flexcpChecked, 1, Col, .Rows - 1, Col) = .Cell(flexcpChecked, Row, Col)
            'reset selection and continue
            Call .Select(1, Col, .Rows - 1, Col)
        End If


        For Row = Min(.Row, .RowSel) To Max(.Row, .RowSel)
            
            
            If IsIn(.ColKey(Col), "PayPoint1", "PayPoint2", "PayPoint3", "PayPoint4", "PayPoint5") Then
            
                'reset grayed checkboxes that the user has overwritten (cant be prevented because grid changes check value before changing selection and before beforeedit event.
                Call SetPayPointCheckboxes(Row, .ColKey(Col))

                'what is the desired percent paid?
                DesiredPcnt = 0
                If .Cell(flexcpChecked, Row, .ColIndex("PayPoint1")) <> flexUnchecked Then DesiredPcnt = DesiredPcnt + .ValueMatrix(Row, .ColIndex("PayPoint1Percent"))
                If .Cell(flexcpChecked, Row, .ColIndex("PayPoint2")) <> flexUnchecked Then DesiredPcnt = DesiredPcnt + .ValueMatrix(Row, .ColIndex("PayPoint2Percent"))
                If .Cell(flexcpChecked, Row, .ColIndex("PayPoint3")) <> flexUnchecked Then DesiredPcnt = DesiredPcnt + .ValueMatrix(Row, .ColIndex("PayPoint3Percent"))
                If .Cell(flexcpChecked, Row, .ColIndex("PayPoint4")) <> flexUnchecked Then DesiredPcnt = DesiredPcnt + .ValueMatrix(Row, .ColIndex("PayPoint4Percent"))
                If .Cell(flexcpChecked, Row, .ColIndex("PayPoint5")) <> flexUnchecked Then DesiredPcnt = DesiredPcnt + .ValueMatrix(Row, .ColIndex("PayPoint5Percent"))

                'specifically handle when nothing checked so we don't reverse what was previously invoiced.
                NothingSelected = .Cell(flexcpChecked, Row, .ColIndex("PayPoint1")) <> flexChecked And _
                                  .Cell(flexcpChecked, Row, .ColIndex("PayPoint2")) <> flexChecked And _
                                  .Cell(flexcpChecked, Row, .ColIndex("PayPoint3")) <> flexChecked And _
                                  .Cell(flexcpChecked, Row, .ColIndex("PayPoint4")) <> flexChecked And _
                                  .Cell(flexcpChecked, Row, .ColIndex("PayPoint5")) <> flexChecked

                'back out and apply amount to be added
                If DesiredPcnt > 99 Then
                    .Cell(flexcpText, Row, .ColIndex("InvoicePretax")) = .Cell(flexcpText, Row, .ColIndex("RemainingPretax"))
                    .Cell(flexcpText, Row, .ColIndex("InvoiceTax")) = .Cell(flexcpText, Row, .ColIndex("RemainingTax"))
                Else
                    If NothingSelected Then
                        NewInvoiceAmt = 0
                    Else
                        NewInvoiceAmt = Round(DesiredPcnt / 100 * .ValueMatrix(Row, .ColIndex("Pretax")), 2) - .ValueMatrix(Row, .ColIndex("InvoicedPretax"))
                    End If
                    .Cell(flexcpText, Row, .ColIndex("InvoicePretax")) = NewInvoiceAmt
                    .Cell(flexcpText, Row, .ColIndex("InvoiceTax")) = Round(NewInvoiceAmt * .ValueMatrix(Row, .ColIndex("TaxRate")) / 100, 2)
                End If

            End If

            Call FormatAmounts(gItems, Row)

        Next

        Call CalcInvoiceTotal(True)
        Call UnloadItems

        .Redraw = flexRDBuffered
    End With
End Sub



Private Sub gItems_AfterMoveColumn(ByVal Col As Long, Position As Long)
    gItems.ColPosition(gItems.ColIndex("PayPoint5")) = 0
    gItems.ColPosition(gItems.ColIndex("PayPoint4")) = 0
    gItems.ColPosition(gItems.ColIndex("PayPoint3")) = 0
    gItems.ColPosition(gItems.ColIndex("PayPoint2")) = 0
    gItems.ColPosition(gItems.ColIndex("PayPoint1")) = 0

End Sub

Private Sub gItems_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim r As Long

    With gItems

        .AutoSearch = flexSearchNone
        .EditMaxLength = 0

        Select Case .ColKey(Col)

            Case "InvoicePretax", "InvoiceTax", "Qty", "Rate", "JointPayee"
                'Cancel = .RowOutlineLevel(Row) = 0

            Case "PayPoint1", "PayPoint2", "PayPoint3", "PayPoint4", "PayPoint5"

            Case Else
                .AutoSearch = flexSearchFromCursor
                Cancel = True

        End Select

    End With

End Sub

Private Sub gItems_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = IsIn(gItems.ColKey(Col), "PayPoint1", "PayPoint2", "PayPoint3", "PayPoint4", "PayPoint5")
End Sub


Private Sub gItems_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim parentRow As Long
    Dim childRow As Long


    With gItems
    Select Case .ColKey(Col)

        Case "Rate"
            .EditText = Val(.EditText)
            .Cell(flexcpText, Row, .ColIndex("InvoicePreTax")) = Round(Val(.EditText) * .Cell(flexcpValue, Row, .ColIndex("Qty")), 2)
            .Cell(flexcpText, Row, .ColIndex("InvoiceTax")) = Round(.Cell(flexcpValue, Row, .ColIndex("InvoicePreTax")) * .Cell(flexcpValue, Row, .ColIndex("TaxRate")) / 100, 2)

        Case "Qty"
            .EditText = Val(.EditText)
            .Cell(flexcpText, Row, .ColIndex("InvoicePreTax")) = Round(Val(.EditText) * .Cell(flexcpValue, Row, .ColIndex("Rate")), 2)
            .Cell(flexcpText, Row, .ColIndex("InvoiceTax")) = Round(.Cell(flexcpValue, Row, .ColIndex("InvoicePreTax")) * .Cell(flexcpValue, Row, .ColIndex("TaxRate")) / 100, 2)

        Case "InvoicePretax"
            .EditText = Round(Val(.EditText), 2)
            .Cell(flexcpText, Row, .ColIndex("InvoiceTax")) = Round(Val(.EditText) * .Cell(flexcpValue, Row, .ColIndex("TaxRate")) / 100, 2)

        Case "InvoiceTax"
            .EditText = Round(Val(.EditText), 2)

    End Select
    End With
End Sub



Private Sub gItems_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
    gItems.ColSel = gItems.Col
End Sub




Private Sub gItems_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then Call ShowColumnMenu(gItems, False)
End Sub

