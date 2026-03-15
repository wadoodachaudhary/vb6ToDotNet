VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Begin VB.Form FPurchaseOrder 
   Caption         =   "Purchase Order"
   ClientHeight    =   9930
   ClientLeft      =   990
   ClientTop       =   810
   ClientWidth     =   17850
   Icon            =   "FPurchaseOrder.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9930
   ScaleWidth      =   17850
   Begin VB.Frame frmScope 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   240
      TabIndex        =   26
      Top             =   3720
      Width           =   14775
      Begin VB.TextBox txtStandardText 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   825
         Left            =   90
         MaxLength       =   2000
         TabIndex        =   13
         Top             =   195
         Width           =   14310
      End
      Begin VB.Label Label1g 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Scope of Work"
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
         Left            =   120
         TabIndex        =   27
         Top             =   0
         Width           =   1290
      End
   End
   Begin VB.Frame frmTerms 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   240
      TabIndex        =   17
      Top             =   8280
      Width           =   7935
      Begin VB.TextBox txtTerms 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   1065
         Left            =   120
         MaxLength       =   2000
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   240
         Width           =   7635
      End
      Begin VB.Label Label124 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Terms"
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
         Left            =   30
         TabIndex        =   22
         Top             =   0
         Width           =   645
      End
   End
   Begin VB.Frame frmTotals 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   11160
      TabIndex        =   16
      Top             =   8280
      Width           =   3855
      Begin VB.Label lblSubtotal 
         Alignment       =   1  'Right Justify
         Height          =   195
         Left            =   2265
         TabIndex        =   43
         Top             =   120
         Width           =   1140
      End
      Begin VB.Label lblTax 
         Alignment       =   1  'Right Justify
         Height          =   195
         Left            =   2265
         TabIndex        =   42
         Top             =   480
         Width           =   1140
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Height          =   195
         Left            =   2265
         TabIndex        =   41
         Top             =   840
         Width           =   1140
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total"
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
         Left            =   1665
         TabIndex        =   40
         Top             =   840
         Width           =   450
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tax"
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
         Left            =   1785
         TabIndex        =   39
         Top             =   480
         Width           =   330
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Subtotal"
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
         Left            =   1395
         TabIndex        =   38
         Top             =   120
         Width           =   720
      End
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   29
      Top             =   0
      Width           =   17850
      _ExtentX        =   31485
      _ExtentY        =   1058
      ButtonWidth     =   1561
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            Key             =   "New"
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
            Caption         =   "Custom"
            Key             =   "TakeoffCustom"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Preview"
            Key             =   "Preview"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Email"
            Key             =   "Send"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "separator"
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Cancel"
            Key             =   "Delete"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   9420
         Top             =   90
      End
   End
   Begin VB.Frame frmHeader 
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   -15
      TabIndex        =   18
      Top             =   690
      Width           =   14895
      Begin VB.TextBox txtPOIndex 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   230
         Left            =   12480
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1608
         Width           =   2070
      End
      Begin VB.TextBox txtDateDue 
         BorderStyle     =   0  'None
         Height          =   230
         Left            =   12480
         TabIndex        =   8
         Text            =   " "
         Top             =   1368
         Width           =   2070
      End
      Begin VB.TextBox txtAddress 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   1065
         Left            =   1200
         MaxLength       =   2000
         MultiLine       =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   480
         Width           =   4155
      End
      Begin VB.TextBox txtDeliveryRecipient 
         BorderStyle     =   0  'None
         Height          =   230
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   4
         Top             =   2280
         Width           =   4155
      End
      Begin VB.TextBox txtDeliveryAddress 
         BorderStyle     =   0  'None
         Height          =   230
         Left            =   1200
         MaxLength       =   250
         TabIndex        =   5
         Top             =   2520
         Width           =   4155
      End
      Begin VB.TextBox txtShipVia 
         BorderStyle     =   0  'None
         Height          =   230
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1680
         Width           =   4155
      End
      Begin VB.TextBox txtFOB 
         BorderStyle     =   0  'None
         Height          =   230
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1920
         Width           =   4155
      End
      Begin VB.TextBox txtRetainagePercent 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   230
         Left            =   12480
         MaxLength       =   50
         TabIndex        =   11
         Top             =   2160
         Width           =   375
      End
      Begin VB.TextBox txtPODescription 
         BorderStyle     =   0  'None
         Height          =   230
         Left            =   10356
         MaxLength       =   100
         TabIndex        =   12
         Text            =   " "
         Top             =   2532
         Width           =   4335
      End
      Begin VB.TextBox txtDateIssued 
         BorderStyle     =   0  'None
         Height          =   230
         Left            =   12480
         TabIndex        =   6
         Top             =   888
         Width           =   2070
      End
      Begin TDBText6Ctl.TDBText txtJob 
         Height          =   230
         Left            =   12480
         TabIndex        =   10
         Top             =   1920
         Width           =   2070
         _Version        =   65536
         _ExtentX        =   3651
         _ExtentY        =   406
         Caption         =   "FPurchaseOrder.frx":000C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FPurchaseOrder.frx":0078
         Key             =   "FPurchaseOrder.frx":0096
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   0
         ShowContextMenu =   -1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MarginBottom    =   1
         Enabled         =   -1
         MousePointer    =   0
         Appearance      =   1
         BorderStyle     =   0
         AlignHorizontal =   0
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   ""
         FormatMode      =   1
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   0
         LengthAsByte    =   0
         Text            =   ""
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin VB.TextBox txtOrderedBy 
         BorderStyle     =   0  'None
         Height          =   230
         Left            =   12480
         MaxLength       =   50
         TabIndex        =   7
         Top             =   1128
         Width           =   2070
      End
      Begin VB.TextBox txtPONumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12348
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "AND10071/001"
         Top             =   408
         Width           =   2115
      End
      Begin TDBText6Ctl.TDBText txtVendor 
         Height          =   230
         Left            =   1200
         TabIndex        =   0
         Top             =   240
         Width           =   4155
         _Version        =   65536
         _ExtentX        =   7329
         _ExtentY        =   406
         Caption         =   "FPurchaseOrder.frx":00DA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FPurchaseOrder.frx":0146
         Key             =   "FPurchaseOrder.frx":0164
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   0
         ShowContextMenu =   -1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MarginBottom    =   1
         Enabled         =   -1
         MousePointer    =   0
         Appearance      =   1
         BorderStyle     =   0
         AlignHorizontal =   0
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   ""
         FormatMode      =   1
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   0
         LengthAsByte    =   0
         Text            =   ""
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   1
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin VB.Image cmdBrowse 
         Height          =   240
         Index           =   1
         Left            =   14565
         Picture         =   "FPurchaseOrder.frx":01A8
         Top             =   1635
         Width           =   240
      End
      Begin VB.Image cmdBrowse 
         Height          =   240
         Index           =   4
         Left            =   14670
         Picture         =   "FPurchaseOrder.frx":02F2
         Top             =   2535
         Width           =   240
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Format"
         Height          =   192
         Left            =   11928
         TabIndex        =   47
         Top             =   1608
         Width           =   504
      End
      Begin VB.Label lblApproval 
         AutoSize        =   -1  'True
         Caption         =   "Approved by ADMIN: Jan 4, 2014"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   5640
         TabIndex        =   44
         Top             =   1920
         Width           =   2400
      End
      Begin VB.Image cmdBrowse 
         Height          =   240
         Index           =   0
         Left            =   14565
         Picture         =   "FPurchaseOrder.frx":043C
         Top             =   885
         Width           =   240
      End
      Begin VB.Image cmdEditVendor 
         Height          =   240
         Left            =   930
         Picture         =   "FPurchaseOrder.frx":0586
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1r 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Recipient"
         Height          =   195
         Left            =   480
         TabIndex        =   37
         Top             =   2280
         Width           =   675
      End
      Begin VB.Label Label120 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Email/Fax"
         Height          =   195
         Left            =   450
         TabIndex        =   36
         Top             =   2520
         Width           =   705
      End
      Begin VB.Label Label123 
         AutoSize        =   -1  'True
         Caption         =   "Ship Via"
         Height          =   195
         Left            =   450
         TabIndex        =   35
         Top             =   1680
         Width           =   705
      End
      Begin VB.Label Label125 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Free On Board"
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   1920
         Width           =   1035
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   300
         TabIndex        =   33
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblSendVia 
         AutoSize        =   -1  'True
         Caption         =   "print"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   5640
         TabIndex        =   32
         Top             =   2520
         Width           =   300
      End
      Begin VB.Label Label155 
         Caption         =   "Retainage"
         Height          =   255
         Left            =   11457
         TabIndex        =   31
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "PURCHASE ORDER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   22.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   732
         Left            =   5280
         TabIndex        =   30
         Top             =   120
         Width           =   6252
      End
      Begin VB.Label Label1d 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Description"
         Height          =   195
         Left            =   9360
         TabIndex        =   28
         Top             =   2520
         Width           =   795
      End
      Begin VB.Image cmdBrowse 
         Height          =   240
         Index           =   2
         Left            =   14565
         Picture         =   "FPurchaseOrder.frx":0B10
         Top             =   1365
         Width           =   240
      End
      Begin VB.Label Label1id 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Due"
         Height          =   192
         Left            =   12132
         TabIndex        =   25
         Top             =   1368
         Width           =   300
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Date"
         Height          =   192
         Left            =   12084
         TabIndex        =   24
         Top             =   888
         Width           =   348
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Job "
         Height          =   195
         Left            =   12132
         TabIndex        =   23
         Top             =   1920
         Width           =   300
      End
      Begin VB.Label Label122 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ordered By"
         Height          =   192
         Left            =   11640
         TabIndex        =   21
         Top             =   1128
         Width           =   792
      End
      Begin VB.Label lblPOIndex 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Purchase Order"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   12348
         TabIndex        =   20
         Top             =   168
         Width           =   2112
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid gItems 
      Height          =   2895
      Left            =   330
      TabIndex        =   14
      Top             =   4980
      Width           =   14535
      _cx             =   1986946086
      _cy             =   1986925554
      Appearance      =   2
      BorderStyle     =   1
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
      Rows            =   5
      Cols            =   80
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FPurchaseOrder.frx":0C5A
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
      TabBehavior     =   1
      OwnerDraw       =   2
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   1
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
         Left            =   2160
         ScaleHeight     =   585
         ScaleWidth      =   3765
         TabIndex        =   45
         Top             =   960
         Visible         =   0   'False
         Width           =   3795
         Begin VB.TextBox lblWarningMessages 
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   360
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   46
            Text            =   "FPurchaseOrder.frx":168F
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
         Picture         =   "FPurchaseOrder.frx":1695
         Top             =   480
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgWarning 
         Height          =   240
         Left            =   0
         Picture         =   "FPurchaseOrder.frx":1C1F
         Top             =   240
         Visible         =   0   'False
         Width           =   240
      End
   End
End
Attribute VB_Name = "FPurchaseOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const SRCFILE = "FPurchaseOrder::"
Public EventTraps As Collection

Private mFunnyFlag As Boolean ' used to ignore the first mouse move event see gitems_mousemove and gitems_rowcolchanged


Private mDirty  As Boolean
Private mNew    As Boolean

Private mTakeoffSettings As String 'csv list of location,wbs1,wbs2...  set in ftakeoffsettings and AddItem()


Private mPostedPO As Boolean

'conditional formatting
Private AcceptedForeColor As Long
Private AcceptedBackColor As Long
Private AcceptedStyle     As String
Private DeclinedForeColor As Long
Private DeclinedBackColor As Long
Private DeclinedStyle     As String
Private MinForeColor      As Long
Private MinBackColor      As Long
Private MinStyle          As String
Private MaxForeColor      As Long
Private MaxBackColor      As Long
Private MaxStyle          As String


    

Private mMode           As FCEMode
Private mJob            As String
Private mCommunity      As String
Private mCommunityPhase As String
Private mRFP            As Long
Private mTakeoffSystemBidID  As Long
Private mTakeoffSystemProjectName As String

Private mEstAssemblyID          As Long 'gets set when the assembly header frame is show. is used if jcextra is changed
Private mAssemblyChanged        As Boolean
Private mCustomerNo             As String 'these get set when the assembly tree is double clicked. see loaditems()
Private mHFLocation             As Long
Private mChangeOrder            As String

Private mViewIndex As Long
Private mLock As Boolean 'stupid flag see txtDateDue.Change,keydown,keyup

Private mTimerTask As String 'stupid menus


'these are for takeoffs - identifys the selected row
Private mCurrentAssemblyType    As AssemblyTypes
Private mCurrentJob             As String
Private mCurrentJCExtra         As String
Private mCurrentChangeOrder     As String
Private mCurrentCommunity       As String
Private mCurrentCommunityPhase  As String
Private mCurrentAssembly        As String
Private mCurrentAssemblyDesc    As String
Public mCurrentEstAssemblyID    As Long
Private mCurrentJobDesc         As String
Private mCurrentModel           As String
Private mCurrentOptionID        As String
Private mCurrentBudgetsLocked   As Boolean
Private mCancelTakeoff          As Boolean

'used purely for debugging
Private mViewQuery As String

Private Const mcASMBlY_ADDCO = 0
Private Const mcASMBlY_ADDCR = 1
Private Const mcASMBlY_ADD = 2
Private Const mcASMBlY_DELETE = 4
Private Const mcASMBlY_RENAME = 5
Private Const mcASMBlY_FINALIZE = 7
Private Const mcASMBlY_ATTACHQUOTE = 8
Private Const mcASMBlY_COPYFROMJOB = 9
Private Const mcASMBlY_PRINT = 11

Private Const mcITEM_COPYITEMS = 0
Private Const mcITEM_SPLITITEMS = 1
Private Const mcITEM_SUBSTITUEITEM = 2
Private Const mcITEM_REMOVEITEMS = 3
Private Const mcITEM_VIEWFILES = 5
Private Const mcITEM_CANCELBUDGETS = 8
Private Const mcITEM_SAVEONETIMETODB = 10
Private Const mcITEM_UPDATEPRICELIST = 11
Private Const mcITEM_COMPAREPRICES = 12
Private Const mcITEM_FORMATTING = 13
                                
Private Const mcRFP_DELETE = 0
                                
                                
Private Const mcPO_PREVIEW = 0
Private Const mcPO_PRINT = 1
Private Const mcPO_SEND = 2
Private Const mcPO_CANCEL = 4
Private Const mcPO_EDITVENDOR = 6
Private Const mcPO_CHANGE = 7
                                
Private Const mcINV_PREVIEW = 0
Private Const mcINV_PRINT = 1
Private Const mcINV_NEW = 3
Private Const mcINV_DELETE = 4
Private Const mcINV_VOID = 5
Private Const mcINV_POST = 7
                                
                                
Private Const mcVENDOR_EDIT = 0
                                
Private Const mcBIDITEM_ACCEPTED = 0
Private Const mcBIDITEM_DECLINED = 1
Private Const mcBIDITEM_COMMENTS = 2
Private Const mcBIDITEM_EXPORT = 4
Private Const mcBIDITEM_IMPORT = 5
Private Const mcBIDITEM_FORMATTING = 7
Private Const mcBIDITEM_VENDOR = 8

Private Const mcRFPITEM_ADD = 0
Private Const mcRFPITEM_REMOVE = 1


Private Const mcBID_ADD = 0
Private Const mcBID_REMOVE = 1
Private Const mcBID_EXPORT = 3
Private Const mcBID_IMPORT = 4
Private Const mcBID_VENDOR = 6


Private Sub cmdBrowse_Click(Index As Integer)
On Error Resume Next
    
    Dim s As String
    
    Select Case Index
        Case 0
            If txtDateIssued.Enabled Then Call DCalendar.Popup(txtDateIssued)
        
        Case 1
            If txtPOIndex.Enabled Then
            If FPickList.Choose(HFApp.Databases(dbHomefront), "PO Index", "select POIndex,description from tblpoindex where DivisionID = " & HFApp.DivisionID, txtPOIndex.Text) Then
                txtPOIndex.Text = FPickList.SelectedItem("POIndex")
            End If
            End If
        
        Case 2
            If txtDateDue.Enabled Then Call DCalendar.Popup(txtDateDue)
            
        Case 4
            If txtPODescription.Enabled Then
                s = txtPODescription.Text
                If FComments.Edit(s, txtPODescription, , "Description", 100) Then txtPODescription.Text = s
            End If
            
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case True
        Case Shift = vbCtrlMask And KeyCode = vbKeyS: Call Toolbar_ButtonClick(Toolbar.Buttons("Save"))
        Case Shift = vbCtrlMask And KeyCode = vbKeyO: Call Toolbar_ButtonClick(Toolbar.Buttons("Open"))
        Case Shift = vbCtrlMask And KeyCode = vbKeyN: Call Toolbar_ButtonClick(Toolbar.Buttons("New"))
    End Select
End Sub

Private Sub Form_Load()

    Dim i As Long
    Dim s As String
    Dim t As String
    Dim rs As Recordset
    Dim c As Control
    
    ClearForm
    FieldsLocked = True
    
    Call IniGetGrid(Me, gItems)
    'configure form
    
    Me.lblApproval.Caption = ""
    Me.lblSendVia.Caption = ""
    Me.Caption = Choose(mMode, "Prepare Quote", "Issue Budgets", "Issue PO's", Me.Caption, Me.Caption, Me.Caption)
        
    'change to normal colors
    Me.BackColor = vbButtonFace
    For Each c In Me.Controls
        If TypeName(c) = "Frame" Then
            c.BackColor = vbButtonFace
        End If
        If TypeName(c) = "Slider" Then
            c.Visible = True
        End If
    Next
        
    txtPOIndex.Text = "" & HFApp.Options.ValueByName("DefaultPOIndex")
    If txtPOIndex.Text = "" Then txtPOIndex.Text = "99 Manual PO"
        
    'get conditional formatting
    AcceptedForeColor = IniGet(AppIni, "Formatting", "AcceptedForeColor", vbWindowText)
    AcceptedBackColor = IniGet(AppIni, "Formatting", "AcceptedBackColor", vbWindowBackground)
    AcceptedStyle = IniGet(AppIni, "Formatting", "AcceptedStyle", "Bold Italic")
    DeclinedForeColor = IniGet(AppIni, "Formatting", "DeclinedForeColor", vbGrayText)
    DeclinedBackColor = IniGet(AppIni, "Formatting", "DeclinedBackColor", vbButtonFace)
    DeclinedStyle = IniGet(AppIni, "Formatting", "DeclinedStyle", "Regular")
    MinForeColor = IniGet(AppIni, "Formatting", "MinForeColor", vbWindowText)
    MinBackColor = IniGet(AppIni, "Formatting", "MinBackColor", 16773869) 'light blue
    MinStyle = IniGet(AppIni, "Formatting", "MinStyle", "Regular")
    MaxForeColor = IniGet(AppIni, "Formatting", "MaxForeColor", vbWindowText)
    MaxBackColor = IniGet(AppIni, "Formatting", "MaxBackColor", 12640511) 'light pink
    MaxStyle = IniGet(AppIni, "Formatting", "MaxStyle", "Regular")
    Call IniGetForm(Me)
    Call SetToolbarIcons(Toolbar, FMain.LargeIcons)
    mDirty = False
    
End Sub

Private Sub Form_Resize()
    gItems.Width = Me.ScaleWidth - 2 * gItems.left
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not SaveData(True) Then
        Cancel = True
        Exit Sub
    End If
    HFApp.Options.ValueByName("DefaultPOIndex") = txtPOIndex.Text
    Call IniPutForm(Me)
    Call IniPutGrid(Me, gItems)
    Unload FComments
End Sub

Private Sub gItems_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
If gItems.Cols > 2 And gItems.Editable = flexEDKbdMouse And gItems.Rows - 1 <> 0 And NewRow <> OldRow Or gItems.Rows - 1 = 1 Then
    If gItems.Row = gItems.Rows - 1 Then
        If gItems.TextMatrix(gItems.Row - 1, gItems.ColIndex("Description")) <> "" And gItems.TextMatrix(LastVisibleRow(gItems.Row), gItems.ColIndex("Description")) = "" Or (gItems.IsSubtotal(gItems.Row) And gItems.TextMatrix(LastVisibleRow(gItems.Row), gItems.ColIndex("Description")) <> "") Then ' gItems.IsSubtotal(gItems.Row) Then
            gItems.AddItem ""
            gItems.TextMatrix(gItems.Rows - 1, gItems.ColIndex("EstAssemblyID")) = mEstAssemblyID
            gItems.TextMatrix(gItems.Rows - 1, gItems.ColIndex("Job_No")) = txtJob.Text
            gItems.TextMatrix(gItems.Rows - 1, gItems.ColIndex("ConversionFactor")) = 1
            gItems.RowData(gItems.Rows - 1) = "NEW"
            gItems.RowPosition(gItems.Rows - 1) = gItems.Rows - 2
        End If
    End If
End If
End Sub
Private Function LastVisibleRow(ByVal Row As Long) As Long
Dim r As Long
For r = Min(Row, gItems.Rows - 2) To 1 Step -1
    If gItems.RowHidden(r) = False Then
        LastVisibleRow = r
        Exit Function
    End If
Next
End Function


Private Sub gItems_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rowPOed As Boolean
    Dim rowEither As Boolean
    
    With gItems
        .ComboList = ""
        .EditMaxLength = 0
        .AutoSearch = flexSearchNone
        
        If .Rows < 2 Then Exit Sub
        
        If .IsSubtotal(Row) Then
            Cancel = True
            Exit Sub
        End If
        
        
        rowPOed = False
        rowEither = rowPOed
        
        If rowPOed Then
            Cancel = True
            Exit Sub
        End If
        
        Select Case .ColKey(Col)
            Case "Comments"
            Case "JCExtra":            Cancel = rowEither:         .ComboList = IIf(HFApp.Options(Use_Timberline), "|...", ""):   .EditMaxLength = 10
            Case "JCCostCode":         Cancel = rowEither:         .ComboList = "|..."
            Case "JCCostCodeDesc":     Cancel = rowEither:         .ComboList = "..."
            Case "JCCategory":         Cancel = rowEither:         .ComboList = "|..."
            Case "JCCategoryDesc":     Cancel = rowEither:         .ComboList = "..."
            Case "Description":                                       .ComboList = "|...":           .EditMaxLength = 200
            Case "ItemComments":
            Case "Job_No":             .ComboList = "|...":
            Case "TakeoffQty":         Cancel = rowPOed

            Case "OrderQty":              Cancel = rowPOed
            Case "Unit":                                                                          .EditMaxLength = 10

            Case "Rate":             Cancel = rowPOed
            Case "OrderUOM":                                       .ComboList = "|...":           .EditMaxLength = 10

            Case "Pretax":           Cancel = rowPOed

            Case "TaxGroup":         Cancel = rowPOed:           .ComboList = "|..."

            Case "JCTax":            Cancel = rowPOed

            Case "NJCTax":           Cancel = rowPOed


            Case "Location":        .ComboList = GetComboList("Location"):                     .EditMaxLength = 50
            Case Else:              Cancel = True
        End Select
        
        If left(.ColKey(Col), 3) = "WBS" Then
            .ComboList = GetComboList(.ColKey(Col))
            .EditMaxLength = 50
            Cancel = False
        End If
        
        .AutoSearch = IIf(Cancel, flexSearchFromCursor, flexSearchNone)
        
    End With
End Sub

Private Function GetComboList(fieldname As String) As String
    Dim i As Long
    Dim s As String
    Dim rs As Recordset
    
    s = "select distinct " & fieldname & " from estimateitems where job=" & DbQuote(Str, mJob) & " order by 1"
    Set rs = HFApp.SqlExec(s, dbHomefront)
    s = ""
    While Not rs.EOF
        If Trim("" & rs(0)) <> "" Then s = s & "|" & rs(0)
        rs.MoveNext
    Wend
    If s = "|" Then s = ""
    GetComboList = s
    
End Function

Private Sub gItems_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single, Cancel As Boolean)
    Dim i As Long
    Dim r As Long
    Dim s As String
    
    With gItems
    If Button = vbRightButton Then
        If .MouseRow = 0 Then
            Cancel = True
            
            
            Call FMain.ShowColumnMenu(gItems, Not IsIn(.ColKey(.MouseCol), "Selected", "WarningMessages"))
            
            
            'this is a wbs column save description in case they changed it.
            If .MouseCol > -1 Then
                If left(.ColKey(.MouseCol), 3) = "WBS" Then
                    i = Val(Mid(.ColKey(.MouseCol), 4))
                    s = "update tbljobs set wbsdesc" & Format(i, "00") & "=" & DbQuote(Str, .TextMatrix(0, .MouseCol)) & " where DivisionID = " & HFApp.DivisionID & " and job_no=" & DbQuote(Str, mJob)
                    Call HFApp.SqlExec(s)
                End If
            End If
            
        Else
            If .Row < 1 And .MouseRow > 1 Then .Row = .MouseRow
            If .Row > 0 Then
            Set MouseCtrl = gItems
            MouseCol = .MouseCol
            
            FMain.mnuEstimateItemsGridSub(mcITEM_CANCELBUDGETS).Enabled = False
            If HFApp.UserPermission("IssuePOs") = False Then
                FMain.mnuEstimateItemsGridSub(mcITEM_SPLITITEMS).Enabled = False
                FMain.mnuEstimateItemsGridSub(mcITEM_COPYITEMS).Enabled = False
                FMain.mnuEstimateItemsGridSub(mcITEM_REMOVEITEMS).Enabled = False
                FMain.mnuEstimateItemsGridSub(mcITEM_SUBSTITUEITEM).Enabled = False
                FMain.mnuEstimateItemsGridSub(mcITEM_SAVEONETIMETODB).Enabled = False
                FMain.mnuEstimateItemsGridSub(mcITEM_UPDATEPRICELIST).Enabled = False
            Else
                FMain.mnuEstimateItemsGridSub(mcITEM_SPLITITEMS).Enabled = txtPONumber.Text = ""
                FMain.mnuEstimateItemsGridSub(mcITEM_COPYITEMS).Enabled = True
                FMain.mnuEstimateItemsGridSub(mcITEM_REMOVEITEMS).Enabled = True
                FMain.mnuEstimateItemsGridSub(mcITEM_SUBSTITUEITEM).Enabled = True
                
                'can do this only if item is onetime
                FMain.mnuEstimateItemsGridSub(mcITEM_SAVEONETIMETODB).Enabled = .TextMatrix(.Row, .ColIndex("EstPhase")) = ""
                
                'can do this if budgeted not generated and po not generated
               ' FMain.mnuEstimateItemsGridSub(mcITEM_SUBSTITUEITEM).Enabled = Trim(txtPONumber.Text) = ""
            End If
            
            
        
        
            
            PopupMenu FMain.mnuEstimateItemsGrid
            End If
        End If
    End If
    End With
End Sub

Private Sub gItems_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    Dim fixd As Boolean
    
    With gItems
        fixd = IsIn(.ColKey(Col), "Selected", "WarningMessages")
        If fixd Or Col < gItems.FrozenCols Then
            Position = Col
        Else
            If Position < gItems.FrozenCols Then
                Position = gItems.FrozenCols
            End If
        End If
    End With

End Sub



Private Sub gItems_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case gItems.ColKey(Col)
        Case "WarningMessages", "Selected"
            Cancel = True
    End Select
End Sub

Private Sub gItems_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim s As String
    Dim i As Long
    Dim OldVendor As String, OldVendorName As String
    Dim rCount As Long
    Dim r As Long
    With gItems
        Select Case .ColKey(Col)
        
            Case "Job_No"
                s = "SELECT Job_No Job,Description,Municipal_Address Address FROM tblJobs where DivisionID = " & HFApp.DivisionID & " and inactive = 0 and isquote = 0"
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Job", s, gItems, , , , , False) Then
                    .Cell(flexcpText, Row, .ColIndex("Job_No"), .RowSel, .ColIndex("Job_No")) = FPickList.SelectedItem("Job")
                End If
            Case "JCExtra"
                If HFApp.Options(AccountingSystem) = asTimberline Then
                    s = "select Extra,Description from jcm_master__extra where job=" & DbQuote(Str, HFApp.FormatJob(.TextMatrix(Row, .ColIndex("Job_No"))))
                    If FPickList.Choose(HFApp.Databases(dbAccounting), "Extra", s, gItems) Then
                        .Cell(flexcpText, Row, .ColIndex("JCExtra"), .RowSel, .ColIndex("JCExtra")) = FPickList.SelectedItem("Extra")
                    End If
                End If
                
            Case "JCCostCode"
                s = ""
                s = s & "Cost Codes" & Chr(1) & vbCrLf
                s = s & "select CostCode,isnull(nullif(Description,''),costcode) Description" & vbCrLf
                s = s & "  from alljobcostcodes" & vbCrLf
                s = s & " where DivisionID = " & HFApp.DivisionID & vbCrLf
                s = s & "   and job=" & DbQuote(Str, .Cell(flexcpText, Row, .ColIndex("Job_No"))) & vbCrLf & Chr(0)
                s = s & "Standard Cost Codes" & Chr(1) & vbCrLf
                s = s & "select CostCode,Description" & vbCrLf
                s = s & "  from StandardCostCodes" & vbCrLf
                s = s & " where DivisionID=" & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Cost Code", s, gItems) Then
                    .Cell(flexcpText, Row, .ColIndex("JCCostCode"), .RowSel, .ColIndex("JCCostCode")) = FPickList.SelectedItem("CostCode")
                    .Cell(flexcpText, Row, .ColIndex("JCCostCodeDesc"), .RowSel, .ColIndex("JCCostCodeDesc")) = FPickList.SelectedItem("Description")
                End If
                
            Case "JCCostCodeDesc"
                s = ""
                s = s & "Cost Codes" & Chr(1) & vbCrLf
                s = s & "select isnull(nullif(Description,''),costcode) Description,CostCode" & vbCrLf
                s = s & "  from alljobcostcodes" & vbCrLf
                s = s & " where DivisionID = " & HFApp.DivisionID & vbCrLf
                s = s & "   and job=" & DbQuote(Str, .Cell(flexcpText, Row, .ColIndex("Job_No"))) & vbCrLf & Chr(0)
                s = s & "Standard Cost Codes" & Chr(1) & vbCrLf
                s = s & "select Description,CostCode" & vbCrLf
                s = s & "  from StandardCostCodes" & vbCrLf
                s = s & " where DivisionID=" & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Cost Code", s, gItems) Then
                    .Cell(flexcpText, Row, .ColIndex("JCCostCode"), .RowSel, .ColIndex("JCCostCode")) = FPickList.SelectedItem("CostCode")
                    .Cell(flexcpText, Row, .ColIndex("JCCostCodeDesc"), .RowSel, .ColIndex("JCCostCodeDesc")) = FPickList.SelectedItem("Description")
                End If
            
            Case "JCCategory"
                s = "SELECT Category,Description FROM StandardCategories where DivisionID = " & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Category", s, gItems) Then
                    .Cell(flexcpText, Row, .ColIndex("JCCategory"), .RowSel, .ColIndex("JCCategory")) = FPickList.SelectedItem("Category")
                    .Cell(flexcpText, Row, .ColIndex("JCCategoryDesc"), .RowSel, .ColIndex("JCCategoryDesc")) = FPickList.SelectedItem("Description")
                End If
                
            Case "JCCategoryDesc"
                s = "SELECT Description,Category FROM StandardCategories where DivisionID = " & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Category", s, gItems, , , , IIf(HFApp.Options(AccountingSystem) = asQuickBooks, "Category", "")) Then
                    .Cell(flexcpText, Row, .ColIndex("JCCategory"), .RowSel, .ColIndex("JCCategory")) = FPickList.SelectedItem("Category")
                    .Cell(flexcpText, Row, .ColIndex("JCCategoryDesc"), .RowSel, .ColIndex("JCCategoryDesc")) = FPickList.SelectedItem("Description")
                End If
'            Case "Job_no"
'                s = "SELECT Job_No Job,Description,Municipal_Address Address FROM tblJobs where inactive = 0 and isquote = 0"
'                If FPickList.Choose(HFApp.Databases(dbHomeFront), "Job", s, gItems, , , , , False) Then
'                    .Cell(flexcpText, Row, .ColIndex("Job_No"), .RowSel, .ColIndex("Job_No")) = FPickList.SelectedItem("Job")
'                End If
            Case "OrderUOM"
                s = "SELECT DISTINCT OrderUOM Unit FROM tblPhaseItem where DivisionID = " & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Units", s, gItems) Then
                    .Cell(flexcpText, Row, .ColIndex("OrderUOM"), .RowSel, .ColIndex("OrderUOM")) = FPickList.SelectedItem("Unit")
                End If
            
            Case "ItemComments"
                s = gItems.Text
                If FComments.Edit(s, gItems) Then
                    gItems.Text = s
                End If
                
            Case "Description":
                s = gItems.Text
                If FComments.Edit(s, gItems, , "Description", 200) Then
                    gItems.Text = s
                End If
                
            Case "TaxGroup"
                s = "SELECT TaxGroup,Description FROM TaxGroups where DivisionID = " & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Tax Groups", s, gItems) Then
                    .Cell(flexcpText, Row, Col, .RowSel, Col) = FPickList.SelectedItem("TaxGroup")
                End If
            


        End Select
        
        
        Call gItems_AfterEdit(Row, Col)
        
    End With
End Sub

Private Sub gItems_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Dim s As String
    Dim r As Long
    Dim CanDeletePO     As Boolean
    Dim MinVal As Long
    Dim MaxVal As Long
    
    
    
    With gItems
        Select Case True
                
            Case KeyCode = vbKeyF And Shift = vbCtrlMask
                Call FFind.ShowForm(gItems)
                
            Case KeyCode = vbKeyDelete And Shift = vbCtrlMask
                MinVal = Min(.Row, .RowSel)
                MaxVal = Max(.Row, .RowSel)
                For r = MinVal To MaxVal
                If r <> .Rows - 1 Then
                    CanDeletePO = True 'Change later to put validation in for checking if invoiced.
                    If CanDeletePO Then
                        .TextMatrix(r, .ColIndex("PODeleted")) = "True"
                        Dirty = True
                        If .IsSubtotal(r) = False Then
                            .RowHidden(r) = True
                             If .RowData(r) <> "NEW" Then .RowData(r) = "DIRTY"
                        End If
                    End If
                End If
                Next
                .Row = GridNextVisibleRow(gItems, MaxVal)
                
        End Select
    End With
    
End Sub

Private Sub gItems_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
    If Button <> 0 Then Exit Sub
    With gItems
        If mFunnyFlag Then
            mFunnyFlag = False
            Exit Sub
        End If
        If .MouseRow >= Min(.Row, .RowSel) And .MouseRow <= Max(.Row, .RowSel) And .MouseCol = .Col Then
            Call gItems_SelChange
        Else
            Call ShowTip(.MouseRow, .MouseCol, MouseX(gItems.hwnd) * Screen.TwipsPerPixelX + 315, MouseY(gItems.hwnd) * Screen.TwipsPerPixelY + 315)
        End If
    End With
End Sub



Private Sub gItems_SelChange()
    Dim r As Long
    Dim prefix As String
    
    Dim pretax As Double
    Dim tax As Double
    Dim tip As String
    Dim Qty As Double

    prefix = ""

    
    With gItems
        If .Row <> .RowSel Then
            For r = Min(.Row, .RowSel) To Max(.Row, .RowSel)
            If .RowHidden(r) = False And Not .IsSubtotal(r) Then
                pretax = pretax + .ValueMatrix(r, .ColIndex(prefix & "Pretax"))
                tax = tax + .ValueMatrix(r, .ColIndex(prefix & "TotalTax"))
                If IsIn(.ColKey(.Col), "TakeoffQty", "OrderQty", "BudgetQty") Then
                    Qty = Qty + .ValueMatrix(r, .Col)
                End If
            End If
            Next
            
                

            
            tip = IIf(IsIn(.ColKey(.Col), "TakeoffQty", "OrderQty"), .ColKey(.Col) & ":" & vbTab & Qty & vbCrLf & vbCrLf, "") & _
                  "Pretax:" & vbTab & Format(pretax, "#,##0.00") & vbCrLf & _
                  "Tax:" & vbTab & Format(tax, "#,##0.00") & vbCrLf & _
                  "Total:" & vbTab & Format(pretax + tax, "#,##0.00")
            
            Call ShowTip(.MouseRow, .MouseCol, MouseX(gItems.hwnd) * Screen.TwipsPerPixelX + 315, MouseY(gItems.hwnd) * Screen.TwipsPerPixelY + 315, tip)
            
        End If
    End With
End Sub


Private Sub gItems_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error Resume Next
    Dim s As String
    Dim i As Long
    Dim Description As String
    Dim rs As Recordset
    Dim r As Long
    
    With gItems
    For r = Min(.Row, .RowSel) To Max(.Row, .RowSel)
        
        s = .EditText
        Select Case .ColKey(Col)
            Case "JCCostCode":            Cancel = Not ValidateField(gItems, s, "Cost Code not found", "SELECT CostCode,Description FROM StandardCostCodes WHERE DivisionID = " & HFApp.DivisionID & " and CostCode=" & DbQuote(Str, s), "JCCostCodeDesc")
            Case "JCCategory":            Cancel = Not ValidateField(gItems, s, "Category not found", "SELECT Category,Description FROM StandardCategories WHERE DivisionID = " & HFApp.DivisionID & " and Category=" & DbQuote(Str, s), "JCCategoryDesc")
            Case "OrderQty":              Cancel = Not IsNumeric(s)
            Case "Rate":                  Cancel = Not IsNumeric(s)
            Case "Pretax":                Cancel = Not IsNumeric(s)
            Case "TaxGroup":              Cancel = Not ValidateField(gItems, s, "Tax Group not found", "SELECT TaxGroup FROM TaxGroups WHERE DivisionID = " & HFApp.DivisionID & " and TaxGroup=" & DbQuote(Str, s))
            Case "JCTax":                 Cancel = Not IsNumeric(s)
            Case "NJCTax":                Cancel = Not IsNumeric(s)
            Case "JCExtra"
                If HFApp.Options(Use_Timberline) And Trim(s) <> "" Then
                    Set rs = HFApp.SqlExec("select extra from jcm_master__extra where job=" & DbQuote(Str, HFApp.FormatJob(.TextMatrix(Row, .ColIndex("Job_No")))) & " AND extra=" & DbQuote(Str, s), dbAccounting)
                    If rs.EOF Then
                        Description = InputBox(vbCrLf & "Extra '" & s & "' could not be found." & vbCrLf & "Do you want to add it?" & vbCrLf & vbCrLf & vbCrLf & "Enter a description for the extra. (30 characters or less)", App.ProductName)
                        If Description = "" Then
                            Cancel = True
                        Else
                            s = ""
                            s = s & "INSERT INTO jcm_master__extra(Job,Extra,Description,Status)" & vbCrLf
                            s = s & "VALUES(" & DbQuote(Str, HFApp.FormatJob(.TextMatrix(Row, .ColIndex("Job_No")))) & vbCrLf
                            s = s & "      ," & DbQuote(Str, .EditText) & vbCrLf
                            s = s & "      ," & DbQuote(Str, left(Description, 30)) & vbCrLf
                            s = s & "      ,'In progress')" & vbCrLf
                            Set rs = HFApp.SqlExec(s, dbAccounting)
                            s = .EditText
                        End If
                    Else
                        s = "" & rs(0)
                    End If
                End If
        End Select
        
        .EditText = s
        
        If r = .Rows - 1 Then
            .RowData(r) = "NEW"
            .AddItem ""
        End If
        If .RowData(r) = "" Then .RowData(r) = "DIRTY"
    
    Next
    End With
    
End Sub


Private Sub CalcItems()
    Dim SubTotal As Double
    Dim tax As Double
    Dim i As Long
    
    With gItems
        For i = 1 To .Rows - 1
            If Not .RowHidden(i) Then
                SubTotal = SubTotal + .ValueMatrix(i, .ColIndex("PreTax"))
                tax = tax + .ValueMatrix(i, .ColIndex("JCTax")) + .ValueMatrix(i, .ColIndex("NJCTax"))
            End If
        Next
        
        lblSubtotal.Caption = Format(SubTotal, "#,###.00")
        lblTax.Caption = Format(tax, "#,###.00")
        lblTotal.Caption = Format(SubTotal + tax, "#,###.00")
        
    End With
End Sub


Private Sub gItems_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error Resume Next
    Dim r           As Long
    Dim rowPOed     As Boolean
    Dim rs          As Recordset
    Dim s           As String
    
    
    With gItems
    
    For r = Min(.Row, .RowSel) To Max(.Row, .RowSel)
    If r > 0 Then
    
        rowPOed = Trim(.TextMatrix(r, .ColIndex("PONumber"))) <> ""
    
        Select Case .ColKey(Col)
            Case "JCCategory", "JCCategoryDesc":
                s = ""
                s = s & "SELECT * FROM TaxGroups WHERE DivisionID = " & HFApp.DivisionID & " and TaxGroup=" & vbCrLf
                s = s & "dbo.Purch_GetDefaultTaxGroup(" & DbQuote(Str, .TextMatrix(r, .ColIndex("Job_No"))) & vbCrLf
                s = s & "   ," & DbQuote(Str, mCommunity) & vbCrLf
                s = s & "   ," & DbQuote(Str, mCommunityPhase) & vbCrLf
                s = s & "   ," & DbQuote(Str, "") & vbCrLf
                s = s & "   ," & DbQuote(Str, "") & vbCrLf
                s = s & "   ," & DbQuote(Str, .TextMatrix(r, .ColIndex("EstPhase"))) & vbCrLf
                s = s & "   ," & DbQuote(Str, .TextMatrix(r, .ColIndex("EstItem"))) & vbCrLf
                s = s & "   ," & DbQuote(Str, txtVendor.Text) & vbCrLf
                s = s & "   ," & DbQuote(Str, .TextMatrix(r, .ColIndex("JCCategory"))) & "," & HFApp.DivisionID & ")"
                Set rs = HFApp.SqlExec(s)
                If Not rs.EOF And Not rowPOed Then
                    .TextMatrix(r, .ColIndex("TaxGroup")) = "" & rs("TaxGroup")
                    .TextMatrix(r, .ColIndex("JCTaxRate")) = "" & rs("JCRate")
                    .TextMatrix(r, .ColIndex("NJCTaxRate")) = "" & rs("NJCRate")
                    .TextMatrix(r, .ColIndex("JCTax")) = Round(.ValueMatrix(r, .ColIndex("Pretax")) * .ValueMatrix(r, .ColIndex("JCTaxRate")) / 100, 2)
                    .TextMatrix(r, .ColIndex("NJCTax")) = Round(.ValueMatrix(r, .ColIndex("Pretax")) * .ValueMatrix(r, .ColIndex("NJCTaxRate")) / 100, 2)
                    .TextMatrix(r, .ColIndex("TotalTax")) = .ValueMatrix(r, .ColIndex("JCTax")) + .ValueMatrix(r, .ColIndex("NJCTax"))
                End If

            Case "TakeoffQty":
                .TextMatrix(r, .ColIndex("OrderQty")) = RoundTo(.ValueMatrix(r, .ColIndex("TakeoffQty")) * (100 + .ValueMatrix(r, .ColIndex("WastePercent"))) / 100 * .ValueMatrix(r, .ColIndex("ConversionFactor")), .ValueMatrix(r, .ColIndex("RoundTo")), .ValueMatrix(r, .ColIndex("RoundDir")))
                .TextMatrix(r, .ColIndex("Pretax")) = Round(.ValueMatrix(r, .ColIndex("OrderQty")) * .ValueMatrix(r, .ColIndex("Rate")), 2)
                .TextMatrix(r, .ColIndex("JCTax")) = Round(.ValueMatrix(r, .ColIndex("Pretax")) * .ValueMatrix(r, .ColIndex("JCTaxRate")) / 100, 2)
                .TextMatrix(r, .ColIndex("NJCTax")) = Round(.ValueMatrix(r, .ColIndex("Pretax")) * .ValueMatrix(r, .ColIndex("NJCTaxRate")) / 100, 2)
                .TextMatrix(r, .ColIndex("TotalTax")) = .ValueMatrix(r, .ColIndex("JCTax")) + .ValueMatrix(r, .ColIndex("NJCTax"))
                .TextMatrix(r, .ColIndex("Total")) = .ValueMatrix(r, .ColIndex("Pretax")) + .ValueMatrix(r, .ColIndex("JCTax")) + .ValueMatrix(r, .ColIndex("NJCTax"))
            
            Case "OrderQty":
                .TextMatrix(r, .ColIndex("TakeoffQty")) = .ValueMatrix(r, .ColIndex("OrderQty")) / .ValueMatrix(r, .ColIndex("ConversionFactor"))
                .TextMatrix(r, .ColIndex("Pretax")) = Round(.ValueMatrix(r, .ColIndex("OrderQty")) * .ValueMatrix(r, .ColIndex("Rate")), 2)
                .TextMatrix(r, .ColIndex("JCTax")) = Round(.ValueMatrix(r, .ColIndex("Pretax")) * .ValueMatrix(r, .ColIndex("JCTaxRate")) / 100, 2)
                .TextMatrix(r, .ColIndex("NJCTax")) = Round(.ValueMatrix(r, .ColIndex("Pretax")) * .ValueMatrix(r, .ColIndex("NJCTaxRate")) / 100, 2)
                .TextMatrix(r, .ColIndex("TotalTax")) = .ValueMatrix(r, .ColIndex("JCTax")) + .ValueMatrix(r, .ColIndex("NJCTax"))
                .TextMatrix(r, .ColIndex("Total")) = .ValueMatrix(r, .ColIndex("Pretax")) + .ValueMatrix(r, .ColIndex("JCTax")) + .ValueMatrix(r, .ColIndex("NJCTax"))
            
            Case "Rate":
                .TextMatrix(r, .ColIndex("POOverridden")) = "True"
                .TextMatrix(r, .ColIndex("Pretax")) = Round(.ValueMatrix(r, .ColIndex("OrderQty")) * .ValueMatrix(r, .ColIndex("Rate")), 2)
                .TextMatrix(r, .ColIndex("JCTax")) = Round(.ValueMatrix(r, .ColIndex("Pretax")) * .ValueMatrix(r, .ColIndex("JCTaxRate")) / 100, 2)
                .TextMatrix(r, .ColIndex("NJCTax")) = Round(.ValueMatrix(r, .ColIndex("Pretax")) * .ValueMatrix(r, .ColIndex("NJCTaxRate")) / 100, 2)
                .TextMatrix(r, .ColIndex("TotalTax")) = .ValueMatrix(r, .ColIndex("JCTax")) + .ValueMatrix(r, .ColIndex("NJCTax"))
                .TextMatrix(r, .ColIndex("Total")) = .ValueMatrix(r, .ColIndex("Pretax")) + .ValueMatrix(r, .ColIndex("JCTax")) + .ValueMatrix(r, .ColIndex("NJCTax"))
                
            Case "Pretax":
                If .ValueMatrix(r, .ColIndex("OrderQty")) = 0 Then .TextMatrix(r, .ColIndex("OrderQty")) = 1
                .TextMatrix(r, .ColIndex("POOverridden")) = "True"
                .TextMatrix(r, .ColIndex("Rate")) = Round(.ValueMatrix(r, .ColIndex("Pretax")) / .ValueMatrix(r, .ColIndex("OrderQty")), 4)
                .TextMatrix(r, .ColIndex("JCTax")) = Round(.ValueMatrix(r, .ColIndex("Pretax")) * .ValueMatrix(r, .ColIndex("JCTaxRate")) / 100, 2)
                .TextMatrix(r, .ColIndex("NJCTax")) = Round(.ValueMatrix(r, .ColIndex("Pretax")) * .ValueMatrix(r, .ColIndex("NJCTaxRate")) / 100, 2)
                .TextMatrix(r, .ColIndex("TotalTax")) = .ValueMatrix(r, .ColIndex("JCTax")) + .ValueMatrix(r, .ColIndex("NJCTax"))
                .TextMatrix(r, .ColIndex("Total")) = .ValueMatrix(r, .ColIndex("Pretax")) + .ValueMatrix(r, .ColIndex("JCTax")) + .ValueMatrix(r, .ColIndex("NJCTax"))
                
            Case "TaxGroup":
                s = "SELECT JCRate,NJCRate FROM TaxGroups WHERE DivisionID = " & HFApp.DivisionID & " and TaxGroup=" & DbQuote(Str, .Cell(flexcpText, r, .ColIndex("TaxGroup")))
                Set rs = HFApp.SqlExec(s)
                .TextMatrix(r, .ColIndex("JCTaxRate")) = Val("" & rs("JCRate"))
                .TextMatrix(r, .ColIndex("NJCTaxRate")) = Val("" & rs("NJCRate"))
                .TextMatrix(r, .ColIndex("JCTax")) = Round(.ValueMatrix(r, .ColIndex("Pretax")) * .ValueMatrix(r, .ColIndex("JCTaxRate")) / 100, 2)
                .TextMatrix(r, .ColIndex("NJCTax")) = Round(.ValueMatrix(r, .ColIndex("Pretax")) * .ValueMatrix(r, .ColIndex("NJCTaxRate")) / 100, 2)
                .TextMatrix(r, .ColIndex("TotalTax")) = .ValueMatrix(r, .ColIndex("JCTax")) + .ValueMatrix(r, .ColIndex("NJCTax"))
                .TextMatrix(r, .ColIndex("Total")) = .ValueMatrix(r, .ColIndex("Pretax")) + .ValueMatrix(r, .ColIndex("JCTax")) + .ValueMatrix(r, .ColIndex("NJCTax"))
            
            Case "JCTax":
                .TextMatrix(r, .ColIndex("TotalTax")) = .ValueMatrix(r, .ColIndex("JCTax")) + .ValueMatrix(r, .ColIndex("NJCTax"))
                .TextMatrix(r, .ColIndex("Total")) = .ValueMatrix(r, .ColIndex("Pretax")) + .ValueMatrix(r, .ColIndex("JCTax")) + .ValueMatrix(r, .ColIndex("NJCTax"))
            
            Case "NJCTax":
                .TextMatrix(r, .ColIndex("TotalTax")) = .ValueMatrix(r, .ColIndex("JCTax")) + .ValueMatrix(r, .ColIndex("NJCTax"))
                .TextMatrix(r, .ColIndex("Total")) = .ValueMatrix(r, .ColIndex("Pretax")) + .ValueMatrix(r, .ColIndex("JCTax")) + .ValueMatrix(r, .ColIndex("NJCTax"))
        
        End Select
        
        If .TextMatrix(r, .ColIndex("Job_No")) = "" Then .TextMatrix(r, .ColIndex("Job_No")) = txtJob.Text
        If .RowData(r) <> "NEW" Then .RowData(r) = "DIRTY"
        
    End If
    Next
    End With
    Call CalcItems
    
    Dirty = True

End Sub

Public Sub mnuEstimateItemsGridSub_Click(Index As Integer)
    Dim rc As Double
    Dim rc2 As Long
    Dim r As Long
    Dim c As Long
    Dim s As String
    Dim NewRow As Long
    Dim Phase As String
    Dim Item As String
    Dim Description As String
    
    With gItems
        If .Row <= 0 Then Exit Sub
        
        Select Case Index
                
            Case mcITEM_VIEWFILES
                Call SaveData(False)
                
                Timer1.Enabled = True
                Timer1.Interval = 10
                
                s = ""
                s = s & "EditItemAttachments|" & .TextMatrix(.Row, .ColIndex("Description")) & "|"
                s = s & "JEI~" & .TextMatrix(.Row, .ColIndex("EstItemID")) & "|Job|"
                s = s & "ASM~" & mCommunity & "~" & "" & "~" & "" & "~" & "" & "~" & .TextMatrix(.Row, .ColIndex("EstPhase")) & "~" & .TextMatrix(.Row, .ColIndex("EstItem")) & "|Model Option Library|"
                s = s & "ITM~" & .TextMatrix(.Row, .ColIndex("EstPhase")) & "~" & .TextMatrix(.Row, .ColIndex("EstItem")) & "|Item Database"
                mTimerTask = s
                
            Case mcITEM_REMOVEITEMS
                Call gItems_KeyDown(vbKeyDelete, vbCtrlMask)
        
            Case mcITEM_SAVEONETIMETODB
                Description = .TextMatrix(.Row, .ColIndex("Description"))
                If FItem.Add(Phase, Item, Description) Then
                    .TextMatrix(.Row, .ColIndex("estphase")) = Phase
                    .TextMatrix(.Row, .ColIndex("estitem")) = Item
                    .TextMatrix(.Row, .ColIndex("Description")) = Description
                    If .RowData(.Row) <> "NEW" Then .RowData(.Row) = "DIRTY"
                    
                    s = ""
                    s = s & "update tblphaseitem" & vbCrLf
                    s = s & "set takeoffuom=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("takeoffuom"))) & vbCrLf
                    s = s & "   ,orderuom=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("orderuom"))) & vbCrLf
                    s = s & "   ,price=" & DbQuote(Num, .TextMatrix(.Row, .ColIndex("Rate"))) & vbCrLf
                    s = s & "   ,poindex=" & DbQuote(Str, "") & vbCrLf
                    s = s & "   ,jccostcode=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("jccostcode"))) & vbCrLf
                    s = s & "   ,jccategory=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("jccategory"))) & vbCrLf
                    s = s & "   ,ustmp=" & DbQuote(Str, HFApp.LoginID) & vbCrLf
                    s = s & "   ,tstmp=getdate()" & vbCrLf
                    s = s & "   ,taxgroup=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("TaxGroup"))) & vbCrLf
                    s = s & "where DivisionID = " & HFApp.DivisionID & " and phase=" & DbQuote(Str, Phase) & vbCrLf
                    s = s & "and item=" & DbQuote(Str, Item) & vbCrLf
                    Call HFApp.SqlExec(s, dbHomefront)
                    
                    Call SaveData(False)
                    
                End If
            
            
            Case mcITEM_SUBSTITUEITEM
                s = ""
                s = s & "select isnull(phase,'')+char(1)+isnull(item,'') phaseitem " & vbCrLf
                s = s & "      ,poindex" & vbCrLf
                s = s & "      ,Phase " & vbCrLf
                s = s & "      ,Item " & vbCrLf
                s = s & "      ,Description " & vbCrLf
                s = s & "      ,takeoffuom,orderuom,jccostcode,jccategory " & vbCrLf
                s = s & "  from tblphaseitem where DivisionID = " & HFApp.DivisionID & vbCrLf
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Item", s, .TextMatrix(.Row, .ColIndex("EstPhase")) & Chr(1) & .TextMatrix(.Row, .ColIndex("EstItem")), , , , "phaseitem,poindex") Then
                    'this can only happen when neither pos or budgets have been generated so we only need to set budget values. the afteredit event will set the po values
                    
                    For r = .Row To .RowSel
                        .TextMatrix(r, .ColIndex("EstPhase")) = FPickList.SelectedItem("Phase")
                        .TextMatrix(r, .ColIndex("EstItem")) = FPickList.SelectedItem("Item")
                        .TextMatrix(r, .ColIndex("Description")) = FPickList.SelectedItem("Description")
                        
                        .TextMatrix(r, .ColIndex("TakeoffUOM")) = FPickList.SelectedItem("TakeoffUOM")
                        .TextMatrix(r, .ColIndex("OrderUOM")) = FPickList.SelectedItem("OrderUOM")
                        .TextMatrix(r, .ColIndex("JCCostCode")) = FPickList.SelectedItem("JCCostCode")
                        .TextMatrix(r, .ColIndex("JCCategory")) = FPickList.SelectedItem("JCCategory")

                        
                        On Error Resume Next
                        rc = 0
                        rc = HFApp.SqlExec("select dbo.Purch_GetItemRate(0,0," & _
                                                          DbQuote(Str, mCurrentCommunity) & "," & _
                                                          DbQuote(Str, mCurrentCommunityPhase) & "," & _
                                                          DbQuote(Str, "") & "," & _
                                                          DbQuote(Str, "") & "," & _
                                                          DbQuote(Str, "") & "," & _
                                                          DbQuote(Str, .TextMatrix(r, .ColIndex("EstPhase"))) & "," & _
                                                          DbQuote(Str, .TextMatrix(r, .ColIndex("EstItem"))) & "," & _
                                                          "0," & _
                                                          DbQuote(Str, txtVendor.Text) & "," & _
                                                          "getdate()," & HFApp.DivisionID & ")", dbHomefront)(0)
                        
                        
                        
                        If rc <> 0 Or HFApp.Options(ZeroRateOnRefreshCosts) = "True" Then
                            .TextMatrix(r, .ColIndex("Rate")) = rc
                            Call gItems_AfterEdit(r, .ColIndex("Rate"))
                        End If
                        If .RowData(r) <> "NEW" Then .RowData(r) = "DIRTY"
                    Next
                    Dirty = True

                End If
                
            
            Case mcITEM_COPYITEMS
                For r = Max(.Row, .RowSel) To Min(.Row, .RowSel) Step -1
                    Dirty = True
                    NewRow = Max(.Row, .RowSel) + 1
                    .AddItem "", NewRow
                    .RowData(NewRow) = "NEW"
                    For c = 0 To .Cols - 1
                        .TextMatrix(NewRow, c) = .TextMatrix(r, c)
                    Next

                    .TextMatrix(NewRow, .ColIndex("OrderQty")) = 0
                    .TextMatrix(NewRow, .ColIndex("TakeoffQty")) = 0
                    .TextMatrix(NewRow, .ColIndex("Pretax")) = 0
                    .TextMatrix(NewRow, .ColIndex("PONumber")) = ""
                    .TextMatrix(NewRow, .ColIndex("EstItemID")) = "0"
                Next
                
            
            Case mcITEM_SPLITITEMS
                Call SplitItems
            
                
            Case mcITEM_COMPAREPRICES
                r = .Row
                If r < 1 Then Exit Sub
                Call FPriceComparison.ShowForm(txtVendor.Text, _
                                               mCommunity, _
                                               mCommunityPhase, _
                                               "", _
                                               "", _
                                               .TextMatrix(r, .ColIndex("EstPhase")), _
                                               .TextMatrix(r, .ColIndex("EstItem")), _
                                               .TextMatrix(r, .ColIndex("Description")))
              
            
        End Select
    End With
End Sub



            
            


Public Sub AddItem(Assembly As String, _
                   AssemblyDescription As String, _
                   Model As String, Phase As String, _
                   Item As String, Description As String, _
                   OrderQty As Double, OrderUOM As String, _
                   TakeoffQty As Double, TakeoffUOM As String, _
                   ConversionFactor As Double, RoundTo As Double, RoundDir As Long, WastePercent As Long, _
                   JCExtra As String, JCCostCode As String, _
                   JCCostCodeDesc As String, _
                   JCCategory As String, JCCategoryDesc As String, _
                   Vendor As String, VendorName As String, price As Double, _
                   TaxGroup As String, _
                   TaxGroupName As String, _
                   JCTaxRate As Double, NJCTaxRate As Double, POIndex As String, _
                   Comments As String, Formula As String, SalesQty As Double, Location As String, WBS, Optional Job As String)
On Error GoTo eh
    
    Dim i As Long
    Dim w As Long
    Dim s As String
    
    Dim rs As Recordset
    Dim RndTo As Double
    Dim RndDir As Long
    Dim POIndexDesc As String
    
    Dim sJobNumber As String

    
    
    mCancelTakeoff = False

    If Job = "" Then
        sJobNumber = "" & HFApp.SqlExec("select job from estimateassemblies where estassemblyid = " & DbQuote(Num, mCurrentEstAssemblyID))(0)
    Else
        sJobNumber = Job
    End If
    Dirty = True
    With gItems
        i = .Rows - 1
        .AddItem "", i
        .RowData(i) = "NEW"
        
        On Error Resume Next
        Set rs = HFApp.SqlExec("select roundto,rounddir from tblphaseitem where DivisionID = " & HFApp.DivisionID & " and phase=" & DbQuote(Str, Phase) & " and item=" & DbQuote(Str, Item), dbHomefront)
        RndTo = Val("" & rs("RoundTo"))
        RndDir = Val("" & rs("RoundDir"))
        On Error GoTo eh
        
        
        .TextMatrix(i, .ColIndex("EstAssemblyID")) = mEstAssemblyID
        .TextMatrix(i, .ColIndex("Job_No")) = sJobNumber
        .TextMatrix(i, .ColIndex("JCExtra")) = JCExtra
        .TextMatrix(i, .ColIndex("EstPhase")) = Phase
        .TextMatrix(i, .ColIndex("EstItem")) = Item
        .TextMatrix(i, .ColIndex("Description")) = Description
        .TextMatrix(i, .ColIndex("ConversionFactor")) = ConversionFactor
        .TextMatrix(i, .ColIndex("RoundTo")) = RoundTo
        .TextMatrix(i, .ColIndex("RoundDir")) = RoundDir
        .TextMatrix(i, .ColIndex("WastePercent")) = WastePercent
        .TextMatrix(i, .ColIndex("JCCostCode")) = JCCostCode
        .TextMatrix(i, .ColIndex("JCCostCodeDesc")) = JCCostCodeDesc
        .TextMatrix(i, .ColIndex("JCCategory")) = JCCategory
        .TextMatrix(i, .ColIndex("OriginalJCCategory")) = JCCategory
        .TextMatrix(i, .ColIndex("JCCategoryDesc")) = JCCategoryDesc
        .TextMatrix(i, .ColIndex("Comments")) = Comments
        .TextMatrix(i, .ColIndex("TakeoffQty")) = TakeoffQty
        .TextMatrix(i, .ColIndex("TakeoffUOM")) = TakeoffUOM
        .TextMatrix(i, .ColIndex("OrderUOM")) = OrderUOM
        .TextMatrix(i, .ColIndex("RoundTo")) = RndTo
        .TextMatrix(i, .ColIndex("RoundDir")) = RndDir
        .TextMatrix(i, .ColIndex("OrderQty")) = OrderQty
        .TextMatrix(i, .ColIndex("Rate")) = price
        .TextMatrix(i, .ColIndex("Pretax")) = Round(OrderQty * price, 2)
        .TextMatrix(i, .ColIndex("TaxGroup")) = TaxGroup
        .TextMatrix(i, .ColIndex("JCTaxRate")) = JCTaxRate
        .TextMatrix(i, .ColIndex("NJCTaxRate")) = NJCTaxRate
        .TextMatrix(i, .ColIndex("JCTax")) = Round(.ValueMatrix(i, .ColIndex("Pretax")) * JCTaxRate / 100, 2)
        .TextMatrix(i, .ColIndex("NJCTax")) = Round(.ValueMatrix(i, .ColIndex("Pretax")) * NJCTaxRate / 100, 2)
        .TextMatrix(i, .ColIndex("TotalTax")) = .ValueMatrix(i, .ColIndex("NJCTax")) + .ValueMatrix(i, .ColIndex("JCTax"))
        .TextMatrix(i, .ColIndex("Total")) = .ValueMatrix(i, .ColIndex("Pretax")) + .ValueMatrix(i, .ColIndex("TotalTax"))
        
        

        If Len(mTakeoffSettings) = 0 Then
            .TextMatrix(i, .ColIndex("Location")) = Location
            For w = 1 To 40
                .TextMatrix(i, .ColIndex("WBS" & Format(w, "00"))) = WBS(w)
            Next
        Else
            .TextMatrix(i, .ColIndex("Location")) = Parse(mTakeoffSettings, 1, Chr(1))
            For w = 1 To 40
                .TextMatrix(i, .ColIndex("WBS" & Format(w, "00"))) = Parse(mTakeoffSettings, w + 1, Chr(1))
            Next
        End If
    End With
    Call CalcItems
    
Exit Sub
eh: Call errHandler(SRCFILE & "AddItem")
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo eh
    Dim s As String
    Dim rs As Recordset
    Dim PONumber As String
    
    Select Case Button.Key
    
        Case "Open"
            s = ""
            s = s & "select p.PONumber, p.PODate, p.Job, p.Vendor,v.Vendor_Name Name, p.Description,pi.Pretax" & vbCrLf
            s = s & "from POMaster p left outer join POTotalInvoiced pi on p.PONumber=pi.PONumber" & vbCrLf
            s = s & "join tblvendors v on p.Vendor = v.Vendor_ID and p.DivisionID = v.DivisionID" & vbCrLf
            s = s & "Where p.Cancelled = 0 and p.DivisionID = " & HFApp.DivisionID
            If FPickList.Choose(HFApp.Databases(dbHomefront), "Open PO", s, txtPONumber.Text, True, False, , "PreTax") Then
                txtPONumber.Text = FPickList.SelectedItem("PONumber")
                txtJob.Text = FPickList.SelectedItem("Job")
                mJob = FPickList.SelectedItem("Job")
                Call LoadPO
                txtPONumber.Locked = True
                mNew = False
                FieldsLocked = False
                SetCtrlFocus txtShipVia
            End If
        
        Case "New"
        
            FieldsLocked = True
            Call ClearForm
            Call ClearFields
            Dirty = False
            
            s = ""
            s = s & "select job_no Job,Description from tblJobs where DivisionID = " & HFApp.DivisionID
            If Not FPickList.Choose(HFApp.Databases(dbHomefront), "Open Job", s, txtJob.Text) Then Exit Sub
            
            txtJob.Text = FPickList.SelectedItem("Job")
            mEstAssemblyID = 0
            txtPONumber.Locked = False
            
            s = "dbo.Purch_CreatePONumber " & DbQuote(Num, HFApp.DivisionID) & "," & DbQuote(Str, txtJob.Text) & "," & DbQuote(Str, txtPOIndex.Text)
            PONumber = "" & HFApp.SqlExec(s)(0)
            txtPONumber.Text = PONumber
            
            mNew = True
            FieldsLocked = False
            Dirty = True
            
            With gItems
                .Rows = 1
                .AddItem ""
            End With
            SetCtrlFocus txtVendor
        
        Case "Save": Call SaveData(False)
        
        Case "Delete"
            PONumber = Trim(txtPONumber.Text)
            If mNew Then
                ClearFields
                FieldsLocked = True
                Dirty = False
            Else
                If CancelPO(PONumber) Then
                    ClearFields
                    FieldsLocked = True
                    Dirty = False
                End If
            End If
        
        Case "Send":
            If SaveData(False) Then
                Call FSendingWizard.ShowForm("PO", , DbQuote(Str, txtPONumber.Text))
            End If
            
            
        Case "Preview":         If SaveData(False) Then Call ShowSelectedPOs(True)
        
        
        Case "TakeoffOneTime", "TakeoffItem", "TakeoffAssembly", "TakeoffCustom"
            'Call FTakeoff.Takeoff(Me, True, Mid(Button.Key, 8), 0, "", mAssemblyType, txtDescription.Text, mCommunity, "", cboModel.Text, "", txtAssembly.Text, "")
                
            If mEstAssemblyID <> -1 And txtJob.Text <> "" Then
                If mEstAssemblyID = 0 Then
                    s = "Insert into EstimateAssemblies(Divisionid,Job,HFLocation,AssemblyType,EstimateIndex) Values("
                    s = s & DbQuote(Num, HFApp.DivisionID) & ","
                    s = s & DbQuote(Str, txtJob.Text) & ",'Manual Estimates',-1,0)"
                    HFApp.SqlExec s
                    mEstAssemblyID = HFApp.SqlIdentity("EstimateAssemblies")
                End If
            
                mCurrentEstAssemblyID = mEstAssemblyID
                
                s = ""
                s = s & "SELECT j.community,j.communityphase,a.assembly,j.job_no,a.jcextra,j.description jobdesc,isnull(nullif(a.model,''),j.model) model,a.optionid,a.hfdescription,a.assemblytype,a.budgetslocked,a.changeorder" & vbCrLf
                s = s & "FROM Estimateassemblies a JOIN tblJobs j on(a.job=j.job_no and j.DivisionID = " & HFApp.DivisionID & ")" & vbCrLf
                s = s & "WHERE EstAssemblyID=" & DbQuote(Num, mCurrentEstAssemblyID)
                Set rs = HFApp.SqlExec(s)
                
                mCurrentAssemblyDesc = "" & rs("HFDescription")
                mCurrentJobDesc = "" & rs("JobDesc")
                mCurrentModel = "" & rs("Model")
                mCurrentOptionID = "" & rs("OptionID")
                mCurrentJob = "" & rs("Job_No")
                mCurrentCommunity = "" & rs("Community")
                mCurrentCommunityPhase = "" & rs("CommunityPhase")
                mCurrentAssembly = "" & rs("Assembly")
                mCurrentJCExtra = "" & rs("JCExtra")
                mCurrentChangeOrder = "" & rs("ChangeOrder")
                mCurrentBudgetsLocked = "" & rs("BudgetsLocked") = "True"
                                
                s = "Call FTakeoff.Takeoff"
                mCurrentAssemblyType = Val("" & rs("AssemblyType"))
                
                'save assembly,model,option back to EstimateAssembly if it doesnt have one
                Dim bReqSave As Boolean
                bReqSave = mCurrentAssembly = ""
                mCancelTakeoff = True
                Call FTakeoff.Takeoff(Me, _
                                      mCurrentChangeOrder <> "", _
                                      Mid(Button.Key, 8), _
                                      mTakeoffSystemBidID, _
                                      mTakeoffSystemProjectName, _
                                      mCurrentAssemblyType, _
                                      mCurrentAssemblyDesc, _
                                      mCurrentCommunity, _
                                      mCurrentCommunityPhase, _
                                      mCurrentModel, _
                                      mCurrentOptionID, _
                                      mCurrentAssembly, _
                                      mCurrentJob, _
                                      txtVendor.tag)
                

            End If
            

    End Select
Exit Sub
eh: Call errHandler(SRCFILE & "Toolbar_ButtonClick", s)
End Sub

Private Sub ClearForm()
    txtVendor.Text = ""
    txtAddress.Text = ""
    txtShipVia.Text = ""
    txtFOB.Text = ""
    txtDeliveryRecipient.Text = ""
    txtDeliveryAddress.Text = ""
    txtDateIssued.Text = ""
    txtOrderedBy.Text = ""
    txtDateDue.Text = ""
    txtJob.Text = ""
    txtRetainagePercent.Text = ""
    txtPODescription.Text = ""
    txtStandardText.Text = ""
    txtTerms.Text = ""
    txtPONumber.Text = ""
    
    gItems.Rows = 1
    Call gItems.AddItem("")
    
    lblSubtotal.Caption = ""
    lblTax.Caption = ""
    lblTotal.Caption = ""
    
End Sub
Private Property Let FieldsLocked(RHS As Boolean)
    
    With Toolbar
        .Buttons("Save").Enabled = Not RHS
        .Buttons("Delete").Enabled = Not RHS
        .Buttons("Send").Enabled = Not RHS
        .Buttons("Preview").Enabled = Not RHS
        .Buttons("TakeoffOneTime").Enabled = Not RHS
        .Buttons("TakeoffItem").Enabled = Not RHS
        .Buttons("TakeoffAssembly").Enabled = Not RHS
        .Buttons("TakeoffCustom").Enabled = Not RHS
    End With

    
    txtVendor.Enabled = Not RHS
    txtShipVia.Enabled = Not RHS
    txtFOB.Enabled = Not RHS
    txtPOIndex.Enabled = Not RHS
    txtDeliveryRecipient.Enabled = Not RHS
    txtDeliveryAddress.Enabled = Not RHS
    txtDateIssued.Enabled = Not RHS
    txtOrderedBy.Enabled = Not RHS
    txtDateDue.Enabled = Not RHS
    txtRetainagePercent.Enabled = Not RHS
    txtPODescription.Enabled = Not RHS
    txtStandardText.Enabled = Not RHS
    gItems.Enabled = Not RHS
    txtTerms.Enabled = Not RHS
    txtJob.Enabled = Not RHS
    
    
    'job is only enabled if using global PO numbers
    If txtJob.Enabled Then txtJob.Enabled = HFApp.Options.ValueByName("PONumberingStyle") = "Sequential"
    
End Property

Private Sub ClearFields()
    txtPONumber = ""
    txtVendor.Text = ""
    txtVendor.tag = ""
    txtAddress.Text = ""
    txtDateIssued.Text = Format(Now(), "medium date")
    lblApproval.Caption = ""
    txtPODescription.Text = ""
    txtStandardText.Text = ""
    txtShipVia.Text = ""
    txtFOB.Text = ""
    txtTerms.Text = ""
    txtOrderedBy.Text = ""
    txtDeliveryAddress.Text = ""
    txtDateDue.Text = ""
    txtRetainagePercent.Text = Val("" & 0) & "%"
    gItems.Rows = 1
End Sub

Private Sub LoadPO()
On Error Resume Next
    Dim r As Long, c As Long, i As Long
    Dim b As Boolean
    Dim s As String
    Dim rs As ADODB.Recordset
    
    s = ""
    s = s & "select PONumber, PODate, POIndex, Job, Vendor,Vendor_Name Name, Description" & vbCrLf
    s = s & ", StandardText, DeliveryMethod, DeliveryDate, DeliveryRecipient, DeliveryAddress, DueDate, ShipVia, FOB, Terms, OrderedBy, CompletedDate, ApprovedBy, ApprovedDate" & vbCrLf
    s = s & ", InvoicedAmount, InvoiceDate,status,approvedby,approveddate" & vbCrLf
    s = s & ", CancellationSent, RetainagePercent,PostingBatch" & vbCrLf
    s = s & "from POMaster p " & vbCrLf
    s = s & "join tblvendors v on p.Vendor = v.Vendor_ID and v.DivisionID = p.DivisionID " & vbCrLf
    s = s & "Where p.Cancelled = 0 and p.PONumber = " & DbQuote(Str, txtPONumber.Text) & " and p.DivisionID = " & HFApp.DivisionID
    Set rs = HFApp.SqlExec(s, dbHomefront)
    If rs.EOF Then Exit Sub
    If rs("PostingBatch") <> 0 Then
        mPostedPO = True
        gItems.Editable = flexEDNone
    Else
        mPostedPO = False
        gItems.Editable = flexEDKbdMouse
    End If
    
    txtJob.Text = "" & rs("Job")
    If txtJob.Text <> "" Then
        mEstAssemblyID = HFApp.SqlExec("Select isnull(min(EstAssemblyID),0) from EstimateAssemblies where Job = " & DbQuote(Str, txtJob.Text))(0)
    End If
    On Error GoTo eh
        
        
    
    
    txtVendor.Text = "" & rs("Name")
    txtVendor.tag = "" & rs("Vendor")
    txtDateIssued.Text = Format("" & rs("PODate"), "medium date")
    txtPOIndex.Text = "" & rs("POIndex")
    txtPODescription.Text = "" & rs("Description")
    txtStandardText.Text = "" & rs("StandardText")
    txtShipVia.Text = "" & rs("ShipVia")
    txtFOB.Text = "" & rs("FOB")
    txtTerms.Text = "" & rs("Terms")
    txtOrderedBy.Text = "" & rs("OrderedBy")
    txtDeliveryAddress.Text = "" & rs("DeliveryAddress")
    txtDateDue.Text = Format("" & rs("DueDate"), "medium date")
    txtRetainagePercent.Text = Val("" & rs("RetainagePercent")) & "%"
    
    
    If "" & rs("status") = "Pending" Then
        lblApproval.FontBold = True
        lblApproval.Caption = "APPROVAL REQUIRED"
    Else
        lblApproval.FontBold = False
        lblApproval.Caption = "Approved by " & rs("approvedby") & ": " & Format("" & rs("approveddate"), "mmm d, yyyy")
    End If
    
    
    Call txtVendor_Validate(b)
    FieldsLocked = False
        
    With gItems
        .Redraw = flexRDNone
        .Rows = 1
        
        
        'EstItemID will be null for custom options or anything that has no assembly...
        s = ""
        s = s & "select i.PONumber, i.ItemSeq, i.Description, i.Comments, i.OrderQty, i.OrderUOM, i.TakeoffQty, i.TakeoffUOM, i.Rate, i.Pretax, i.TaxGroup, i.JCTax, i.JCTaxRate" & vbCrLf
        s = s & ", i.NJCTax, i.NJCTaxRate, i.DontPrint, i.Job, i.JCExtra, i.JCCostCode,cc.description JCCostCodeDesc, i.JCCategory,c.description JCCategoryDesc, i.EstPhase, i.EstItemID, i.CustomerChangeOrder, i.GenBatch " & vbCrLf
        s = s & ", i.SortOrder, i.EstItem, i.PartNumber, i.LineNumber, i.LineDescription,j.Description JobDesc" & vbCrLf
        s = s & "from POItems i" & vbCrLf
        s = s & "left outer join tblJobs j on i.Job = j.Job_No and i.DivisionID = j.DivisionID" & vbCrLf
        s = s & "left outer join standardcostcodes cc on i.DivisionID = cc.DivisionID and i.jccostcode=cc.costcode" & vbCrLf
        s = s & "left outer join standardcategories c on i.DivisionID = c.DivisionID and i.jccategory=c.category" & vbCrLf
        s = s & " WHERE PONumber = " & DbQuote(Str, txtPONumber.Text) & " and i.DivisionID = " & HFApp.DivisionID
        Set rs = HFApp.SqlExec(s)
        While Not rs.EOF
                .AddItem ""
                r = .Rows - 1
                For c = 0 To .Cols - 1
                    Select Case .ColKey(c)
                            
                        Case "EstAssemblyID"
                            .Cell(flexcpText, r, c) = mEstAssemblyID
                        Case "ExcludeFromPO"
                            .Cell(flexcpChecked, r, c) = flexUnchecked
                        Case "Seq"
                            .Cell(flexcpText, r, c) = "" & rs("ItemSeq")
                        Case "Description"
                            .Cell(flexcpText, r, c) = "" & rs(.ColKey(c))
                        Case "Total"
                            .Cell(flexcpText, r, c) = "" & (rs("PreTax") + rs("JCTax") + rs("NJCTax"))
                        Case "Job_No"
                            .Cell(flexcpText, r, c) = "" & rs("Job")
                        Case Else
                            On Error Resume Next
                            .Cell(flexcpText, r, c) = "" & rs(.ColKey(c))
                            On Error GoTo eh
                    End Select
                Next
                .TextMatrix(r, .ColIndex("TotalTax")) = .ValueMatrix(r, .ColIndex("JCTax")) + .ValueMatrix(r, .ColIndex("NJCTax"))
            
            rs.MoveNext
        Wend

        .AddItem ""
        .Redraw = flexRDBuffered
    End With
    
    Dirty = False
    
Exit Sub
eh: Call errHandler(SRCFILE & "LoadPO", s)
End Sub

Private Sub ShowSelectedPOs(Preview As Boolean)
On Error GoTo eh
    Dim s As String
    Dim pos As String
    Dim rs As Recordset
    Dim lastRpt As String
    'Dim f As FRptViewer
    Dim c As ZybUtil.Crystal
    Dim PrinterName As String
    Dim i As Long
    
    pos = DbQuote(Str, txtPONumber.Text)
    If pos = "" Then Exit Sub
    
    s = ""
    s = s & "SELECT PONumber,POFormat" & vbCrLf
    s = s & "  FROM PurchaseOrders" & vbCrLf
    s = s & " WHERE PONumber IN(" & pos & ")" & " and DivisionID = " & HFApp.DivisionID & vbCrLf
    s = s & "ORDER BY POFormat" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    'loop thru this set building csv list of ponumbers in each format then showing rpt
    While Not rs.EOF
        If lastRpt <> "" & rs("POFormat") Then
            If pos <> "" And lastRpt <> "" Then
                'Set f = New FRptViewer
                s = PathAppend(HFApp.SystemFolder, "Estimating\PO Formats", lastRpt & ".rpt")
                'Call f.ShowReport(s, Preview, False, "PONumber", Mid(pos, 2))
                
                Set c = New ZybUtil.Crystal
                Call c.LoadODBCReport(s, HFApp.LoginDSN, HFApp.LoginDBUID, HFApp.LoginDBPWD)
                On Error Resume Next
                Call c.ParameterValue("DivisionID", HFApp.DivisionID)
                Call c.ParameterValue("PONumber", Mid(pos, 2))
                On Error GoTo 0
                If Preview Then
                    Call c.PrintPreview("Print Preview")
                Else
                    If Not VBPrintDlg(i, eprAll, True, , , True, , , , , , , Me.hwnd, , PrinterName) Then Exit Sub
                    Call c.PrintReport(PrinterName)
                End If
            
            End If
            lastRpt = "" & rs("POFormat")
            pos = ""
        End If
        pos = pos & "," & DbQuote(Str, "" & rs("PONumber"))
        rs.MoveNext
    Wend
    If pos <> "" And lastRpt <> "" Then
        'Set f = New FRptViewer
        s = PathAppend(HFApp.SystemFolder, "Estimating\PO Formats", lastRpt & ".rpt")
        'Call f.ShowReport(s, Preview, False, "PONumber", Mid(pos, 2))
    
        Set c = New ZybUtil.Crystal
        Call c.LoadODBCReport(s, HFApp.LoginDSN, HFApp.LoginDBUID, HFApp.LoginDBPWD)
        On Error Resume Next
        Call c.ParameterValue("DivisionID", HFApp.DivisionID)
        Call c.ParameterValue("PONumber", Mid(pos, 2))
        On Error GoTo 0
        If Preview Then
            Call c.PrintPreview("Print Preview")
        Else
            Call c.PrintReport
        End If
        
    End If
Exit Sub
eh:
    Select Case Err.Number
    Case 438 'File not found.
        MsgBox "Unable to open PO format """ & lastRpt & """." & vbCrLf & vbCrLf & "File not found:" & vbCrLf & PathAppend(HFApp.SystemFolder, "Estimating\PO Formats", lastRpt & ".rpt"), vbExclamation, App.ProductName
    Case Else: Call errHandler(SRCFILE & "ShowSelectedPOs", s)
    End Select
End Sub



Public Sub ShowForm(Mode As FCEMode, Job As String)
    'ProjectBased = False
    mMode = Mode
    mJob = Job
    If Mode = fceQuote Then ProjectBased = False
    If ProjectBased And Mode <> fceQuote Then
        mCommunity = "" & HFApp.SqlExec("Select Community from tblJobs where DivisionID = " & HFApp.DivisionID & " and job_no = " & DbQuote(Str, mJob))(0)
        
    End If
    Me.Show
    Me.tag = Mode & Chr(0) & mJob
    
End Sub


Private Sub ShowTip(Row As Long, Col As Long, x As Long, Y As Long, Optional tip As String)
    With gItems
        If Row < 0 Or Col < 0 Or (.RowSel = .Row And tip <> "") Then
            picWarningMessages.Visible = False
        Else
            lblWarningMessages = IIf(tip <> "", tip, gItems.Cell(flexcpData, Row, Col))
            
            imgTipIcon.Picture = IIf(tip <> "", imgInfo.Picture, imgWarning.Picture)
            
            
            picWarningMessages.Visible = lblWarningMessages <> ""
            
            Set Me.Font = lblWarningMessages.Font
            
            lblWarningMessages.Alignment = IIf(tip <> "", 1, 0)
            lblWarningMessages.Width = Me.TextWidth(lblWarningMessages) + lblWarningMessages.left
            lblWarningMessages.Height = Me.TextHeight(lblWarningMessages) + lblWarningMessages.Top
            
            picWarningMessages.Width = lblWarningMessages.Width + lblWarningMessages.left + lblWarningMessages.Top
            picWarningMessages.Height = lblWarningMessages.Height + 2 * lblWarningMessages.Top
            If .Height - Y - picWarningMessages.Height - 365 < 0 Then Y = .Height - picWarningMessages.Height - 365
            If .Width - x - picWarningMessages.Width - 365 < 0 Then x = .Width - picWarningMessages.Width - 365
            picWarningMessages.Top = Y
            picWarningMessages.left = x
            
        End If
    End With
End Sub


Private Function SaveData(prompt As Boolean) As Boolean
On Error GoTo eh
    Dim s As String
    Dim r As Long
    Dim community As String
    Dim CommunityPhase As String

    
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
    
        
        
    If Not ValidateData() Then Exit Function
        
    
    Screen.MousePointer = vbHourglass
    community = mCommunity
    CommunityPhase = mCommunityPhase
    
    If SaveItems Then
        SaveData = True
        Dirty = False
    End If
    Screen.MousePointer = vbDefault
    
Exit Function
eh: Call errHandler(SRCFILE & "SaveData", s)
End Function

Private Function ValidateData() As Boolean
Dim r As Long

    txtJob.Text = Trim(txtJob.Text)
    If Trim(txtJob.Text) = "" Then
        MsgBox "A Job number is required.", vbExclamation, App.ProductName
        Call SetCtrlFocus(txtJob)
        Exit Function
    End If
    
    
    With gItems
    For r = 1 To .Rows - 2
        If Trim(.TextMatrix(r, .ColIndex("JCCostCode"))) = "" Or Trim(.TextMatrix(r, .ColIndex("JCCategory"))) = "" Then
            MsgBox "Cost code and category are required.", vbExclamation, App.ProductName
            Call SetCtrlFocus(gItems)
            .Row = r
            Exit Function
        End If
    Next
    End With
    
    
    ValidateData = True
    
End Function

Private Sub SplitItems()
On Error GoTo eh
    Dim Percents() As Double
    Dim WriteUnits As Boolean
    Dim UnitStart As Long
    Dim UnitIncrement As Long
    
    Dim r As Long
    Dim i As Long
    Dim c As Long
    Dim NewRow As Long
    
    If Not FSplitItems.ShowForm(Percents(), WriteUnits, UnitStart, UnitIncrement) Then Exit Sub
    
    With gItems
    For r = Max(.Row, .RowSel) To Min(.Row, .RowSel) Step -1
        'for each copy to be made
        For i = 1 To UBound(Percents)
            Dirty = True
            
            'create new row. copy all values from this row
            NewRow = Max(.Row, .RowSel) + 1
            .AddItem "", NewRow
            .RowData(NewRow) = "NEW"
            For c = 0 To .Cols - 1
                .TextMatrix(NewRow, c) = .TextMatrix(r, c)
            Next
            
            If WriteUnits Then
                .TextMatrix(NewRow, .ColIndex("Unit")) = UnitStart + ((i - 1) * UnitIncrement)
            End If
            
            'change id of new item
            .TextMatrix(NewRow, .ColIndex("EstItemID")) = "0"
                .TextMatrix(NewRow, .ColIndex("OrderQty")) = .ValueMatrix(r, .ColIndex("OrderQty")) * Percents(i)
                Call .Select(NewRow, .ColIndex("OrderQty"))
                Call gItems_AfterEdit(NewRow, .ColIndex("OrderQty"))

            
            
        Next
        
        
        'delete source po item
        If .RowData(r) = "NEW" Then
            Call .RemoveItem(r)
        Else
            .RowHidden(r) = True
            If .RowData(r) <> "NEW" Then .RowData(r) = "DIRTY"
        End If
        
        .TextMatrix(r, .ColIndex("PODeleted")) = "True"
        .TextMatrix(r, .ColIndex("OrderQty")) = 0
        .TextMatrix(r, .ColIndex("Rate")) = 0
    Next
    End With

Exit Sub
eh: MsgBox Err.Description: Resume
End Sub


Private Property Get Dirty() As Boolean
    Dirty = mDirty
End Property
Private Property Let Dirty(RHS As Boolean)
    mDirty = RHS
End Property

Private Function CleanJob(FormatedJob As String) As String
    CleanJob = Trim(Replace(Replace(Replace(Replace(Replace(FormatedJob, "\", ""), ",", ""), "/", ""), ".", ""), "-", ""))
End Function


Public Function SaveItems() As Boolean
On Error GoTo eh
    Dim w As Long
    Dim i As Long
    Dim s As String
    Dim AssmID As Integer
    Dim sJobNumber As String
    Dim CanDeletePO As Boolean
    Dim rs As Recordset
    
    If mNew Then
        s = "insert into POMaster(PONumber,DivisionID, POIndex, Job, Vendor, Description, StandardText, DeliveryMethod,DeliveryRecipient, DeliveryAddress, DueDate, ShipVia, FOB, Terms, OrderedBy"
        s = s & ", VariancePO, UStmp, TStmp, PODate, RetainagePercent,Summarized)"
        s = s & " Values(" & DbQuote(Str, txtPONumber.Text) & ","
        s = s & HFApp.DivisionID & ","
        s = s & DbQuote(Str, txtPOIndex.Text) & ","
        s = s & DbQuote(Str, txtJob.Text) & ","
        s = s & DbQuote(Str, txtVendor.tag) & ","
        s = s & DbQuote(Str, txtPODescription) & ","
        s = s & DbQuote(Str, txtStandardText.Text) & ","
        s = s & DbQuote(Num, IIf(InStr(1, txtDeliveryAddress.Text, "@") <> 0, 1, 0)) & ","
        s = s & DbQuote(Str, txtDeliveryRecipient.Text) & ","
        s = s & DbQuote(Str, txtDeliveryAddress.Text) & ","
        s = s & DbQuote(Date, txtDateDue.Text) & ","
        s = s & DbQuote(Str, txtShipVia.Text) & ","
        s = s & DbQuote(Str, txtFOB.Text) & ","
        s = s & DbQuote(Str, txtTerms.Text) & ","
        s = s & DbQuote(Str, txtOrderedBy.Text) & ","
        s = s & DbQuote(Num, 0) & ","
        s = s & DbQuote(Str, HFApp.LoginID) & ",getdate(),"
        s = s & DbQuote(Date, txtDateIssued.Text) & ","
        s = s & DbQuote(Num, txtRetainagePercent.Text) & ","
        s = s & DbQuote(Bit, HFApp.Options(PostSummarizedPOs))
        s = s & ")"
    Else
        'save to pomaster
        s = ""
        s = s & "UPDATE POMaster" & vbCrLf
        s = s & "   SET Description=" & DbQuote(Str, txtPODescription.Text) & vbCrLf
        s = s & "      ,StandardText=" & DbQuote(Str, txtStandardText.Text) & vbCrLf
        s = s & "      ,ShipVia=" & DbQuote(Str, txtShipVia.Text) & vbCrLf
        s = s & "      ,FOB=" & DbQuote(Str, txtFOB.Text) & vbCrLf
        s = s & "      ,Terms=" & DbQuote(Str, txtTerms.Text) & vbCrLf
        s = s & "      ,OrderedBy=" & DbQuote(Str, txtOrderedBy.Text) & vbCrLf
        s = s & "      ,PODate=" & DbQuote(Date, txtDateIssued.Text) & vbCrLf
        s = s & "      ,DueDate=" & DbQuote(Date, txtDateDue.Text) & vbCrLf
        s = s & "      ,DeliveryRecipient=" & DbQuote(Str, txtDeliveryRecipient.Text) & vbCrLf
        s = s & "      ,DeliveryAddress=" & DbQuote(Str, txtDeliveryAddress.Text) & vbCrLf
        s = s & "      ,RetainagePercent=" & DbQuote(Num, Val(Replace(txtRetainagePercent.Text, "%", ""))) & vbCrLf
        Select Case True
            Case InStr(1, txtDeliveryAddress.Text, "@") > 0:  s = s & "      ,DeliveryMethod=" & DbQuote(Num, dtEmail) & vbCrLf
            Case Trim(txtDeliveryAddress.Text) = "":          s = s & "      ,DeliveryMethod=" & DbQuote(Num, dtPrint) & vbCrLf
            Case Else:                                        s = s & "      ,DeliveryMethod=" & DbQuote(Num, dtFax) & vbCrLf
        End Select
        s = s & "WHERE PONumber=" & DbQuote(Str, txtPONumber.Text) & vbCrLf
    End If
    Call HFApp.SqlExec(s)
    txtPONumber.Locked = True
    mNew = False
    
    
    With gItems
    For i = .Rows - 2 To 1 Step -1
    Select Case True
        Case .RowHidden(i)
            Call HFApp.SqlExec("Update EstimateItems set ponumber='',pogenbatch=0,podeleted=1 where ponumber=" & DbQuote(Str, txtPONumber.Text) & " and estitemid = " & DbQuote(Num, .TextMatrix(i, .ColIndex("EstItemID"))))
            Call HFApp.SqlExec("delete from poitems where DivisionID = " & HFApp.DivisionID & " and POnumber=" & DbQuote(Str, txtPONumber.Text) & " and ItemSeq = " & DbQuote(Num, .ValueMatrix(i, .ColIndex("Seq"))))
            Call .RemoveItem(i)
    End Select
    Next
    
    For i = 1 To .Rows - 2
    Select Case True
        Case .RowData(i) = "NEW"
            If "" & .TextMatrix(i, .ColIndex("Job_No")) = "" Then
                 If Val(.TextMatrix(i, .ColIndex("EstAssemblyID"))) <> AssmID Then
                      AssmID = Val(.TextMatrix(i, .ColIndex("EstAssemblyID")))
                      sJobNumber = "" & HFApp.SqlExec("select Job from estimateassemblies where EstAssemblyid = " & DbQuote(Num, .TextMatrix(i, .ColIndex("EstAssemblyID"))))(0)
                 End If
            Else
                 sJobNumber = .TextMatrix(i, .ColIndex("Job_No"))
            End If
            If .ValueMatrix(i, .ColIndex("EstAssemblyID")) = 0 Then
                If mEstAssemblyID = 0 Then
                    s = "Insert into EstimateAssemblies(divisionid,Job,HFLocation,AssemblyType,EstimateIndex) Values("
                    s = s & DbQuote(Num, HFApp.DivisionID) & vbCrLf
                    s = s & "," & DbQuote(Str, txtJob.Text) & ",'Manual Estimates',-1,0)"
                    HFApp.SqlExec s
                    mEstAssemblyID = HFApp.SqlIdentity("EstimateAssemblies")
                End If
                .TextMatrix(i, .ColIndex("EstAssemblyID")) = mEstAssemblyID
            End If
            s = ""
            s = s & "INSERT INTO EstimateItems(EstAssemblyID,POGenBatch,Assembly,AssemblyDescription,POIndex,Model,Phase,Item,Job,divisionid,JCExtra,JCCostCode,JCCategory,SortOrder,Description,Comments,formula,TakeoffQty,TakeoffUOM,ConversionFactor,OrderUOM,BudgetDeleted,PODeleted,BudgetVendor,BudgetQty,BudgetRate,BudgetPretax,BudgetTaxGroup,BudgetJCTax,BudgetJCTaxRate,BudgetNJCTax,BudgetNJCTaxRate,BudgetGenerated,POVendor,POQty,PORate,POPretax,POTaxGroup,POJCTax,POJCTaxRate,PONJCTax,PONJCTaxRate,POOverridden,BudgetOverridden,ExcludeFromPO,OriginalJCCategory,SalesQty,Location,Unit,PONumber"
            For w = 1 To 40
                s = s & ",WBS" & Format(w, "00")
            Next
            s = s & ")" & vbCrLf
            s = s & "VALUES(" & DbQuote(Num, .TextMatrix(i, .ColIndex("EstAssemblyID"))) & vbCrLf
            s = s & "      ,-99," & DbQuote(Str, "") & vbCrLf
            s = s & "      ," & DbQuote(Str, "") & vbCrLf
            s = s & "      ," & DbQuote(Str, txtPOIndex.Text) & vbCrLf
            s = s & "      ," & DbQuote(Str, "") & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("EstPhase"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("EstItem"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, sJobNumber) & vbCrLf
            s = s & "      ," & DbQuote(Num, HFApp.DivisionID) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("JCExtra"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("JCCostCode"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("JCCategory"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, i) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("Description"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("Comments"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("Formula"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("TakeoffQty"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("TakeoffUOM"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("ConversionFactor"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("OrderUOM"))) & vbCrLf
            s = s & "      ," & DbQuote(Bit, 1) & vbCrLf
            s = s & "      ," & DbQuote(Bit, .TextMatrix(i, .ColIndex("PODeleted"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, txtVendor.tag) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("OrderQty"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("Rate"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("Pretax"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("TaxGroup"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("JCTax"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("JCTaxRate"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("NJCTax"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("NJCTaxRate"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, 1) & vbCrLf
            s = s & "      ," & DbQuote(Str, txtVendor.tag) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("OrderQty"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("Rate"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("Pretax"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("TaxGroup"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("JCTax"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("JCTaxRate"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("NJCTax"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("NJCTaxRate"))) & vbCrLf
            s = s & "      ," & DbQuote(Bit, .TextMatrix(i, .ColIndex("POOverridden")) = "True") & vbCrLf
            s = s & "      ," & DbQuote(Bit, 1) & vbCrLf
            s = s & "      ," & IIf(1 = 2, 1, 0) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("OriginalJCCategory"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, 1) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("Location"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("Unit"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, txtPONumber.Text) & vbCrLf
            For w = 1 To 40
                s = s & "," & DbQuote(Str, .TextMatrix(i, .ColIndex("WBS" & Format(w, "00")))) & vbCrLf
            Next
            s = s & ")" & vbCrLf
            HFApp.SqlExec s
            .TextMatrix(i, .ColIndex("EstItemID")) = HFApp.SqlIdentity("EstimateItems")
            
            'Insert New Po Items
            s = "insert into POItems(PONumber,DivisionID, Description, Comments, OrderQty, OrderUOM, TakeoffQty, TakeoffUOM, Rate, Pretax, TaxGroup, JCTax, JCTaxRate, NJCTax, NJCTaxRate"
            s = s & ", Job, JCExtra, JCCostCode, JCCategory, EstPhase, EstItem, EstItemID, GenBatch, SortOrder, PartNumber) Values("
            s = s & DbQuote(Str, txtPONumber.Text)
            s = s & "," & HFApp.DivisionID
            s = s & "," & DbQuote(Str, .TextMatrix(i, .ColIndex("Description")))
            s = s & "," & DbQuote(Str, .TextMatrix(i, .ColIndex("Comments")))
            s = s & "," & DbQuote(Num, .ValueMatrix(i, .ColIndex("OrderQty")))
            s = s & "," & DbQuote(Str, .TextMatrix(i, .ColIndex("OrderUOM")))
            s = s & "," & DbQuote(Num, .ValueMatrix(i, .ColIndex("TakeoffQty")))
            s = s & "," & DbQuote(Str, .TextMatrix(i, .ColIndex("TakeoffUOM")))
            s = s & "," & DbQuote(Num, .ValueMatrix(i, .ColIndex("Rate")))
            s = s & "," & DbQuote(Num, .ValueMatrix(i, .ColIndex("Pretax")))
            s = s & "," & DbQuote(Str, .TextMatrix(i, .ColIndex("TaxGroup")))
            s = s & "," & DbQuote(Num, .ValueMatrix(i, .ColIndex("JCTax")))
            s = s & "," & DbQuote(Num, .ValueMatrix(i, .ColIndex("JCTaxRate")))
            s = s & "," & DbQuote(Num, .ValueMatrix(i, .ColIndex("NJCTax")))
            s = s & "," & DbQuote(Num, .ValueMatrix(i, .ColIndex("NJCTaxRate")))
            s = s & "," & DbQuote(Str, txtJob.Text)
            s = s & "," & DbQuote(Str, .TextMatrix(i, .ColIndex("JCExtra")))
            s = s & "," & DbQuote(Str, .TextMatrix(i, .ColIndex("JCCostCode")))
            s = s & "," & DbQuote(Str, .TextMatrix(i, .ColIndex("JCCategory")))
            s = s & "," & DbQuote(Str, .TextMatrix(i, .ColIndex("EstPhase")))
            s = s & "," & DbQuote(Str, .TextMatrix(i, .ColIndex("EstItem")))
            s = s & "," & DbQuote(Num, .ValueMatrix(i, .ColIndex("EstItemID")))
            s = s & "," & DbQuote(Num, -99)
            s = s & "," & DbQuote(Num, i)
            s = s & "," & DbQuote(Str, "")
            s = s & ")"
            Call HFApp.SqlExec(s, dbHomefront)
            .TextMatrix(i, .ColIndex("Seq")) = HFApp.SqlIdentity("POItems")
            .RowData(i) = ""
         
        
        Case .RowData(i) = "DIRTY"
            If "" & .TextMatrix(i, .ColIndex("Job_No")) = "" Then
                 If Val(.TextMatrix(i, .ColIndex("EstAssemblyID"))) <> AssmID Then
                      AssmID = Val(.TextMatrix(i, .ColIndex("EstAssemblyID")))
                      sJobNumber = "" & HFApp.SqlExec("select Job from estimateassemblies where EstAssemblyid = " & DbQuote(Num, .TextMatrix(i, .ColIndex("EstAssemblyID"))))(0)
                      .TextMatrix(i, .ColIndex("Job_No")) = sJobNumber
                 End If
            End If
            s = ""
            s = s & "UPDATE EstimateItems" & vbCrLf
            s = s & "SET Phase=" & DbQuote(Str, .TextMatrix(i, .ColIndex("EstPhase"))) & vbCrLf
            s = s & "   ,Item=" & DbQuote(Str, .TextMatrix(i, .ColIndex("EstItem"))) & vbCrLf
            s = s & "   ,TakeoffQty=" & DbQuote(Num, .TextMatrix(i, .ColIndex("TakeoffQty"))) & vbCrLf
            s = s & "   ,BudgetQty=" & DbQuote(Num, .TextMatrix(i, .ColIndex("OrderQty"))) & vbCrLf
            s = s & "   ,POQty=" & DbQuote(Num, .TextMatrix(i, .ColIndex("OrderQty"))) & vbCrLf
            s = s & "   ,Job=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Job_No"))) & vbCrLf
            s = s & "   ,EstAssemblyID=" & DbQuote(Num, .TextMatrix(i, .ColIndex("EstAssemblyID"))) & vbCrLf
            s = s & "   ,POIndex=" & DbQuote(Str, txtPOIndex.Text) & vbCrLf
            s = s & "   ,JCExtra=" & DbQuote(Str, .TextMatrix(i, .ColIndex("JCExtra")), , True) & vbCrLf
            s = s & "   ,JCCostCode=" & DbQuote(Str, .TextMatrix(i, .ColIndex("JCCostCode"))) & vbCrLf
            s = s & "   ,JCCategory=" & DbQuote(Str, .TextMatrix(i, .ColIndex("JCCategory"))) & vbCrLf
            s = s & "   ,Description=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Description"))) & vbCrLf
            s = s & "   ,Comments=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Comments"))) & vbCrLf
            s = s & "   ,OrderUOM=" & DbQuote(Str, .TextMatrix(i, .ColIndex("OrderUOM"))) & vbCrLf
            s = s & "   ,BudgetVendor=" & DbQuote(Str, txtVendor.tag) & vbCrLf
            s = s & "   ,BudgetRate=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Rate"))) & vbCrLf
            s = s & "   ,BudgetPretax=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Pretax"))) & vbCrLf
            s = s & "   ,BudgetTaxGroup=" & DbQuote(Str, .TextMatrix(i, .ColIndex("TaxGroup"))) & vbCrLf
            s = s & "   ,BudgetJCTax=" & DbQuote(Num, .TextMatrix(i, .ColIndex("JCTax"))) & vbCrLf
            s = s & "   ,BudgetJCTaxRate=" & DbQuote(Num, .TextMatrix(i, .ColIndex("JCTaxRate"))) & vbCrLf
            s = s & "   ,BudgetNJCTax=" & DbQuote(Num, .TextMatrix(i, .ColIndex("NJCTax"))) & vbCrLf
            s = s & "   ,BudgetNJCTaxRate=" & DbQuote(Num, .TextMatrix(i, .ColIndex("NJCTaxRate"))) & vbCrLf
            s = s & "   ,POVendor=" & DbQuote(Str, txtVendor.tag) & vbCrLf
            s = s & "   ,PORate=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Rate"))) & vbCrLf
            s = s & "   ,POPretax=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Pretax"))) & vbCrLf
            s = s & "   ,POTaxGroup=" & DbQuote(Str, .TextMatrix(i, .ColIndex("TaxGroup"))) & vbCrLf
            s = s & "   ,POJCTax=" & DbQuote(Num, .TextMatrix(i, .ColIndex("JCTax"))) & vbCrLf
            s = s & "   ,POJCTaxRate=" & DbQuote(Num, .TextMatrix(i, .ColIndex("JCTaxRate"))) & vbCrLf
            s = s & "   ,PONJCTax=" & DbQuote(Num, .TextMatrix(i, .ColIndex("NJCTax"))) & vbCrLf
            s = s & "   ,PONJCTaxRate=" & DbQuote(Num, .TextMatrix(i, .ColIndex("NJCTaxRate"))) & vbCrLf
            s = s & "   ,Location=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Location"))) & vbCrLf
            s = s & "   ,Unit=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Unit"))) & vbCrLf
            For w = 1 To 40
                s = s & "   ,WBS" & Format(w, "00") & "=" & DbQuote(Str, .TextMatrix(i, .ColIndex("WBS" & Format(w, "00")))) & vbCrLf
            Next
            s = s & "WHERE EstItemID=" & DbQuote(Num, .TextMatrix(i, .ColIndex("EstItemID"))) & vbCrLf
            Call HFApp.SqlExec(s)
            
            'Update PO Items
            s = ""
            s = s & "UPDATE POItems" & vbCrLf
            s = s & "SET EstPhase=" & DbQuote(Str, .TextMatrix(i, .ColIndex("EstPhase"))) & vbCrLf
            s = s & "   ,EstItem=" & DbQuote(Str, .TextMatrix(i, .ColIndex("EstItem"))) & vbCrLf
            s = s & "   ,TakeoffQty=" & DbQuote(Num, .TextMatrix(i, .ColIndex("TakeoffQty"))) & vbCrLf
            s = s & "   ,OrderQty=" & DbQuote(Num, .TextMatrix(i, .ColIndex("OrderQty"))) & vbCrLf
            s = s & "   ,Job=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Job_No"))) & vbCrLf
            s = s & "   ,JCExtra=" & DbQuote(Str, .TextMatrix(i, .ColIndex("JCExtra")), , True) & vbCrLf
            s = s & "   ,JCCostCode=" & DbQuote(Str, .TextMatrix(i, .ColIndex("JCCostCode"))) & vbCrLf
            s = s & "   ,JCCategory=" & DbQuote(Str, .TextMatrix(i, .ColIndex("JCCategory"))) & vbCrLf
            s = s & "   ,Description=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Description"))) & vbCrLf
            s = s & "   ,Comments=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Comments"))) & vbCrLf
            s = s & "   ,OrderUOM=" & DbQuote(Str, .TextMatrix(i, .ColIndex("OrderUOM"))) & vbCrLf
            s = s & "   ,Rate=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Rate"))) & vbCrLf
            s = s & "   ,Pretax=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Pretax"))) & vbCrLf
            s = s & "   ,TaxGroup=" & DbQuote(Str, .TextMatrix(i, .ColIndex("TaxGroup"))) & vbCrLf
            s = s & "   ,JCTax=" & DbQuote(Num, .TextMatrix(i, .ColIndex("JCTax"))) & vbCrLf
            s = s & "   ,JCTaxRate=" & DbQuote(Num, .TextMatrix(i, .ColIndex("JCTaxRate"))) & vbCrLf
            s = s & "   ,NJCTax=" & DbQuote(Num, .TextMatrix(i, .ColIndex("NJCTax"))) & vbCrLf
            s = s & "   ,NJCTaxRate=" & DbQuote(Num, .TextMatrix(i, .ColIndex("NJCTaxRate"))) & vbCrLf
            's = s & "   ,POVendor=" & DbQuote(Str, txtVendor.tag) & vbCrLf
            s = s & "WHERE ItemSeq=" & DbQuote(Num, .ValueMatrix(i, .ColIndex("Seq"))) & vbCrLf
            Call HFApp.SqlExec(s)
            .RowData(i) = ""
        
        
        
        Case .RowHidden(i)
            HFApp.SqlExec ("Update EstimateItems set ponumber='',pogenbatch=0,podeleted=1 where ponumber=" & DbQuote(Str, txtPONumber.Text) & " and estitemid = " & DbQuote(Num, .TextMatrix(i, .ColIndex("EstItemID"))))
            HFApp.SqlExec ("delete from poitems where DivisionID = " & HFApp.DivisionID & " and POnumber=" & DbQuote(Str, txtPONumber.Text) & " and ItemSeq = " & DbQuote(Num, .ValueMatrix(i, .ColIndex("Seq"))))
    
    
    
    End Select
    Next
    End With
    
    
    'delete any lines that are both budget and podeleted
    s = ""
    s = s & "delete estimateitems " & vbCrLf
    s = s & "where budgetdeleted=1" & vbCrLf
    s = s & "  and podeleted=1" & vbCrLf
    s = s & "  and DivisionID=" & HFApp.DivisionID & vbCrLf
    s = s & "  and job=" & DbQuote(Str, txtJob.Text) & vbCrLf
    Call HFApp.SqlExec(s)
    
   
    'set linenumbers
    If HFApp.Options(PostSummarizedPOs) Then
        s = ""
        s = s & "update poitems" & vbCrLf
        s = s & "set linedescription=isnull(c.description,a.Description)" & vbCrLf
        s = s & "   ,linenumber=(select LineNumber" & vbCrLf
        s = s & "                from (select row_number() over (PARTITION BY ponumber order by ponumber,job,jcextra,jccostcode,jccategory,taxgroup) LineNumber,ponumber,job,jcextra,jccostcode,jccategory,taxgroup " & vbCrLf
        s = s & "                      from poitems x where x.DivisionID = " & HFApp.DivisionID & " and x.ponumber=a.ponumber" & vbCrLf
        s = s & "                      group by ponumber,job,jcextra,jccostcode,jccategory,taxgroup) b " & vbCrLf
        s = s & "                where a.ponumber=b.ponumber" & vbCrLf
        s = s & "                and a.taxgroup=b.taxgroup" & vbCrLf
        s = s & "                and a.job=b.job" & vbCrLf
        s = s & "                and a.jcextra=b.jcextra" & vbCrLf
        s = s & "                and a.jccostcode=b.jccostcode" & vbCrLf
        s = s & "                and a.jccategory=b.jccategory)" & vbCrLf
        s = s & "from poitems a" & vbCrLf
        s = s & "left outer join standardcostcodes c on a.DivisionID = c.DivisionID and a.jccostcode=c.costcode" & vbCrLf
        s = s & "where a.DivisionID = " & HFApp.DivisionID & " and a.ponumber=" & DbQuote(Str, txtPONumber.Text)
        Call HFApp.SqlExec(s)
    Else
        s = ""
        s = s & "update poitems" & vbCrLf
        s = s & "set linedescription=a.description" & vbCrLf
        s = s & "   ,linenumber=(select LineNumber" & vbCrLf
        s = s & "                from (select row_number() over (PARTITION BY ponumber order by ponumber,ItemSeq) LineNumber,ponumber,ItemSeq" & vbCrLf
        s = s & "                      from poitems x where x.DivisionID = " & HFApp.DivisionID & " and x.ponumber=a.ponumber" & vbCrLf
        s = s & "                      group by ponumber,ItemSeq) b " & vbCrLf
        s = s & "                where a.ponumber=b.ponumber" & vbCrLf
        s = s & "                and a.itemseq=b.itemseq)" & vbCrLf
        s = s & "from poitems a" & vbCrLf
        s = s & "where a.DivisionID = " & HFApp.DivisionID & " and a.ponumber=" & DbQuote(Str, txtPONumber.Text)
        Call HFApp.SqlExec(s)
    End If

    'set po status
    s = ""
    s = s & "select p.ponumber" & vbCrLf
    s = s & "from estimateitems e" & vbCrLf
    s = s & "join pomaster p on e.ponumber=p.ponumber" & vbCrLf
    s = s & "where p.ponumber=" & DbQuote(Str, txtPONumber.Text) & vbCrLf
    s = s & "group by p.ponumber,p.status" & vbCrLf
    s = s & "having sum(e.popretax) + sum(isnull(e.pojctax,0) + isnull(e.ponjctax,0)) <= " & DbQuote(Num, MaxPOAmount) & vbCrLf
    Set rs = HFApp.SqlExec(s)
    If rs.EOF Then
        lblApproval.FontBold = True
        lblApproval.Caption = "APPROVAL REQUIRED"
    Else
        s = ""
        s = s & "update pomaster set status='Approved'" & vbCrLf
        s = s & " ,approvedby=" & DbQuote(Str, HFApp.LoginID) & vbCrLf
        s = s & "where ponumber=" & DbQuote(Str, txtPONumber.Text) & vbCrLf
        Call HFApp.SqlExec(s)
        lblApproval.FontBold = False
        lblApproval.Caption = "Approved by " & HFApp.LoginID & ": " & Format(Now, "mmm d, yyyy")
    End If
    
    
    
    
    SaveItems = True
Exit Function
eh: If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        MsgBox "There is already a PO numbered """ & txtPONumber.Text & """. You must specify a different PO number.", vbExclamation, App.ProductName
    Else
        Call errHandler(SRCFILE & "SaveItems", s)
    End If
End Function

Private Sub txtDateDue_Change()
    Dirty = True
End Sub

Private Sub txtDateDue_GotFocus()
    SelectAll txtDateDue
End Sub

Private Sub txtDateDue_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 And Shift = 0 Then Call cmdBrowse_Click(2)
End Sub

Private Sub txtDateDue_Validate(Cancel As Boolean)
    If IsDate(txtDateDue.Text) Or txtDateDue.Text = "" Then
        txtDateDue.Text = Format(txtDateDue.Text, "medium date")
    Else
        Cancel = True
    End If
End Sub

Private Sub txtDateIssued_Change()
    Dirty = True
End Sub

Private Sub txtDateIssued_GotFocus()
    SelectAll txtDateIssued
End Sub

Private Sub txtDateIssued_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 And Shift = 0 Then Call cmdBrowse_Click(0)
End Sub

Private Sub txtDateIssued_Validate(Cancel As Boolean)
    If IsDate(txtDateIssued.Text) Or txtDateIssued.Text = "" Then
        txtDateIssued.Text = Format(txtDateIssued.Text, "medium date")
    Else
        Cancel = True
    End If
End Sub

Private Sub txtDeliveryAddress_Change()
    Dirty = True
    Select Case True
        Case InStr(1, txtDeliveryAddress.Text, "@") > 0:  lblSendVia.Caption = "Email"
        Case Trim(txtDeliveryAddress.Text) = "":          lblSendVia.Caption = "Print"
        Case Else:                                        lblSendVia.Caption = "Fax"
    End Select
End Sub

Private Sub txtDeliveryAddress_GotFocus()
    SelectAll txtDeliveryAddress
End Sub

Private Sub txtDeliveryRecipient_Change()
    Dirty = True
End Sub

Private Sub txtDeliveryRecipient_GotFocus()
    SelectAll txtDeliveryRecipient
End Sub

Private Sub txtFOB_Change()
    Dirty = True
End Sub

Private Sub txtFOB_GotFocus()
    SelectAll txtFOB
End Sub

Private Sub txtJob_Change()
    Dirty = True
End Sub

Private Sub txtJob_GotFocus()
    SelectAll txtJob
End Sub

Private Sub txtJob_Validate(Cancel As Boolean)
    Dim s As String
    Dim rs As Recordset
        
    If txtJob.Text <> "" Then
        
        s = ""
        s = s & "SELECT job_no" & vbCrLf
        s = s & "  FROM tbljobs" & vbCrLf
        s = s & " WHERE job_no=" & DbQuote(Str, txtJob.Text) & vbCrLf
        s = s & "   AND divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
        Set rs = HFApp.SqlExec(s)
        If rs.EOF Then
            Cancel = True
        Else
            txtJob.Text = "" & rs("job_no")
        End If
        
    End If

End Sub

Private Sub txtOrderedBy_Change()
    Dirty = True
End Sub

Private Sub txtOrderedBy_GotFocus()
    SelectAll txtOrderedBy
End Sub

Private Sub txtPODescription_Change()
    Dirty = True
End Sub

Private Sub txtPODescription_GotFocus()
    SelectAll txtPODescription
End Sub

Private Sub txtPODescription_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 And Shift = 0 Then Call cmdBrowse_Click(4)
End Sub

Private Sub txtPOIndex_GotFocus()
    Call SelectAll(txtPOIndex)
End Sub


Private Sub txtPOIndex_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 And Shift = 0 Then Call cmdBrowse_Click(1)
End Sub

Private Sub txtRetainagePercent_Change()
    Dirty = True
End Sub

Private Sub txtRetainagePercent_GotFocus()
    SelectAll txtRetainagePercent
End Sub

Private Sub txtRetainagePercent_Validate(Cancel As Boolean)
    txtRetainagePercent.Text = Val(txtRetainagePercent.Text) & "%"
End Sub

Private Sub txtShipVia_Change()
    Dirty = True
End Sub

Private Sub txtShipVia_GotFocus()
    SelectAll txtShipVia
End Sub

Private Sub txtStandardText_Change()
    Dirty = True
End Sub

Private Sub txtStandardText_GotFocus()
    SelectAll txtStandardText
End Sub

Private Sub txtTerms_Change()
    Dirty = True
End Sub

Private Sub txtVendor_Change()
    txtVendor.tag = ""
End Sub

Private Sub txtVendor_DropOpen(NoDefault As Boolean)
    Dim s As String
    Dim b As Boolean
    
    s = ""
    s = s & "SELECT DISTINCT" & vbCrLf
    s = s & "  isnull(v.TradeType,'') TradeType" & vbCrLf
    s = s & "      ,v.Vendor_Name Company" & vbCrLf
    s = s & "      ,v.Vendor_ID Vendor, v.Addr1, v.Addr2, v.City, v.State, v.Zip" & vbCrLf
    s = s & "      ,v.Phone,e.Email " & vbCrLf
    s = s & "  FROM tblVendors v" & vbCrLf
    s = s & "  Left outer Join ContactPurchasingEmails e on v.Vendor_ID = e.VendorCode and v.DivisionID = e.DivisionID"
    s = s & " WHERE v.inactive=0 and v.DivisionID =  " & HFApp.DivisionID & vbCrLf
    s = s & "order by 1 desc,2" & vbCrLf
    If FPickList.Choose(HFApp.Databases(dbHomefront), "Vendor", s, gItems, , , , IIf(HFApp.Options(AccountingSystem) = asQuickBooks, "Vendor,Email,Addr1,Addr2,State,Zip", "Email,Addr1,Addr2,State,Zip")) Then
        txtVendor.Text = FPickList.SelectedItem("company")
        txtVendor.tag = FPickList.SelectedItem("vendor")
        Call txtVendor_Validate(b)
        
        Dirty = True
        
    End If
End Sub


Private Sub txtVendor_GotFocus()
    SelectAll txtVendor
End Sub

Private Sub txtVendor_Validate(Cancel As Boolean)
    Dim s As String
    Dim rs As Recordset
        
    If txtVendor.Text = "" Then Exit Sub
    
        
    s = ""
    s = s & "SELECT v.TradeType" & vbCrLf
    s = s & "      ,v.Vendor_Name Company" & vbCrLf
    s = s & "      ,v.Vendor_ID Vendor, v.Addr1, v.Addr2, v.City, v.State, v.Zip" & vbCrLf
    s = s & "      ,v.Phone,e.Email " & vbCrLf
    s = s & "  FROM tblVendors v" & vbCrLf
    s = s & "  Left outer Join ContactPurchasingEmails e on v.Vendor_ID = e.VendorCode and v.DivisionID = e.DivisionID" & vbCrLf
    s = s & " WHERE v.inactive=0 and v.DivisionID =  " & HFApp.DivisionID & vbCrLf
    
    If txtVendor.tag <> "" Then
        s = s & " and v.vendor_id=" & DbQuote(Str, txtVendor.tag) & vbCrLf
    Else
        If HFApp.Options(AccountingSystem) = asQuickBooks Then
            s = s & " and v.vendor_name=" & DbQuote(Str, txtVendor.Text) & vbCrLf
        Else
            s = s & " and v.vendor_id=" & DbQuote(Str, txtVendor.Text) & vbCrLf
        End If
    End If
    
    
    Set rs = HFApp.SqlExec(s)
    If rs.EOF Then
        Cancel = True
    Else
        txtVendor.Text = "" & rs("Company")
        txtVendor.tag = "" & rs("Vendor")
        txtDeliveryAddress.Text = "" & rs("Email")
        s = "" & rs("Addr1")
        If "" & rs("Addr2") <> "" Then
            s = s & vbCrLf & rs("Addr2") & vbCrLf
        End If
        s = s & rs("City") & " " & rs("State") & vbCrLf
        s = s & rs("Zip")
        txtAddress.Text = s
    End If
    
End Sub
